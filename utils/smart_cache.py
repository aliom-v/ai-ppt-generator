"""智能缓存系统 - 支持多层缓存和智能预热"""
import os
import json
import time
import hashlib
import pickle
from typing import Any, Optional, Dict, List, Union, Callable
from functools import wraps
from threading import RLock
import redis
from utils.logger import get_logger

logger = get_logger("smart_cache")


class CacheEntry:
    """缓存条目"""

    def __init__(self, value: Any, ttl: Optional[int] = None):
        self.value = value
        self.created_at = time.time()
        self.expires_at = self.created_at + ttl if ttl else None
        self.access_count = 0
        self.last_accessed = self.created_at

    def is_expired(self) -> bool:
        """检查是否过期"""
        if self.expires_at is None:
            return False
        return time.time() > self.expires_at

    def access(self) -> Any:
        """访问缓存项"""
        self.access_count += 1
        self.last_accessed = time.time()
        return self.value


class MemoryCache:
    """内存缓存实现"""

    def __init__(self, max_size: int = 1000, default_ttl: int = 3600):
        self.max_size = max_size
        self.default_ttl = default_ttl
        self._cache: Dict[str, CacheEntry] = {}
        self._lock = RLock()

    def _evict_expired(self):
        """清理过期项"""
        now = time.time()
        expired_keys = [k for k, v in self._cache.items()
                       if v.expires_at and now > v.expires_at]
        for key in expired_keys:
            del self._cache[key]

    def _evict_lru(self):
        """清理最少使用的项"""
        if len(self._cache) < self.max_size:
            return

        # 按最后访问时间排序，删除最旧的
        sorted_items = sorted(self._cache.items(),
                            key=lambda x: x[1].last_accessed)
        for key, _ in sorted_items[:len(self._cache) - self.max_size + 1]:
            del self._cache[key]

    def get(self, key: str) -> Optional[Any]:
        """获取缓存值"""
        with self._lock:
            self._evict_expired()

            entry = self._cache.get(key)
            if entry is None:
                return None

            if entry.is_expired():
                del self._cache[key]
                return None

            return entry.access()

    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> bool:
        """设置缓存值"""
        with self._lock:
            self._evict_expired()
            self._evict_lru()

            ttl = ttl or self.default_ttl
            self._cache[key] = CacheEntry(value, ttl)
            return True

    def delete(self, key: str) -> bool:
        """删除缓存项"""
        with self._lock:
            if key in self._cache:
                del self._cache[key]
                return True
            return False

    def clear(self):
        """清空缓存"""
        with self._lock:
            self._cache.clear()

    def size(self) -> int:
        """获取缓存大小"""
        with self._lock:
            self._evict_expired()
            return len(self._cache)

    def get_stats(self) -> Dict[str, Any]:
        """获取缓存统计"""
        with self._lock:
            self._evict_expired()
            return {
                "size": len(self._cache),
                "max_size": self.max_size,
                "entries": [
                    {
                        "key": key,
                        "access_count": entry.access_count,
                        "created_at": entry.created_at,
                        "last_accessed": entry.last_accessed,
                        "expires_at": entry.expires_at
                    }
                    for key, entry in self._cache.items()
                ]
            }


class SmartCache:
    """智能缓存系统 - 支持多层缓存"""

    def __init__(self, redis_url: Optional[str] = None,
                 memory_max_size: int = 1000,
                 default_ttl: int = 3600,
                 enable_compression: bool = True):
        self.memory_cache = MemoryCache(memory_max_size, default_ttl)
        self.redis_client = None
        self.enable_compression = enable_compression
        self.default_ttl = default_ttl

        # 尝试连接Redis
        if redis_url:
            try:
                self.redis_client = redis.from_url(redis_url)
                self.redis_client.ping()
                logger.info("Redis缓存已连接")
            except Exception as e:
                logger.warning(f"Redis连接失败，仅使用内存缓存: {e}")

        # 缓存统计
        self.stats = {
            "hits": 0,
            "misses": 0,
            "sets": 0,
            "memory_hits": 0,
            "redis_hits": 0
        }
        self._stats_lock = RLock()

    def _generate_key(self, prefix: str, *args, **kwargs) -> str:
        """生成缓存键"""
        # 创建确定性的键
        key_data = {
            "args": args,
            "kwargs": sorted(kwargs.items())
        }
        key_str = json.dumps(key_data, sort_keys=True, default=str)
        key_hash = hashlib.sha256(key_str.encode()).hexdigest()[:16]
        return f"{prefix}:{key_hash}"

    def _serialize(self, value: Any) -> bytes:
        """序列化值"""
        if self.enable_compression:
            import zlib
            return zlib.compress(pickle.dumps(value))
        return pickle.dumps(value)

    def _deserialize(self, data: bytes) -> Any:
        """反序列化值"""
        if self.enable_compression:
            import zlib
            return pickle.loads(zlib.decompress(data))
        return pickle.loads(data)

    def get(self, key: str) -> Optional[Any]:
        """获取缓存值（多层缓存）"""
        # 先查内存缓存
        value = self.memory_cache.get(key)
        if value is not None:
            with self._stats_lock:
                self.stats["hits"] += 1
                self.stats["memory_hits"] += 1
            return value

        # 再查Redis
        if self.redis_client:
            try:
                data = self.redis_client.get(key)
                if data:
                    value = self._deserialize(data)
                    # 回填到内存缓存
                    self.memory_cache.set(key, value, self.default_ttl)

                    with self._stats_lock:
                        self.stats["hits"] += 1
                        self.stats["redis_hits"] += 1
                    return value
            except Exception as e:
                logger.error(f"Redis获取失败: {e}")

        # 都没命中
        with self._stats_lock:
            self.stats["misses"] += 1
        return None

    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> bool:
        """设置缓存值（多层缓存）"""
        ttl = ttl or self.default_ttl

        # 设置内存缓存
        memory_success = self.memory_cache.set(key, value, ttl)

        # 设置Redis
        redis_success = True
        if self.redis_client:
            try:
                data = self._serialize(value)
                self.redis_client.setex(key, ttl, data)
            except Exception as e:
                logger.error(f"Redis设置失败: {e}")
                redis_success = False

        with self._stats_lock:
            self.stats["sets"] += 1

        return memory_success and redis_success

    def delete(self, key: str) -> bool:
        """删除缓存项"""
        memory_success = self.memory_cache.delete(key)

        redis_success = True
        if self.redis_client:
            try:
                self.redis_client.delete(key)
            except Exception as e:
                logger.error(f"Redis删除失败: {e}")
                redis_success = False

        return memory_success or redis_success

    def clear(self):
        """清空所有缓存"""
        self.memory_cache.clear()

        if self.redis_client:
            try:
                self.redis_client.flushdb()
            except Exception as e:
                logger.error(f"Redis清空失败: {e}")

    def get_stats(self) -> Dict[str, Any]:
        """获取缓存统计"""
        with self._stats_lock:
            stats = self.stats.copy()

        total_requests = stats["hits"] + stats["misses"]
        hit_rate = stats["hits"] / total_requests if total_requests > 0 else 0

        stats.update({
            "hit_rate": round(hit_rate, 4),
            "memory_cache": self.memory_cache.get_stats()
        })

        return stats

    def warm_up(self, data: Dict[str, Any], ttl: Optional[int] = None):
        """缓存预热"""
        logger.info(f"开始缓存预热，共 {len(data)} 项")
        success_count = 0

        for key, value in data.items():
            if self.set(key, value, ttl):
                success_count += 1

        logger.info(f"缓存预热完成，成功 {success_count}/{len(data)} 项")

    def invalidate_pattern(self, pattern: str) -> int:
        """按模式删除缓存"""
        count = 0

        # 内存缓存模式删除
        keys_to_delete = [k for k in self.memory_cache._cache.keys()
                         if pattern in k]
        for key in keys_to_delete:
            if self.memory_cache.delete(key):
                count += 1

        # Redis模式删除
        if self.redis_client:
            try:
                keys = self.redis_client.keys(f"*{pattern}*")
                if keys:
                    count += self.redis_client.delete(*keys)
            except Exception as e:
                logger.error(f"Redis模式删除失败: {e}")

        logger.info(f"按模式删除缓存: {pattern}, 删除了 {count} 项")
        return count


# 全局缓存实例
_cache: Optional[SmartCache] = None


def get_cache() -> SmartCache:
    """获取全局缓存实例"""
    global _cache
    if _cache is None:
        redis_url = os.getenv('REDIS_URL')
        max_size = int(os.getenv('CACHE_MAX_SIZE', '1000'))
        default_ttl = int(os.getenv('CACHE_TTL', '3600'))

        _cache = SmartCache(
            redis_url=redis_url,
            memory_max_size=max_size,
            default_ttl=default_ttl
        )
    return _cache


def cache_result(prefix: str, ttl: Optional[int] = None,
                key_func: Optional[Callable] = None):
    """缓存装饰器"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # 生成缓存键
            if key_func:
                cache_key = key_func(*args, **kwargs)
            else:
                cache_key = get_cache()._generate_key(prefix, *args, **kwargs)

            # 尝试从缓存获取
            result = get_cache().get(cache_key)
            if result is not None:
                return result

            # 执行函数并缓存结果
            result = func(*args, **kwargs)
            get_cache().set(cache_key, result, ttl)
            return result

        # 添加缓存控制方法
        wrapper.cache_key = lambda *args, **kwargs: (
            key_func(*args, **kwargs) if key_func
            else get_cache()._generate_key(prefix, *args, **kwargs)
        )
        wrapper.invalidate = lambda *args, **kwargs: (
            get_cache().delete(wrapper.cache_key(*args, **kwargs))
        )
        wrapper.get_cached = lambda *args, **kwargs: (
            get_cache().get(wrapper.cache_key(*args, **kwargs))
        )

        return wrapper
    return decorator


def cache_ttl(ttl: int):
    """设置TTL的装饰器辅助函数"""
    return lambda prefix: cache_result(prefix, ttl=ttl)


# 缓存预热函数
def warm_common_templates():
    """预热常用模板缓存"""
    from ppt.template_manager import get_template_manager
    cache = get_cache()

    # 预热模板列表
    templates = get_template_manager().list_templates()
    cache.set("templates:list", templates, ttl=86400)  # 24小时

    # 预热每个模板的详细信息
    for template_name in templates:
        template_info = get_template_manager().get_template_info(template_name)
        cache.set(f"template:{template_name}", template_info, ttl=86400)

    logger.info("模板缓存预热完成")


def clear_all_cache():
    """清空所有缓存"""
    get_cache().clear()
    logger.info("所有缓存已清空")


def get_cache_stats() -> Dict[str, Any]:
    """获取缓存统计信息"""
    return get_cache().get_stats()