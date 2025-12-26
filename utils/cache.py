"""生成缓存模块 - 避免重复调用 AI（带 LRU 淘汰机制）"""
import os
import json
import hashlib
import threading
import time
from pathlib import Path
from typing import Optional, Dict, Any, List
from datetime import datetime, timedelta
from collections import OrderedDict

from utils.logger import get_logger

logger = get_logger("cache")


class GenerationCache:
    """PPT 生成结果缓存（带 LRU 淘汰机制）

    根据主题、受众、页数等参数缓存 AI 生成的结果，
    相同参数的请求直接返回缓存，避免重复调用 API。

    特性：
    - LRU 淘汰：当缓存数量超过限制时，自动删除最少使用的条目
    - 过期清理：自动清理超过有效期的缓存
    - 访问计数：记录每个缓存条目的访问次数
    - 线程安全：支持多线程并发访问
    - 延迟写入：减少频繁的文件 I/O
    """

    # 延迟写入间隔（秒）
    WRITE_DELAY = 5.0

    def __init__(
        self,
        cache_dir: str = "cache",
        max_age_hours: int = 24 * 7,
        max_entries: int = 10000,
        auto_cleanup: bool = True
    ):
        """初始化缓存

        Args:
            cache_dir: 缓存目录
            max_age_hours: 缓存有效期（小时），默认 7 天
            max_entries: 最大缓存条目数，默认 10000
            auto_cleanup: 是否在启动时自动清理过期缓存
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.max_age = timedelta(hours=max_age_hours)
        self.max_entries = max_entries
        self.index_file = self.cache_dir / "index.json"
        self._index: OrderedDict[str, Dict] = OrderedDict()
        self._lock = threading.RLock()
        self._stats = {"hits": 0, "misses": 0, "evictions": 0}
        self._dirty = False  # 标记索引是否需要保存
        self._last_save_time = time.time()
        self._load_index()

        # 启动时清理过期缓存
        if auto_cleanup:
            self.cleanup_expired()

    def _load_index(self) -> None:
        """加载缓存索引"""
        if self.index_file.exists():
            try:
                with open(self.index_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # 转换为 OrderedDict 并按最后访问时间排序
                sorted_items = sorted(
                    data.items(),
                    key=lambda x: x[1].get("last_accessed", x[1].get("created_at", "2000-01-01"))
                )
                self._index = OrderedDict(sorted_items)
            except Exception as e:
                logger.warning(f"加载缓存索引失败: {e}")
                self._index = OrderedDict()

    def _save_index(self, force: bool = False) -> None:
        """保存缓存索引（延迟写入）

        Args:
            force: 是否强制立即保存
        """
        self._dirty = True
        current_time = time.time()

        # 延迟写入：只有超过间隔或强制保存时才写入
        if not force and (current_time - self._last_save_time) < self.WRITE_DELAY:
            return

        try:
            with open(self.index_file, 'w', encoding='utf-8') as f:
                json.dump(dict(self._index), f, ensure_ascii=False, indent=2)
            self._dirty = False
            self._last_save_time = current_time
        except Exception as e:
            logger.warning(f"保存缓存索引失败: {e}")

    def flush(self) -> None:
        """强制保存索引到磁盘"""
        with self._lock:
            if self._dirty:
                self._save_index(force=True)

    def _make_key(self, topic: str, audience: str, page_count: int,
                  description: str = "", model: str = "") -> str:
        """生成缓存键

        使用描述的前500字符生成key，避免超长描述导致key不稳定
        """
        # 截断描述以提高缓存命中率
        desc_truncated = description[:500] if description else ""
        content = f"{topic}|{audience}|{page_count}|{desc_truncated}|{model}"
        return hashlib.md5(content.encode()).hexdigest()[:16]

    def _make_simple_key(self, topic: str, audience: str, page_count: int) -> str:
        """生成简单缓存键（用于模糊匹配）"""
        content = f"{topic.lower().strip()}|{audience.lower().strip()}|{page_count}"
        return hashlib.md5(content.encode()).hexdigest()[:12]

    def get(self, topic: str, audience: str, page_count: int,
            description: str = "", model: str = "") -> Optional[Dict[str, Any]]:
        """获取缓存的生成结果

        Args:
            topic: PPT 主题
            audience: 目标受众
            page_count: 页数
            description: 详细描述
            model: 模型名称

        Returns:
            缓存的 PPT 结构字典，如果没有缓存或已过期则返回 None
        """
        with self._lock:
            key = self._make_key(topic, audience, page_count, description, model)

            if key not in self._index:
                self._stats["misses"] += 1
                return None

            entry = self._index[key]
            cache_file = self.cache_dir / f"{key}.json"

            # 检查文件是否存在
            if not cache_file.exists():
                del self._index[key]
                self._save_index()
                self._stats["misses"] += 1
                return None

            # 检查是否过期
            created_at = datetime.fromisoformat(entry.get("created_at", "2000-01-01"))
            if datetime.now() - created_at > self.max_age:
                logger.debug(f"缓存已过期: {topic}")
                self._remove_entry(key)
                self._stats["misses"] += 1
                return None

            # 读取缓存
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # 更新访问信息（LRU）
                entry["last_accessed"] = datetime.now().isoformat()
                entry["access_count"] = entry.get("access_count", 0) + 1
                # 移动到末尾（最近访问）
                self._index.move_to_end(key)
                self._save_index()

                self._stats["hits"] += 1
                logger.info(f"使用缓存: {topic} ({page_count}页)")
                return data
            except Exception as e:
                logger.error(f"读取缓存失败: {e}")
                self._remove_entry(key)
                self._stats["misses"] += 1
                return None

    def set(self, topic: str, audience: str, page_count: int,
            data: Dict[str, Any], description: str = "", model: str = "") -> None:
        """保存生成结果到缓存

        Args:
            topic: PPT 主题
            audience: 目标受众
            page_count: 页数
            data: PPT 结构字典
            description: 详细描述
            model: 模型名称
        """
        with self._lock:
            key = self._make_key(topic, audience, page_count, description, model)
            cache_file = self.cache_dir / f"{key}.json"

            try:
                # 保存数据
                with open(cache_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)

                now = datetime.now().isoformat()
                # 更新索引
                self._index[key] = {
                    "topic": topic,
                    "audience": audience,
                    "page_count": page_count,
                    "model": model,
                    "created_at": now,
                    "last_accessed": now,
                    "access_count": 1,
                    "file": str(cache_file)
                }
                # 移动到末尾
                self._index.move_to_end(key)

                # 检查是否需要淘汰
                self._evict_if_needed()

                self._save_index()
                logger.debug(f"已缓存: {topic}")
            except Exception as e:
                logger.error(f"保存缓存失败: {e}")

    def _evict_if_needed(self) -> None:
        """如果超出容量限制，淘汰最旧的条目（LRU）"""
        while len(self._index) > self.max_entries:
            # 淘汰最早的条目（OrderedDict 头部）
            oldest_key = next(iter(self._index))
            logger.debug(f"LRU 淘汰缓存: {oldest_key}")
            self._remove_entry(oldest_key, save_index=False)
            self._stats["evictions"] += 1

    def _remove_entry(self, key: str, save_index: bool = True) -> None:
        """删除缓存条目"""
        if key in self._index:
            del self._index[key]
            if save_index:
                self._save_index()

        cache_file = self.cache_dir / f"{key}.json"
        if cache_file.exists():
            try:
                cache_file.unlink()
            except OSError as e:
                logger.warning(f"删除缓存文件失败 {cache_file}: {e}")

    def clear(self) -> int:
        """清空所有缓存

        Returns:
            删除的缓存数量
        """
        with self._lock:
            count = len(self._index)
            for key in list(self._index.keys()):
                self._remove_entry(key, save_index=False)
            self._save_index(force=True)
            logger.info(f"已清空 {count} 条缓存")
            return count

    def cleanup_expired(self) -> int:
        """清理过期缓存

        Returns:
            清理的缓存数量
        """
        with self._lock:
            count = 0
            for key in list(self._index.keys()):
                entry = self._index[key]
                created_at = datetime.fromisoformat(entry.get("created_at", "2000-01-01"))
                if datetime.now() - created_at > self.max_age:
                    self._remove_entry(key, save_index=False)
                    count += 1

            if count > 0:
                self._save_index(force=True)
                logger.info(f"清理了 {count} 条过期缓存")
            return count

    def cleanup_orphaned_files(self) -> int:
        """清理孤立的缓存文件（索引中不存在的文件）

        Returns:
            清理的文件数量
        """
        with self._lock:
            count = 0
            valid_keys = set(self._index.keys())

            for cache_file in self.cache_dir.glob("*.json"):
                if cache_file.name == "index.json":
                    continue
                key = cache_file.stem
                if key not in valid_keys:
                    try:
                        cache_file.unlink()
                        count += 1
                        logger.debug(f"删除孤立文件: {cache_file}")
                    except OSError as e:
                        logger.warning(f"删除孤立文件失败 {cache_file}: {e}")

            if count > 0:
                logger.info(f"清理了 {count} 个孤立缓存文件")
            return count

    def stats(self) -> Dict[str, Any]:
        """获取缓存统计信息"""
        with self._lock:
            total = self._stats["hits"] + self._stats["misses"]
            hit_rate = self._stats["hits"] / total if total > 0 else 0

            # 计算缓存大小
            total_size = 0
            for cache_file in self.cache_dir.glob("*.json"):
                if cache_file.name != "index.json":
                    total_size += cache_file.stat().st_size

            return {
                "total_entries": len(self._index),
                "max_entries": self.max_entries,
                "hits": self._stats["hits"],
                "misses": self._stats["misses"],
                "hit_rate": f"{hit_rate:.2%}",
                "evictions": self._stats["evictions"],
                "total_size_mb": round(total_size / (1024 * 1024), 2),
                "max_age_hours": self.max_age.total_seconds() / 3600
            }

    def get_recent(self, limit: int = 10) -> List[Dict]:
        """获取最近使用的缓存条目

        Args:
            limit: 返回数量

        Returns:
            缓存条目列表
        """
        with self._lock:
            # 从末尾取（最近访问的在末尾）
            items = list(self._index.items())[-limit:]
            return [
                {"key": k, **v}
                for k, v in reversed(items)
            ]


# 全局缓存实例
_cache: Optional[GenerationCache] = None
_cache_lock = threading.Lock()


def get_cache() -> GenerationCache:
    """获取全局缓存实例（线程安全）"""
    global _cache
    if _cache is None:
        with _cache_lock:
            # Double-check locking
            if _cache is None:
                _cache = GenerationCache()
    return _cache


def cleanup_cache() -> Dict[str, int]:
    """执行缓存清理（过期 + 孤立文件）

    Returns:
        清理统计 {"expired": n, "orphaned": n}
    """
    cache = get_cache()
    return {
        "expired": cache.cleanup_expired(),
        "orphaned": cache.cleanup_orphaned_files()
    }
