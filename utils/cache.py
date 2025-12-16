"""生成缓存模块 - 避免重复调用 AI"""
import os
import json
import hashlib
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime, timedelta

from utils.logger import get_logger

logger = get_logger("cache")


class GenerationCache:
    """PPT 生成结果缓存
    
    根据主题、受众、页数等参数缓存 AI 生成的结果，
    相同参数的请求直接返回缓存，避免重复调用 API。
    """
    
    def __init__(self, cache_dir: str = "cache", max_age_hours: int = 24 * 7):
        """初始化缓存
        
        Args:
            cache_dir: 缓存目录
            max_age_hours: 缓存有效期（小时），默认 7 天
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.max_age = timedelta(hours=max_age_hours)
        self.index_file = self.cache_dir / "index.json"
        self._index: Dict[str, Dict] = {}
        self._load_index()
    
    def _load_index(self) -> None:
        """加载缓存索引"""
        if self.index_file.exists():
            try:
                with open(self.index_file, 'r', encoding='utf-8') as f:
                    self._index = json.load(f)
            except Exception as e:
                logger.warning(f"加载缓存索引失败: {e}")
                self._index = {}
    
    def _save_index(self) -> None:
        """保存缓存索引"""
        try:
            with open(self.index_file, 'w', encoding='utf-8') as f:
                json.dump(self._index, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"保存缓存索引失败: {e}")
    
    def _make_key(self, topic: str, audience: str, page_count: int, 
                  description: str = "", model: str = "") -> str:
        """生成缓存键"""
        content = f"{topic}|{audience}|{page_count}|{description}|{model}"
        return hashlib.md5(content.encode()).hexdigest()[:16]
    
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
        key = self._make_key(topic, audience, page_count, description, model)
        
        if key not in self._index:
            return None
        
        entry = self._index[key]
        cache_file = self.cache_dir / f"{key}.json"
        
        # 检查文件是否存在
        if not cache_file.exists():
            del self._index[key]
            self._save_index()
            return None
        
        # 检查是否过期
        created_at = datetime.fromisoformat(entry.get("created_at", "2000-01-01"))
        if datetime.now() - created_at > self.max_age:
            logger.debug(f"缓存已过期: {topic}")
            self._remove_entry(key)
            return None
        
        # 读取缓存
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            logger.info(f"使用缓存: {topic} ({page_count}页)")
            return data
        except Exception as e:
            logger.warning(f"读取缓存失败: {e}")
            self._remove_entry(key)
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
        key = self._make_key(topic, audience, page_count, description, model)
        cache_file = self.cache_dir / f"{key}.json"
        
        try:
            # 保存数据
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            # 更新索引
            self._index[key] = {
                "topic": topic,
                "audience": audience,
                "page_count": page_count,
                "model": model,
                "created_at": datetime.now().isoformat(),
                "file": str(cache_file)
            }
            self._save_index()
            logger.debug(f"已缓存: {topic}")
        except Exception as e:
            logger.warning(f"保存缓存失败: {e}")
    
    def _remove_entry(self, key: str) -> None:
        """删除缓存条目"""
        if key in self._index:
            del self._index[key]
            self._save_index()
        
        cache_file = self.cache_dir / f"{key}.json"
        if cache_file.exists():
            try:
                cache_file.unlink()
            except:
                pass
    
    def clear(self) -> int:
        """清空所有缓存
        
        Returns:
            删除的缓存数量
        """
        count = len(self._index)
        for key in list(self._index.keys()):
            self._remove_entry(key)
        logger.info(f"已清空 {count} 条缓存")
        return count
    
    def cleanup_expired(self) -> int:
        """清理过期缓存
        
        Returns:
            清理的缓存数量
        """
        count = 0
        for key in list(self._index.keys()):
            entry = self._index[key]
            created_at = datetime.fromisoformat(entry.get("created_at", "2000-01-01"))
            if datetime.now() - created_at > self.max_age:
                self._remove_entry(key)
                count += 1
        
        if count > 0:
            logger.info(f"清理了 {count} 条过期缓存")
        return count


# 全局缓存实例
_cache: Optional[GenerationCache] = None


def get_cache() -> GenerationCache:
    """获取全局缓存实例"""
    global _cache
    if _cache is None:
        _cache = GenerationCache()
    return _cache
