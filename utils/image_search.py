"""网络图片搜索模块 - 优化版（支持并行下载和缓存）"""
import os
import hashlib
import json
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional, List, Dict
from pathlib import Path
from functools import lru_cache

from config.settings import ImageConfig


class ImageCache:
    """图片缓存管理"""
    
    def __init__(self, cache_dir: str = "images/cache"):
        self.cache_dir = cache_dir
        self.cache_file = os.path.join(cache_dir, "cache_index.json")
        self._cache: Dict[str, str] = {}
        self._load_cache()
    
    def _load_cache(self):
        """加载缓存索引"""
        Path(self.cache_dir).mkdir(parents=True, exist_ok=True)
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
            except:
                self._cache = {}
    
    def _save_cache(self):
        """保存缓存索引"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self._cache, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def get(self, keyword: str) -> Optional[str]:
        """获取缓存的图片路径"""
        key = self._make_key(keyword)
        path = self._cache.get(key)
        if path and os.path.exists(path):
            return path
        return None
    
    def set(self, keyword: str, path: str):
        """设置缓存"""
        key = self._make_key(keyword)
        self._cache[key] = path
        self._save_cache()
    
    def _make_key(self, keyword: str) -> str:
        """生成缓存键"""
        return hashlib.md5(keyword.lower().encode()).hexdigest()[:16]


class ImageSearcher:
    """图片搜索器 - 支持并行下载和缓存"""
    
    def __init__(self, config: Optional[ImageConfig] = None):
        """初始化"""
        self.config = config or ImageConfig.from_env()
        self.api_key = self.config.unsplash_key
        self.base_url = "https://api.unsplash.com"
        self.download_dir = self.config.download_dir
        self.cache = ImageCache() if self.config.cache_enabled else None
        
        Path(self.download_dir).mkdir(parents=True, exist_ok=True)
    
    def search_images(
        self,
        keyword: str,
        per_page: int = 5,
        orientation: str = "landscape"
    ) -> List[Dict]:
        """搜索图片"""
        if not self.api_key:
            print("⚠️ 未设置 UNSPLASH_ACCESS_KEY，无法搜索图片")
            return []
        
        try:
            response = requests.get(
                f"{self.base_url}/search/photos",
                params={"query": keyword, "per_page": per_page, "orientation": orientation},
                headers={"Authorization": f"Client-ID {self.api_key}"},
                timeout=10
            )
            response.raise_for_status()
            
            return [
                {
                    "id": item["id"],
                    "description": item.get("description") or item.get("alt_description", ""),
                    "url": item["urls"]["regular"],
                    "thumb_url": item["urls"]["thumb"],
                    "author": item["user"]["name"],
                }
                for item in response.json().get("results", [])
            ]
        except Exception as e:
            print(f"搜索图片失败: {e}")
            return []
    
    def download_image(self, image_url: str, filename: Optional[str] = None) -> Optional[str]:
        """下载单张图片"""
        try:
            response = requests.get(image_url, timeout=30)
            response.raise_for_status()
            
            if not filename:
                filename = f"image_{hash(image_url) % 100000}.jpg"
            
            filepath = os.path.join(self.download_dir, filename)
            with open(filepath, "wb") as f:
                f.write(response.content)
            
            return filepath
        except Exception as e:
            print(f"下载图片失败: {e}")
            return None
    
    def search_and_download(self, keyword: str, index: int = 0) -> Optional[str]:
        """搜索并下载图片（带缓存）"""
        # 检查缓存
        if self.cache:
            cached = self.cache.get(keyword)
            if cached:
                print(f"✓ 使用缓存图片: {cached}")
                return cached
        
        results = self.search_images(keyword, per_page=index + 1)
        if not results or index >= len(results):
            print(f"未找到关键词 '{keyword}' 的图片")
            return None
        
        image = results[index]
        safe_keyword = "".join(c for c in keyword if c.isalnum() or c in (' ', '-', '_'))
        safe_keyword = safe_keyword.replace(' ', '_')[:30]
        filename = f"{safe_keyword}_{image['id']}.jpg"
        
        filepath = self.download_image(image["url"], filename)
        
        # 更新缓存
        if filepath and self.cache:
            self.cache.set(keyword, filepath)
        
        return filepath
    
    def download_multiple(self, keywords: List[str]) -> Dict[str, Optional[str]]:
        """并行下载多张图片"""
        results = {}
        max_workers = self.config.max_concurrent_downloads
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_keyword = {
                executor.submit(self.search_and_download, kw): kw
                for kw in keywords
            }
            
            for future in as_completed(future_to_keyword):
                keyword = future_to_keyword[future]
                try:
                    results[keyword] = future.result()
                except Exception as e:
                    print(f"下载 '{keyword}' 失败: {e}")
                    results[keyword] = None
        
        return results


# 全局实例（延迟初始化）
_searcher: Optional[ImageSearcher] = None


def get_searcher() -> ImageSearcher:
    """获取全局搜索器实例"""
    global _searcher
    if _searcher is None:
        _searcher = ImageSearcher()
    return _searcher


def search_and_download_image(keyword: str, index: int = 0) -> Optional[str]:
    """搜索并下载图片（便捷函数）"""
    return get_searcher().search_and_download(keyword, index)


def download_images_parallel(keywords: List[str]) -> Dict[str, Optional[str]]:
    """并行下载多张图片（便捷函数）"""
    return get_searcher().download_multiple(keywords)
