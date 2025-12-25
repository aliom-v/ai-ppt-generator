"""网络图片搜索模块 - 优化版（支持并行下载、缓存和重试机制）"""
import os
import hashlib
import json
import time
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional, List, Dict
from pathlib import Path

from config.settings import ImageConfig
from utils.logger import get_logger

logger = get_logger("image_search")

# 重试配置
MAX_RETRIES = 3
RETRY_DELAY = 1  # 秒


class ImageCache:
    """图片缓存管理（带延迟写入优化）"""

    def __init__(self, cache_dir: str = "images/cache", write_delay: int = 5):
        self.cache_dir = cache_dir
        self.cache_file = os.path.join(cache_dir, "cache_index.json")
        self._cache: Dict[str, str] = {}
        self._dirty = False  # 标记是否有未保存的更改
        self._last_save = time.time()
        self._write_delay = write_delay  # 延迟写入时间（秒）
        self._load_cache()

    def _load_cache(self):
        """加载缓存索引"""
        Path(self.cache_dir).mkdir(parents=True, exist_ok=True)
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
            except (json.JSONDecodeError, IOError):
                self._cache = {}  # 缓存文件损坏或无法读取

    def _save_cache(self, force: bool = False):
        """保存缓存索引（延迟写入）"""
        if not self._dirty:
            return

        # 延迟写入：只有超过指定时间才写入
        if not force and (time.time() - self._last_save) < self._write_delay:
            return

        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self._cache, f, ensure_ascii=False, indent=2)
            self._dirty = False
            self._last_save = time.time()
        except IOError:
            pass  # 无法写入缓存文件

    def get(self, keyword: str) -> Optional[str]:
        """获取缓存的图片路径"""
        key = self._make_key(keyword)
        path = self._cache.get(key)
        if path and os.path.exists(path):
            return path
        return None

    def set(self, keyword: str, path: str):
        """设置缓存（延迟写入）"""
        key = self._make_key(keyword)
        self._cache[key] = path
        self._dirty = True
        self._save_cache()  # 尝试保存（会检查延迟）

    def flush(self):
        """强制保存缓存"""
        self._save_cache(force=True)

    def _make_key(self, keyword: str) -> str:
        """生成缓存键（使用 SHA256 避免碰撞）"""
        return hashlib.sha256(keyword.lower().encode()).hexdigest()[:24]


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
            logger.warning("未设置 UNSPLASH_ACCESS_KEY，无法搜索图片")
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
            logger.error(f"搜索图片失败: {e}")
            return []
    
    def download_image(
        self,
        image_url: str,
        filename: Optional[str] = None,
        max_size_mb: float = 10.0
    ) -> Optional[str]:
        """下载单张图片（带重试机制和大小限制）

        Args:
            image_url: 图片 URL
            filename: 保存的文件名
            max_size_mb: 最大文件大小（MB），默认 10MB

        Returns:
            下载后的文件路径，失败返回 None
        """
        last_error = None
        max_size_bytes = int(max_size_mb * 1024 * 1024)

        for attempt in range(MAX_RETRIES):
            try:
                # 使用流式下载，限制大小
                with requests.get(image_url, timeout=self.config.download_timeout, stream=True) as response:
                    response.raise_for_status()

                    # 检查 Content-Type
                    content_type = response.headers.get('Content-Type', '')
                    if not content_type.startswith('image/'):
                        logger.warning(f"非图片类型: {content_type}")
                        return None

                    # 检查 Content-Length（如果有）
                    content_length = response.headers.get('Content-Length')
                    if content_length and int(content_length) > max_size_bytes:
                        logger.warning(f"图片过大: {int(content_length) / 1024 / 1024:.1f}MB > {max_size_mb}MB")
                        return None

                    if not filename:
                        url_hash = hashlib.sha256(image_url.encode()).hexdigest()[:12]
                        filename = f"image_{url_hash}.jpg"

                    filepath = os.path.join(self.download_dir, filename)

                    # 流式写入，同时检查大小
                    downloaded_size = 0
                    with open(filepath, "wb") as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            if chunk:
                                downloaded_size += len(chunk)
                                if downloaded_size > max_size_bytes:
                                    f.close()
                                    os.remove(filepath)
                                    logger.warning(f"图片下载超过大小限制: {max_size_mb}MB")
                                    return None
                                f.write(chunk)

                    # 验证文件是否为有效图片
                    if not self._validate_image_file(filepath):
                        os.remove(filepath)
                        logger.warning(f"下载的文件不是有效图片")
                        return None

                    return filepath

            except requests.exceptions.Timeout as e:
                last_error = e
                if attempt < MAX_RETRIES - 1:
                    logger.warning(f"下载超时，第 {attempt + 1} 次重试...")
                    time.sleep(RETRY_DELAY * (attempt + 1))
            except requests.exceptions.RequestException as e:
                last_error = e
                if attempt < MAX_RETRIES - 1:
                    logger.warning(f"下载失败: {e}，第 {attempt + 1} 次重试...")
                    time.sleep(RETRY_DELAY * (attempt + 1))
            except Exception as e:
                logger.error(f"下载图片失败: {e}")
                return None

        logger.error(f"下载图片失败，已重试 {MAX_RETRIES} 次: {last_error}")
        return None

    def _validate_image_file(self, filepath: str) -> bool:
        """验证文件是否为有效图片（检查文件头魔数）"""
        try:
            with open(filepath, 'rb') as f:
                header = f.read(16)

            # 检查常见图片格式的魔数
            # JPEG: FF D8 FF
            if header[:3] == b'\xff\xd8\xff':
                return True
            # PNG: 89 50 4E 47 0D 0A 1A 0A
            if header[:8] == b'\x89PNG\r\n\x1a\n':
                return True
            # GIF: 47 49 46 38
            if header[:4] in (b'GIF8', b'GIF9'):
                return True
            # WebP: 52 49 46 46 ... 57 45 42 50
            if header[:4] == b'RIFF' and header[8:12] == b'WEBP':
                return True
            # BMP: 42 4D
            if header[:2] == b'BM':
                return True

            return False
        except Exception:
            return False
    
    def search_and_download(self, keyword: str, index: int = 0) -> Optional[str]:
        """搜索并下载图片（带缓存）"""
        # 检查缓存
        if self.cache:
            cached = self.cache.get(keyword)
            if cached:
                logger.debug(f"使用缓存图片: {cached}")
                return cached
        
        results = self.search_images(keyword, per_page=index + 1)
        if not results or index >= len(results):
            logger.warning(f"未找到关键词 '{keyword}' 的图片")
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
                    logger.error(f"下载 '{keyword}' 失败: {e}")
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


def reset_searcher(config: Optional[ImageConfig] = None) -> None:
    """重置全局搜索器（用于更新配置）"""
    global _searcher
    _searcher = ImageSearcher(config) if config else None


def search_and_download_image(keyword: str, index: int = 0) -> Optional[str]:
    """搜索并下载图片（便捷函数）"""
    return get_searcher().search_and_download(keyword, index)


def download_images_parallel(keywords: List[str]) -> Dict[str, Optional[str]]:
    """并行下载多张图片（便捷函数）"""
    return get_searcher().download_multiple(keywords)
