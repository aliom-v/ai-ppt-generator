"""多图片源提供者模块 - 支持 Unsplash、Pexels、Pixabay 和 AI 生图"""
import os
import hashlib
import json
import requests
import base64
from abc import ABC, abstractmethod
from typing import Optional, List, Dict, Any
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from enum import Enum


class ImageProvider(Enum):
    """图片提供者枚举"""
    UNSPLASH = "unsplash"
    PEXELS = "pexels"
    PIXABAY = "pixabay"
    DALLE = "dalle"
    LOCAL = "local"


@dataclass
class ImageResult:
    """图片搜索结果"""
    id: str
    url: str
    thumb_url: str
    description: str
    author: str
    provider: str
    width: int = 0
    height: int = 0


class BaseImageProvider(ABC):
    """图片提供者基类"""

    def __init__(self, api_key: str, download_dir: str = "images/downloaded"):
        self.api_key = api_key
        self.download_dir = download_dir
        Path(download_dir).mkdir(parents=True, exist_ok=True)

    @abstractmethod
    def search(self, keyword: str, per_page: int = 5) -> List[ImageResult]:
        """搜索图片"""
        pass

    def download(self, url: str, filename: str) -> Optional[str]:
        """下载图片"""
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()

            filepath = os.path.join(self.download_dir, filename)
            with open(filepath, "wb") as f:
                f.write(response.content)
            return filepath
        except Exception as e:
            print(f"下载图片失败: {e}")
            return None


class UnsplashProvider(BaseImageProvider):
    """Unsplash 图片提供者"""

    def __init__(self, api_key: str, download_dir: str = "images/downloaded"):
        super().__init__(api_key, download_dir)
        self.base_url = "https://api.unsplash.com"

    def search(self, keyword: str, per_page: int = 5) -> List[ImageResult]:
        if not self.api_key:
            return []

        try:
            response = requests.get(
                f"{self.base_url}/search/photos",
                params={"query": keyword, "per_page": per_page, "orientation": "landscape"},
                headers={"Authorization": f"Client-ID {self.api_key}"},
                timeout=10
            )
            response.raise_for_status()

            results = []
            for item in response.json().get("results", []):
                results.append(ImageResult(
                    id=item["id"],
                    url=item["urls"]["regular"],
                    thumb_url=item["urls"]["thumb"],
                    description=item.get("description") or item.get("alt_description", ""),
                    author=item["user"]["name"],
                    provider="unsplash",
                    width=item.get("width", 0),
                    height=item.get("height", 0)
                ))
            return results
        except Exception as e:
            print(f"Unsplash 搜索失败: {e}")
            return []


class PexelsProvider(BaseImageProvider):
    """Pexels 图片提供者"""

    def __init__(self, api_key: str, download_dir: str = "images/downloaded"):
        super().__init__(api_key, download_dir)
        self.base_url = "https://api.pexels.com/v1"

    def search(self, keyword: str, per_page: int = 5) -> List[ImageResult]:
        if not self.api_key:
            return []

        try:
            response = requests.get(
                f"{self.base_url}/search",
                params={"query": keyword, "per_page": per_page, "orientation": "landscape"},
                headers={"Authorization": self.api_key},
                timeout=10
            )
            response.raise_for_status()

            results = []
            for item in response.json().get("photos", []):
                results.append(ImageResult(
                    id=str(item["id"]),
                    url=item["src"]["large"],
                    thumb_url=item["src"]["tiny"],
                    description=item.get("alt", ""),
                    author=item["photographer"],
                    provider="pexels",
                    width=item.get("width", 0),
                    height=item.get("height", 0)
                ))
            return results
        except Exception as e:
            print(f"Pexels 搜索失败: {e}")
            return []


class PixabayProvider(BaseImageProvider):
    """Pixabay 图片提供者"""

    def __init__(self, api_key: str, download_dir: str = "images/downloaded"):
        super().__init__(api_key, download_dir)
        self.base_url = "https://pixabay.com/api/"

    def search(self, keyword: str, per_page: int = 5) -> List[ImageResult]:
        if not self.api_key:
            return []

        try:
            response = requests.get(
                self.base_url,
                params={
                    "key": self.api_key,
                    "q": keyword,
                    "per_page": per_page,
                    "orientation": "horizontal",
                    "image_type": "photo"
                },
                timeout=10
            )
            response.raise_for_status()

            results = []
            for item in response.json().get("hits", []):
                results.append(ImageResult(
                    id=str(item["id"]),
                    url=item["largeImageURL"],
                    thumb_url=item["previewURL"],
                    description=item.get("tags", ""),
                    author=item["user"],
                    provider="pixabay",
                    width=item.get("imageWidth", 0),
                    height=item.get("imageHeight", 0)
                ))
            return results
        except Exception as e:
            print(f"Pixabay 搜索失败: {e}")
            return []


class DalleProvider(BaseImageProvider):
    """DALL-E AI 生图提供者"""

    def __init__(self, api_key: str, api_base: str = "https://api.openai.com/v1",
                 download_dir: str = "images/downloaded"):
        super().__init__(api_key, download_dir)
        self.api_base = api_base

    def search(self, keyword: str, per_page: int = 1) -> List[ImageResult]:
        """生成图片（DALL-E 是生成而非搜索）"""
        if not self.api_key:
            return []

        try:
            response = requests.post(
                f"{self.api_base}/images/generations",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": "dall-e-3",
                    "prompt": f"Professional presentation slide image about: {keyword}. High quality, clean, corporate style.",
                    "n": 1,
                    "size": "1792x1024",
                    "quality": "standard"
                },
                timeout=60
            )
            response.raise_for_status()

            results = []
            data = response.json()
            for i, item in enumerate(data.get("data", [])):
                results.append(ImageResult(
                    id=f"dalle_{hashlib.md5(keyword.encode()).hexdigest()[:8]}_{i}",
                    url=item.get("url", ""),
                    thumb_url=item.get("url", ""),
                    description=f"AI generated: {keyword}",
                    author="DALL-E",
                    provider="dalle",
                    width=1792,
                    height=1024
                ))
            return results
        except Exception as e:
            print(f"DALL-E 生成失败: {e}")
            return []

    def generate_image(self, prompt: str, size: str = "1792x1024") -> Optional[str]:
        """直接生成图片并保存"""
        results = self.search(prompt, 1)
        if results and results[0].url:
            safe_prompt = "".join(c for c in prompt if c.isalnum() or c in (' ', '-', '_'))[:30]
            filename = f"dalle_{safe_prompt.replace(' ', '_')}_{results[0].id}.png"
            return self.download(results[0].url, filename)
        return None


class LocalImageProvider(BaseImageProvider):
    """本地图片库提供者"""

    def __init__(self, library_dir: str = "images/library", download_dir: str = "images/downloaded"):
        super().__init__("", download_dir)
        self.library_dir = library_dir
        self.index_file = os.path.join(library_dir, "index.json")
        self._index: Dict[str, List[Dict]] = {}
        self._load_index()

    def _load_index(self):
        """加载图片索引"""
        Path(self.library_dir).mkdir(parents=True, exist_ok=True)
        if os.path.exists(self.index_file):
            try:
                with open(self.index_file, 'r', encoding='utf-8') as f:
                    self._index = json.load(f)
            except:
                self._index = {}

    def _save_index(self):
        """保存图片索引"""
        try:
            with open(self.index_file, 'w', encoding='utf-8') as f:
                json.dump(self._index, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存索引失败: {e}")

    def add_image(self, filepath: str, keywords: List[str], description: str = ""):
        """添加图片到库"""
        if not os.path.exists(filepath):
            return False

        filename = os.path.basename(filepath)
        dest_path = os.path.join(self.library_dir, filename)

        # 复制文件到库
        if filepath != dest_path:
            import shutil
            shutil.copy2(filepath, dest_path)

        # 更新索引
        image_info = {
            "id": hashlib.md5(filename.encode()).hexdigest()[:12],
            "filename": filename,
            "path": dest_path,
            "description": description,
            "keywords": keywords
        }

        for keyword in keywords:
            kw_lower = keyword.lower()
            if kw_lower not in self._index:
                self._index[kw_lower] = []
            self._index[kw_lower].append(image_info)

        self._save_index()
        return True

    def search(self, keyword: str, per_page: int = 5) -> List[ImageResult]:
        """搜索本地图片库"""
        keyword_lower = keyword.lower()
        results = []
        seen_ids = set()

        # 精确匹配
        if keyword_lower in self._index:
            for img in self._index[keyword_lower][:per_page]:
                if img["id"] not in seen_ids:
                    results.append(ImageResult(
                        id=img["id"],
                        url=img["path"],
                        thumb_url=img["path"],
                        description=img.get("description", ""),
                        author="Local Library",
                        provider="local"
                    ))
                    seen_ids.add(img["id"])

        # 模糊匹配
        if len(results) < per_page:
            for kw, images in self._index.items():
                if keyword_lower in kw or kw in keyword_lower:
                    for img in images:
                        if img["id"] not in seen_ids and len(results) < per_page:
                            results.append(ImageResult(
                                id=img["id"],
                                url=img["path"],
                                thumb_url=img["path"],
                                description=img.get("description", ""),
                                author="Local Library",
                                provider="local"
                            ))
                            seen_ids.add(img["id"])

        return results

    def list_all(self) -> List[Dict]:
        """列出所有图片"""
        all_images = []
        seen_ids = set()

        for images in self._index.values():
            for img in images:
                if img["id"] not in seen_ids:
                    all_images.append(img)
                    seen_ids.add(img["id"])

        return all_images

    def remove_image(self, image_id: str) -> bool:
        """从库中移除图片"""
        removed = False
        for keyword in list(self._index.keys()):
            self._index[keyword] = [
                img for img in self._index[keyword] if img["id"] != image_id
            ]
            if not self._index[keyword]:
                del self._index[keyword]
            else:
                removed = True

        if removed:
            self._save_index()
        return removed


class MultiSourceImageSearcher:
    """多源图片搜索器"""

    def __init__(self, config: Dict[str, str] = None):
        """初始化多源搜索器

        Args:
            config: 配置字典，包含各平台的 API Key
                - unsplash_key
                - pexels_key
                - pixabay_key
                - openai_key (用于 DALL-E)
                - openai_base (可选)
        """
        config = config or {}
        self.download_dir = config.get("download_dir", "images/downloaded")

        self.providers: Dict[str, BaseImageProvider] = {}

        # 初始化各个提供者
        if config.get("unsplash_key"):
            self.providers["unsplash"] = UnsplashProvider(
                config["unsplash_key"], self.download_dir
            )

        if config.get("pexels_key"):
            self.providers["pexels"] = PexelsProvider(
                config["pexels_key"], self.download_dir
            )

        if config.get("pixabay_key"):
            self.providers["pixabay"] = PixabayProvider(
                config["pixabay_key"], self.download_dir
            )

        if config.get("openai_key"):
            self.providers["dalle"] = DalleProvider(
                config["openai_key"],
                config.get("openai_base", "https://api.openai.com/v1"),
                self.download_dir
            )

        # 本地图片库始终可用
        self.providers["local"] = LocalImageProvider(
            config.get("library_dir", "images/library"),
            self.download_dir
        )

        # 缓存
        self._cache: Dict[str, str] = {}
        self._cache_file = os.path.join(self.download_dir, "../cache/multi_cache.json")
        self._load_cache()

    def _load_cache(self):
        """加载缓存"""
        try:
            cache_dir = os.path.dirname(self._cache_file)
            Path(cache_dir).mkdir(parents=True, exist_ok=True)
            if os.path.exists(self._cache_file):
                with open(self._cache_file, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
        except:
            self._cache = {}

    def _save_cache(self):
        """保存缓存"""
        try:
            with open(self._cache_file, 'w', encoding='utf-8') as f:
                json.dump(self._cache, f, ensure_ascii=False, indent=2)
        except:
            pass

    def _make_cache_key(self, keyword: str, provider: str = None) -> str:
        """生成缓存键"""
        key = f"{keyword.lower()}_{provider or 'any'}"
        return hashlib.md5(key.encode()).hexdigest()[:16]

    def search(self, keyword: str, provider: str = None, per_page: int = 5) -> List[ImageResult]:
        """搜索图片

        Args:
            keyword: 搜索关键词
            provider: 指定提供者（可选），如 'unsplash', 'pexels', 'pixabay', 'dalle', 'local'
            per_page: 每页数量

        Returns:
            图片结果列表
        """
        if provider and provider in self.providers:
            return self.providers[provider].search(keyword, per_page)

        # 搜索所有可用提供者
        all_results = []
        for name, prov in self.providers.items():
            if name != "dalle":  # DALL-E 单独处理（收费）
                results = prov.search(keyword, per_page)
                all_results.extend(results)

        return all_results[:per_page * 2]  # 返回多一些结果供选择

    def search_and_download(self, keyword: str, provider: str = None, index: int = 0) -> Optional[str]:
        """搜索并下载图片"""
        # 检查缓存
        cache_key = self._make_cache_key(keyword, provider)
        if cache_key in self._cache:
            cached_path = self._cache[cache_key]
            if os.path.exists(cached_path):
                print(f"✓ 使用缓存图片: {cached_path}")
                return cached_path

        # 搜索图片
        results = self.search(keyword, provider, index + 3)
        if not results or index >= len(results):
            print(f"未找到关键词 '{keyword}' 的图片")
            return None

        # 下载图片
        result = results[index]
        safe_keyword = "".join(c for c in keyword if c.isalnum() or c in (' ', '-', '_'))
        safe_keyword = safe_keyword.replace(' ', '_')[:30]
        ext = "png" if result.provider == "dalle" else "jpg"
        filename = f"{safe_keyword}_{result.provider}_{result.id}.{ext}"

        # 如果是本地图片，直接返回路径
        if result.provider == "local":
            return result.url

        # 下载
        filepath = self.providers[result.provider].download(result.url, filename)

        # 更新缓存
        if filepath:
            self._cache[cache_key] = filepath
            self._save_cache()

        return filepath

    def generate_ai_image(self, prompt: str) -> Optional[str]:
        """使用 DALL-E 生成图片"""
        if "dalle" not in self.providers:
            print("DALL-E 未配置，请设置 OpenAI API Key")
            return None

        return self.providers["dalle"].generate_image(prompt)

    def download_multiple(self, keywords: List[str], provider: str = None) -> Dict[str, Optional[str]]:
        """并行下载多张图片"""
        results = {}

        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_keyword = {
                executor.submit(self.search_and_download, kw, provider): kw
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

    def get_available_providers(self) -> List[str]:
        """获取可用的图片提供者列表"""
        return list(self.providers.keys())

    def add_to_local_library(self, filepath: str, keywords: List[str], description: str = "") -> bool:
        """添加图片到本地库"""
        if "local" in self.providers:
            return self.providers["local"].add_image(filepath, keywords, description)
        return False

    def list_local_library(self) -> List[Dict]:
        """列出本地图片库"""
        if "local" in self.providers:
            return self.providers["local"].list_all()
        return []


# 便捷函数
_multi_searcher: Optional[MultiSourceImageSearcher] = None


def get_multi_searcher(config: Dict[str, str] = None) -> MultiSourceImageSearcher:
    """获取多源搜索器实例"""
    global _multi_searcher
    if _multi_searcher is None or config:
        _multi_searcher = MultiSourceImageSearcher(config)
    return _multi_searcher


def search_image_multi_source(keyword: str, provider: str = None) -> Optional[str]:
    """多源搜索并下载图片"""
    return get_multi_searcher().search_and_download(keyword, provider)
