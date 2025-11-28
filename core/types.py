"""类型定义模块"""
from typing import TypedDict, List, Optional, Literal


# Slide 类型
SlideType = Literal["bullets", "image_with_text", "ending"]


class SlideDict(TypedDict, total=False):
    """幻灯片字典类型"""
    type: SlideType
    title: str
    bullets: List[str]
    text: str
    image_path: Optional[str]
    image_keyword: Optional[str]
    subtitle: str


class PptPlanDict(TypedDict):
    """PPT 计划字典类型"""
    title: str
    subtitle: str
    slides: List[SlideDict]


class GenerateRequest(TypedDict, total=False):
    """生成请求类型"""
    topic: str
    audience: str
    page_count: int
    description: str
    auto_page_count: bool
    auto_download: bool
    template_id: str
    api_key: str
    api_base: str
    model_name: str
    unsplash_key: str


class GenerateResponse(TypedDict):
    """生成响应类型"""
    success: bool
    filename: str
    title: str
    subtitle: str
    slide_count: int
    download_url: str


class ErrorResponse(TypedDict):
    """错误响应类型"""
    error: str
    details: Optional[str]


class TemplateInfo(TypedDict):
    """模板信息类型"""
    id: str
    name: str
    path: str
    category: str
    description: str
    preview: str


class ImageSearchResult(TypedDict):
    """图片搜索结果类型"""
    id: str
    description: str
    url: str
    thumb_url: str
    download_url: str
    author: str
    author_url: str
