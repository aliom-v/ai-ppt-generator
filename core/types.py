"""类型定义模块"""
from typing import TypedDict, List, Optional, Literal, Dict, Any


# Slide 类型 - 支持7种页面类型
SlideType = Literal[
    "bullets",        # 要点页
    "image_with_text", # 图文混排页
    "two_column",     # 双栏布局页
    "timeline",       # 时间线页
    "comparison",     # 对比页
    "quote",          # 引用/名言页
    "ending"          # 结束页
]


class SlideDict(TypedDict, total=False):
    """幻灯片字典类型"""
    type: SlideType
    title: str
    bullets: List[str]
    text: str
    image_path: Optional[str]
    image_keyword: Optional[str]
    subtitle: str
    # two_column 类型字段
    left_title: str
    left_items: List[str]
    right_title: str
    right_items: List[str]
    # timeline 类型字段
    timeline_items: List[Dict[str, str]]  # [{time: str, event: str}, ...]
    # comparison 类型字段
    comparison_left: Dict[str, Any]  # {title: str, items: List[str]}
    comparison_right: Dict[str, Any]
    # quote 类型字段
    quote_text: str
    quote_author: str


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
