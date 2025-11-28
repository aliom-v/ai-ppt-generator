"""PPT 结构数据模型"""
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional


@dataclass
class Slide:
    """单页幻灯片数据模型
    
    支持的页面类型 (slide_type):
    - bullets: 要点页，包含标题和要点列表
    - image_with_text: 图文混排页，左侧图片 + 右侧文字
    - two_column: 双栏布局页，要点分左右两栏显示
    - timeline: 时间线页，展示流程或时间节点
    - comparison: 对比页，左右对比两个方案
    - quote: 引用/名言页，展示重要引用
    - ending: 结束页，用于致谢或 Q&A
    """
    title: str
    slide_type: str = "bullets"
    bullets: List[str] = field(default_factory=list)
    text: str = ""  # 用于 image_with_text 和 quote 类型的描述文字
    image_path: Optional[str] = None  # 本地图片路径
    image_keyword: Optional[str] = None  # 图片关键词（用于网络搜索）
    subtitle: str = ""  # 用于 title、ending 和 quote 类型


@dataclass
class PptPlan:
    """PPT 整体结构数据模型"""
    title: str
    subtitle: str
    slides: List[Slide] = field(default_factory=list)


def ppt_plan_from_dict(data: Dict[str, Any]) -> PptPlan:
    """从字典创建 PptPlan 对象
    
    Args:
        data: 包含 PPT 结构的字典
        
    Returns:
        PptPlan 对象
        
    Raises:
        ValueError: 当必需字段缺失时
    """
    if not isinstance(data, dict):
        raise ValueError("输入数据必须是字典类型")
    
    title = data.get("title", "")
    subtitle = data.get("subtitle", "")
    
    if not title:
        raise ValueError("PPT 标题不能为空")
    
    slides_data = data.get("slides", [])
    slides = []
    
    for slide_data in slides_data:
        if not isinstance(slide_data, dict):
            continue
            
        slide_title = slide_data.get("title", "")
        slide_type = slide_data.get("type", "bullets")
        bullets = slide_data.get("bullets", [])
        text = slide_data.get("text", "")
        image_path = slide_data.get("image_path")
        image_keyword = slide_data.get("image_keyword")
        slide_subtitle = slide_data.get("subtitle", "")
        
        if not isinstance(bullets, list):
            bullets = []
        
        slides.append(Slide(
            title=slide_title,
            slide_type=slide_type,
            bullets=bullets,
            text=text,
            image_path=image_path,
            image_keyword=image_keyword,
            subtitle=slide_subtitle
        ))
    
    return PptPlan(title=title, subtitle=subtitle, slides=slides)


def ppt_plan_to_dict(plan: PptPlan) -> Dict[str, Any]:
    """将 PptPlan 对象转换为字典（可选功能）"""
    return {
        "title": plan.title,
        "subtitle": plan.subtitle,
        "slides": [
            {
                "title": slide.title,
                "type": slide.slide_type,
                "bullets": slide.bullets,
                "text": slide.text,
                "image_path": slide.image_path,
                "image_keyword": slide.image_keyword,
                "subtitle": slide.subtitle
            }
            for slide in plan.slides
        ]
    }
