"""统一的 PPT 构建器 - 增强版（支持更多页面类型、跨平台字体、统一样式）"""
import os
import platform
from typing import Optional, List, Tuple
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from core.ppt_plan import PptPlan, Slide

# 尝试导入图片搜索模块
try:
    from utils.image_search import search_and_download_image, download_images_parallel
    IMAGE_SEARCH_AVAILABLE = True
except ImportError:
    IMAGE_SEARCH_AVAILABLE = False


# 跨平台字体配置
def get_default_fonts() -> Tuple[str, str]:
    """获取跨平台默认字体"""
    system = platform.system()
    if system == "Windows":
        return ("微软雅黑", "Microsoft YaHei")
    elif system == "Darwin":  # macOS
        return ("PingFang SC", "Helvetica Neue")
    else:  # Linux
        return ("Noto Sans CJK SC", "DejaVu Sans")


FONT_CN, FONT_EN = get_default_fonts()


# 颜色主题
class ColorTheme:
    """颜色主题配置"""
    PRIMARY = RGBColor(25, 118, 210)     # 主色-蓝色
    SECONDARY = RGBColor(66, 165, 245)   # 辅色-浅蓝
    ACCENT = RGBColor(255, 152, 0)       # 强调色-橙色
    TEXT_DARK = RGBColor(51, 51, 51)     # 深色文字
    TEXT_LIGHT = RGBColor(127, 127, 127) # 浅色文字
    BG_LIGHT = RGBColor(240, 240, 240)   # 浅色背景
    WHITE = RGBColor(255, 255, 255)
    SUCCESS = RGBColor(76, 175, 80)      # 绿色
    WARNING = RGBColor(255, 193, 7)      # 黄色


# 全局计数器，用于交替样式
_bullets_style_counter = 0


def build_ppt_from_plan(
    plan: PptPlan,
    template_path: Optional[str],
    output_path: str,
    auto_download_images: bool = False
) -> None:
    """根据 PptPlan 生成 PPTX 文件"""
    global _bullets_style_counter
    _bullets_style_counter = 0  # 重置计数器
    
    use_template = template_path and os.path.exists(template_path) and template_path.endswith('.pptx')
    
    if use_template:
        print(f"✓ 使用模板: {template_path}")
        prs = Presentation(template_path)
    else:
        print("✓ 使用默认样式")
        prs = Presentation()
    
    # 预下载所有图片（并行）
    if auto_download_images and IMAGE_SEARCH_AVAILABLE:
        _predownload_images(plan.slides)
    
    # 创建封面页
    _create_title_slide(prs, plan.title, plan.subtitle)
    
    # 创建内容页
    for slide_data in plan.slides:
        slide_type = slide_data.slide_type.lower()
        
        if slide_type == "bullets":
            _create_bullets_slide(prs, slide_data)
        elif slide_type == "image_with_text":
            _create_image_with_text_slide(prs, slide_data)
        elif slide_type == "two_column":
            _create_two_column_slide(prs, slide_data)
        elif slide_type == "timeline":
            _create_timeline_slide(prs, slide_data)
        elif slide_type == "comparison":
            _create_comparison_slide(prs, slide_data)
        elif slide_type == "quote":
            _create_quote_slide(prs, slide_data)
        elif slide_type == "ending":
            _create_ending_slide(prs, slide_data)
        else:
            _create_bullets_slide(prs, slide_data)
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    prs.save(output_path)
    print(f"✓ PPT 已保存: {output_path}")


def _predownload_images(slides: List[Slide]) -> None:
    """预下载所有需要的图片（并行）"""
    keywords = []
    keyword_to_slides: dict = {}
    
    for slide in slides:
        if slide.slide_type.lower() == "image_with_text":
            if slide.image_keyword and not (slide.image_path and os.path.exists(slide.image_path)):
                kw = slide.image_keyword
                if kw not in keyword_to_slides:
                    keywords.append(kw)
                    keyword_to_slides[kw] = []
                keyword_to_slides[kw].append(slide)
    
    if not keywords:
        return
    
    print(f"📥 并行下载 {len(keywords)} 张图片...")
    results = download_images_parallel(keywords)
    
    for keyword, path in results.items():
        if path and keyword in keyword_to_slides:
            for slide in keyword_to_slides[keyword]:
                slide.image_path = path
            print(f"  ✓ {keyword}")


def _set_font(text_frame, font_name: str = None, font_size: int = None, bold: bool = False, color: RGBColor = None):
    """设置文本框字体（跨平台兼容）"""
    if font_name is None:
        font_name = FONT_CN
    
    if text_frame and text_frame.paragraphs:
        for paragraph in text_frame.paragraphs:
            paragraph.font.name = font_name
            if font_size:
                paragraph.font.size = Pt(font_size)
            if bold:
                paragraph.font.bold = True
            if color:
                paragraph.font.color.rgb = color
            for run in paragraph.runs:
                run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if bold:
                    run.font.bold = True
                if color:
                    run.font.color.rgb = color


def _add_header_decoration(slide, prs: Presentation):
    """添加页面顶部装饰条"""
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Pt(8)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = ColorTheme.PRIMARY
    header.line.fill.background()


def _add_footer_decoration(slide, prs: Presentation):
    """添加页面底部装饰"""
    footer = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), prs.slide_height - Pt(4),
        prs.slide_width, Pt(4)
    )
    footer.fill.solid()
    footer.fill.fore_color.rgb = ColorTheme.SECONDARY
    footer.line.fill.background()


def _add_side_accent(slide, prs: Presentation):
    """添加左侧装饰条"""
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0.8),
        Pt(6), Inches(0.6)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = ColorTheme.ACCENT
    accent.line.fill.background()


def _create_title_slide(prs: Presentation, title: str, subtitle: str) -> None:
    """创建封面页"""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    
    try:
        slide.shapes.title.text = title
        _set_font(slide.shapes.title.text_frame, font_size=44, bold=True)
    except:
        pass
    
    try:
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = subtitle
            _set_font(slide.placeholders[1].text_frame, font_size=24)
    except:
        pass


def _get_slide_width_inches(prs: Presentation) -> float:
    """获取幻灯片宽度（英寸）"""
    return prs.slide_width.inches if hasattr(prs.slide_width, 'inches') else prs.slide_width / 914400


def _create_bullets_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建要点页 - 多种样式交替"""
    global _bullets_style_counter
    
    bullets = [b for b in (slide_data.bullets or []) if b]
    num_bullets = len(bullets)
    
    if num_bullets == 0:
        return
    
    # 根据计数器选择样式（4种样式交替）
    style = _bullets_style_counter % 4
    _bullets_style_counter += 1
    
    if style == 0:
        _create_bullets_style_cards(prs, slide_data, bullets)
    elif style == 1:
        _create_bullets_style_list(prs, slide_data, bullets)
    elif style == 2:
        _create_bullets_style_icons(prs, slide_data, bullets)
    else:
        _create_bullets_style_gradient(prs, slide_data, bullets)


def _create_bullets_style_cards(prs: Presentation, slide_data: Slide, bullets: List[str]) -> None:
    """样式1: 卡片式布局 - 居中"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=28, bold=True, color=ColorTheme.TEXT_DARK)
    
    num_bullets = len(bullets)
    card_colors = [
        RGBColor(227, 242, 253), RGBColor(232, 245, 233),
        RGBColor(255, 243, 224), RGBColor(243, 229, 245),
        RGBColor(255, 235, 238), RGBColor(224, 247, 250),
    ]
    accent_colors = [
        ColorTheme.PRIMARY, ColorTheme.SUCCESS, ColorTheme.ACCENT,
        RGBColor(156, 39, 176), RGBColor(244, 67, 54), RGBColor(0, 188, 212),
    ]
    
    # 统一使用列表式布局，更适合长文本
    content_width = slide_w - 1.0  # 左右各留0.5英寸
    start_x = 0.5
    spacing = min(1.1, 5.5 / max(num_bullets, 1))
    
    for i, bullet in enumerate(bullets[:5]):
        y = 1.0 + i * spacing
        _draw_bullet_card_horizontal(slide, start_x, y, content_width, spacing - 0.1, i + 1, bullet,
                                    card_colors[i % len(card_colors)], accent_colors[i % len(accent_colors)])


def _create_bullets_style_list(prs: Presentation, slide_data: Slide, bullets: List[str]) -> None:
    """样式2: 简洁列表式 - 居中"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=28, bold=True, color=ColorTheme.TEXT_DARK)
    
    # 居中计算
    content_width = slide_w - 1.0
    start_x = 0.5
    
    # 内容区域背景 - 居中
    content_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(start_x), Inches(0.95), Inches(content_width), Inches(5.8)
    )
    content_bg.fill.solid()
    content_bg.fill.fore_color.rgb = RGBColor(250, 250, 250)
    content_bg.line.fill.background()
    
    num_bullets = len(bullets)
    total_chars = sum(len(b) for b in bullets)
    
    # 根据内容量动态调整
    if total_chars > 500 or num_bullets > 5:
        font_size, spacing = 12, 1.05
    elif total_chars > 350 or num_bullets > 4:
        font_size, spacing = 13, 1.1
    else:
        font_size, spacing = 14, 1.15
    
    colors = [ColorTheme.PRIMARY, ColorTheme.SUCCESS, ColorTheme.ACCENT,
              RGBColor(156, 39, 176), RGBColor(0, 188, 212)]
    
    for i, bullet in enumerate(bullets[:5]):
        y = 1.1 + i * spacing
        
        # 序号圆点
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(start_x + 0.15), Inches(y + 0.08), Inches(0.32), Inches(0.32))
        dot.fill.solid()
        dot.fill.fore_color.rgb = colors[i % len(colors)]
        dot.line.fill.background()
        
        num_box = slide.shapes.add_textbox(Inches(start_x + 0.15), Inches(y + 0.1), Inches(0.32), Inches(0.32))
        num_frame = num_box.text_frame
        num_frame.text = str(i + 1)
        num_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        _set_font(num_frame, font_size=12, bold=True, color=ColorTheme.WHITE)
        
        # 文字
        text_box = slide.shapes.add_textbox(Inches(start_x + 0.6), Inches(y), Inches(content_width - 0.75), Inches(spacing))
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.text = bullet
        _set_font(text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)


def _create_bullets_style_icons(prs: Presentation, slide_data: Slide, bullets: List[str]) -> None:
    """样式3: 图标式布局（带大序号）- 居中优化"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=26, bold=True, color=ColorTheme.TEXT_DARK)
    
    num_bullets = len(bullets)
    colors = [ColorTheme.PRIMARY, ColorTheme.SUCCESS, ColorTheme.ACCENT,
              RGBColor(156, 39, 176), RGBColor(0, 188, 212)]
    
    # 统一使用列表式布局
    content_width = slide_w - 1.0
    start_x = 0.5
    spacing = min(1.1, 5.5 / max(num_bullets, 1))
    
    for i, bullet in enumerate(bullets[:5]):
        y = 1.0 + i * spacing
        
        # 序号方块
        num_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(start_x), Inches(y), Inches(0.65), Inches(spacing - 0.15)
        )
        num_bg.fill.solid()
        num_bg.fill.fore_color.rgb = colors[i % len(colors)]
        num_bg.line.fill.background()
        
        num_box = slide.shapes.add_textbox(Inches(start_x), Inches(y + 0.12), Inches(0.65), Inches(0.5))
        num_frame = num_box.text_frame
        num_frame.text = str(i + 1)
        num_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        _set_font(num_frame, font_size=22, bold=True, color=ColorTheme.WHITE)
        
        # 内容背景
        content_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(start_x + 0.75), Inches(y), Inches(content_width - 0.75), Inches(spacing - 0.15)
        )
        content_bg.fill.solid()
        content_bg.fill.fore_color.rgb = RGBColor(250, 250, 250)
        content_bg.line.fill.background()
        
        # 内容文字
        text_len = len(bullet)
        font_size = 11 if text_len > 80 else 12 if text_len > 50 else 13
        text_box = slide.shapes.add_textbox(Inches(start_x + 0.85), Inches(y + 0.08), Inches(content_width - 0.95), Inches(spacing - 0.2))
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.text = bullet
        _set_font(text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)


def _create_bullets_style_gradient(prs: Presentation, slide_data: Slide, bullets: List[str]) -> None:
    """样式4: 渐变色条式 - 居中优化"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=26, bold=True, color=ColorTheme.TEXT_DARK)
    
    num_bullets = len(bullets)
    
    # 渐变色系
    gradient_colors = [
        RGBColor(25, 118, 210),   # 蓝
        RGBColor(56, 142, 60),    # 绿
        RGBColor(245, 124, 0),    # 橙
        RGBColor(123, 31, 162),   # 紫
        RGBColor(211, 47, 47),    # 红
    ]
    
    bg_colors = [
        RGBColor(227, 242, 253),
        RGBColor(232, 245, 233),
        RGBColor(255, 243, 224),
        RGBColor(243, 229, 245),
        RGBColor(255, 235, 238),
    ]
    
    # 居中计算
    content_width = slide_w - 1.0
    start_x = 0.5
    spacing = min(1.1, 5.5 / max(num_bullets, 1))
    
    for i, bullet in enumerate(bullets[:5]):
        y = 1.0 + i * spacing
        
        # 背景条 - 居中
        bar_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(start_x), Inches(y), Inches(content_width), Inches(spacing - 0.1)
        )
        bar_bg.fill.solid()
        bar_bg.fill.fore_color.rgb = bg_colors[i % len(bg_colors)]
        bar_bg.line.fill.background()
        
        # 左侧色条
        left_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(start_x), Inches(y), Pt(8), Inches(spacing - 0.1)
        )
        left_bar.fill.solid()
        left_bar.fill.fore_color.rgb = gradient_colors[i % len(gradient_colors)]
        left_bar.line.fill.background()
        
        # 序号
        num_box = slide.shapes.add_textbox(Inches(start_x + 0.2), Inches(y + 0.05), Inches(0.4), Inches(spacing - 0.2))
        num_frame = num_box.text_frame
        num_frame.text = str(i + 1)
        _set_font(num_frame, font_size=18, bold=True, color=gradient_colors[i % len(gradient_colors)])
        
        # 内容
        text_len = len(bullet)
        font_size = 11 if text_len > 100 else 12 if text_len > 60 else 13
        text_box = slide.shapes.add_textbox(Inches(start_x + 0.65), Inches(y + 0.1), Inches(content_width - 0.85), Inches(spacing - 0.2))
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.text = bullet
        _set_font(text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)


def _draw_bullet_card(slide, x, y, width, height, num, text, bg_color, accent_color, horizontal=False):
    """绘制单个要点卡片（垂直布局）"""
    # 卡片背景
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y),
        Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.fill.background()
    
    # 序号圆圈
    circle_size = 0.4
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x + 0.15), Inches(y + 0.15),
        Inches(circle_size), Inches(circle_size)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = accent_color
    circle.line.fill.background()
    
    # 序号文字
    num_box = slide.shapes.add_textbox(
        Inches(x + 0.15), Inches(y + 0.18),
        Inches(circle_size), Inches(circle_size)
    )
    num_frame = num_box.text_frame
    num_frame.text = str(num)
    num_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(num_frame, font_size=14, bold=True, color=ColorTheme.WHITE)
    
    # 内容文字
    text_len = len(text)
    if horizontal:
        # 横向布局：序号在左，文字在右
        if text_len > 80:
            font_size = 11
        elif text_len > 50:
            font_size = 12
        else:
            font_size = 13
        text_box = slide.shapes.add_textbox(
            Inches(x + 0.7), Inches(y + 0.2),
            Inches(width - 0.9), Inches(height - 0.4)
        )
    else:
        # 垂直布局：序号在上，文字在下
        if text_len > 100:
            font_size = 11
        elif text_len > 60:
            font_size = 12
        else:
            font_size = 13
        text_box = slide.shapes.add_textbox(
            Inches(x + 0.15), Inches(y + 0.7),
            Inches(width - 0.3), Inches(height - 0.9)
        )
    
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    text_frame.text = text
    _set_font(text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)


def _draw_bullet_card_horizontal(slide, x, y, width, height, num, text, bg_color, accent_color):
    """绘制横向要点卡片（列表式）"""
    # 卡片背景
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y),
        Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.fill.background()
    
    # 左侧色条
    accent_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y),
        Pt(6), Inches(height)
    )
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = accent_color
    accent_bar.line.fill.background()
    
    # 序号圆圈
    circle_size = 0.35
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x + 0.15), Inches(y + (height - 0.35) / 2),
        Inches(circle_size), Inches(circle_size)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = accent_color
    circle.line.fill.background()
    
    # 序号文字
    num_box = slide.shapes.add_textbox(
        Inches(x + 0.15), Inches(y + (height - 0.35) / 2 + 0.02),
        Inches(circle_size), Inches(circle_size)
    )
    num_frame = num_box.text_frame
    num_frame.text = str(num)
    num_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(num_frame, font_size=12, bold=True, color=ColorTheme.WHITE)
    
    # 内容文字
    text_len = len(text)
    if text_len > 80:
        font_size = 12
    elif text_len > 50:
        font_size = 13
    else:
        font_size = 14
    
    text_box = slide.shapes.add_textbox(
        Inches(x + 0.6), Inches(y + 0.1),
        Inches(width - 0.8), Inches(height - 0.2)
    )
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    text_frame.text = text
    text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
    _set_font(text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)


def _create_image_with_text_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建图文混排页 - 居中布局"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=28, bold=True, color=ColorTheme.TEXT_DARK)
    
    # 计算居中位置
    content_width = slide_w - 1.0
    start_x = 0.5
    img_width = content_width * 0.45
    text_width = content_width * 0.52
    gap = content_width * 0.03
    
    # 图片或占位符 - 左侧
    img_left = start_x
    if slide_data.image_path and os.path.exists(slide_data.image_path):
        try:
            slide.shapes.add_picture(
                slide_data.image_path,
                Inches(img_left), Inches(1.1),
                width=Inches(img_width)
            )
        except Exception as e:
            print(f"⚠️ 无法插入图片: {e}")
            _add_image_placeholder(slide, slide_data, img_left, img_width)
    else:
        _add_image_placeholder(slide, slide_data, img_left, img_width)
    
    # 文字说明（带背景框）- 右侧
    text_left = img_left + img_width + gap
    text_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(text_left), Inches(1.1),
        Inches(text_width), Inches(5.5)
    )
    text_bg.fill.solid()
    text_bg.fill.fore_color.rgb = ColorTheme.BG_LIGHT
    text_bg.line.fill.background()
    
    # 根据文字长度动态调整字体
    text_content = slide_data.text or "（图片说明）"
    text_len = len(text_content)
    if text_len > 300:
        font_size = 12
    elif text_len > 220:
        font_size = 13
    elif text_len > 150:
        font_size = 14
    else:
        font_size = 15
    
    text_box = slide.shapes.add_textbox(Inches(text_left + 0.15), Inches(1.25), Inches(text_width - 0.3), Inches(5.2))
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    text_frame.text = text_content
    _set_font(text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)


def _add_image_placeholder(slide, slide_data: Slide, left: float = 0.75, width: float = 4.0) -> None:
    """添加图片占位符"""
    # 占位符背景
    bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(1.2),
        Inches(width), Inches(5.3)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = ColorTheme.BG_LIGHT
    bg.line.color.rgb = RGBColor(200, 200, 200)
    bg.line.width = Pt(1)
    
    box = slide.shapes.add_textbox(Inches(left), Inches(1.2), Inches(width), Inches(5.3))
    frame = box.text_frame
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    text = f"🖼️ 图片占位符\n\n关键词: {slide_data.image_keyword}" if slide_data.image_keyword else "🖼️ 图片占位符\n\n请在此处插入图片"
    
    frame.text = text
    p = frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(14)
    p.font.name = FONT_CN
    p.font.color.rgb = ColorTheme.TEXT_LIGHT


def _create_two_column_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建双栏布局页 - 居中优化"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=26, bold=True, color=ColorTheme.TEXT_DARK)
    
    # 分割要点为左右两栏
    bullets = slide_data.bullets or []
    mid = (len(bullets) + 1) // 2
    left_bullets = bullets[:mid]
    right_bullets = bullets[mid:]
    
    # 居中计算
    content_width = slide_w - 1.0
    start_x = 0.5
    col_width = (content_width - 0.3) / 2
    gap = 0.3
    
    # 左栏背景
    left_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(start_x), Inches(1.0),
        Inches(col_width), Inches(5.6)
    )
    left_bg.fill.solid()
    left_bg.fill.fore_color.rgb = ColorTheme.BG_LIGHT
    left_bg.line.fill.background()
    
    # 左栏内容
    left_box = slide.shapes.add_textbox(Inches(start_x + 0.12), Inches(1.15), Inches(col_width - 0.24), Inches(5.3))
    left_frame = left_box.text_frame
    left_frame.word_wrap = True
    for i, bullet in enumerate(left_bullets):
        p = left_frame.paragraphs[0] if i == 0 else left_frame.add_paragraph()
        p.text = f"• {bullet}"
        p.font.name = FONT_CN
        p.font.size = Pt(12)
        p.font.color.rgb = ColorTheme.TEXT_DARK
        p.space_after = Pt(10)
    
    # 右栏背景
    right_x = start_x + col_width + gap
    right_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(right_x), Inches(1.0),
        Inches(col_width), Inches(5.6)
    )
    right_bg.fill.solid()
    right_bg.fill.fore_color.rgb = RGBColor(240, 248, 255)  # 浅蓝色
    right_bg.line.fill.background()
    
    # 右栏内容
    right_box = slide.shapes.add_textbox(Inches(right_x + 0.12), Inches(1.15), Inches(col_width - 0.24), Inches(5.3))
    right_frame = right_box.text_frame
    right_frame.word_wrap = True
    for i, bullet in enumerate(right_bullets):
        p = right_frame.paragraphs[0] if i == 0 else right_frame.add_paragraph()
        p.text = f"• {bullet}"
        p.font.name = FONT_CN
        p.font.size = Pt(12)
        p.font.color.rgb = ColorTheme.TEXT_DARK
        p.space_after = Pt(10)


def _create_timeline_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建时间线页 - 居中优化版"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=26, bold=True, color=ColorTheme.TEXT_DARK)
    
    bullets = slide_data.bullets or []
    num_items = min(len(bullets), 5)  # 最多5个时间点
    
    if num_items == 0:
        return
    
    # 计算最长文本长度，决定卡片大小
    max_text_len = max(len(b) for b in bullets[:num_items])
    
    # 绘制时间线主轴 - 居中
    content_width = slide_w - 1.0
    start_x = 0.5
    line_y = Inches(3.75)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(start_x), line_y,
        Inches(content_width), Pt(4)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = ColorTheme.PRIMARY
    line.line.fill.background()
    
    # 根据数量计算间距 - 基于实际宽度
    item_spacing = content_width / (num_items + 1)
    positions = [start_x + item_spacing * (i + 1) for i in range(num_items)]
    
    colors = [ColorTheme.PRIMARY, ColorTheme.SECONDARY, ColorTheme.ACCENT, ColorTheme.SUCCESS, ColorTheme.WARNING]
    
    # 根据文本长度决定卡片宽度
    if max_text_len > 50:
        card_width = 1.8
    elif max_text_len > 35:
        card_width = 1.6
    else:
        card_width = 1.5
    
    for i, bullet in enumerate(bullets[:num_items]):
        x = Inches(positions[i])
        
        # 圆点
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x - Inches(0.12), line_y - Inches(0.1),
            Inches(0.24), Inches(0.24)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = colors[i % len(colors)]
        circle.line.color.rgb = ColorTheme.WHITE
        circle.line.width = Pt(2)
        
        # 序号
        num_box = slide.shapes.add_textbox(x - Inches(0.08), line_y - Inches(0.06), Inches(0.16), Inches(0.16))
        num_frame = num_box.text_frame
        num_frame.text = str(i + 1)
        num_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        _set_font(num_frame, font_size=10, bold=True, color=ColorTheme.WHITE)
        
        # 文字卡片（交替上下）
        is_top = i % 2 == 0
        card_y = Inches(1.0) if is_top else Inches(4.1)
        card_height = 2.4
        
        # 卡片背景
        card_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x - Inches(card_width / 2), card_y,
            Inches(card_width), Inches(card_height)
        )
        card_bg.fill.solid()
        card_bg.fill.fore_color.rgb = colors[i % len(colors)]
        card_bg.line.fill.background()
        
        # 连接线
        conn_y = Inches(3.4) if is_top else line_y + Pt(4)
        conn_height = Inches(0.35) if is_top else Inches(0.35)
        conn = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            x - Pt(2), conn_y,
            Pt(4), conn_height
        )
        conn.fill.solid()
        conn.fill.fore_color.rgb = colors[i % len(colors)]
        conn.line.fill.background()
        
        # 卡片文字 - 根据内容长度调整字体
        text_len = len(bullet)
        if text_len > 60:
            font_size = 8
        elif text_len > 45:
            font_size = 9
        elif text_len > 30:
            font_size = 10
        else:
            font_size = 11
        
        text_box = slide.shapes.add_textbox(
            x - Inches(card_width / 2 - 0.08), card_y + Inches(0.08),
            Inches(card_width - 0.16), Inches(card_height - 0.16)
        )
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.text = bullet
        text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        _set_font(text_frame, font_size=font_size, color=ColorTheme.WHITE)


def _create_comparison_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建对比页（左右对比）- 居中优化"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 标题 - 居中
    margin = 0.4
    title_box = slide.shapes.add_textbox(Inches(margin), Inches(0.25), Inches(slide_w - 2*margin), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = slide_data.title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(title_frame, font_size=26, bold=True, color=ColorTheme.TEXT_DARK)
    
    bullets = slide_data.bullets or []
    mid = (len(bullets) + 1) // 2
    left_items = bullets[:mid]
    right_items = bullets[mid:]
    
    # 居中计算
    content_width = slide_w - 1.0
    start_x = 0.5
    col_width = (content_width - 0.3) / 2
    gap = 0.3
    right_x = start_x + col_width + gap
    
    # 左侧标题框
    left_header = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(start_x), Inches(0.9),
        Inches(col_width), Inches(0.45)
    )
    left_header.fill.solid()
    left_header.fill.fore_color.rgb = ColorTheme.PRIMARY
    left_header.line.fill.background()
    left_header.text_frame.text = "方案 A"
    left_header.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(left_header.text_frame, font_size=15, bold=True, color=ColorTheme.WHITE)
    
    # 左侧内容背景
    left_content_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(start_x), Inches(1.4),
        Inches(col_width), Inches(5.2)
    )
    left_content_bg.fill.solid()
    left_content_bg.fill.fore_color.rgb = RGBColor(240, 248, 255)
    left_content_bg.line.fill.background()
    
    # 左侧内容
    left_box = slide.shapes.add_textbox(Inches(start_x + 0.15), Inches(1.65), Inches(col_width - 0.3), Inches(4.8))
    left_frame = left_box.text_frame
    left_frame.word_wrap = True
    for i, item in enumerate(left_items):
        p = left_frame.paragraphs[0] if i == 0 else left_frame.add_paragraph()
        p.text = f"• {item}"
        p.font.name = FONT_CN
        p.font.size = Pt(12)
        p.font.color.rgb = ColorTheme.TEXT_DARK
        p.space_after = Pt(10)
    
    # 右侧标题框
    right_header = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(right_x), Inches(0.9),
        Inches(col_width), Inches(0.45)
    )
    right_header.fill.solid()
    right_header.fill.fore_color.rgb = ColorTheme.ACCENT
    right_header.line.fill.background()
    right_header.text_frame.text = "方案 B"
    right_header.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(right_header.text_frame, font_size=15, bold=True, color=ColorTheme.WHITE)
    
    # 右侧内容背景
    right_content_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(right_x), Inches(1.4),
        Inches(col_width), Inches(5.2)
    )
    right_content_bg.fill.solid()
    right_content_bg.fill.fore_color.rgb = RGBColor(255, 248, 240)
    right_content_bg.line.fill.background()
    
    # 右侧内容
    right_box = slide.shapes.add_textbox(Inches(right_x + 0.12), Inches(1.55), Inches(col_width - 0.24), Inches(4.9))
    right_frame = right_box.text_frame
    right_frame.word_wrap = True
    for i, item in enumerate(right_items):
        p = right_frame.paragraphs[0] if i == 0 else right_frame.add_paragraph()
        p.text = f"• {item}"
        p.font.name = FONT_CN
        p.font.size = Pt(12)
        p.font.color.rgb = ColorTheme.TEXT_DARK
        p.space_after = Pt(10)


def _create_quote_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建引用/名言页 - 居中优化"""
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    slide_w = _get_slide_width_inches(prs)
    _add_header_decoration(slide, prs)
    
    # 居中计算
    content_width = slide_w - 1.0
    start_x = 0.5
    
    # 背景装饰 - 居中
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(start_x), Inches(1.0),
        Inches(content_width), Inches(5.6)
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = ColorTheme.BG_LIGHT
    bg_shape.line.fill.background()
    
    # 大引号装饰 - 居中
    quote_mark = slide.shapes.add_textbox(Inches(start_x + 0.25), Inches(0.8), Inches(1.2), Inches(1.2))
    quote_frame = quote_mark.text_frame
    quote_frame.text = "\u201C"  # 左双引号装饰
    quote_frame.paragraphs[0].font.size = Pt(80)
    quote_frame.paragraphs[0].font.color.rgb = ColorTheme.PRIMARY
    quote_frame.paragraphs[0].font.name = "Georgia"
    
    # 引用内容 - 根据长度调整字体
    quote_text = slide_data.text or slide_data.title
    text_len = len(quote_text)
    if text_len > 100:
        font_size = 17
    elif text_len > 70:
        font_size = 19
    else:
        font_size = 22
    
    quote_box = slide.shapes.add_textbox(Inches(start_x + 0.4), Inches(2.0), Inches(content_width - 0.8), Inches(3.2))
    quote_text_frame = quote_box.text_frame
    quote_text_frame.word_wrap = True
    quote_text_frame.text = quote_text
    quote_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    _set_font(quote_text_frame, font_size=font_size, color=ColorTheme.TEXT_DARK)
    
    # 作者/来源
    if slide_data.subtitle:
        author_box = slide.shapes.add_textbox(Inches(start_x + 0.4), Inches(5.6), Inches(content_width - 0.8), Inches(0.5))
        author_frame = author_box.text_frame
        author_frame.text = f"— {slide_data.subtitle}"
        author_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        _set_font(author_frame, font_size=15, color=ColorTheme.PRIMARY)


def _create_ending_slide(prs: Presentation, slide_data: Slide) -> None:
    """创建结束页"""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    
    try:
        slide.shapes.title.text = slide_data.title
        _set_font(slide.shapes.title.text_frame, font_size=44, bold=True)
    except:
        pass
    
    try:
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = slide_data.subtitle or ""
            _set_font(slide.placeholders[1].text_frame, font_size=24)
    except:
        pass
