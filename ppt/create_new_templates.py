"""创建新模板 - 暗色主题、极简风格、中国风"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os


def create_dark_theme_template():
    """创建暗色主题模板"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 暗色配色方案
    bg_color = RGBColor(18, 18, 18)  # 深黑背景
    primary = RGBColor(96, 165, 250)  # 蓝色
    accent = RGBColor(244, 114, 182)  # 粉色
    text_primary = RGBColor(248, 250, 252)  # 白色文字
    text_secondary = RGBColor(148, 163, 184)  # 灰色文字

    # 创建封面布局
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局

    # 深色背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background()

    # 渐变装饰
    accent_shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(-2), Inches(-2), Inches(8), Inches(8)
    )
    accent_shape.fill.solid()
    accent_shape.fill.fore_color.rgb = RGBColor(59, 130, 246)
    accent_shape.fill.fore_color.brightness = -0.3
    accent_shape.line.fill.background()

    prs.save('ppt/pptx_templates/dark_theme.pptx')
    print("✓ 暗色主题模板已创建: dark_theme.pptx")


def create_minimalist_template():
    """创建极简风格模板"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 极简配色
    bg_color = RGBColor(255, 255, 255)  # 纯白背景
    primary = RGBColor(17, 17, 17)  # 纯黑
    accent = RGBColor(229, 231, 235)  # 浅灰线条

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 白色背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background()

    # 简洁底部线条
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), prs.slide_height - Inches(0.5),
        prs.slide_width - Inches(1), Pt(2)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = primary
    line.line.fill.background()

    prs.save('ppt/pptx_templates/minimalist.pptx')
    print("✓ 极简风格模板已创建: minimalist.pptx")


def create_chinese_style_template():
    """创建中国风模板"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 中国风配色
    bg_color = RGBColor(254, 252, 248)  # 米白色
    primary = RGBColor(180, 50, 50)  # 中国红
    gold = RGBColor(184, 134, 11)  # 金色
    ink = RGBColor(45, 45, 45)  # 墨色

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 米白背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background()

    # 红色边框装饰
    top_border = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.3), Inches(0.3),
        prs.slide_width - Inches(0.6), Pt(4)
    )
    top_border.fill.solid()
    top_border.fill.fore_color.rgb = primary
    top_border.line.fill.background()

    bottom_border = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.3), prs.slide_height - Inches(0.4),
        prs.slide_width - Inches(0.6), Pt(4)
    )
    bottom_border.fill.solid()
    bottom_border.fill.fore_color.rgb = primary
    bottom_border.line.fill.background()

    # 角落装饰
    corner_size = Inches(0.8)
    corners = [
        (Inches(0.3), Inches(0.3)),  # 左上
        (prs.slide_width - Inches(1.1), Inches(0.3)),  # 右上
        (Inches(0.3), prs.slide_height - Inches(1.1)),  # 左下
        (prs.slide_width - Inches(1.1), prs.slide_height - Inches(1.1)),  # 右下
    ]

    for x, y in corners:
        corner = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, x, y, corner_size, corner_size
        )
        corner.fill.solid()
        corner.fill.fore_color.rgb = gold
        corner.line.fill.background()

    prs.save('ppt/pptx_templates/chinese_style.pptx')
    print("✓ 中国风模板已创建: chinese_style.pptx")


def create_gradient_blue_template():
    """创建渐变蓝模板"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 渐变背景效果（使用多个矩形模拟）
    colors = [
        RGBColor(30, 64, 175),   # 深蓝
        RGBColor(59, 130, 246),  # 中蓝
        RGBColor(96, 165, 250),  # 浅蓝
    ]

    height_per_section = prs.slide_height / len(colors)
    for i, color in enumerate(colors):
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, height_per_section * i,
            prs.slide_width, height_per_section + Pt(2)
        )
        rect.fill.solid()
        rect.fill.fore_color.rgb = color
        rect.line.fill.background()

    prs.save('ppt/pptx_templates/gradient_blue.pptx')
    print("✓ 渐变蓝模板已创建: gradient_blue.pptx")


def create_tech_modern_template():
    """创建科技现代模板"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 深色科技背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(15, 23, 42)
    bg.line.fill.background()

    # 网格线装饰（水平线）
    for i in range(1, 8):
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, Inches(i),
            prs.slide_width, Pt(1)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(51, 65, 85)
        line.line.fill.background()

    # 左侧发光条
    glow = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        Pt(4), prs.slide_height
    )
    glow.fill.solid()
    glow.fill.fore_color.rgb = RGBColor(34, 211, 238)  # 青色
    glow.line.fill.background()

    prs.save('ppt/pptx_templates/tech_modern.pptx')
    print("✓ 科技现代模板已创建: tech_modern.pptx")


def create_warm_sunset_template():
    """创建暖色夕阳模板"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 暖色渐变
    colors = [
        RGBColor(251, 146, 60),   # 橙色
        RGBColor(249, 115, 22),   # 深橙
        RGBColor(234, 88, 12),    # 红橙
    ]

    height_per_section = prs.slide_height / len(colors)
    for i, color in enumerate(colors):
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, height_per_section * i,
            prs.slide_width, height_per_section + Pt(2)
        )
        rect.fill.solid()
        rect.fill.fore_color.rgb = color
        rect.line.fill.background()

    prs.save('ppt/pptx_templates/warm_sunset.pptx')
    print("✓ 暖色夕阳模板已创建: warm_sunset.pptx")


def create_all_templates():
    """创建所有新模板"""
    os.makedirs('ppt/pptx_templates', exist_ok=True)

    print("\n创建新模板...")
    print("=" * 40)

    create_dark_theme_template()
    create_minimalist_template()
    create_chinese_style_template()
    create_gradient_blue_template()
    create_tech_modern_template()
    create_warm_sunset_template()

    print("=" * 40)
    print("所有新模板创建完成！\n")


if __name__ == "__main__":
    create_all_templates()
