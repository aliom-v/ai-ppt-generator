"""图片增强处理模块 - 支持质量验证、增强处理和视觉效果"""
import os
from typing import Optional, Tuple, Dict, Any
from dataclasses import dataclass
from pathlib import Path

from utils.logger import get_logger

logger = get_logger("image_enhancer")

# 尝试导入 Pillow
try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageDraw, ImageOps
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    logger.warning("Pillow 未安装，图片增强功能不可用")


@dataclass
class ImageQualityReport:
    """图片质量报告"""
    path: str
    is_valid: bool = False
    width: int = 0
    height: int = 0
    format: str = ""
    file_size: int = 0
    aspect_ratio: float = 0.0
    is_too_small: bool = False
    is_too_large: bool = False
    issues: list = None

    def __post_init__(self):
        if self.issues is None:
            self.issues = []


@dataclass
class EnhanceConfig:
    """图片增强配置"""
    # 基础增强
    brightness: float = 1.0       # 亮度 (0.5-1.5, 1.0=不变)
    contrast: float = 1.05        # 对比度 (0.5-1.5, 1.0=不变)
    saturation: float = 1.1       # 饱和度 (0.0-2.0, 1.0=不变)
    sharpness: float = 1.1        # 锐度 (0.0-2.0, 1.0=不变)

    # 边框和圆角
    add_border: bool = False      # 是否添加边框
    border_color: Tuple[int, int, int] = (255, 255, 255)  # 边框颜色
    border_width: int = 2         # 边框宽度
    corner_radius: int = 0        # 圆角半径 (0=无圆角)

    # 阴影效果
    add_shadow: bool = False      # 是否添加阴影
    shadow_offset: Tuple[int, int] = (5, 5)  # 阴影偏移
    shadow_blur: int = 10         # 阴影模糊

    # 自动增强
    auto_enhance: bool = True     # 是否自动增强
    auto_contrast: bool = True    # 自动对比度
    auto_color: bool = False      # 自动颜色平衡


# 图片质量标准
MIN_WIDTH = 400
MIN_HEIGHT = 300
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB
PREFERRED_ASPECT_RATIOS = [(16, 9), (4, 3), (3, 2), (1, 1)]


def validate_image(image_path: str) -> ImageQualityReport:
    """验证图片质量

    Args:
        image_path: 图片文件路径

    Returns:
        ImageQualityReport 质量报告
    """
    report = ImageQualityReport(path=image_path)

    if not os.path.exists(image_path):
        report.issues.append("文件不存在")
        return report

    if not PIL_AVAILABLE:
        # 没有 Pillow 时，只做基础检查
        report.file_size = os.path.getsize(image_path)
        report.is_valid = report.file_size > 0
        if report.file_size > MAX_FILE_SIZE:
            report.is_too_large = True
            report.issues.append(f"文件过大: {report.file_size / 1024 / 1024:.1f}MB")
        return report

    try:
        with Image.open(image_path) as img:
            report.width, report.height = img.size
            report.format = img.format or "UNKNOWN"
            report.file_size = os.path.getsize(image_path)
            report.aspect_ratio = report.width / report.height if report.height > 0 else 0

            # 检查尺寸
            if report.width < MIN_WIDTH or report.height < MIN_HEIGHT:
                report.is_too_small = True
                report.issues.append(f"图片过小: {report.width}x{report.height}")

            # 检查文件大小
            if report.file_size > MAX_FILE_SIZE:
                report.is_too_large = True
                report.issues.append(f"文件过大: {report.file_size / 1024 / 1024:.1f}MB")

            # 检查格式
            supported_formats = ["JPEG", "PNG", "GIF", "BMP", "WEBP"]
            if report.format not in supported_formats:
                report.issues.append(f"不推荐的格式: {report.format}")

            # 验证图片完整性
            img.verify()

            report.is_valid = len(report.issues) == 0 or not report.is_too_small

    except Exception as e:
        report.issues.append(f"图片损坏或无法读取: {e}")
        report.is_valid = False

    return report


def enhance_image(
    image_path: str,
    output_path: Optional[str] = None,
    config: Optional[EnhanceConfig] = None
) -> Optional[str]:
    """增强图片

    Args:
        image_path: 原图路径
        output_path: 输出路径（None 则覆盖原图）
        config: 增强配置

    Returns:
        处理后的图片路径，失败返回 None
    """
    if not PIL_AVAILABLE:
        logger.warning("Pillow 未安装，跳过图片增强")
        return image_path

    if not os.path.exists(image_path):
        logger.error(f"图片不存在: {image_path}")
        return None

    if config is None:
        config = EnhanceConfig()

    if output_path is None:
        # 生成增强后的文件名
        path = Path(image_path)
        output_path = str(path.parent / f"{path.stem}_enhanced{path.suffix}")

    try:
        with Image.open(image_path) as img:
            # 转换为 RGB（如果需要）
            if img.mode in ('RGBA', 'P'):
                # 保留透明通道
                if img.mode == 'P':
                    img = img.convert('RGBA')
                # 创建白色背景
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'RGBA':
                    background.paste(img, mask=img.split()[3])
                    img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            # 自动增强
            if config.auto_enhance:
                if config.auto_contrast:
                    img = ImageOps.autocontrast(img, cutoff=0.5)

            # 基础增强
            if config.brightness != 1.0:
                enhancer = ImageEnhance.Brightness(img)
                img = enhancer.enhance(config.brightness)

            if config.contrast != 1.0:
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(config.contrast)

            if config.saturation != 1.0:
                enhancer = ImageEnhance.Color(img)
                img = enhancer.enhance(config.saturation)

            if config.sharpness != 1.0:
                enhancer = ImageEnhance.Sharpness(img)
                img = enhancer.enhance(config.sharpness)

            # 添加圆角
            if config.corner_radius > 0:
                img = _add_rounded_corners(img, config.corner_radius)

            # 添加边框
            if config.add_border and config.border_width > 0:
                img = _add_border(img, config.border_width, config.border_color)

            # 添加阴影
            if config.add_shadow:
                img = _add_shadow(img, config.shadow_offset, config.shadow_blur)

            # 保存
            img.save(output_path, quality=95, optimize=True)
            logger.debug(f"图片增强完成: {output_path}")
            return output_path

    except Exception as e:
        logger.error(f"图片增强失败: {e}")
        return None


def _add_rounded_corners(img: "Image.Image", radius: int) -> "Image.Image":
    """添加圆角

    Args:
        img: 原始图片
        radius: 圆角半径

    Returns:
        处理后的图片
    """
    # 创建圆角蒙版
    mask = Image.new('L', img.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([(0, 0), img.size], radius=radius, fill=255)

    # 应用蒙版
    output = Image.new('RGBA', img.size, (255, 255, 255, 0))
    output.paste(img, (0, 0))
    output.putalpha(mask)

    return output


def _add_border(
    img: "Image.Image",
    width: int,
    color: Tuple[int, int, int]
) -> "Image.Image":
    """添加边框

    Args:
        img: 原始图片
        width: 边框宽度
        color: 边框颜色

    Returns:
        处理后的图片
    """
    # 创建带边框的新图片
    new_size = (img.size[0] + 2 * width, img.size[1] + 2 * width)
    bordered = Image.new('RGB', new_size, color)
    bordered.paste(img, (width, width))
    return bordered


def _add_shadow(
    img: "Image.Image",
    offset: Tuple[int, int],
    blur_radius: int
) -> "Image.Image":
    """添加阴影效果

    Args:
        img: 原始图片
        offset: 阴影偏移 (x, y)
        blur_radius: 模糊半径

    Returns:
        处理后的图片
    """
    # 创建更大的画布
    padding = blur_radius * 2 + max(abs(offset[0]), abs(offset[1]))
    new_size = (img.size[0] + padding * 2, img.size[1] + padding * 2)

    # 创建阴影层
    shadow = Image.new('RGBA', new_size, (255, 255, 255, 0))

    # 创建阴影形状
    shadow_shape = Image.new('RGBA', img.size, (0, 0, 0, 100))
    shadow_pos = (padding + offset[0], padding + offset[1])
    shadow.paste(shadow_shape, shadow_pos)

    # 模糊阴影
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur_radius))

    # 合并图片
    img_pos = (padding, padding)
    if img.mode == 'RGBA':
        shadow.paste(img, img_pos, img)
    else:
        shadow.paste(img, img_pos)

    return shadow


def resize_for_ppt(
    image_path: str,
    max_width: int = 1920,
    max_height: int = 1080,
    output_path: Optional[str] = None
) -> Optional[str]:
    """调整图片大小以适合 PPT

    Args:
        image_path: 原图路径
        max_width: 最大宽度
        max_height: 最大高度
        output_path: 输出路径

    Returns:
        处理后的图片路径
    """
    if not PIL_AVAILABLE:
        return image_path

    if not os.path.exists(image_path):
        return None

    if output_path is None:
        path = Path(image_path)
        output_path = str(path.parent / f"{path.stem}_resized{path.suffix}")

    try:
        with Image.open(image_path) as img:
            # 计算缩放比例
            ratio = min(max_width / img.width, max_height / img.height)

            if ratio >= 1:
                # 图片已经够小，无需缩放
                return image_path

            new_size = (int(img.width * ratio), int(img.height * ratio))
            resized = img.resize(new_size, Image.Resampling.LANCZOS)
            resized.save(output_path, quality=95, optimize=True)
            logger.debug(f"图片已调整大小: {img.size} -> {new_size}")
            return output_path

    except Exception as e:
        logger.error(f"调整图片大小失败: {e}")
        return None


def optimize_for_web(
    image_path: str,
    output_path: Optional[str] = None,
    max_size_kb: int = 500
) -> Optional[str]:
    """优化图片以减小文件大小

    Args:
        image_path: 原图路径
        output_path: 输出路径
        max_size_kb: 目标最大文件大小（KB）

    Returns:
        处理后的图片路径
    """
    if not PIL_AVAILABLE:
        return image_path

    if not os.path.exists(image_path):
        return None

    if output_path is None:
        path = Path(image_path)
        output_path = str(path.parent / f"{path.stem}_optimized.jpg")

    try:
        with Image.open(image_path) as img:
            # 转换为 RGB
            if img.mode != 'RGB':
                img = img.convert('RGB')

            # 逐步降低质量直到达到目标大小
            quality = 95
            while quality > 20:
                img.save(output_path, 'JPEG', quality=quality, optimize=True)
                file_size = os.path.getsize(output_path) / 1024
                if file_size <= max_size_kb:
                    break
                quality -= 5

            logger.debug(f"图片已优化: {os.path.getsize(image_path) / 1024:.1f}KB -> {file_size:.1f}KB")
            return output_path

    except Exception as e:
        logger.error(f"图片优化失败: {e}")
        return None


def create_placeholder_image(
    width: int = 800,
    height: int = 600,
    text: str = "图片占位符",
    output_path: str = "placeholder.png",
    bg_color: Tuple[int, int, int] = (240, 240, 240),
    text_color: Tuple[int, int, int] = (128, 128, 128)
) -> Optional[str]:
    """创建占位图片

    Args:
        width: 宽度
        height: 高度
        text: 占位文字
        output_path: 输出路径
        bg_color: 背景颜色
        text_color: 文字颜色

    Returns:
        图片路径
    """
    if not PIL_AVAILABLE:
        logger.warning("Pillow 未安装，无法创建占位图片")
        return None

    try:
        img = Image.new('RGB', (width, height), bg_color)
        draw = ImageDraw.Draw(img)

        # 绘制边框
        draw.rectangle(
            [(10, 10), (width - 10, height - 10)],
            outline=(200, 200, 200),
            width=2
        )

        # 绘制对角线
        draw.line([(10, 10), (width - 10, height - 10)], fill=(200, 200, 200), width=1)
        draw.line([(width - 10, 10), (10, height - 10)], fill=(200, 200, 200), width=1)

        # 绘制文字（使用默认字体）
        text_bbox = draw.textbbox((0, 0), text)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        text_x = (width - text_width) // 2
        text_y = (height - text_height) // 2
        draw.text((text_x, text_y), text, fill=text_color)

        # 保存
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        img.save(output_path)
        return output_path

    except Exception as e:
        logger.error(f"创建占位图片失败: {e}")
        return None


def batch_enhance_images(
    image_paths: list,
    config: Optional[EnhanceConfig] = None,
    output_dir: Optional[str] = None
) -> Dict[str, Optional[str]]:
    """批量增强图片

    Args:
        image_paths: 图片路径列表
        config: 增强配置
        output_dir: 输出目录

    Returns:
        {原路径: 增强后路径} 的映射
    """
    results = {}

    if output_dir:
        Path(output_dir).mkdir(parents=True, exist_ok=True)

    for path in image_paths:
        if output_dir:
            filename = Path(path).name
            output_path = str(Path(output_dir) / filename)
        else:
            output_path = None

        results[path] = enhance_image(path, output_path, config)

    return results


# 预设增强配置
PRESET_CONFIGS = {
    "default": EnhanceConfig(),
    "vivid": EnhanceConfig(
        contrast=1.15,
        saturation=1.25,
        sharpness=1.2,
    ),
    "soft": EnhanceConfig(
        contrast=0.95,
        saturation=0.9,
        brightness=1.05,
    ),
    "sharp": EnhanceConfig(
        sharpness=1.5,
        contrast=1.1,
    ),
    "rounded": EnhanceConfig(
        corner_radius=20,
        add_border=True,
        border_width=2,
        border_color=(255, 255, 255),
    ),
    "shadow": EnhanceConfig(
        add_shadow=True,
        shadow_offset=(8, 8),
        shadow_blur=15,
    ),
}


def get_preset_config(preset_name: str) -> EnhanceConfig:
    """获取预设配置

    Args:
        preset_name: 预设名称

    Returns:
        EnhanceConfig 配置对象
    """
    return PRESET_CONFIGS.get(preset_name, PRESET_CONFIGS["default"])
