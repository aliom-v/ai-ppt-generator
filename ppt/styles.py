"""PPT 样式配置模块 - 颜色主题、字体配置、自定义配色方案"""
import platform
import json
import os
import threading
from typing import Tuple, Dict, Optional, Any, List
from dataclasses import dataclass, asdict, field
from pptx.dml.color import RGBColor
from pathlib import Path

from utils.logger import get_logger

logger = get_logger("styles")


def get_default_fonts() -> Tuple[str, str]:
    """获取跨平台默认字体

    Returns:
        (中文字体, 英文字体) 元组
    """
    system = platform.system()
    if system == "Windows":
        return ("微软雅黑", "Microsoft YaHei")
    elif system == "Darwin":  # macOS
        return ("PingFang SC", "Helvetica Neue")
    else:  # Linux
        return ("Noto Sans CJK SC", "DejaVu Sans")


# 默认字体
FONT_CN, FONT_EN = get_default_fonts()


def hex_to_rgb(hex_color: str) -> RGBColor:
    """将十六进制颜色转换为 RGBColor

    Args:
        hex_color: 十六进制颜色字符串，如 "#1976D2" 或 "1976D2"

    Returns:
        RGBColor 对象
    """
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


def rgb_to_hex(color: RGBColor) -> str:
    """将 RGBColor 转换为十六进制字符串"""
    return f"#{color[0]:02X}{color[1]:02X}{color[2]:02X}"


@dataclass
class ThemeConfig:
    """主题配置数据类"""
    name: str
    display_name: str
    primary: str        # 主色调
    secondary: str      # 次要色
    accent: str         # 强调色
    background: str     # 背景色
    text_primary: str   # 主要文字色
    text_secondary: str # 次要文字色
    success: str        # 成功色
    warning: str        # 警告色
    font_title: str = ""   # 标题字体（空则使用默认）
    font_body: str = ""    # 正文字体（空则使用默认）
    description: str = ""

    def to_dict(self) -> Dict[str, str]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ThemeConfig":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


# ==================== 预设主题 ====================

PRESET_THEMES: Dict[str, ThemeConfig] = {
    "professional": ThemeConfig(
        name="professional",
        display_name="专业蓝",
        primary="#1976D2",
        secondary="#42A5F5",
        accent="#FF9800",
        background="#FFFFFF",
        text_primary="#333333",
        text_secondary="#7F7F7F",
        success="#4CAF50",
        warning="#FFC107",
        description="经典商务蓝色主题，适合正式场合"
    ),
    "creative": ThemeConfig(
        name="creative",
        display_name="创意彩虹",
        primary="#E91E63",
        secondary="#9C27B0",
        accent="#00BCD4",
        background="#FAFAFA",
        text_primary="#212121",
        text_secondary="#757575",
        success="#8BC34A",
        warning="#FF9800",
        description="活力多彩主题，适合创意展示"
    ),
    "minimal": ThemeConfig(
        name="minimal",
        display_name="极简黑白",
        primary="#212121",
        secondary="#424242",
        accent="#757575",
        background="#FFFFFF",
        text_primary="#212121",
        text_secondary="#9E9E9E",
        success="#4CAF50",
        warning="#FF9800",
        description="简洁黑白主题，适合学术和正式文档"
    ),
    "tech": ThemeConfig(
        name="tech",
        display_name="科技暗黑",
        primary="#00E5FF",
        secondary="#18FFFF",
        accent="#FF4081",
        background="#1A1A2E",
        text_primary="#FFFFFF",
        text_secondary="#B0B0B0",
        success="#00E676",
        warning="#FFEA00",
        description="深色科技主题，适合技术演示"
    ),
    "nature": ThemeConfig(
        name="nature",
        display_name="自然绿意",
        primary="#388E3C",
        secondary="#66BB6A",
        accent="#FF7043",
        background="#F1F8E9",
        text_primary="#33691E",
        text_secondary="#689F38",
        success="#4CAF50",
        warning="#FFC107",
        description="自然绿色主题，清新自然"
    ),
    "warm": ThemeConfig(
        name="warm",
        display_name="温暖橙调",
        primary="#FF5722",
        secondary="#FF8A65",
        accent="#FFC107",
        background="#FFF3E0",
        text_primary="#5D4037",
        text_secondary="#8D6E63",
        success="#4CAF50",
        warning="#FF9800",
        description="温暖橙色主题，热情活力"
    ),
}


class ColorTheme:
    """颜色主题配置

    提供统一的配色方案，确保 PPT 风格一致性。
    支持从预设主题或自定义配置加载颜色。
    """

    def __init__(self, theme_config: Optional[ThemeConfig] = None):
        """初始化颜色主题

        Args:
            theme_config: 主题配置，为 None 时使用默认 professional 主题
        """
        if theme_config is None:
            theme_config = PRESET_THEMES["professional"]

        self.config = theme_config
        self._load_colors()

    def _load_colors(self):
        """从配置加载颜色"""
        self.PRIMARY = hex_to_rgb(self.config.primary)
        self.SECONDARY = hex_to_rgb(self.config.secondary)
        self.ACCENT = hex_to_rgb(self.config.accent)
        self.BACKGROUND = hex_to_rgb(self.config.background)
        self.TEXT_DARK = hex_to_rgb(self.config.text_primary)
        self.TEXT_LIGHT = hex_to_rgb(self.config.text_secondary)
        self.SUCCESS = hex_to_rgb(self.config.success)
        self.WARNING = hex_to_rgb(self.config.warning)

        # 兼容旧代码
        self.WHITE = RGBColor(255, 255, 255)
        self.BG_LIGHT = RGBColor(240, 240, 240)

        # 生成派生颜色
        self._generate_derived_colors()

    def _generate_derived_colors(self):
        """生成派生的颜色组合"""
        # 卡片背景色（基于主题色的浅色版本）
        self.CARD_BG_COLORS = [
            self._lighten_color(self.PRIMARY, 0.85),
            self._lighten_color(self.SUCCESS, 0.85),
            self._lighten_color(self.ACCENT, 0.85),
            self._lighten_color(self.SECONDARY, 0.85),
            self._lighten_color(hex_to_rgb("#9C27B0"), 0.85),  # 紫色
            self._lighten_color(hex_to_rgb("#00BCD4"), 0.85),  # 青色
        ]

        # 强调色组
        self.ACCENT_COLORS = [
            self.PRIMARY,
            self.SUCCESS,
            self.ACCENT,
            hex_to_rgb("#9C27B0"),  # 紫色
            hex_to_rgb("#F44336"),  # 红色
            hex_to_rgb("#00BCD4"),  # 青色
        ]

        # 渐变色
        self.GRADIENT_COLORS = [
            self.PRIMARY,
            self.SUCCESS,
            self.ACCENT,
            self.SECONDARY,
            hex_to_rgb("#F44336"),  # 红色
        ]

    def _lighten_color(self, color: RGBColor, factor: float) -> RGBColor:
        """将颜色变浅

        Args:
            color: 原始颜色
            factor: 变浅因子（0-1），越大越接近白色

        Returns:
            变浅后的颜色
        """
        r = int(color[0] + (255 - color[0]) * factor)
        g = int(color[1] + (255 - color[1]) * factor)
        b = int(color[2] + (255 - color[2]) * factor)
        return RGBColor(r, g, b)

    @property
    def font_title(self) -> str:
        """标题字体"""
        return self.config.font_title or FONT_CN

    @property
    def font_body(self) -> str:
        """正文字体"""
        return self.config.font_body or FONT_CN


class ThemeManager:
    """主题管理器 - 管理预设和自定义主题"""

    def __init__(self, custom_themes_dir: str = "data/themes"):
        self.custom_themes_dir = Path(custom_themes_dir)
        self.custom_themes_dir.mkdir(parents=True, exist_ok=True)
        self._custom_themes: Dict[str, ThemeConfig] = {}
        self._lock = threading.Lock()
        self._load_custom_themes()

    def _load_custom_themes(self):
        """加载自定义主题"""
        themes_file = self.custom_themes_dir / "custom_themes.json"
        if themes_file.exists():
            try:
                with open(themes_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for name, config in data.items():
                    self._custom_themes[name] = ThemeConfig.from_dict(config)
                logger.info(f"加载了 {len(self._custom_themes)} 个自定义主题")
            except Exception as e:
                logger.warning(f"加载自定义主题失败: {e}")

    def _save_custom_themes(self):
        """保存自定义主题"""
        themes_file = self.custom_themes_dir / "custom_themes.json"
        try:
            data = {name: config.to_dict() for name, config in self._custom_themes.items()}
            with open(themes_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存自定义主题失败: {e}")

    def get_theme(self, name: str) -> Optional[ThemeConfig]:
        """获取主题配置

        Args:
            name: 主题名称（预设或自定义）

        Returns:
            主题配置，如果不存在返回 None
        """
        # 先查预设
        if name in PRESET_THEMES:
            return PRESET_THEMES[name]
        # 再查自定义
        with self._lock:
            return self._custom_themes.get(name)

    def get_color_theme(self, name: str) -> ColorTheme:
        """获取 ColorTheme 实例

        Args:
            name: 主题名称

        Returns:
            ColorTheme 实例，如果主题不存在则返回默认主题
        """
        config = self.get_theme(name)
        return ColorTheme(config)

    def list_themes(self) -> List[Dict[str, Any]]:
        """列出所有可用主题

        Returns:
            主题列表，每个元素包含 name, display_name, description, is_preset
        """
        themes = []
        # 预设主题
        for name, config in PRESET_THEMES.items():
            themes.append({
                "name": name,
                "display_name": config.display_name,
                "description": config.description,
                "is_preset": True,
                "primary_color": config.primary
            })
        # 自定义主题
        with self._lock:
            for name, config in self._custom_themes.items():
                themes.append({
                    "name": name,
                    "display_name": config.display_name,
                    "description": config.description,
                    "is_preset": False,
                    "primary_color": config.primary
                })
        return themes

    def create_custom_theme(self, config: ThemeConfig) -> bool:
        """创建自定义主题

        Args:
            config: 主题配置

        Returns:
            是否创建成功
        """
        if config.name in PRESET_THEMES:
            logger.warning(f"不能覆盖预设主题: {config.name}")
            return False

        with self._lock:
            self._custom_themes[config.name] = config
            self._save_custom_themes()
            logger.info(f"创建自定义主题: {config.name}")
        return True

    def delete_custom_theme(self, name: str) -> bool:
        """删除自定义主题

        Args:
            name: 主题名称

        Returns:
            是否删除成功
        """
        if name in PRESET_THEMES:
            logger.warning(f"不能删除预设主题: {name}")
            return False

        with self._lock:
            if name in self._custom_themes:
                del self._custom_themes[name]
                self._save_custom_themes()
                logger.info(f"删除自定义主题: {name}")
                return True
        return False


# ==================== 全局实例 ====================

_theme_manager: Optional[ThemeManager] = None
_theme_manager_lock = threading.Lock()


def get_theme_manager() -> ThemeManager:
    """获取全局主题管理器实例"""
    global _theme_manager
    if _theme_manager is None:
        with _theme_manager_lock:
            if _theme_manager is None:
                _theme_manager = ThemeManager()
    return _theme_manager


# ==================== 兼容旧代码的默认颜色 ====================

# 默认主题实例（兼容旧代码）
_default_theme = ColorTheme()

# 导出默认颜色（兼容旧代码）
CARD_BG_COLORS = _default_theme.CARD_BG_COLORS
ACCENT_COLORS = _default_theme.ACCENT_COLORS
GRADIENT_COLORS = _default_theme.GRADIENT_COLORS


# 兼容旧的 ColorTheme 类属性访问
class _LegacyColorTheme:
    """兼容旧代码的静态 ColorTheme"""
    PRIMARY = _default_theme.PRIMARY
    SECONDARY = _default_theme.SECONDARY
    ACCENT = _default_theme.ACCENT
    TEXT_DARK = _default_theme.TEXT_DARK
    TEXT_LIGHT = _default_theme.TEXT_LIGHT
    BG_LIGHT = _default_theme.BG_LIGHT
    WHITE = _default_theme.WHITE
    SUCCESS = _default_theme.SUCCESS
    WARNING = _default_theme.WARNING


# 替换原有的 ColorTheme 以保持向后兼容
# 注意：新代码应该使用 get_theme_manager().get_color_theme(name)
ColorTheme = _LegacyColorTheme
