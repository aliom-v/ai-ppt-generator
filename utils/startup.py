"""启动配置验证"""
import os
import sys
from typing import List, Tuple
from pathlib import Path

from utils.logger import get_logger

logger = get_logger("startup")


class ConfigValidator:
    """配置验证器

    在应用启动时验证关键配置，提前发现问题。
    """

    def __init__(self):
        self.errors: List[str] = []
        self.warnings: List[str] = []

    def validate_all(self) -> Tuple[bool, List[str], List[str]]:
        """执行所有验证

        Returns:
            (is_valid, errors, warnings)
        """
        self._validate_directories()
        self._validate_dependencies()
        self._check_optional_configs()

        return len(self.errors) == 0, self.errors, self.warnings

    def _validate_directories(self):
        """验证必需的目录"""
        required_dirs = [
            "web/uploads",
            "web/outputs",
            "images/downloaded",
            "cache",
        ]

        for dir_path in required_dirs:
            path = Path(dir_path)
            try:
                path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.errors.append(f"无法创建目录 {dir_path}: {e}")

    def _validate_dependencies(self):
        """验证必需的依赖"""
        required_packages = [
            ("openai", "openai"),
            ("pptx", "python-pptx"),
            ("flask", "flask"),
            ("PIL", "Pillow"),
        ]

        for import_name, package_name in required_packages:
            try:
                __import__(import_name)
            except ImportError:
                self.errors.append(f"缺少依赖: {package_name}，请运行 pip install {package_name}")

    def _check_optional_configs(self):
        """检查可选配置"""
        # 检查 API Key
        if not os.getenv("AI_API_KEY"):
            self.warnings.append("AI_API_KEY 未设置，需要在使用时提供")

        # 检查图片搜索
        if not os.getenv("UNSPLASH_ACCESS_KEY"):
            self.warnings.append("UNSPLASH_ACCESS_KEY 未设置，图片自动搜索功能将不可用")

        # 检查 SECRET_KEY
        if not os.getenv("SECRET_KEY"):
            self.warnings.append("SECRET_KEY 未设置，将使用随机生成的密钥（生产环境应设置固定密钥）")

        # 检查模板目录
        template_dir = Path("ppt/pptx_templates")
        if not template_dir.exists() or not list(template_dir.glob("*.pptx")):
            self.warnings.append("未找到 PPT 模板文件，将使用默认样式")


def validate_startup_config() -> bool:
    """验证启动配置

    Returns:
        验证是否通过
    """
    validator = ConfigValidator()
    is_valid, errors, warnings = validator.validate_all()

    # 输出警告
    for warning in warnings:
        logger.warning(f"⚠️  {warning}")

    # 输出错误
    for error in errors:
        logger.error(f"❌ {error}")

    if is_valid:
        logger.info("✅ 配置验证通过")
    else:
        logger.error("❌ 配置验证失败，请修复上述错误")

    return is_valid


def print_startup_banner():
    """打印启动横幅"""
    banner = """
╔═══════════════════════════════════════════════════════╗
║             AI PPT Generator v1.0.0                   ║
║         智能 PPT 自动生成工具                           ║
╚═══════════════════════════════════════════════════════╝
    """
    print(banner)


def initialize_app():
    """初始化应用

    执行启动检查和初始化任务。

    Returns:
        是否初始化成功
    """
    print_startup_banner()

    # 验证配置
    if not validate_startup_config():
        return False

    # 设置清理任务
    try:
        from utils.scheduler import setup_cleanup_tasks
        setup_cleanup_tasks()
        logger.info("✅ 后台清理任务已启动")
    except Exception as e:
        logger.warning(f"⚠️  启动后台任务失败: {e}")

    return True
