"""配置模块：管理 API 配置和项目设置"""
from dataclasses import dataclass, field
from typing import Optional
import os


# ==================== 常量定义 ====================

# API 默认值
DEFAULT_API_BASE_URL = "https://api.openai.com/v1"
DEFAULT_MODEL_NAME = "gpt-4o-mini"
DEFAULT_TEMPERATURE = 0.7
DEFAULT_MAX_RETRIES = 3
DEFAULT_TIMEOUT = 180

# 图片配置
DEFAULT_MAX_CONCURRENT_DOWNLOADS = 8
DEFAULT_DOWNLOAD_TIMEOUT = 30

# 应用配置
DEFAULT_MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10MB
DEFAULT_PORT = 5000

# 限制
MAX_TOPIC_LENGTH = 500
MAX_AUDIENCE_LENGTH = 200
MAX_DESCRIPTION_LENGTH = 100000
MAX_MODEL_NAME_LENGTH = 100
MAX_PAGE_COUNT = 100
MAX_BATCH_ITEMS = 20
MAX_FILENAME_LENGTH = 50

# 速率限制
RATE_LIMIT_PER_MINUTE = 5
RATE_LIMIT_PER_HOUR = 50
BATCH_RATE_LIMIT_PER_MINUTE = 2
BATCH_RATE_LIMIT_PER_HOUR = 20

# 缓存时间（秒）
CACHE_CONFIG_TTL = 300  # 5分钟
CACHE_TEMPLATES_TTL = 86400  # 24小时
CACHE_STATS_TTL = 3600  # 1小时

# 文件清理
FILE_CLEANUP_MAX_AGE_HOURS = 24
FILE_CLEANUP_MAX_FILES = 100


@dataclass
class AIConfig:
    """AI 配置类 - 支持实例化传递，解决多用户并发问题"""
    api_key: str
    api_base_url: str = DEFAULT_API_BASE_URL
    model_name: str = DEFAULT_MODEL_NAME
    temperature: float = DEFAULT_TEMPERATURE
    max_retries: int = DEFAULT_MAX_RETRIES
    timeout: int = DEFAULT_TIMEOUT

    def __post_init__(self):
        """初始化后处理 - 规范化 API Base URL"""
        self.api_base_url = self._normalize_base_url(self.api_base_url)

    @staticmethod
    def _normalize_base_url(url: str) -> str:
        """规范化 API Base URL - 只做基本清理，不自动添加路径"""
        if not url:
            return DEFAULT_API_BASE_URL

        url = url.strip()

        # 只移除末尾斜杠
        url = url.rstrip('/')

        return url

    def validate(self) -> bool:
        """验证配置是否有效"""
        if not self.api_key:
            raise ValueError("API_KEY 未设置")
        if not self.api_base_url:
            raise ValueError("API_BASE_URL 未设置")
        return True

    @classmethod
    def from_env(cls) -> "AIConfig":
        """从环境变量创建配置"""
        return cls(
            api_key=os.getenv("AI_API_KEY", ""),
            api_base_url=os.getenv("AI_API_BASE", DEFAULT_API_BASE_URL),
            model_name=os.getenv("AI_MODEL_NAME", DEFAULT_MODEL_NAME),
        )

    @classmethod
    def from_dict(cls, data: dict) -> "AIConfig":
        """从字典创建配置"""
        return cls(
            api_key=data.get("api_key", ""),
            api_base_url=data.get("api_base", DEFAULT_API_BASE_URL),
            model_name=data.get("model_name", DEFAULT_MODEL_NAME),
            temperature=data.get("temperature", DEFAULT_TEMPERATURE),
        )


@dataclass
class ImageConfig:
    """图片搜索配置"""
    unsplash_key: str = ""
    pexels_key: str = ""
    pixabay_key: str = ""
    download_dir: str = "images/downloaded"
    cache_enabled: bool = True
    max_concurrent_downloads: int = DEFAULT_MAX_CONCURRENT_DOWNLOADS
    download_timeout: int = DEFAULT_DOWNLOAD_TIMEOUT

    @classmethod
    def from_env(cls) -> "ImageConfig":
        """从环境变量创建配置"""
        return cls(
            unsplash_key=os.getenv("UNSPLASH_ACCESS_KEY", ""),
            pexels_key=os.getenv("PEXELS_API_KEY", ""),
            pixabay_key=os.getenv("PIXABAY_API_KEY", ""),
            max_concurrent_downloads=int(os.getenv("IMAGE_MAX_CONCURRENT", str(DEFAULT_MAX_CONCURRENT_DOWNLOADS))),
        )


def _get_or_create_secret_key() -> str:
    """获取或创建持久化的 SECRET_KEY"""
    from pathlib import Path

    # 优先使用环境变量
    key = os.getenv("SECRET_KEY")
    if key:
        return key

    # 尝试从文件读取
    key_file = Path(".secret_key")
    if key_file.exists():
        try:
            return key_file.read_text().strip()
        except Exception:
            pass

    # 生成新 key 并持久化
    key = os.urandom(24).hex()
    try:
        key_file.write_text(key)
    except Exception:
        pass  # 无法写入时仍返回生成的 key
    return key


@dataclass
class AppConfig:
    """应用配置"""
    secret_key: str = field(default_factory=_get_or_create_secret_key)
    upload_folder: str = "web/uploads"
    output_folder: str = "web/outputs"
    max_upload_size: int = DEFAULT_MAX_UPLOAD_SIZE
    debug: bool = False
    host: str = "0.0.0.0"
    port: int = DEFAULT_PORT

    @classmethod
    def from_env(cls) -> "AppConfig":
        """从环境变量创建配置"""
        return cls(
            secret_key=_get_or_create_secret_key(),
            debug=os.getenv("DEBUG", "false").lower() == "true",
            port=int(os.getenv("PORT", str(DEFAULT_PORT))),
        )


# 全局默认配置（仅用于 CLI 模式）
class Settings:
    """兼容旧版的配置类"""

    def __init__(self):
        self._config: Optional[AIConfig] = None

    @property
    def api_key(self) -> str:
        return os.getenv("AI_API_KEY", "")

    @property
    def api_base_url(self) -> str:
        return os.getenv("AI_API_BASE", DEFAULT_API_BASE_URL)

    @property
    def model_name(self) -> str:
        return os.getenv("AI_MODEL_NAME", DEFAULT_MODEL_NAME)

    @property
    def default_template(self) -> str:
        return "ppt/pptx_templates/default.pptx"

    @property
    def default_output(self) -> str:
        return "output.pptx"

    def validate(self) -> bool:
        if not self.api_key:
            raise ValueError("API_KEY 未设置，请设置环境变量 AI_API_KEY")
        return True

    def to_ai_config(self) -> AIConfig:
        """转换为 AIConfig 对象"""
        return AIConfig(
            api_key=self.api_key,
            api_base_url=self.api_base_url,
            model_name=self.model_name,
        )


settings = Settings()
