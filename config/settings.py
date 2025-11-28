"""配置模块：管理 API 配置和项目设置"""
from dataclasses import dataclass, field
from typing import Optional
import os


@dataclass
class AIConfig:
    """AI 配置类 - 支持实例化传递，解决多用户并发问题"""
    api_key: str
    api_base_url: str = "https://api.openai.com/v1"
    model_name: str = "gpt-4o-mini"
    temperature: float = 0.7
    max_retries: int = 3
    timeout: int = 180
    
    def __post_init__(self):
        """初始化后处理 - 规范化 API Base URL"""
        self.api_base_url = self._normalize_base_url(self.api_base_url)
    
    @staticmethod
    def _normalize_base_url(url: str) -> str:
        """规范化 API Base URL - 只做基本清理，不自动添加路径"""
        if not url:
            return "https://api.openai.com/v1"
        
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
            api_base_url=os.getenv("AI_API_BASE", "https://api.openai.com/v1"),
            model_name=os.getenv("AI_MODEL_NAME", "gpt-4o-mini"),
        )
    
    @classmethod
    def from_dict(cls, data: dict) -> "AIConfig":
        """从字典创建配置"""
        return cls(
            api_key=data.get("api_key", ""),
            api_base_url=data.get("api_base", "https://api.openai.com/v1"),
            model_name=data.get("model_name", "gpt-4o-mini"),
            temperature=data.get("temperature", 0.7),
        )


@dataclass
class ImageConfig:
    """图片搜索配置"""
    unsplash_key: str = ""
    download_dir: str = "images/downloaded"
    cache_enabled: bool = True
    max_concurrent_downloads: int = 3
    
    @classmethod
    def from_env(cls) -> "ImageConfig":
        """从环境变量创建配置"""
        return cls(
            unsplash_key=os.getenv("UNSPLASH_ACCESS_KEY", ""),
        )


@dataclass  
class AppConfig:
    """应用配置"""
    secret_key: str = field(default_factory=lambda: os.getenv("SECRET_KEY", os.urandom(24).hex()))
    upload_folder: str = "web/uploads"
    output_folder: str = "web/outputs"
    max_upload_size: int = 10 * 1024 * 1024  # 10MB
    debug: bool = False
    host: str = "0.0.0.0"
    port: int = 5000
    
    @classmethod
    def from_env(cls) -> "AppConfig":
        """从环境变量创建配置"""
        return cls(
            secret_key=os.getenv("SECRET_KEY", os.urandom(24).hex()),
            debug=os.getenv("DEBUG", "false").lower() == "true",
            port=int(os.getenv("PORT", "5000")),
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
        return os.getenv("AI_API_BASE", "https://api.openai.com/v1")
    
    @property
    def model_name(self) -> str:
        return os.getenv("AI_MODEL_NAME", "gpt-4o-mini")
    
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
