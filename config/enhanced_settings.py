"""配置管理器 - 支持加密敏感配置"""
import os
import json
from typing import Dict, Any, Optional, Union
from dataclasses import dataclass, asdict
from pathlib import Path
import base64
from cryptography.fernet import Fernet
from utils.logger import get_logger

logger = get_logger("config")


def decrypt_config_value(encrypted_value: Optional[str]) -> Optional[str]:
    """解密配置值"""
    if not encrypted_value:
        return None

    try:
        # 检查是否为加密值（以enc:开头）
        if encrypted_value.startswith("enc:"):
            # 从环境变量获取解密密钥
            key = os.getenv('CONFIG_DECRYPT_KEY')
            if not key:
                logger.warning("未设置CONFIG_DECRYPT_KEY，无法解密配置")
                return encrypted_value[4:]  # 返回加密前的值（去掉enc:前缀）

            # 解密
            cipher = Fernet(key.encode())
            encrypted_data = base64.urlsafe_b64decode(encrypted_value[4:].encode())
            decrypted = cipher.decrypt(encrypted_data).decode()
            return decrypted
        else:
            return encrypted_value
    except Exception as e:
        logger.error(f"解密配置失败: {e}")
        return encrypted_value


@dataclass
class DatabaseConfig:
    """数据库配置"""
    host: str = "localhost"
    port: int = 5432
    name: str = "ai_ppt"
    username: str = ""
    password: str = ""
    pool_size: int = 5
    ssl_mode: str = "prefer"

    def __post_init__(self):
        # 解密敏感字段
        self.password = decrypt_config_value(self.password) or self.password


@dataclass
class RedisConfig:
    """Redis配置"""
    scheme: str = "redis"
    host: str = "localhost"
    port: int = 6379
    password: Optional[str] = None
    database: int = 0
    max_connections: int = 10
    socket_timeout: int = 5
    socket_connect_timeout: int = 5

    def __post_init__(self):
        self.password = decrypt_config_value(self.password)


@dataclass
class AIConfig:
    """AI配置"""
    api_key: str = ""
    api_base_url: str = "https://api.openai.com/v1"
    model_name: str = "gpt-4o-mini"
    temperature: float = 0.7
    max_retries: int = 3
    timeout: int = 60
    rate_limit: int = 100  # 每分钟请求数

    def __post_init__(self):
        self.api_key = decrypt_config_value(self.api_key) or self.api_key

    def validate(self):
        """验证配置"""
        if not self.api_key:
            raise ValueError("AI API Key 未配置")
        if not self.api_base_url:
            raise ValueError("API Base URL 未配置")
        if self.temperature < 0 or self.temperature > 2:
            raise ValueError("temperature 必须在 0-2 之间")
        if self.max_retries < 0 or self.max_retries > 10:
            raise ValueError("max_retries 必须在 0-10 之间")


@dataclass
class UnsplashConfig:
    """Unsplash配置"""
    access_key: str = ""
    secret_key: Optional[str] = None
    enable: bool = False
    daily_limit: int = 50

    def __post_init__(self):
        self.access_key = decrypt_config_value(self.access_key) or self.access_key
        self.secret_key = decrypt_config_value(self.secret_key) if self.secret_key else None


@dataclass
class SecurityConfig:
    """安全配置"""
    secret_key: str = ""
    jwt_secret: str = ""
    encryption_key: str = ""
    password_salt: str = ""
    csrf_protection: bool = True
    rate_limiting: bool = True
    max_login_attempts: int = 5
    session_timeout: int = 3600

    def __post_init__(self):
        self.secret_key = decrypt_config_value(self.secret_key) or self.secret_key
        self.jwt_secret = decrypt_config_value(self.jwt_secret) or self.jwt_secret
        self.encryption_key = decrypt_config_value(self.encryption_key) or self.encryption_key
        self.password_salt = decrypt_config_value(self.password_salt) or self.password_salt

    def validate(self):
        """验证安全配置"""
        if not self.secret_key:
            # 生成默认密钥
            self.secret_key = Fernet.generate_key().decode()
            logger.warning("使用临时密钥，重启后将失效。请设置SECRET_KEY环境变量")

        if not self.encryption_key:
            self.encryption_key = Fernet.generate_key().decode()
            logger.warning("使用临时加密密钥，重启后将失效。请设置ENCRYPTION_KEY环境变量")


@dataclass
class SystemConfig:
    """系统配置"""
    debug: bool = False
    log_level: str = "INFO"
    max_workers: int = 3
    enable_async: bool = True
    cache_enabled: bool = True
    cache_ttl: int = 3600
    output_dir: str = "web/outputs"
    temp_dir: str = "temp"
    max_ppt_size: int = 100 * 1024 * 1024  # 100MB
    cleanup_interval: int = 3600
    metrics_enabled: bool = True

    def __post_init__(self):
        # 确保目录存在
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        Path(self.temp_dir).mkdir(parents=True, exist_ok=True)


@dataclass
class AppConfig:
    """应用完整配置"""
    database: DatabaseConfig = DatabaseConfig()
    redis: RedisConfig = RedisConfig()
    ai: AIConfig = AIConfig()
    unsplash: UnsplashConfig = UnsplashConfig()
    security: SecurityConfig = SecurityConfig()
    system: SystemConfig = SystemConfig()

    @classmethod
    def from_env(cls) -> "AppConfig":
        """从环境变量加载配置"""
        # 数据库配置
        db_config = DatabaseConfig(
            host=os.getenv("DB_HOST", "localhost"),
            port=int(os.getenv("DB_PORT", "5432")),
            name=os.getenv("DB_NAME", "ai_ppt"),
            username=os.getenv("DB_USERNAME", ""),
            password=os.getenv("DB_PASSWORD", ""),
            pool_size=int(os.getenv("DB_POOL_SIZE", "5")),
            ssl_mode=os.getenv("DB_SSL_MODE", "prefer")
        )

        # Redis配置
        redis_config = RedisConfig(
            scheme=os.getenv("REDIS_SCHEME", "redis"),
            host=os.getenv("REDIS_HOST", "localhost"),
            port=int(os.getenv("REDIS_PORT", "6379")),
            password=os.getenv("REDIS_PASSWORD"),
            database=int(os.getenv("REDIS_DB", "0")),
            max_connections=int(os.getenv("REDIS_MAX_CONNECTIONS", "10"))
        )

        # AI配置
        ai_config = AIConfig(
            api_key=os.getenv("AI_API_KEY", ""),
            api_base_url=os.getenv("AI_API_BASE", "https://api.openai.com/v1"),
            model_name=os.getenv("AI_MODEL_NAME", "gpt-4o-mini"),
            temperature=float(os.getenv("AI_TEMPERATURE", "0.7")),
            max_retries=int(os.getenv("AI_MAX_RETRIES", "3")),
            timeout=int(os.getenv("AI_TIMEOUT", "60")),
            rate_limit=int(os.getenv("AI_RATE_LIMIT", "100"))
        )

        # Unsplash配置
        unsplash_config = UnsplashConfig(
            access_key=os.getenv("UNSPLASH_ACCESS_KEY", ""),
            secret_key=os.getenv("UNSPLASH_SECRET_KEY"),
            enable=os.getenv("UNSPLASH_ENABLE", "").lower() == "true",
            daily_limit=int(os.getenv("UNSPLASH_DAILY_LIMIT", "50"))
        )

        # 安全配置
        security_config = SecurityConfig(
            secret_key=os.getenv("SECRET_KEY", ""),
            jwt_secret=os.getenv("JWT_SECRET", ""),
            encryption_key=os.getenv("ENCRYPTION_KEY", ""),
            password_salt=os.getenv("PASSWORD_SALT", ""),
            csrf_protection=os.getenv("CSRF_PROTECTION", "").lower() != "false",
            rate_limiting=os.getenv("RATE_LIMITING", "").lower() != "false",
            max_login_attempts=int(os.getenv("MAX_LOGIN_ATTEMPTS", "5")),
            session_timeout=int(os.getenv("SESSION_TIMEOUT", "3600"))
        )

        # 系统配置
        system_config = SystemConfig(
            debug=os.getenv("DEBUG", "").lower() == "true",
            log_level=os.getenv("LOG_LEVEL", "INFO"),
            max_workers=int(os.getenv("MAX_WORKERS", "3")),
            enable_async=os.getenv("ENABLE_ASYNC", "true").lower() == "true",
            cache_enabled=os.getenv("CACHE_ENABLED", "true").lower() == "true",
            cache_ttl=int(os.getenv("CACHE_TTL", "3600")),
            output_dir=os.getenv("OUTPUT_DIR", "web/outputs"),
            temp_dir=os.getenv("TEMP_DIR", "temp"),
            max_ppt_size=int(os.getenv("MAX_PPT_SIZE", str(100 * 1024 * 1024))),
            cleanup_interval=int(os.getenv("CLEANUP_INTERVAL", "3600")),
            metrics_enabled=os.getenv("METRICS_ENABLED", "true").lower() == "true"
        )

        return cls(
            database=db_config,
            redis=redis_config,
            ai=ai_config,
            unsplash=unsplash_config,
            security=security_config,
            system=system_config
        )

    @classmethod
    def from_file(cls, config_path: str) -> "AppConfig":
        """从配置文件加载"""
        path = Path(config_path)
        if not path.exists():
            raise FileNotFoundError(f"配置文件不存在: {config_path}")

        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        return cls(**data)

    def save(self, config_path: str):
        """保存配置到文件（加密敏感信息）"""
        path = Path(config_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        # 获取加密密钥
        encrypt_key = os.getenv('CONFIG_ENCRYPT_KEY')
        if encrypt_key:
            cipher = Fernet(encrypt_key.encode())

            # 加密敏感字段
            config_dict = asdict(self)

            # AI配置
            if config_dict['ai']['api_key']:
                encrypted = cipher.encrypt(config_dict['ai']['api_key'].encode())
                config_dict['ai']['api_key'] = "enc:" + base64.urlsafe_b64encode(encrypted).decode()

            # Unsplash配置
            if config_dict['unsplash']['access_key']:
                encrypted = cipher.encrypt(config_dict['unsplash']['access_key'].encode())
                config_dict['unsplash']['access_key'] = "enc:" + base64.urlsafe_b64encode(encrypted).decode()

            if config_dict['unsplash']['secret_key']:
                encrypted = cipher.encrypt(config_dict['unsplash']['secret_key'].encode())
                config_dict['unsplash']['secret_key'] = "enc:" + base64.urlsafe_b64encode(encrypted).decode()

            # 安全配置
            for field in ['secret_key', 'jwt_secret', 'encryption_key', 'password_salt']:
                if config_dict['security'][field]:
                    encrypted = cipher.encrypt(config_dict['security'][field].encode())
                    config_dict['security'][field] = "enc:" + base64.urlsafe_b64encode(encrypted).decode()

            # 数据库密码
            if config_dict['database']['password']:
                encrypted = cipher.encrypt(config_dict['database']['password'].encode())
                config_dict['database']['password'] = "enc:" + base64.urlsafe_b64encode(encrypted).decode()

            # Redis密码
            if config_dict['redis']['password']:
                encrypted = cipher.encrypt(config_dict['redis']['password'].encode())
                config_dict['redis']['password'] = "enc:" + base64.urlsafe_b64encode(encrypted).decode()
        else:
            # 不加密
            config_dict = asdict(self)

        # 保存文件
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=2)

        # 设置文件权限（仅所有者可读写）
        if os.name != 'nt':  # Unix系统
            path.chmod(0o600)

        logger.info(f"配置已保存到: {config_path}")

    def validate_all(self):
        """验证所有配置"""
        self.ai.validate()
        self.security.validate()

        # 验证路径
        if not Path(self.system.output_dir).exists():
            Path(self.system.output_dir).mkdir(parents=True, exist_ok=True)

        if not Path(self.system.temp_dir).exists():
            Path(self.system.temp_dir).mkdir(parents=True, exist_ok=True)

        logger.info("配置验证通过")

    def to_ai_config(self) -> AIConfig:
        """转换为AI配置（保持向后兼容）"""
        return self.ai

    def get_redis_url(self) -> Optional[str]:
        """获取Redis连接URL"""
        if self.redis.password:
            return f"{self.redis.scheme}://:{self.redis.password}@{self.redis.host}:{self.redis.port}/{self.redis.database}"
        else:
            return f"{self.redis.scheme}://{self.redis.host}:{self.redis.port}/{self.redis.database}"

    def get_database_url(self) -> str:
        """获取数据库连接URL"""
        if self.database.password:
            return f"postgresql://{self.database.username}:{self.database.password}@{self.database.host}:{self.database.port}/{self.database.name}"
        else:
            return f"postgresql://{self.database.username}@{self.database.host}:{self.database.port}/{self.database.name}"


# 全局配置实例
_config: Optional[AppConfig] = None


def load_config(config_path: Optional[str] = None) -> AppConfig:
    """加载配置"""
    global _config

    if _config is None:
        if config_path and Path(config_path).exists():
            _config = AppConfig.from_file(config_path)
            logger.info(f"从文件加载配置: {config_path}")
        else:
            _config = AppConfig.from_env()
            logger.info("从环境变量加载配置")

        # 验证配置
        _config.validate_all()

    return _config


def get_config() -> AppConfig:
    """获取配置实例"""
    if _config is None:
        return load_config()
    return _config


def reload_config(config_path: Optional[str] = None):
    """重新加载配置"""
    global _config
    _config = None
    return load_config(config_path)


# 向后兼容
settings = get_config()