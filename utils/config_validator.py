"""配置验证模块

提供配置 Schema 定义和验证功能。
"""
import os
from dataclasses import dataclass, field, fields
from typing import Any, Dict, List, Optional, Type, TypeVar, get_type_hints
from pathlib import Path

T = TypeVar("T")


class ConfigValidationError(Exception):
    """配置验证错误"""

    def __init__(self, errors: List[str]):
        self.errors = errors
        super().__init__(f"配置验证失败: {'; '.join(errors)}")


def validate_config(config_class: Type[T], data: Dict[str, Any]) -> T:
    """验证并创建配置对象

    Args:
        config_class: 配置类（dataclass）
        data: 配置数据字典

    Returns:
        验证后的配置对象
    """
    errors = []
    validated_data = {}

    # 获取类型提示
    hints = get_type_hints(config_class)

    for f in fields(config_class):
        name = f.name
        value = data.get(name, f.default if f.default is not f.default_factory else None)

        # 处理 default_factory
        if value is None and f.default_factory is not f.default_factory:
            value = f.default_factory()

        # 获取字段元数据
        metadata = f.metadata or {}
        required = metadata.get("required", False)
        min_value = metadata.get("min")
        max_value = metadata.get("max")
        min_length = metadata.get("min_length")
        max_length = metadata.get("max_length")
        pattern = metadata.get("pattern")
        choices = metadata.get("choices")
        validator = metadata.get("validator")

        # 必填检查
        if required and value is None:
            errors.append(f"{name} 是必填项")
            continue

        if value is None:
            validated_data[name] = f.default if f.default is not f.default_factory else None
            continue

        # 类型转换和验证
        expected_type = hints.get(name, str)
        try:
            # 处理 Optional 类型
            origin = getattr(expected_type, "__origin__", None)
            if origin is type(None) or (hasattr(expected_type, "__args__") and type(None) in expected_type.__args__):
                # 提取非 None 类型
                if hasattr(expected_type, "__args__"):
                    expected_type = next(t for t in expected_type.__args__ if t is not type(None))

            # 类型转换
            if expected_type == int:
                value = int(value)
            elif expected_type == float:
                value = float(value)
            elif expected_type == bool:
                if isinstance(value, str):
                    value = value.lower() in ("true", "1", "yes", "on")
                else:
                    value = bool(value)
            elif expected_type == str:
                value = str(value)

        except (ValueError, TypeError) as e:
            errors.append(f"{name} 类型错误: {e}")
            continue

        # 范围验证
        if min_value is not None and value < min_value:
            errors.append(f"{name} 不能小于 {min_value}")
        if max_value is not None and value > max_value:
            errors.append(f"{name} 不能大于 {max_value}")

        # 长度验证
        if isinstance(value, str):
            if min_length is not None and len(value) < min_length:
                errors.append(f"{name} 长度不能小于 {min_length}")
            if max_length is not None and len(value) > max_length:
                errors.append(f"{name} 长度不能超过 {max_length}")

        # 正则验证
        if pattern and isinstance(value, str):
            import re
            if not re.match(pattern, value):
                errors.append(f"{name} 格式不正确")

        # 枚举验证
        if choices and value not in choices:
            errors.append(f"{name} 必须是以下值之一: {choices}")

        # 自定义验证器
        if validator:
            try:
                if not validator(value):
                    errors.append(f"{name} 验证失败")
            except Exception as e:
                errors.append(f"{name} 验证错误: {e}")

        validated_data[name] = value

    if errors:
        raise ConfigValidationError(errors)

    return config_class(**validated_data)


def config_field(
    default: Any = None,
    required: bool = False,
    min: Any = None,
    max: Any = None,
    min_length: int = None,
    max_length: int = None,
    pattern: str = None,
    choices: List[Any] = None,
    validator: callable = None,
    env_var: str = None,
    description: str = "",
):
    """创建带验证元数据的配置字段

    用法:
        @dataclass
        class MyConfig:
            port: int = config_field(
                default=8080,
                min=1,
                max=65535,
                env_var="APP_PORT",
                description="服务端口"
            )
    """
    metadata = {
        "required": required,
        "min": min,
        "max": max,
        "min_length": min_length,
        "max_length": max_length,
        "pattern": pattern,
        "choices": choices,
        "validator": validator,
        "env_var": env_var,
        "description": description,
    }

    # 过滤掉 None 值
    metadata = {k: v for k, v in metadata.items() if v is not None}

    if default is None:
        return field(default=None, metadata=metadata)
    return field(default=default, metadata=metadata)


def load_config_from_env(config_class: Type[T], prefix: str = "") -> T:
    """从环境变量加载配置

    Args:
        config_class: 配置类
        prefix: 环境变量前缀

    Returns:
        配置对象
    """
    data = {}

    for f in fields(config_class):
        # 获取环境变量名
        env_var = f.metadata.get("env_var") if f.metadata else None
        if not env_var:
            env_var = f"{prefix}{f.name}".upper()

        value = os.getenv(env_var)
        if value is not None:
            data[f.name] = value
        elif f.default is not f.default_factory:
            data[f.name] = f.default
        elif f.default_factory is not f.default_factory:
            data[f.name] = f.default_factory()

    return validate_config(config_class, data)


# 常用验证器
def is_valid_port(value: int) -> bool:
    """验证端口号"""
    return 1 <= value <= 65535


def is_valid_url(value: str) -> bool:
    """验证 URL"""
    from urllib.parse import urlparse
    try:
        result = urlparse(value)
        return all([result.scheme in ("http", "https"), result.netloc])
    except Exception:
        return False


def is_valid_path(value: str) -> bool:
    """验证路径是否存在"""
    return Path(value).exists()


def is_valid_email(value: str) -> bool:
    """验证邮箱格式"""
    import re
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, value))


# 示例配置类
@dataclass
class ServerConfig:
    """服务器配置示例"""

    host: str = config_field(
        default="0.0.0.0",
        description="监听地址"
    )
    port: int = config_field(
        default=5000,
        min=1,
        max=65535,
        env_var="PORT",
        description="监听端口"
    )
    debug: bool = config_field(
        default=False,
        env_var="DEBUG",
        description="调试模式"
    )
    workers: int = config_field(
        default=4,
        min=1,
        max=32,
        description="工作进程数"
    )
    log_level: str = config_field(
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        description="日志级别"
    )


def print_config_help(config_class: Type) -> str:
    """打印配置帮助信息"""
    lines = [f"配置项 ({config_class.__name__}):", "=" * 50]

    hints = get_type_hints(config_class)

    for f in fields(config_class):
        metadata = f.metadata or {}
        name = f.name
        type_name = hints.get(name, type(f.default)).__name__
        default = f.default if f.default is not f.default_factory else "(factory)"
        env_var = metadata.get("env_var", name.upper())
        description = metadata.get("description", "")
        required = metadata.get("required", False)

        line = f"  {name} ({type_name})"
        if required:
            line += " [必填]"
        if description:
            line += f": {description}"

        lines.append(line)
        lines.append(f"    默认值: {default}")
        lines.append(f"    环境变量: {env_var}")

        if metadata.get("min") is not None or metadata.get("max") is not None:
            range_str = f"    范围: {metadata.get('min', '-∞')} ~ {metadata.get('max', '+∞')}"
            lines.append(range_str)

        if metadata.get("choices"):
            lines.append(f"    可选值: {metadata.get('choices')}")

        lines.append("")

    return "\n".join(lines)
