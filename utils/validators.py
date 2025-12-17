"""输入验证模块

提供可复用的验证器，用于验证用户输入和 API 请求参数。
"""
import re
from typing import Any, Dict, List, Optional, Tuple, Union
from dataclasses import dataclass
from urllib.parse import urlparse


@dataclass
class ValidationError:
    """验证错误"""
    field: str
    message: str
    code: str = "INVALID"

    def to_dict(self) -> Dict[str, str]:
        return {
            "field": self.field,
            "message": self.message,
            "code": self.code,
        }


class ValidationResult:
    """验证结果"""

    def __init__(self):
        self.errors: List[ValidationError] = []
        self._data: Dict[str, Any] = {}

    @property
    def is_valid(self) -> bool:
        return len(self.errors) == 0

    def add_error(self, field: str, message: str, code: str = "INVALID"):
        self.errors.append(ValidationError(field, message, code))

    def set_value(self, field: str, value: Any):
        self._data[field] = value

    def get_value(self, field: str, default: Any = None) -> Any:
        return self._data.get(field, default)

    @property
    def data(self) -> Dict[str, Any]:
        return self._data.copy()

    def to_error_response(self) -> Dict[str, Any]:
        return {
            "success": False,
            "error": self.errors[0].message if self.errors else "验证失败",
            "errors": [e.to_dict() for e in self.errors],
        }


class Validators:
    """验证器集合"""

    @staticmethod
    def required(value: Any, field: str) -> Optional[ValidationError]:
        """必填验证"""
        if value is None or (isinstance(value, str) and not value.strip()):
            return ValidationError(field, f"{field} 不能为空", "REQUIRED")
        return None

    @staticmethod
    def string_length(
        value: str,
        field: str,
        min_length: int = 0,
        max_length: int = 10000,
    ) -> Optional[ValidationError]:
        """字符串长度验证"""
        if not isinstance(value, str):
            return ValidationError(field, f"{field} 必须是字符串", "TYPE_ERROR")

        length = len(value.strip())
        if length < min_length:
            return ValidationError(
                field, f"{field} 至少需要 {min_length} 个字符", "TOO_SHORT"
            )
        if length > max_length:
            return ValidationError(
                field, f"{field} 不能超过 {max_length} 个字符", "TOO_LONG"
            )
        return None

    @staticmethod
    def integer_range(
        value: Any,
        field: str,
        min_value: int = None,
        max_value: int = None,
        default: int = None,
    ) -> Tuple[Optional[int], Optional[ValidationError]]:
        """整数范围验证，返回 (转换后的值, 错误)"""
        try:
            num = int(value)
        except (ValueError, TypeError):
            if default is not None:
                return default, None
            return None, ValidationError(field, f"{field} 必须是整数", "TYPE_ERROR")

        if min_value is not None and num < min_value:
            num = min_value
        if max_value is not None and num > max_value:
            num = max_value

        return num, None

    @staticmethod
    def url(value: str, field: str, allow_localhost: bool = False) -> Optional[ValidationError]:
        """URL 验证"""
        if not value:
            return ValidationError(field, f"{field} 不能为空", "REQUIRED")

        try:
            parsed = urlparse(value)
            if parsed.scheme not in ("http", "https"):
                return ValidationError(
                    field, f"{field} 必须是 HTTP 或 HTTPS URL", "INVALID_SCHEME"
                )

            if not parsed.hostname:
                return ValidationError(field, f"{field} URL 格式无效", "INVALID_FORMAT")

            if not allow_localhost:
                host = parsed.hostname.lower()
                blocked = ("localhost", "127.0.0.1", "0.0.0.0", "::1")
                if host in blocked:
                    return ValidationError(
                        field, f"{field} 不允许使用本地地址", "BLOCKED_HOST"
                    )

                # 检查内网地址
                if host.startswith("10.") or host.startswith("192.168."):
                    return ValidationError(
                        field, f"{field} 不允许使用内网地址", "BLOCKED_HOST"
                    )
                if host.startswith("172."):
                    try:
                        second = int(host.split(".")[1])
                        if 16 <= second <= 31:
                            return ValidationError(
                                field, f"{field} 不允许使用内网地址", "BLOCKED_HOST"
                            )
                    except (IndexError, ValueError):
                        pass

        except Exception:
            return ValidationError(field, f"{field} URL 格式无效", "INVALID_FORMAT")

        return None

    @staticmethod
    def api_key(value: str, field: str = "api_key") -> Optional[ValidationError]:
        """API Key 验证"""
        if not value or not value.strip():
            return ValidationError(field, "请提供 API Key", "REQUIRED")

        value = value.strip()
        if len(value) < 10:
            return ValidationError(field, "API Key 格式无效", "INVALID_FORMAT")

        return None

    @staticmethod
    def model_name(value: str, field: str = "model_name") -> Optional[ValidationError]:
        """模型名称验证"""
        if not value:
            return None  # 可选字段

        value = value.strip()
        if len(value) > 100:
            return ValidationError(field, "模型名称过长", "TOO_LONG")

        # 只允许字母、数字、连字符、下划线和点
        if not re.match(r"^[\w\-\.]+$", value):
            return ValidationError(field, "模型名称包含无效字符", "INVALID_CHARS")

        return None

    @staticmethod
    def filename(value: str, field: str = "filename") -> Tuple[str, Optional[ValidationError]]:
        """文件名验证和清理"""
        if not value:
            return "", ValidationError(field, "文件名不能为空", "REQUIRED")

        # 移除路径分隔符和危险字符
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", value)
        safe = safe.strip(". ")

        if not safe:
            return "", ValidationError(field, "文件名无效", "INVALID")

        if len(safe) > 200:
            safe = safe[:200]

        return safe, None

    @staticmethod
    def email(value: str, field: str = "email") -> Optional[ValidationError]:
        """邮箱验证"""
        if not value:
            return ValidationError(field, "邮箱不能为空", "REQUIRED")

        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        if not re.match(pattern, value.strip()):
            return ValidationError(field, "邮箱格式无效", "INVALID_FORMAT")

        return None


class RequestValidator:
    """请求验证器

    用法:
        validator = RequestValidator(request.json)
        validator.require("topic").string(min_length=1, max_length=500)
        validator.optional("page_count").integer(min_value=1, max_value=100, default=5)
        validator.require("api_key").api_key()

        if not validator.is_valid:
            return jsonify(validator.to_error_response()), 400

        data = validator.data
    """

    def __init__(self, data: Dict[str, Any] = None):
        self._data = data or {}
        self._result = ValidationResult()
        self._current_field: Optional[str] = None
        self._current_value: Any = None
        self._is_required: bool = False

    def require(self, field: str) -> "RequestValidator":
        """标记必填字段"""
        self._current_field = field
        self._current_value = self._data.get(field)
        self._is_required = True

        error = Validators.required(self._current_value, field)
        if error:
            self._result.add_error(error.field, error.message, error.code)

        return self

    def optional(self, field: str, default: Any = None) -> "RequestValidator":
        """标记可选字段"""
        self._current_field = field
        self._current_value = self._data.get(field, default)
        self._is_required = False
        return self

    def string(
        self, min_length: int = 0, max_length: int = 10000, strip: bool = True
    ) -> "RequestValidator":
        """字符串验证"""
        if self._current_value is None:
            if not self._is_required:
                self._result.set_value(self._current_field, "")
            return self

        value = str(self._current_value)
        if strip:
            value = value.strip()

        error = Validators.string_length(value, self._current_field, min_length, max_length)
        if error:
            self._result.add_error(error.field, error.message, error.code)
        else:
            self._result.set_value(self._current_field, value)

        return self

    def integer(
        self,
        min_value: int = None,
        max_value: int = None,
        default: int = None,
    ) -> "RequestValidator":
        """整数验证"""
        value, error = Validators.integer_range(
            self._current_value,
            self._current_field,
            min_value,
            max_value,
            default,
        )
        if error:
            self._result.add_error(error.field, error.message, error.code)
        else:
            self._result.set_value(self._current_field, value)

        return self

    def boolean(self, default: bool = False) -> "RequestValidator":
        """布尔值验证"""
        value = self._current_value
        if value is None:
            value = default
        else:
            value = bool(value)

        self._result.set_value(self._current_field, value)
        return self

    def url(self, allow_localhost: bool = False) -> "RequestValidator":
        """URL 验证"""
        if self._current_value is None:
            if not self._is_required:
                self._result.set_value(self._current_field, "")
            return self

        value = str(self._current_value).strip()
        error = Validators.url(value, self._current_field, allow_localhost)
        if error:
            self._result.add_error(error.field, error.message, error.code)
        else:
            self._result.set_value(self._current_field, value)

        return self

    def api_key(self) -> "RequestValidator":
        """API Key 验证"""
        if self._current_value is None:
            if self._is_required:
                self._result.add_error(self._current_field, "请提供 API Key", "REQUIRED")
            return self

        value = str(self._current_value).strip()
        error = Validators.api_key(value, self._current_field)
        if error:
            self._result.add_error(error.field, error.message, error.code)
        else:
            self._result.set_value(self._current_field, value)

        return self

    def model_name(self, default: str = "gpt-4o-mini") -> "RequestValidator":
        """模型名称验证"""
        value = self._current_value
        if not value:
            value = default

        value = str(value).strip()
        error = Validators.model_name(value, self._current_field)
        if error:
            self._result.add_error(error.field, error.message, error.code)
        else:
            self._result.set_value(self._current_field, value)

        return self

    def one_of(self, choices: List[Any], default: Any = None) -> "RequestValidator":
        """枚举验证"""
        value = self._current_value
        if value is None:
            value = default

        if value not in choices:
            self._result.add_error(
                self._current_field,
                f"{self._current_field} 必须是以下值之一: {', '.join(map(str, choices))}",
                "INVALID_CHOICE",
            )
        else:
            self._result.set_value(self._current_field, value)

        return self

    @property
    def is_valid(self) -> bool:
        return self._result.is_valid

    @property
    def errors(self) -> List[ValidationError]:
        return self._result.errors

    @property
    def data(self) -> Dict[str, Any]:
        return self._result.data

    def to_error_response(self) -> Dict[str, Any]:
        return self._result.to_error_response()
