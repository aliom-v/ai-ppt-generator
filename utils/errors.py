"""增强错误处理模块

提供统一的错误码、错误分类和用户友好的错误消息。
"""
from enum import Enum
from typing import Any, Dict, Optional
from dataclasses import dataclass
from http import HTTPStatus


class ErrorCategory(str, Enum):
    """错误分类"""
    VALIDATION = "validation"       # 输入验证错误
    AUTHENTICATION = "auth"         # 认证错误
    AUTHORIZATION = "permission"    # 权限错误
    RESOURCE = "resource"           # 资源错误
    EXTERNAL = "external"           # 外部服务错误
    RATE_LIMIT = "rate_limit"       # 限流错误
    INTERNAL = "internal"           # 内部错误
    BUSINESS = "business"           # 业务逻辑错误


class ErrorCode(str, Enum):
    """标准错误码"""

    # 通用错误 (1xxx)
    UNKNOWN_ERROR = "E1000"
    INTERNAL_ERROR = "E1001"
    SERVICE_UNAVAILABLE = "E1002"
    MAINTENANCE_MODE = "E1003"

    # 验证错误 (2xxx)
    VALIDATION_ERROR = "E2000"
    MISSING_FIELD = "E2001"
    INVALID_FORMAT = "E2002"
    VALUE_TOO_LONG = "E2003"
    VALUE_TOO_SHORT = "E2004"
    VALUE_OUT_OF_RANGE = "E2005"
    INVALID_TYPE = "E2006"
    INVALID_URL = "E2007"
    BLOCKED_URL = "E2008"

    # 认证错误 (3xxx)
    AUTH_REQUIRED = "E3000"
    INVALID_API_KEY = "E3001"
    EXPIRED_API_KEY = "E3002"
    INVALID_TOKEN = "E3003"

    # 权限错误 (4xxx)
    PERMISSION_DENIED = "E4000"
    RESOURCE_FORBIDDEN = "E4001"

    # 资源错误 (5xxx)
    RESOURCE_NOT_FOUND = "E5000"
    FILE_NOT_FOUND = "E5001"
    TEMPLATE_NOT_FOUND = "E5002"
    TASK_NOT_FOUND = "E5003"

    # 外部服务错误 (6xxx)
    EXTERNAL_SERVICE_ERROR = "E6000"
    AI_API_ERROR = "E6001"
    AI_API_TIMEOUT = "E6002"
    AI_API_RATE_LIMIT = "E6003"
    AI_RESPONSE_INVALID = "E6004"
    IMAGE_API_ERROR = "E6005"
    CIRCUIT_BREAKER_OPEN = "E6006"

    # 限流错误 (7xxx)
    RATE_LIMIT_EXCEEDED = "E7000"
    QUOTA_EXCEEDED = "E7001"
    CONCURRENT_LIMIT = "E7002"

    # 业务错误 (8xxx)
    GENERATION_FAILED = "E8000"
    EXPORT_FAILED = "E8001"
    BATCH_FAILED = "E8002"
    TASK_CANCELLED = "E8003"
    FILE_TOO_LARGE = "E8004"
    UNSUPPORTED_FORMAT = "E8005"


# 错误码到 HTTP 状态码的映射
ERROR_HTTP_STATUS = {
    ErrorCode.UNKNOWN_ERROR: HTTPStatus.INTERNAL_SERVER_ERROR,
    ErrorCode.INTERNAL_ERROR: HTTPStatus.INTERNAL_SERVER_ERROR,
    ErrorCode.SERVICE_UNAVAILABLE: HTTPStatus.SERVICE_UNAVAILABLE,
    ErrorCode.MAINTENANCE_MODE: HTTPStatus.SERVICE_UNAVAILABLE,

    ErrorCode.VALIDATION_ERROR: HTTPStatus.BAD_REQUEST,
    ErrorCode.MISSING_FIELD: HTTPStatus.BAD_REQUEST,
    ErrorCode.INVALID_FORMAT: HTTPStatus.BAD_REQUEST,
    ErrorCode.VALUE_TOO_LONG: HTTPStatus.BAD_REQUEST,
    ErrorCode.VALUE_TOO_SHORT: HTTPStatus.BAD_REQUEST,
    ErrorCode.VALUE_OUT_OF_RANGE: HTTPStatus.BAD_REQUEST,
    ErrorCode.INVALID_TYPE: HTTPStatus.BAD_REQUEST,
    ErrorCode.INVALID_URL: HTTPStatus.BAD_REQUEST,
    ErrorCode.BLOCKED_URL: HTTPStatus.BAD_REQUEST,

    ErrorCode.AUTH_REQUIRED: HTTPStatus.UNAUTHORIZED,
    ErrorCode.INVALID_API_KEY: HTTPStatus.UNAUTHORIZED,
    ErrorCode.EXPIRED_API_KEY: HTTPStatus.UNAUTHORIZED,
    ErrorCode.INVALID_TOKEN: HTTPStatus.UNAUTHORIZED,

    ErrorCode.PERMISSION_DENIED: HTTPStatus.FORBIDDEN,
    ErrorCode.RESOURCE_FORBIDDEN: HTTPStatus.FORBIDDEN,

    ErrorCode.RESOURCE_NOT_FOUND: HTTPStatus.NOT_FOUND,
    ErrorCode.FILE_NOT_FOUND: HTTPStatus.NOT_FOUND,
    ErrorCode.TEMPLATE_NOT_FOUND: HTTPStatus.NOT_FOUND,
    ErrorCode.TASK_NOT_FOUND: HTTPStatus.NOT_FOUND,

    ErrorCode.EXTERNAL_SERVICE_ERROR: HTTPStatus.BAD_GATEWAY,
    ErrorCode.AI_API_ERROR: HTTPStatus.BAD_GATEWAY,
    ErrorCode.AI_API_TIMEOUT: HTTPStatus.GATEWAY_TIMEOUT,
    ErrorCode.AI_API_RATE_LIMIT: HTTPStatus.TOO_MANY_REQUESTS,
    ErrorCode.AI_RESPONSE_INVALID: HTTPStatus.BAD_GATEWAY,
    ErrorCode.IMAGE_API_ERROR: HTTPStatus.BAD_GATEWAY,
    ErrorCode.CIRCUIT_BREAKER_OPEN: HTTPStatus.SERVICE_UNAVAILABLE,

    ErrorCode.RATE_LIMIT_EXCEEDED: HTTPStatus.TOO_MANY_REQUESTS,
    ErrorCode.QUOTA_EXCEEDED: HTTPStatus.TOO_MANY_REQUESTS,
    ErrorCode.CONCURRENT_LIMIT: HTTPStatus.TOO_MANY_REQUESTS,

    ErrorCode.GENERATION_FAILED: HTTPStatus.INTERNAL_SERVER_ERROR,
    ErrorCode.EXPORT_FAILED: HTTPStatus.INTERNAL_SERVER_ERROR,
    ErrorCode.BATCH_FAILED: HTTPStatus.INTERNAL_SERVER_ERROR,
    ErrorCode.TASK_CANCELLED: HTTPStatus.CONFLICT,
    ErrorCode.FILE_TOO_LARGE: HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
    ErrorCode.UNSUPPORTED_FORMAT: HTTPStatus.UNSUPPORTED_MEDIA_TYPE,
}

# 用户友好的错误消息
ERROR_MESSAGES = {
    ErrorCode.UNKNOWN_ERROR: "发生未知错误，请稍后重试",
    ErrorCode.INTERNAL_ERROR: "服务器内部错误，请稍后重试",
    ErrorCode.SERVICE_UNAVAILABLE: "服务暂时不可用，请稍后重试",
    ErrorCode.MAINTENANCE_MODE: "系统维护中，请稍后重试",

    ErrorCode.VALIDATION_ERROR: "输入数据验证失败",
    ErrorCode.MISSING_FIELD: "缺少必填字段",
    ErrorCode.INVALID_FORMAT: "数据格式无效",
    ErrorCode.VALUE_TOO_LONG: "输入内容过长",
    ErrorCode.VALUE_TOO_SHORT: "输入内容过短",
    ErrorCode.VALUE_OUT_OF_RANGE: "数值超出允许范围",
    ErrorCode.INVALID_TYPE: "数据类型错误",
    ErrorCode.INVALID_URL: "URL 格式无效",
    ErrorCode.BLOCKED_URL: "该 URL 不允许访问",

    ErrorCode.AUTH_REQUIRED: "需要认证",
    ErrorCode.INVALID_API_KEY: "API Key 无效",
    ErrorCode.EXPIRED_API_KEY: "API Key 已过期",
    ErrorCode.INVALID_TOKEN: "访问令牌无效",

    ErrorCode.PERMISSION_DENIED: "没有权限执行此操作",
    ErrorCode.RESOURCE_FORBIDDEN: "无权访问此资源",

    ErrorCode.RESOURCE_NOT_FOUND: "请求的资源不存在",
    ErrorCode.FILE_NOT_FOUND: "文件不存在",
    ErrorCode.TEMPLATE_NOT_FOUND: "模板不存在",
    ErrorCode.TASK_NOT_FOUND: "任务不存在",

    ErrorCode.EXTERNAL_SERVICE_ERROR: "外部服务错误",
    ErrorCode.AI_API_ERROR: "AI 服务调用失败",
    ErrorCode.AI_API_TIMEOUT: "AI 服务响应超时",
    ErrorCode.AI_API_RATE_LIMIT: "AI 服务请求过于频繁",
    ErrorCode.AI_RESPONSE_INVALID: "AI 返回的数据格式无效",
    ErrorCode.IMAGE_API_ERROR: "图片服务调用失败",
    ErrorCode.CIRCUIT_BREAKER_OPEN: "服务暂时不可用，请稍后重试",

    ErrorCode.RATE_LIMIT_EXCEEDED: "请求过于频繁，请稍后重试",
    ErrorCode.QUOTA_EXCEEDED: "已超出使用配额",
    ErrorCode.CONCURRENT_LIMIT: "并发请求数超出限制",

    ErrorCode.GENERATION_FAILED: "PPT 生成失败",
    ErrorCode.EXPORT_FAILED: "导出失败",
    ErrorCode.BATCH_FAILED: "批量处理失败",
    ErrorCode.TASK_CANCELLED: "任务已取消",
    ErrorCode.FILE_TOO_LARGE: "文件过大",
    ErrorCode.UNSUPPORTED_FORMAT: "不支持的文件格式",
}


@dataclass
class AppError(Exception):
    """应用错误

    统一的错误类，包含错误码、消息和详情。

    用法:
        raise AppError(
            code=ErrorCode.MISSING_FIELD,
            message="主题不能为空",
            details={"field": "topic"},
        )
    """
    code: ErrorCode
    message: str = ""
    details: Optional[Dict[str, Any]] = None
    original_error: Optional[Exception] = None

    def __post_init__(self):
        if not self.message:
            self.message = ERROR_MESSAGES.get(self.code, "未知错误")
        super().__init__(self.message)

    @property
    def http_status(self) -> int:
        """获取对应的 HTTP 状态码"""
        return ERROR_HTTP_STATUS.get(self.code, HTTPStatus.INTERNAL_SERVER_ERROR).value

    @property
    def category(self) -> ErrorCategory:
        """获取错误分类"""
        code_prefix = self.code.value[1]  # E1xxx -> 1
        category_map = {
            "1": ErrorCategory.INTERNAL,
            "2": ErrorCategory.VALIDATION,
            "3": ErrorCategory.AUTHENTICATION,
            "4": ErrorCategory.AUTHORIZATION,
            "5": ErrorCategory.RESOURCE,
            "6": ErrorCategory.EXTERNAL,
            "7": ErrorCategory.RATE_LIMIT,
            "8": ErrorCategory.BUSINESS,
        }
        return category_map.get(code_prefix, ErrorCategory.INTERNAL)

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        result = {
            "error": True,
            "code": self.code.value,
            "message": self.message,
            "category": self.category.value,
        }
        if self.details:
            result["details"] = self.details
        return result

    def to_response(self) -> tuple:
        """转换为 Flask 响应"""
        from flask import jsonify
        return jsonify(self.to_dict()), self.http_status


# 便捷异常类
class ValidationError(AppError):
    """验证错误"""
    def __init__(self, message: str, field: str = None, **kwargs):
        details = {"field": field} if field else None
        super().__init__(
            code=ErrorCode.VALIDATION_ERROR,
            message=message,
            details=details,
            **kwargs,
        )


class AuthenticationError(AppError):
    """认证错误"""
    def __init__(self, message: str = None, **kwargs):
        super().__init__(
            code=ErrorCode.AUTH_REQUIRED,
            message=message,
            **kwargs,
        )


class NotFoundError(AppError):
    """资源不存在错误"""
    def __init__(self, resource: str, resource_id: str = None, **kwargs):
        message = f"{resource}不存在"
        if resource_id:
            message = f"{resource} '{resource_id}' 不存在"
        super().__init__(
            code=ErrorCode.RESOURCE_NOT_FOUND,
            message=message,
            details={"resource": resource, "id": resource_id},
            **kwargs,
        )


class ExternalServiceError(AppError):
    """外部服务错误"""
    def __init__(self, service: str, message: str = None, **kwargs):
        super().__init__(
            code=ErrorCode.EXTERNAL_SERVICE_ERROR,
            message=message or f"{service}服务调用失败",
            details={"service": service},
            **kwargs,
        )


class RateLimitError(AppError):
    """限流错误"""
    def __init__(self, retry_after: int = None, **kwargs):
        details = {"retry_after": retry_after} if retry_after else None
        super().__init__(
            code=ErrorCode.RATE_LIMIT_EXCEEDED,
            details=details,
            **kwargs,
        )


def handle_app_error(error: AppError):
    """Flask 错误处理器"""
    from utils.logger import get_logger
    logger = get_logger("error")

    # 记录错误日志
    if error.category == ErrorCategory.INTERNAL:
        logger.error(f"[{error.code.value}] {error.message}", exc_info=error.original_error)
    else:
        logger.warning(f"[{error.code.value}] {error.message}")

    return error.to_response()


def register_error_handlers(app):
    """注册 Flask 错误处理器"""
    app.register_error_handler(AppError, handle_app_error)

    @app.errorhandler(404)
    def not_found(error):
        return AppError(code=ErrorCode.RESOURCE_NOT_FOUND).to_response()

    @app.errorhandler(500)
    def internal_error(error):
        return AppError(code=ErrorCode.INTERNAL_ERROR).to_response()
