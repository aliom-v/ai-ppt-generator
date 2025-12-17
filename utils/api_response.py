"""统一 API 响应格式"""
from typing import Any, Optional, Dict
from flask import jsonify, Response
from utils.request_context import get_request_id


class APIResponse:
    """统一 API 响应构建器

    所有 API 响应都应该使用这个类来构建，确保格式一致。

    成功响应格式:
    {
        "success": true,
        "data": {...},
        "request_id": "abc12345"
    }

    错误响应格式:
    {
        "success": false,
        "error": {
            "code": "ERROR_CODE",
            "message": "错误描述"
        },
        "request_id": "abc12345"
    }
    """

    @staticmethod
    def success(
        data: Any = None,
        message: Optional[str] = None,
        status_code: int = 200,
        **extra
    ) -> tuple:
        """构建成功响应

        Args:
            data: 响应数据
            message: 可选的成功消息
            status_code: HTTP 状态码
            **extra: 额外字段

        Returns:
            (Response, status_code) 元组
        """
        response = {
            "success": True,
            "request_id": get_request_id(),
        }

        if data is not None:
            response["data"] = data

        if message:
            response["message"] = message

        # 添加额外字段
        response.update(extra)

        return jsonify(response), status_code

    @staticmethod
    def error(
        message: str,
        code: str = "ERROR",
        status_code: int = 400,
        details: Optional[Dict] = None
    ) -> tuple:
        """构建错误响应

        Args:
            message: 错误消息
            code: 错误代码
            status_code: HTTP 状态码
            details: 错误详情

        Returns:
            (Response, status_code) 元组
        """
        error_info = {
            "code": code,
            "message": message,
        }

        if details:
            error_info["details"] = details

        response = {
            "success": False,
            "error": error_info,
            "request_id": get_request_id(),
        }

        return jsonify(response), status_code

    @staticmethod
    def validation_error(message: str, field: Optional[str] = None) -> tuple:
        """构建验证错误响应"""
        details = {"field": field} if field else None
        return APIResponse.error(
            message=message,
            code="VALIDATION_ERROR",
            status_code=400,
            details=details
        )

    @staticmethod
    def not_found(message: str = "资源不存在") -> tuple:
        """构建 404 响应"""
        return APIResponse.error(
            message=message,
            code="NOT_FOUND",
            status_code=404
        )

    @staticmethod
    def unauthorized(message: str = "未授权访问") -> tuple:
        """构建 401 响应"""
        return APIResponse.error(
            message=message,
            code="UNAUTHORIZED",
            status_code=401
        )

    @staticmethod
    def rate_limited(message: str, retry_after: int) -> tuple:
        """构建限流响应"""
        response, _ = APIResponse.error(
            message=message,
            code="RATE_LIMITED",
            status_code=429,
            details={"retry_after": retry_after}
        )
        response.headers['Retry-After'] = str(retry_after)
        return response, 429

    @staticmethod
    def server_error(message: str = "服务器内部错误") -> tuple:
        """构建 500 响应"""
        return APIResponse.error(
            message=message,
            code="SERVER_ERROR",
            status_code=500
        )


# 便捷函数别名
api_success = APIResponse.success
api_error = APIResponse.error
