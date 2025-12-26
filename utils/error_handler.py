"""全局异常处理中间件"""
from functools import wraps
from typing import Callable
from flask import Flask

from utils.api_response import APIResponse
from utils.logger import get_logger
from utils.errors import AppError, ErrorCode, ErrorCategory

logger = get_logger("error_handler")


class ErrorHandlerMiddleware:
    """Flask 全局异常处理中间件

    捕获所有未处理的异常，返回统一格式的错误响应。

    用法:
        app = Flask(__name__)
        ErrorHandlerMiddleware(app)
    """

    def __init__(self, app: Flask = None):
        self.app = app
        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        """初始化 Flask 应用"""
        # 注册 AppError 处理器
        app.register_error_handler(AppError, self._handle_app_error)

        # 注册 HTTP 错误处理器
        app.register_error_handler(400, self._handle_400)
        app.register_error_handler(401, self._handle_401)
        app.register_error_handler(403, self._handle_403)
        app.register_error_handler(404, self._handle_404)
        app.register_error_handler(405, self._handle_405)
        app.register_error_handler(413, self._handle_413)
        app.register_error_handler(429, self._handle_429)
        app.register_error_handler(500, self._handle_500)
        app.register_error_handler(Exception, self._handle_exception)

    def _handle_app_error(self, error: AppError):
        """处理 AppError"""
        if error.category == ErrorCategory.INTERNAL:
            logger.error(f"[{error.code.value}] {error.message}", exc_info=error.original_error)
        else:
            logger.warning(f"[{error.code.value}] {error.message}")
        return error.to_response()

    def _handle_400(self, error):
        """处理 400 错误"""
        return APIResponse.error(
            message=str(error.description) if hasattr(error, 'description') else "请求参数错误",
            code="BAD_REQUEST",
            status_code=400
        )

    def _handle_401(self, error):
        """处理 401 错误"""
        return APIResponse.unauthorized()

    def _handle_403(self, error):
        """处理 403 错误"""
        return APIResponse.error(
            message="禁止访问",
            code="FORBIDDEN",
            status_code=403
        )

    def _handle_404(self, error):
        """处理 404 错误"""
        return APIResponse.not_found()

    def _handle_405(self, error):
        """处理 405 错误"""
        return APIResponse.error(
            message="请求方法不允许",
            code="METHOD_NOT_ALLOWED",
            status_code=405
        )

    def _handle_413(self, error):
        """处理 413 错误（文件过大）"""
        return APIResponse.error(
            message="上传文件过大",
            code="FILE_TOO_LARGE",
            status_code=413
        )

    def _handle_429(self, error):
        """处理 429 错误（限流）"""
        return APIResponse.error(
            message="请求过于频繁，请稍后再试",
            code="RATE_LIMITED",
            status_code=429
        )

    def _handle_500(self, error):
        """处理 500 错误"""
        logger.error(f"服务器错误: {error}", exc_info=True)
        return APIResponse.server_error()

    def _handle_exception(self, error):
        """处理所有未捕获的异常"""
        from utils.request_context import get_request_id

        request_id = get_request_id()
        logger.error(f"未处理异常 [request_id={request_id}]: {error}", exc_info=True)

        # 在生产环境不暴露详细错误信息
        return APIResponse.error(
            message="服务器内部错误，请稍后重试",
            code="INTERNAL_ERROR",
            status_code=500,
            details={"request_id": request_id}
        )


def handle_exceptions(f: Callable) -> Callable:
    """装饰器：捕获函数异常并返回统一错误响应

    用于需要单独处理异常的路由。

    用法:
        @app.route('/api/test')
        @handle_exceptions
        def test():
            ...
    """
    @wraps(f)
    def wrapper(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except ValueError as e:
            return APIResponse.validation_error(str(e))
        except PermissionError as e:
            return APIResponse.error(str(e), code="PERMISSION_DENIED", status_code=403)
        except FileNotFoundError as e:
            return APIResponse.not_found(str(e))
        except Exception as e:
            logger.error(f"路由异常: {e}", exc_info=True)
            return APIResponse.server_error()
    return wrapper
