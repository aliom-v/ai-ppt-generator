"""请求上下文管理 - 请求 ID 追踪"""
import uuid
import threading
from typing import Optional
from functools import wraps

# 线程本地存储，用于存储请求上下文
_request_context = threading.local()


def generate_request_id() -> str:
    """生成唯一请求 ID"""
    return str(uuid.uuid4())[:8]


def get_request_id() -> Optional[str]:
    """获取当前请求 ID"""
    return getattr(_request_context, 'request_id', None)


def set_request_id(request_id: str) -> None:
    """设置当前请求 ID"""
    _request_context.request_id = request_id


def clear_request_id() -> None:
    """清除当前请求 ID"""
    if hasattr(_request_context, 'request_id'):
        delattr(_request_context, 'request_id')


class RequestContextMiddleware:
    """Flask 请求上下文中间件

    自动为每个请求生成唯一 ID，并添加到响应头中。

    用法:
        app = Flask(__name__)
        RequestContextMiddleware(app)
    """

    def __init__(self, app=None):
        self.app = app
        if app is not None:
            self.init_app(app)

    def init_app(self, app):
        """初始化 Flask 应用"""
        app.before_request(self._before_request)
        app.after_request(self._after_request)
        app.teardown_request(self._teardown_request)

    def _before_request(self):
        """请求前：生成请求 ID"""
        from flask import request
        # 优先使用客户端传入的请求 ID
        request_id = request.headers.get('X-Request-ID')
        if not request_id:
            request_id = generate_request_id()
        set_request_id(request_id)

    def _after_request(self, response):
        """请求后：添加请求 ID 到响应头"""
        request_id = get_request_id()
        if request_id:
            response.headers['X-Request-ID'] = request_id
        return response

    def _teardown_request(self, exception=None):
        """请求结束：清理上下文"""
        clear_request_id()


def with_request_id(f):
    """装饰器：确保函数执行时有请求 ID 上下文

    用于非 Flask 请求上下文中的函数（如后台任务）。
    """
    @wraps(f)
    def wrapper(*args, **kwargs):
        if get_request_id() is None:
            set_request_id(generate_request_id())
        try:
            return f(*args, **kwargs)
        finally:
            pass  # 不在这里清理，由调用方决定
    return wrapper
