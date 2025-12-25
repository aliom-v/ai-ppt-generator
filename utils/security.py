"""安全中间件 - 添加安全响应头和安全策略"""
import os
import secrets
from functools import wraps
from typing import Optional, List, Dict

from flask import Flask, request, g, Response


class SecurityHeadersMiddleware:
    """安全响应头中间件

    添加常见的安全响应头，防止 XSS、点击劫持等攻击。

    用法:
        app = Flask(__name__)
        SecurityHeadersMiddleware(app)
    """

    def __init__(
        self,
        app: Flask = None,
        csp_policy: Optional[str] = None,
        hsts_max_age: int = 31536000,
        frame_options: str = "DENY",
        content_type_options: bool = True,
        xss_protection: bool = True,
        referrer_policy: str = "strict-origin-when-cross-origin",
    ):
        self.csp_policy = csp_policy
        self.hsts_max_age = hsts_max_age
        self.frame_options = frame_options
        self.content_type_options = content_type_options
        self.xss_protection = xss_protection
        self.referrer_policy = referrer_policy

        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        app.after_request(self._add_security_headers)

    def _add_security_headers(self, response: Response) -> Response:
        # X-Frame-Options: 防止点击劫持
        if self.frame_options:
            response.headers["X-Frame-Options"] = self.frame_options

        # X-Content-Type-Options: 防止 MIME 类型嗅探
        if self.content_type_options:
            response.headers["X-Content-Type-Options"] = "nosniff"

        # X-XSS-Protection: 启用浏览器 XSS 过滤
        if self.xss_protection:
            response.headers["X-XSS-Protection"] = "1; mode=block"

        # Referrer-Policy: 控制 Referer 头
        if self.referrer_policy:
            response.headers["Referrer-Policy"] = self.referrer_policy

        # Content-Security-Policy: 内容安全策略
        if self.csp_policy:
            response.headers["Content-Security-Policy"] = self.csp_policy

        # Strict-Transport-Security: 强制 HTTPS
        if self.hsts_max_age > 0 and request.is_secure:
            response.headers["Strict-Transport-Security"] = (
                f"max-age={self.hsts_max_age}; includeSubDomains"
            )

        # Permissions-Policy: 限制 API 访问权限
        response.headers["Permissions-Policy"] = (
            "accelerometer=(), camera=(), geolocation=(), gyroscope=(), "
            "magnetometer=(), microphone=(), payment=(), usb=()"
        )

        # Cache-Control: API 响应不缓存
        if request.path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
            response.headers["Pragma"] = "no-cache"

        return response


class CSRFProtection:
    """CSRF 保护

    为表单请求添加 CSRF Token 验证。
    Token 存储在 session 中，确保跨请求有效。

    用法:
        app = Flask(__name__)
        csrf = CSRFProtection(app)

        @app.route('/form', methods=['POST'])
        @csrf.protect
        def handle_form():
            ...
    """

    def __init__(self, app: Flask = None, token_length: int = 32):
        self.token_length = token_length
        self._exempt_views: List[str] = []

        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        # 确保 session 可用
        if not app.secret_key:
            import warnings
            warnings.warn("Flask secret_key 未设置，CSRF 保护可能无法正常工作")

        @app.context_processor
        def csrf_token_processor():
            return {"csrf_token": self._get_or_create_token}

    def _get_or_create_token(self) -> str:
        """获取或创建 CSRF Token（存储在 session 中）"""
        from flask import session
        if "csrf_token" not in session:
            session["csrf_token"] = secrets.token_hex(self.token_length)
        return session["csrf_token"]

    def exempt(self, view_func):
        """标记视图函数免于 CSRF 保护"""
        self._exempt_views.append(view_func.__name__)
        return view_func

    def protect(self, f):
        """CSRF 保护装饰器"""
        @wraps(f)
        def wrapper(*args, **kwargs):
            if f.__name__ in self._exempt_views:
                return f(*args, **kwargs)

            if request.method in ("POST", "PUT", "DELETE", "PATCH"):
                from flask import session
                token = (
                    request.form.get("csrf_token")
                    or request.headers.get("X-CSRF-Token")
                    or (request.json.get("csrf_token") if request.is_json and request.json else None)
                )

                session_token = session.get("csrf_token")
                if not token or not session_token or not secrets.compare_digest(token, session_token):
                    from flask import jsonify
                    return jsonify({"error": "CSRF Token 无效"}), 403

            return f(*args, **kwargs)
        return wrapper


class RateLimitByIP:
    """基于 IP 的速率限制

    更细粒度的速率限制，支持不同端点不同限制。
    """

    def __init__(
        self,
        default_limit: int = 100,
        default_window: int = 60,
    ):
        self.default_limit = default_limit
        self.default_window = default_window
        self._counters: Dict[str, Dict] = {}
        self._limits: Dict[str, tuple] = {}

    def limit(self, requests: int, per_seconds: int):
        """限制装饰器

        Args:
            requests: 允许的请求数
            per_seconds: 时间窗口（秒）
        """
        def decorator(f):
            endpoint = f.__name__
            self._limits[endpoint] = (requests, per_seconds)

            @wraps(f)
            def wrapper(*args, **kwargs):
                import time

                ip = request.remote_addr or "unknown"
                key = f"{endpoint}:{ip}"
                now = time.time()

                limit, window = self._limits.get(endpoint, (self.default_limit, self.default_window))

                # 获取或创建计数器
                if key not in self._counters:
                    self._counters[key] = {"count": 0, "reset_at": now + window}

                counter = self._counters[key]

                # 检查是否需要重置
                if now >= counter["reset_at"]:
                    counter["count"] = 0
                    counter["reset_at"] = now + window

                # 检查是否超限
                if counter["count"] >= limit:
                    from flask import jsonify
                    remaining = int(counter["reset_at"] - now)
                    response = jsonify({
                        "error": f"请求过于频繁，请 {remaining} 秒后重试",
                        "retry_after": remaining,
                    })
                    response.headers["Retry-After"] = str(remaining)
                    return response, 429

                counter["count"] += 1
                return f(*args, **kwargs)

            return wrapper
        return decorator

    def cleanup(self):
        """清理过期的计数器"""
        import time
        now = time.time()
        expired = [k for k, v in self._counters.items() if now >= v["reset_at"]]
        for k in expired:
            del self._counters[k]


def setup_security(
    app: Flask,
    enable_headers: bool = True,
    enable_csrf: bool = False,
    csp_policy: Optional[str] = None,
):
    """设置安全中间件

    Args:
        app: Flask 应用
        enable_headers: 是否启用安全响应头
        enable_csrf: 是否启用 CSRF 保护（API 模式通常不需要）
        csp_policy: 自定义 CSP 策略
    """
    if enable_headers:
        # 默认 CSP 策略（适合 API 和简单前端）
        default_csp = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' cdn.jsdelivr.net; "
            "style-src 'self' 'unsafe-inline' cdn.jsdelivr.net; "
            "img-src 'self' data: https:; "
            "font-src 'self' cdn.jsdelivr.net; "
            "connect-src 'self'; "
            "frame-ancestors 'none';"
        )
        SecurityHeadersMiddleware(app, csp_policy=csp_policy or default_csp)

    if enable_csrf:
        return CSRFProtection(app)

    return None
