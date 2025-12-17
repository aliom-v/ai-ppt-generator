"""HTTP 中间件 - CORS 支持和 Gzip 压缩"""
import gzip
import os
from io import BytesIO
from functools import wraps
from typing import List, Optional

from flask import Flask, request, make_response


class CORSMiddleware:
    """CORS 跨域支持中间件

    用法:
        app = Flask(__name__)
        CORSMiddleware(app, origins=["http://localhost:3000"])
    """

    def __init__(
        self,
        app: Flask = None,
        origins: Optional[List[str]] = None,
        methods: Optional[List[str]] = None,
        headers: Optional[List[str]] = None,
        expose_headers: Optional[List[str]] = None,
        max_age: int = 86400,
    ):
        self.origins = origins or ["*"]
        self.methods = methods or ["GET", "POST", "PUT", "DELETE", "OPTIONS"]
        self.headers = headers or ["Content-Type", "Authorization", "X-Request-ID"]
        self.expose_headers = expose_headers or ["X-Request-ID", "X-Response-Time"]
        self.max_age = max_age

        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        app.after_request(self._add_cors_headers)

    def _add_cors_headers(self, response):
        origin = request.headers.get("Origin", "")

        # 检查是否允许该来源
        if "*" in self.origins or origin in self.origins:
            response.headers["Access-Control-Allow-Origin"] = origin or "*"
        elif self.origins:
            response.headers["Access-Control-Allow-Origin"] = self.origins[0]

        response.headers["Access-Control-Allow-Methods"] = ", ".join(self.methods)
        response.headers["Access-Control-Allow-Headers"] = ", ".join(self.headers)
        response.headers["Access-Control-Expose-Headers"] = ", ".join(self.expose_headers)
        response.headers["Access-Control-Max-Age"] = str(self.max_age)

        # 处理预检请求
        if request.method == "OPTIONS":
            response.status_code = 204

        return response


class GzipMiddleware:
    """Gzip 压缩中间件

    自动压缩大于阈值的响应。

    用法:
        app = Flask(__name__)
        GzipMiddleware(app, min_size=500)
    """

    def __init__(
        self,
        app: Flask = None,
        min_size: int = 500,
        compression_level: int = 6,
    ):
        self.min_size = min_size
        self.compression_level = compression_level
        self.compressible_types = {
            "text/html",
            "text/css",
            "text/javascript",
            "application/javascript",
            "application/json",
            "application/xml",
            "text/xml",
            "text/plain",
        }

        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        app.after_request(self._compress_response)

    def _compress_response(self, response):
        # 检查客户端是否支持 gzip
        if "gzip" not in request.headers.get("Accept-Encoding", ""):
            return response

        # 检查响应是否已经压缩
        if response.headers.get("Content-Encoding"):
            return response

        # 检查内容类型是否可压缩
        content_type = response.headers.get("Content-Type", "")
        if not any(ct in content_type for ct in self.compressible_types):
            return response

        # 检查响应大小
        data = response.get_data()
        if len(data) < self.min_size:
            return response

        # 压缩数据
        compressed = gzip.compress(data, compresslevel=self.compression_level)

        # 只有压缩后更小才使用压缩版本
        if len(compressed) >= len(data):
            return response

        response.set_data(compressed)
        response.headers["Content-Encoding"] = "gzip"
        response.headers["Content-Length"] = len(compressed)
        response.headers["Vary"] = "Accept-Encoding"

        return response


def setup_http_middleware(
    app: Flask,
    enable_cors: bool = True,
    enable_gzip: bool = True,
    cors_origins: Optional[List[str]] = None,
):
    """设置 HTTP 中间件

    Args:
        app: Flask 应用
        enable_cors: 是否启用 CORS
        enable_gzip: 是否启用 Gzip 压缩
        cors_origins: CORS 允许的来源列表
    """
    if enable_cors:
        origins = cors_origins or os.getenv("CORS_ORIGINS", "*").split(",")
        CORSMiddleware(app, origins=origins)

    if enable_gzip:
        GzipMiddleware(app)
