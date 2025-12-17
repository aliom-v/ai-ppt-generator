"""API 版本控制模块

支持通过 URL 前缀或请求头进行 API 版本控制。
"""
from functools import wraps
from typing import Callable, Dict, List, Optional, Any

from flask import Flask, Blueprint, request, jsonify


class APIVersion:
    """API 版本管理器

    用法:
        app = Flask(__name__)
        api = APIVersion(app, default_version="v1")

        # 注册版本化的蓝图
        api.register_version("v1", v1_blueprint)
        api.register_version("v2", v2_blueprint)

        # 或者使用装饰器
        @api.route("/users", versions=["v1", "v2"])
        def get_users():
            return {"users": []}

        @api.route("/users", versions=["v2"])
        def get_users_v2():
            return {"users": [], "pagination": {}}
    """

    def __init__(
        self,
        app: Flask = None,
        default_version: str = "v1",
        version_header: str = "X-API-Version",
        supported_versions: List[str] = None,
    ):
        self.default_version = default_version
        self.version_header = version_header
        self.supported_versions = supported_versions or ["v1"]
        self._version_blueprints: Dict[str, Blueprint] = {}
        self._app = app

        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        self._app = app

        # 添加版本信息到响应头
        @app.after_request
        def add_version_header(response):
            version = getattr(request, "api_version", self.default_version)
            response.headers["X-API-Version"] = version
            return response

        # 注册版本端点
        @app.route("/api/versions")
        def list_versions():
            return jsonify({
                "versions": self.supported_versions,
                "default": self.default_version,
                "current": getattr(request, "api_version", self.default_version),
            })

    def register_version(self, version: str, blueprint: Blueprint):
        """注册版本化的蓝图

        Args:
            version: 版本号（如 "v1", "v2"）
            blueprint: Flask Blueprint
        """
        if version not in self.supported_versions:
            self.supported_versions.append(version)

        self._version_blueprints[version] = blueprint

        if self._app:
            self._app.register_blueprint(blueprint, url_prefix=f"/api/{version}")

    def get_version(self) -> str:
        """获取当前请求的 API 版本

        优先级：URL 前缀 > 请求头 > 默认版本
        """
        # 从 URL 获取
        path = request.path
        for version in self.supported_versions:
            if path.startswith(f"/api/{version}/") or path == f"/api/{version}":
                return version

        # 从请求头获取
        header_version = request.headers.get(self.version_header)
        if header_version and header_version in self.supported_versions:
            return header_version

        return self.default_version

    def version_required(self, *versions: str):
        """版本要求装饰器

        限制端点只在指定版本可用。

        用法:
            @app.route("/api/new-feature")
            @api.version_required("v2", "v3")
            def new_feature():
                ...
        """
        def decorator(f: Callable) -> Callable:
            @wraps(f)
            def wrapper(*args, **kwargs):
                current = self.get_version()
                request.api_version = current

                if current not in versions:
                    return jsonify({
                        "error": f"此功能在 API {current} 中不可用",
                        "available_versions": list(versions),
                    }), 404

                return f(*args, **kwargs)
            return wrapper
        return decorator

    def deprecated(self, message: str = None, sunset_version: str = None):
        """标记端点为已弃用

        用法:
            @app.route("/api/v1/old-endpoint")
            @api.deprecated(message="请使用 /api/v2/new-endpoint", sunset_version="v3")
            def old_endpoint():
                ...
        """
        def decorator(f: Callable) -> Callable:
            @wraps(f)
            def wrapper(*args, **kwargs):
                response = f(*args, **kwargs)

                # 如果是元组，提取 response 对象
                if isinstance(response, tuple):
                    resp, *rest = response
                else:
                    resp = response
                    rest = []

                # 添加弃用警告头
                from flask import make_response
                resp = make_response(resp)
                resp.headers["Deprecation"] = "true"

                if message:
                    resp.headers["X-Deprecation-Notice"] = message
                if sunset_version:
                    resp.headers["Sunset"] = f"API {sunset_version}"

                if rest:
                    return (resp, *rest)
                return resp
            return wrapper
        return decorator


def create_versioned_blueprint(name: str, version: str) -> Blueprint:
    """创建版本化的蓝图

    Args:
        name: 蓝图名称
        version: 版本号

    Returns:
        配置好的 Blueprint
    """
    bp = Blueprint(f"{name}_{version}", __name__)

    @bp.before_request
    def set_version():
        request.api_version = version

    return bp


# 便捷函数：创建 v1 和 v2 蓝图
def create_api_blueprints() -> Dict[str, Blueprint]:
    """创建标准的 API 版本蓝图

    Returns:
        {"v1": Blueprint, "v2": Blueprint}
    """
    return {
        "v1": create_versioned_blueprint("api", "v1"),
        "v2": create_versioned_blueprint("api", "v2"),
    }
