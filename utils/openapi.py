"""OpenAPI 文档生成模块

自动生成 API 文档，支持 Swagger UI 展示。
"""
import json
from typing import Any, Dict, List, Optional
from dataclasses import dataclass, field, asdict
from functools import wraps

from flask import Flask, Blueprint, jsonify, render_template_string


# OpenAPI 3.0 数据结构
@dataclass
class APIParameter:
    """API 参数"""
    name: str
    location: str  # query, path, header, cookie
    description: str = ""
    required: bool = False
    schema: Dict = field(default_factory=lambda: {"type": "string"})


@dataclass
class APIRequestBody:
    """请求体"""
    description: str = ""
    required: bool = True
    content_type: str = "application/json"
    schema: Dict = field(default_factory=dict)
    example: Any = None


@dataclass
class APIResponse:
    """API 响应"""
    status_code: int
    description: str
    content_type: str = "application/json"
    schema: Dict = field(default_factory=dict)
    example: Any = None


@dataclass
class APIEndpoint:
    """API 端点"""
    path: str
    method: str
    summary: str
    description: str = ""
    tags: List[str] = field(default_factory=list)
    parameters: List[APIParameter] = field(default_factory=list)
    request_body: Optional[APIRequestBody] = None
    responses: List[APIResponse] = field(default_factory=list)
    deprecated: bool = False


class OpenAPIGenerator:
    """OpenAPI 文档生成器

    用法:
        app = Flask(__name__)
        openapi = OpenAPIGenerator(app, title="AI PPT Generator API")

        @app.route('/api/users')
        @openapi.document(
            summary="获取用户列表",
            tags=["用户"],
            responses=[
                APIResponse(200, "成功", example={"users": []})
            ]
        )
        def get_users():
            return {"users": []}
    """

    def __init__(
        self,
        app: Flask = None,
        title: str = "API",
        version: str = "1.0.0",
        description: str = "",
        servers: List[Dict] = None,
    ):
        self.title = title
        self.version = version
        self.description = description
        self.servers = servers or [{"url": "/", "description": "当前服务器"}]
        self._endpoints: List[APIEndpoint] = []
        self._tags: Dict[str, str] = {}

        if app:
            self.init_app(app)

    def init_app(self, app: Flask):
        """初始化 Flask 应用"""
        # 注册 OpenAPI JSON 端点
        @app.route("/openapi.json")
        def openapi_json():
            return jsonify(self.generate_spec())

        # 注册 Swagger UI
        @app.route("/docs")
        def swagger_ui():
            return render_template_string(SWAGGER_UI_TEMPLATE, title=self.title)

        # 注册 ReDoc
        @app.route("/redoc")
        def redoc():
            return render_template_string(REDOC_TEMPLATE, title=self.title)

    def add_tag(self, name: str, description: str = ""):
        """添加标签"""
        self._tags[name] = description

    def document(
        self,
        summary: str,
        description: str = "",
        tags: List[str] = None,
        parameters: List[APIParameter] = None,
        request_body: APIRequestBody = None,
        responses: List[APIResponse] = None,
        deprecated: bool = False,
    ):
        """文档装饰器"""
        def decorator(f):
            # 存储文档信息到函数属性
            f._openapi = {
                "summary": summary,
                "description": description or f.__doc__ or "",
                "tags": tags or [],
                "parameters": parameters or [],
                "request_body": request_body,
                "responses": responses or [APIResponse(200, "成功")],
                "deprecated": deprecated,
            }

            @wraps(f)
            def wrapper(*args, **kwargs):
                return f(*args, **kwargs)

            return wrapper
        return decorator

    def register_endpoint(self, endpoint: APIEndpoint):
        """注册端点"""
        self._endpoints.append(endpoint)

    def generate_spec(self) -> Dict:
        """生成 OpenAPI 规范"""
        spec = {
            "openapi": "3.0.3",
            "info": {
                "title": self.title,
                "version": self.version,
                "description": self.description,
            },
            "servers": self.servers,
            "paths": {},
            "components": {
                "schemas": self._generate_schemas(),
                "securitySchemes": {
                    "ApiKeyAuth": {
                        "type": "apiKey",
                        "in": "header",
                        "name": "X-API-Key",
                    }
                }
            },
        }

        # 添加标签
        if self._tags:
            spec["tags"] = [
                {"name": name, "description": desc}
                for name, desc in self._tags.items()
            ]

        # 添加路径
        for endpoint in self._endpoints:
            if endpoint.path not in spec["paths"]:
                spec["paths"][endpoint.path] = {}

            operation = {
                "summary": endpoint.summary,
                "description": endpoint.description,
                "responses": {},
            }

            if endpoint.tags:
                operation["tags"] = endpoint.tags

            if endpoint.deprecated:
                operation["deprecated"] = True

            if endpoint.parameters:
                operation["parameters"] = [
                    {
                        "name": p.name,
                        "in": p.location,
                        "description": p.description,
                        "required": p.required,
                        "schema": p.schema,
                    }
                    for p in endpoint.parameters
                ]

            if endpoint.request_body:
                rb = endpoint.request_body
                operation["requestBody"] = {
                    "description": rb.description,
                    "required": rb.required,
                    "content": {
                        rb.content_type: {
                            "schema": rb.schema,
                        }
                    }
                }
                if rb.example:
                    operation["requestBody"]["content"][rb.content_type]["example"] = rb.example

            for resp in endpoint.responses:
                operation["responses"][str(resp.status_code)] = {
                    "description": resp.description,
                }
                if resp.schema or resp.example:
                    operation["responses"][str(resp.status_code)]["content"] = {
                        resp.content_type: {}
                    }
                    if resp.schema:
                        operation["responses"][str(resp.status_code)]["content"][resp.content_type]["schema"] = resp.schema
                    if resp.example:
                        operation["responses"][str(resp.status_code)]["content"][resp.content_type]["example"] = resp.example

            spec["paths"][endpoint.path][endpoint.method.lower()] = operation

        return spec

    def _generate_schemas(self) -> Dict:
        """生成通用 Schema"""
        return {
            "Error": {
                "type": "object",
                "properties": {
                    "error": {"type": "string"},
                    "code": {"type": "string"},
                },
            },
            "Success": {
                "type": "object",
                "properties": {
                    "success": {"type": "boolean"},
                    "message": {"type": "string"},
                },
            },
            "GenerateRequest": {
                "type": "object",
                "required": ["topic", "api_key"],
                "properties": {
                    "topic": {"type": "string", "description": "PPT 主题", "maxLength": 500},
                    "audience": {"type": "string", "description": "目标受众", "default": "通用受众"},
                    "page_count": {"type": "integer", "description": "页数", "minimum": 1, "maximum": 100, "default": 5},
                    "api_key": {"type": "string", "description": "AI API Key"},
                    "api_base": {"type": "string", "description": "API Base URL"},
                    "model_name": {"type": "string", "description": "模型名称", "default": "gpt-4o-mini"},
                    "template_id": {"type": "string", "description": "模板 ID"},
                    "auto_download": {"type": "boolean", "description": "自动下载图片", "default": False},
                },
            },
            "GenerateResponse": {
                "type": "object",
                "properties": {
                    "success": {"type": "boolean"},
                    "filename": {"type": "string"},
                    "title": {"type": "string"},
                    "subtitle": {"type": "string"},
                    "slide_count": {"type": "integer"},
                    "duration_ms": {"type": "integer"},
                    "download_url": {"type": "string"},
                },
            },
            "TaskStatus": {
                "type": "object",
                "properties": {
                    "task_id": {"type": "string"},
                    "status": {"type": "string", "enum": ["pending", "running", "success", "failed", "cancelled"]},
                    "progress": {"type": "integer", "minimum": 0, "maximum": 100},
                    "message": {"type": "string"},
                    "result": {"type": "object"},
                    "error": {"type": "string"},
                },
            },
        }


def setup_openapi(app: Flask) -> OpenAPIGenerator:
    """设置 OpenAPI 文档

    自动为应用注册 API 端点文档。
    """
    openapi = OpenAPIGenerator(
        app,
        title="AI PPT Generator API",
        version="1.0.0",
        description="AI 驱动的 PPT 自动生成服务 API 文档",
    )

    # 添加标签
    openapi.add_tag("生成", "PPT 生成相关接口")
    openapi.add_tag("任务", "异步任务管理")
    openapi.add_tag("历史", "生成历史记录")
    openapi.add_tag("系统", "系统状态和配置")

    # 注册端点文档
    _register_endpoints(openapi)

    return openapi


def _register_endpoints(openapi: OpenAPIGenerator):
    """注册所有端点文档"""

    # 生成 API
    openapi.register_endpoint(APIEndpoint(
        path="/api/generate",
        method="POST",
        summary="生成 PPT（同步）",
        description="同步生成 PPT，等待完成后返回结果。适合快速生成少量页面。",
        tags=["生成"],
        request_body=APIRequestBody(
            description="生成参数",
            schema={"$ref": "#/components/schemas/GenerateRequest"},
            example={
                "topic": "人工智能简介",
                "audience": "技术团队",
                "page_count": 5,
                "api_key": "sk-xxx",
            },
        ),
        responses=[
            APIResponse(200, "生成成功", schema={"$ref": "#/components/schemas/GenerateResponse"}),
            APIResponse(400, "参数错误", schema={"$ref": "#/components/schemas/Error"}),
            APIResponse(500, "服务器错误", schema={"$ref": "#/components/schemas/Error"}),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/generate/async",
        method="POST",
        summary="生成 PPT（异步）",
        description="异步生成 PPT，立即返回任务 ID，通过轮询获取状态。",
        tags=["生成", "任务"],
        request_body=APIRequestBody(
            description="生成参数",
            schema={"$ref": "#/components/schemas/GenerateRequest"},
        ),
        responses=[
            APIResponse(200, "任务创建成功", example={
                "success": True,
                "task_id": "abc123",
                "status_url": "/api/tasks/abc123",
            }),
            APIResponse(400, "参数错误"),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/generate/stream",
        method="POST",
        summary="生成 PPT（SSE 实时进度）",
        description="异步生成 PPT，通过 SSE 实时推送进度。",
        tags=["生成", "任务"],
        request_body=APIRequestBody(
            description="生成参数",
            schema={"$ref": "#/components/schemas/GenerateRequest"},
        ),
        responses=[
            APIResponse(200, "任务创建成功", example={
                "success": True,
                "task_id": "abc123",
                "events_url": "/api/events/abc123",
            }),
        ],
    ))

    # 任务 API
    openapi.register_endpoint(APIEndpoint(
        path="/api/tasks/{task_id}",
        method="GET",
        summary="获取任务状态",
        tags=["任务"],
        parameters=[
            APIParameter("task_id", "path", "任务 ID", required=True),
        ],
        responses=[
            APIResponse(200, "任务状态", schema={"$ref": "#/components/schemas/TaskStatus"}),
            APIResponse(404, "任务不存在"),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/tasks/{task_id}",
        method="DELETE",
        summary="取消任务",
        tags=["任务"],
        parameters=[
            APIParameter("task_id", "path", "任务 ID", required=True),
        ],
        responses=[
            APIResponse(200, "取消成功"),
            APIResponse(400, "无法取消"),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/tasks",
        method="GET",
        summary="列出任务",
        tags=["任务"],
        parameters=[
            APIParameter("limit", "query", "返回数量", schema={"type": "integer", "default": 20}),
        ],
        responses=[
            APIResponse(200, "任务列表"),
        ],
    ))

    # 历史 API
    openapi.register_endpoint(APIEndpoint(
        path="/api/history",
        method="GET",
        summary="获取生成历史",
        tags=["历史"],
        parameters=[
            APIParameter("limit", "query", "返回数量", schema={"type": "integer", "default": 20}),
            APIParameter("offset", "query", "偏移量", schema={"type": "integer", "default": 0}),
        ],
        responses=[
            APIResponse(200, "历史记录列表"),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/history/stats",
        method="GET",
        summary="获取统计信息",
        tags=["历史"],
        responses=[
            APIResponse(200, "统计信息", example={
                "total_generations": 100,
                "successful": 95,
                "success_rate": 95.0,
            }),
        ],
    ))

    # 系统 API
    openapi.register_endpoint(APIEndpoint(
        path="/health",
        method="GET",
        summary="健康检查",
        tags=["系统"],
        responses=[
            APIResponse(200, "健康", example={"status": "healthy"}),
            APIResponse(503, "不健康"),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/metrics",
        method="GET",
        summary="性能指标",
        tags=["系统"],
        responses=[
            APIResponse(200, "指标数据"),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/templates",
        method="GET",
        summary="获取模板列表",
        tags=["系统"],
        responses=[
            APIResponse(200, "模板列表", example={
                "success": True,
                "templates": [{"id": "default", "name": "默认模板"}],
            }),
        ],
    ))

    openapi.register_endpoint(APIEndpoint(
        path="/api/download/{filename}",
        method="GET",
        summary="下载 PPT 文件",
        tags=["生成"],
        parameters=[
            APIParameter("filename", "path", "文件名", required=True),
        ],
        responses=[
            APIResponse(200, "PPT 文件", content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"),
            APIResponse(404, "文件不存在"),
        ],
    ))


# Swagger UI HTML 模板
SWAGGER_UI_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>{{ title }} - API 文档</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/swagger-ui-dist@5/swagger-ui.css">
</head>
<body>
    <div id="swagger-ui"></div>
    <script src="https://cdn.jsdelivr.net/npm/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
    <script>
        SwaggerUIBundle({
            url: '/openapi.json',
            dom_id: '#swagger-ui',
            presets: [SwaggerUIBundle.presets.apis, SwaggerUIBundle.SwaggerUIStandalonePreset],
            layout: "StandaloneLayout"
        });
    </script>
</body>
</html>
"""

# ReDoc HTML 模板
REDOC_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>{{ title }} - API 文档</title>
    <link href="https://fonts.googleapis.com/css?family=Montserrat:300,400,700|Roboto:300,400,700" rel="stylesheet">
    <style>body { margin: 0; padding: 0; }</style>
</head>
<body>
    <redoc spec-url='/openapi.json'></redoc>
    <script src="https://cdn.jsdelivr.net/npm/redoc@latest/bundles/redoc.standalone.js"></script>
</body>
</html>
"""
