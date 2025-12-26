"""Web 界面 - Flask 应用（重构版）"""
# -*- coding: utf-8 -*-
import sys
import os
from pathlib import Path

# 设置编码
if sys.platform == 'win32':
    try:
        import locale
        locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
    except locale.Error:
        pass

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from flask import Flask, render_template, jsonify, send_from_directory

from utils.logger import get_logger
from utils.request_context import RequestContextMiddleware
from utils.error_handler import ErrorHandlerMiddleware
from utils.metrics import PerformanceMiddleware, get_metrics_collector
from utils.http_middleware import setup_http_middleware
from utils.security import setup_security
from utils.openapi import setup_openapi
from config.settings import AppConfig

# 导入蓝图
from web.blueprints import api_bp, tasks_bp, batch_bp, ppt_edit_bp, export_bp
from web.blueprints.common import cleanup_old_files

logger = get_logger("web")

# 应用配置
app_config = AppConfig.from_env()

# 检查是否存在新版前端静态文件
static_folder = Path(__file__).parent / 'static'
use_new_frontend = (static_folder / 'index.html').exists()

if use_new_frontend:
    # 使用新版 React 前端
    app = Flask(__name__, static_folder='static', static_url_path='')
    logger.info("使用新版 React 前端")
else:
    # 使用旧版模板
    app = Flask(__name__)
    logger.info("使用旧版模板前端")
app.config['SECRET_KEY'] = app_config.secret_key
app.config['UPLOAD_FOLDER'] = app_config.upload_folder
app.config['OUTPUT_FOLDER'] = app_config.output_folder
app.config['MAX_CONTENT_LENGTH'] = app_config.max_upload_size
app.config['JSON_AS_ASCII'] = False

# 初始化中间件
RequestContextMiddleware(app)  # 请求 ID 追踪
ErrorHandlerMiddleware(app)    # 全局异常处理
PerformanceMiddleware(app)     # 性能监控
setup_http_middleware(app)     # CORS 和 Gzip
setup_security(app)            # 安全响应头
setup_openapi(app)             # API 文档

# 注册蓝图
app.register_blueprint(api_bp)
app.register_blueprint(tasks_bp)
app.register_blueprint(batch_bp)
app.register_blueprint(ppt_edit_bp)
app.register_blueprint(export_bp)

# 确保目录存在
Path(app.config['UPLOAD_FOLDER']).mkdir(parents=True, exist_ok=True)
Path(app.config['OUTPUT_FOLDER']).mkdir(parents=True, exist_ok=True)

# 启动时清理旧文件
cleanup_old_files(app.config['OUTPUT_FOLDER'])


@app.route('/')
def index():
    """首页"""
    if use_new_frontend:
        return send_from_directory(app.static_folder, 'index.html')
    return render_template('index.html')


# SPA 路由支持 - 处理前端路由
@app.errorhandler(404)
def not_found(e):
    """处理 404，支持 SPA 路由"""
    if use_new_frontend:
        return send_from_directory(app.static_folder, 'index.html')
    return jsonify({'error': 'Not found'}), 404


@app.route('/health')
def health():
    """基础健康检查"""
    from utils.health import check_health
    result = check_health()
    status_code = 200 if result["status"] == "healthy" else 503
    return jsonify(result), status_code


@app.route('/health/detailed')
def health_detailed():
    """详细健康检查（包含系统信息和指标）"""
    from utils.health import get_detailed_health
    return jsonify(get_detailed_health())


@app.route('/metrics')
def metrics():
    """获取性能指标"""
    return jsonify(get_metrics_collector().get_stats())


if __name__ == '__main__':
    logger.info(f"AI PPT 生成器启动 | 访问地址: http://localhost:{app_config.port}")
    app.run(debug=app_config.debug, host=app_config.host, port=app_config.port)
