"""Web 界面 - Flask 应用（优化版）"""
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
        pass  # Windows 上可能不支持 zh_CN.UTF-8 locale

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from flask import Flask, render_template, request, jsonify, send_file, Response, make_response
from werkzeug.utils import secure_filename
import json
import hashlib
from datetime import datetime
from functools import wraps
from urllib.parse import urlparse

from core.ai_client import generate_ppt_plan, AIClientError, JSONParseError, test_api_connection
from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from ppt.template_manager import template_manager
from utils.file_parser import parse_file, get_text_summary, validate_file
from utils.logger import get_logger
from utils.rate_limit import rate_limit
from utils.request_context import RequestContextMiddleware
from utils.error_handler import ErrorHandlerMiddleware
from utils.api_response import APIResponse
from utils.metrics import PerformanceMiddleware, get_metrics_collector
from utils.http_middleware import setup_http_middleware
from utils.security import setup_security
from utils.openapi import setup_openapi
from config.settings import AIConfig, ImageConfig, AppConfig

logger = get_logger("web")


# 应用配置
app_config = AppConfig.from_env()

app = Flask(__name__)
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

# 确保目录存在
Path(app.config['UPLOAD_FOLDER']).mkdir(parents=True, exist_ok=True)
Path(app.config['OUTPUT_FOLDER']).mkdir(parents=True, exist_ok=True)


def cleanup_old_files(folder: str, max_age_hours: int = 24, max_files: int = 100):
    """清理旧文件，防止目录无限增长"""
    import time
    try:
        files = []
        for f in Path(folder).glob('*.pptx'):
            files.append((f, f.stat().st_mtime))
        
        # 按修改时间排序
        files.sort(key=lambda x: x[1], reverse=True)
        
        now = time.time()
        for i, (filepath, mtime) in enumerate(files):
            # 删除超过 max_age_hours 小时的文件，或超过 max_files 数量的文件
            age_hours = (now - mtime) / 3600
            if age_hours > max_age_hours or i >= max_files:
                try:
                    filepath.unlink()
                except OSError:
                    pass  # 文件可能正在使用中
    except OSError:
        pass  # 目录不存在或无法访问


# 启动时清理旧文件
cleanup_old_files(app.config['OUTPUT_FOLDER'])


def validate_request(required_fields: list):
    """请求验证装饰器"""
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            data = request.json or {}
            missing = [field for field in required_fields if not data.get(field)]
            if missing:
                return jsonify({'error': f'缺少必填字段: {", ".join(missing)}'}), 400
            return f(*args, **kwargs)
        return wrapper
    return decorator


def sanitize_filename(filename: str, max_length: int = 50) -> str:
    """安全处理文件名"""
    safe = "".join(c for c in filename if c.isalnum() or c in (' ', '-', '_', '.'))
    safe = safe.replace(' ', '_')
    return safe[:max_length]


def cached_json_response(data: dict, max_age: int = 3600, etag: str = None):
    """创建带缓存头的 JSON 响应

    Args:
        data: 响应数据
        max_age: 缓存时间（秒）
        etag: ETag 值（可选）

    Returns:
        Flask Response 对象
    """
    response = make_response(jsonify(data))
    response.headers['Cache-Control'] = f'public, max-age={max_age}'
    if etag:
        response.headers['ETag'] = etag
    return response


def validate_api_url(url: str) -> bool:
    """验证 API URL 安全性（防止 SSRF）"""
    if not url:
        return False
    try:
        parsed = urlparse(url)
        # 只允许 http/https
        if parsed.scheme not in ('http', 'https'):
            return False
        # 禁止本地地址（防止 SSRF）
        host = parsed.hostname or ''
        host_lower = host.lower()

        # 禁止本地回环地址
        blocked_hosts = ('localhost', '127.0.0.1', '0.0.0.0', '::1')
        if host_lower in blocked_hosts:
            return False

        # 禁止内网地址（RFC 1918）
        # 10.0.0.0 - 10.255.255.255
        if host.startswith('10.'):
            return False
        # 192.168.0.0 - 192.168.255.255
        if host.startswith('192.168.'):
            return False
        # 172.16.0.0 - 172.31.255.255（注意：172.0-15 和 172.32+ 是公网）
        if host.startswith('172.'):
            try:
                second_octet = int(host.split('.')[1])
                if 16 <= second_octet <= 31:
                    return False
            except (IndexError, ValueError):
                pass

        # 禁止链路本地地址
        if host.startswith('169.254.'):
            return False

        # 禁止以 . 结尾的域名（可能绕过检查）
        if host_lower.endswith('.'):
            return False

        # 禁止带端口的本地服务
        if host_lower.endswith('.local') or host_lower.endswith('.internal'):
            return False

        return True
    except Exception:
        return False


def _validate_generation_params(data: dict) -> tuple:
    """验证生成 PPT 的通用参数

    Args:
        data: 请求数据字典

    Returns:
        (params_dict, error_response) - 成功时 error_response 为 None，失败时 params_dict 为 None
    """
    errors = []

    # 主题验证
    topic = data.get('topic', '').strip()
    if not topic:
        errors.append('主题不能为空')
    elif len(topic) > 500:
        errors.append('主题不能超过 500 字符')

    # 受众验证
    audience = data.get('audience', '通用受众').strip() or '通用受众'
    if len(audience) > 200:
        errors.append('受众描述不能超过 200 字符')

    # 页数验证
    try:
        page_count = int(data.get('page_count', 5))
        page_count = max(1, min(page_count, 100))
    except (ValueError, TypeError):
        page_count = 5

    # 其他参数
    auto_download = bool(data.get('auto_download', False))
    description = (data.get('description') or '').strip()
    if len(description) > 100000:
        description = description[:100000]

    auto_page_count = bool(data.get('auto_page_count', False))
    template_id = data.get('template_id', '')

    # API Key 验证
    api_key = data.get('api_key', '')
    if not api_key:
        errors.append('请先配置 AI API Key')

    # API URL 验证
    api_base = data.get('api_base', 'https://api.openai.com/v1')
    if api_base and not validate_api_url(api_base):
        errors.append('API URL 无效或不允许访问')

    # 模型名称验证
    model_name = data.get('model_name', 'gpt-4o-mini')
    if len(model_name) > 100:
        errors.append('模型名称过长')

    # 图片搜索配置
    unsplash_key = data.get('unsplash_key', '')

    # 如果有错误，返回错误响应
    if errors:
        return None, (jsonify({'error': errors[0] if len(errors) == 1 else errors}), 400)

    # 返回验证后的参数
    params = {
        'topic': topic,
        'audience': audience,
        'page_count': page_count,
        'auto_download': auto_download,
        'description': description,
        'auto_page_count': auto_page_count,
        'template_id': template_id,
        'api_key': api_key,
        'api_base': api_base,
        'model_name': model_name,
        'unsplash_key': unsplash_key,
    }

    return params, None


@app.route('/')
def index():
    """首页"""
    return render_template('index.html')


@app.route('/api/generate', methods=['POST'])
@rate_limit(requests_per_minute=5, requests_per_hour=50)
@validate_request(['topic'])
def generate_ppt():
    """生成 PPT API"""
    try:
        data = request.json

        # 使用公共验证函数
        params, error = _validate_generation_params(data)
        if error:
            return error

        # 解构参数
        topic = params['topic']
        audience = params['audience']
        page_count = params['page_count']
        auto_download = params['auto_download']
        description = params['description']
        auto_page_count = params['auto_page_count']
        template_id = params['template_id']
        api_key = params['api_key']
        api_base = params['api_base']
        model_name = params['model_name']
        unsplash_key = params['unsplash_key']

        # 创建配置对象
        ai_config = AIConfig(
            api_key=api_key,
            api_base_url=api_base,
            model_name=model_name
        )

        # 设置图片搜索配置
        if unsplash_key:
            from utils.image_search import reset_searcher
            from config.settings import ImageConfig
            image_config = ImageConfig(unsplash_key=unsplash_key)
            reset_searcher(image_config)

        # 日志
        logger.info(f"生成 PPT: {topic} | 受众: {audience} | 模型: {ai_config.model_name} | 页数: {'自动' if auto_page_count else page_count}")

        # 记录开始时间
        import time
        start_time = time.time()

        # 生成 PPT 结构
        plan_dict = generate_ppt_plan(
            topic, audience, page_count, description, auto_page_count, ai_config
        )
        plan = ppt_plan_from_dict(plan_dict)

        # 生成文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_topic = sanitize_filename(topic, 30)
        filename = f"{safe_topic}_{timestamp}.pptx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)

        # 获取模板
        template_path = template_manager.get_template(template_id) if template_id else None

        # 生成 PPT
        build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)

        # 计算耗时和文件大小
        duration_ms = int((time.time() - start_time) * 1000)
        file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0

        # 记录历史
        try:
            from utils.history import get_history
            from utils.request_context import get_request_id
            get_history().add(
                topic=topic,
                audience=audience,
                page_count=page_count,
                model_name=model_name,
                template_id=template_id,
                filename=filename,
                file_size=file_size,
                slide_count=len(plan.slides),
                duration_ms=duration_ms,
                status="success",
                request_id=get_request_id() or "",
                client_ip=request.remote_addr or "",
            )
        except Exception as e:
            logger.warning(f"记录历史失败: {e}")

        return jsonify({
            'success': True,
            'filename': filename,
            'title': plan.title,
            'subtitle': plan.subtitle,
            'slide_count': len(plan.slides),
            'duration_ms': duration_ms,
            'download_url': f'/api/download/{filename}'
        })

    except JSONParseError as e:
        return jsonify({'error': str(e), 'type': 'json_error'}), 400
    except AIClientError as e:
        return jsonify({'error': str(e), 'type': 'ai_error'}), 500
    except Exception as e:
        logger.error(f"生成失败: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<path:filename>')
def download_ppt(filename):
    """下载 PPT"""
    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])
        
        # 安全检查
        if not filepath.startswith(output_folder):
            return jsonify({'error': '非法的文件路径'}), 403
        
        if not os.path.exists(filepath):
            return jsonify({'error': f'文件不存在: {filename}'}), 404
        
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/preview', methods=['POST'])
@validate_request(['topic'])
def preview_structure():
    """预览 PPT 结构"""
    try:
        data = request.json
        
        api_key = data.get('api_key', '')
        if not api_key:
            return jsonify({'error': '请先配置 AI API Key'}), 400
        
        ai_config = AIConfig(
            api_key=api_key,
            api_base_url=data.get('api_base', 'https://api.openai.com/v1'),
            model_name=data.get('model_name', 'gpt-4o-mini')
        )
        
        plan_dict = generate_ppt_plan(
            data.get('topic', '').strip(),
            data.get('audience', '通用受众').strip(),
            int(data.get('page_count', 5)),
            config=ai_config
        )
        plan = ppt_plan_from_dict(plan_dict)
        
        slides_preview = [
            {
                'index': i,
                'type': slide.slide_type,
                'title': slide.title,
                'bullets': slide.bullets or [],
                'text': slide.text or '',
                'image_keyword': slide.image_keyword or '',
                'subtitle': slide.subtitle or ''
            }
            for i, slide in enumerate(plan.slides, 1)
        ]
        
        return jsonify({
            'success': True,
            'title': plan.title,
            'subtitle': plan.subtitle,
            'slides': slides_preview
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/config')
def get_config():
    """获取配置信息（缓存5分钟）"""
    data = {
        'ai_configured': bool(os.getenv('AI_API_KEY')),
        'image_search_available': bool(os.getenv('UNSPLASH_ACCESS_KEY'))
    }
    # 配置信息缓存5分钟
    return cached_json_response(data, max_age=300)


@app.route('/api/test-connection', methods=['POST'])
def test_connection():
    """测试 API 连通性"""
    try:
        data = request.json or {}
        
        api_key = data.get('api_key', '')
        if not api_key:
            return jsonify({
                'success': False,
                'message': '请输入 API Key'
            })
        
        ai_config = AIConfig(
            api_key=api_key,
            api_base_url=data.get('api_base', 'https://api.openai.com/v1'),
            model_name=data.get('model_name', 'gpt-4o-mini')
        )
        
        result = test_api_connection(ai_config)
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'测试失败: {e}'
        })


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传并解析文件"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '没有文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'}), 400
        
        if not validate_file(file.filename):
            return jsonify({
                'success': False,
                'error': '不支持的文件格式。支持：TXT, MD, DOCX, PDF'
            }), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            content = parse_file(filepath)
            summary = get_text_summary(content)
            
            return jsonify({
                'success': True,
                'content': content,
                'summary': summary,
                'filename': filename
            })
        finally:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except OSError:
                    pass  # 文件可能正在使用中
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/templates')
def get_templates():
    """获取所有可用模板（缓存24小时）"""
    try:
        templates = template_manager.list_templates()
        data = {
            'success': True,
            'templates': templates,
            'count': len(templates)
        }
        # 模板列表变化不频繁，缓存24小时
        return cached_json_response(data, max_age=86400)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'templates': []})


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


@app.route('/api/history')
def get_history_list():
    """获取生成历史列表"""
    try:
        from utils.history import get_history
        limit = min(int(request.args.get('limit', 20)), 100)
        offset = max(int(request.args.get('offset', 0)), 0)
        records = get_history().get_recent(limit=limit, offset=offset)
        return jsonify({
            'success': True,
            'records': records,
            'count': len(records)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/history/stats')
def get_history_stats():
    """获取生成统计信息（缓存1小时）"""
    try:
        from utils.history import get_history
        stats = get_history().get_stats()
        data = {
            'success': True,
            **stats
        }
        # 统计信息缓存1小时
        return cached_json_response(data, max_age=3600)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/history/search')
def search_history():
    """搜索历史记录"""
    try:
        from utils.history import get_history
        keyword = request.args.get('keyword', '')
        status = request.args.get('status', '')
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')
        limit = min(int(request.args.get('limit', 20)), 100)

        records = get_history().search(
            keyword=keyword,
            status=status,
            start_date=start_date,
            end_date=end_date,
            limit=limit
        )
        return jsonify({
            'success': True,
            'records': records,
            'count': len(records)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/history/<int:record_id>')
def get_history_detail(record_id):
    """获取单条历史记录详情"""
    try:
        from utils.history import get_history
        record = get_history().get_by_id(record_id)
        if record:
            return jsonify({'success': True, 'record': record})
        return jsonify({'success': False, 'error': '记录不存在'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ==================== 异步生成 API ====================

@app.route('/api/generate/async', methods=['POST'])
@rate_limit(requests_per_minute=5, requests_per_hour=50)
@validate_request(['topic'])
def generate_ppt_async():
    """异步生成 PPT API

    返回任务 ID，可通过 /api/tasks/<task_id> 查询进度。
    """
    from utils.async_tasks import get_task_manager, TaskStatus

    try:
        data = request.json

        # 使用公共验证函数
        params, error = _validate_generation_params(data)
        if error:
            return error

        # 解构参数
        topic = params['topic']
        audience = params['audience']
        page_count = params['page_count']
        auto_download = params['auto_download']
        description = params['description']
        auto_page_count = params['auto_page_count']
        template_id = params['template_id']
        api_key = params['api_key']
        api_base = params['api_base']
        model_name = params['model_name']
        unsplash_key = params['unsplash_key']

        # 创建任务
        task_manager = get_task_manager()
        task_id = task_manager.create_task()

        # 定义后台执行的生成函数
        def generate_task(progress_callback=None):
            import time as _time
            start_time = _time.time()

            # 设置进度
            def set_progress(pct, msg):
                if progress_callback:
                    progress_callback(pct, msg)

            set_progress(5, "初始化配置...")

            # 创建 AI 配置
            ai_config = AIConfig(
                api_key=api_key,
                api_base_url=api_base,
                model_name=model_name
            )

            # 设置图片搜索配置
            if unsplash_key:
                from utils.image_search import reset_searcher
                from config.settings import ImageConfig
                image_config = ImageConfig(unsplash_key=unsplash_key)
                reset_searcher(image_config)

            set_progress(10, "正在生成 PPT 结构...")

            # 生成 PPT 结构
            plan_dict = generate_ppt_plan(
                topic, audience, page_count, description, auto_page_count, ai_config
            )
            plan = ppt_plan_from_dict(plan_dict)

            set_progress(50, f"生成了 {len(plan.slides)} 页幻灯片，正在构建...")

            # 生成文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_topic = sanitize_filename(topic, 30)
            filename = f"{safe_topic}_{timestamp}.pptx"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)

            # 获取模板
            template_path = template_manager.get_template(template_id) if template_id else None

            set_progress(60, "正在生成 PPT 文件...")

            # 生成 PPT
            build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)

            set_progress(90, "正在保存...")

            # 计算耗时和文件大小
            duration_ms = int((_time.time() - start_time) * 1000)
            file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0

            # 记录历史
            try:
                from utils.history import get_history
                get_history().add(
                    topic=topic,
                    audience=audience,
                    page_count=page_count,
                    model_name=model_name,
                    template_id=template_id,
                    filename=filename,
                    file_size=file_size,
                    slide_count=len(plan.slides),
                    duration_ms=duration_ms,
                    status="success",
                    request_id=task_id,
                    client_ip="",
                )
            except Exception as e:
                logger.warning(f"记录历史失败: {e}")

            return {
                'filename': filename,
                'title': plan.title,
                'subtitle': plan.subtitle,
                'slide_count': len(plan.slides),
                'duration_ms': duration_ms,
                'download_url': f'/api/download/{filename}'
            }

        # 启动后台任务
        task_manager.run_task(task_id, generate_task)

        logger.info(f"异步生成任务已创建: {task_id} | 主题: {topic}")

        return jsonify({
            'success': True,
            'task_id': task_id,
            'message': '任务已创建，请通过任务 ID 查询进度',
            'status_url': f'/api/tasks/{task_id}'
        })

    except Exception as e:
        logger.error(f"创建异步任务失败: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/tasks/<task_id>')
def get_task_status(task_id):
    """获取任务状态"""
    from utils.async_tasks import get_task_manager

    task = get_task_manager().get_task(task_id)
    if not task:
        return jsonify({'success': False, 'error': '任务不存在'}), 404

    return jsonify({
        'success': True,
        **task.to_dict()
    })


@app.route('/api/tasks/<task_id>', methods=['DELETE'])
def cancel_task(task_id):
    """取消任务"""
    from utils.async_tasks import get_task_manager

    success = get_task_manager().cancel_task(task_id)
    if success:
        return jsonify({'success': True, 'message': '任务已取消'})
    return jsonify({
        'success': False,
        'error': '无法取消任务（可能已在执行或已完成）'
    }), 400


@app.route('/api/tasks')
def list_tasks():
    """列出最近的任务"""
    from utils.async_tasks import get_task_manager

    limit = min(int(request.args.get('limit', 20)), 100)
    tasks = get_task_manager().list_tasks(limit=limit)
    stats = get_task_manager().get_stats()

    return jsonify({
        'success': True,
        'tasks': tasks,
        'stats': stats
    })


# ==================== SSE 实时进度推送 ====================

@app.route('/api/events/<task_id>')
def task_events(task_id):
    """任务进度 SSE 流

    客户端连接后实时接收任务进度更新。

    用法（前端）:
        const eventSource = new EventSource('/api/events/task123');
        eventSource.onmessage = (e) => console.log(e.data);
        eventSource.addEventListener('progress', (e) => {
            const data = JSON.parse(e.data);
            console.log(`进度: ${data.progress}%`);
        });
        eventSource.addEventListener('complete', (e) => {
            const result = JSON.parse(e.data);
            console.log('完成:', result);
            eventSource.close();
        });
    """
    from utils.sse import create_sse_response
    return create_sse_response(task_id)


@app.route('/api/generate/stream', methods=['POST'])
@rate_limit(requests_per_minute=5, requests_per_hour=50)
@validate_request(['topic'])
def generate_ppt_stream():
    """带 SSE 进度推送的异步生成

    返回任务 ID，客户端可通过 /api/events/<task_id> 接收实时进度。
    """
    from utils.async_tasks import get_task_manager
    from utils.sse import get_sse_manager

    try:
        data = request.json

        # 使用公共验证函数
        params, error = _validate_generation_params(data)
        if error:
            return error

        # 解构参数
        topic = params['topic']
        audience = params['audience']
        page_count = params['page_count']
        auto_download = params['auto_download']
        description = params['description']
        auto_page_count = params['auto_page_count']
        template_id = params['template_id']
        api_key = params['api_key']
        api_base = params['api_base']
        model_name = params['model_name']
        unsplash_key = params['unsplash_key']

        # 创建任务和 SSE 通道
        task_manager = get_task_manager()
        sse_manager = get_sse_manager()
        task_id = task_manager.create_task()
        sse_channel = sse_manager.create_channel(task_id)

        def generate_with_sse(progress_callback=None):
            import time as _time
            start_time = _time.time()

            def set_progress(pct, msg):
                if progress_callback:
                    progress_callback(pct, msg)
                # 同时推送到 SSE
                sse_channel.send_progress(pct, msg)

            set_progress(5, "初始化配置...")

            ai_config = AIConfig(
                api_key=api_key,
                api_base_url=api_base,
                model_name=model_name
            )

            if unsplash_key:
                from utils.image_search import reset_searcher
                from config.settings import ImageConfig
                reset_searcher(ImageConfig(unsplash_key=unsplash_key))

            set_progress(10, "正在生成 PPT 结构...")

            plan_dict = generate_ppt_plan(
                topic, audience, page_count, description, auto_page_count, ai_config
            )
            plan = ppt_plan_from_dict(plan_dict)

            set_progress(50, f"生成了 {len(plan.slides)} 页幻灯片，正在构建...")

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_topic = sanitize_filename(topic, 30)
            filename = f"{safe_topic}_{timestamp}.pptx"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)

            template_path = template_manager.get_template(template_id) if template_id else None

            set_progress(60, "正在生成 PPT 文件...")
            build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)

            set_progress(90, "正在保存...")

            duration_ms = int((_time.time() - start_time) * 1000)
            file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0

            # 记录历史
            try:
                from utils.history import get_history
                get_history().add(
                    topic=topic, audience=audience, page_count=page_count,
                    model_name=model_name, template_id=template_id,
                    filename=filename, file_size=file_size,
                    slide_count=len(plan.slides), duration_ms=duration_ms,
                    status="success", request_id=task_id, client_ip="",
                )
            except Exception as e:
                logger.warning(f"记录历史失败: {e}")

            result = {
                'filename': filename,
                'title': plan.title,
                'subtitle': plan.subtitle,
                'slide_count': len(plan.slides),
                'duration_ms': duration_ms,
                'download_url': f'/api/download/{filename}'
            }

            # 通过 SSE 推送完成事件
            sse_channel.send_complete(result)
            return result

        # 启动后台任务
        task_manager.run_task(task_id, generate_with_sse)

        logger.info(f"带 SSE 的异步生成任务已创建: {task_id} | 主题: {topic}")

        return jsonify({
            'success': True,
            'task_id': task_id,
            'message': '任务已创建，请连接 SSE 获取实时进度',
            'events_url': f'/api/events/{task_id}',
            'status_url': f'/api/tasks/{task_id}'
        })

    except Exception as e:
        logger.error(f"创建 SSE 任务失败: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


# ==================== 批量生成 API ====================

@app.route('/api/batch', methods=['POST'])
@rate_limit(requests_per_minute=2, requests_per_hour=20)
def create_batch_job():
    """创建批量生成任务

    请求体:
    {
        "items": [
            {"topic": "主题1", "audience": "受众", "page_count": 5},
            {"topic": "主题2"}
        ],
        "api_key": "...",
        "api_base": "...",
        "model_name": "...",
        "template_id": "..."
    }
    """
    from utils.batch import get_batch_generator

    try:
        data = request.json or {}

        items = data.get('items', [])
        if not items:
            return jsonify({'error': '请提供要生成的项目列表'}), 400

        if len(items) > 20:
            return jsonify({'error': '单次最多生成 20 个 PPT'}), 400

        # 验证每个项目
        for i, item in enumerate(items):
            if not item.get('topic'):
                return jsonify({'error': f'项目 {i+1} 缺少主题'}), 400

        api_key = data.get('api_key', '')
        if not api_key:
            return jsonify({'error': '请先配置 AI API Key'}), 400

        api_base = data.get('api_base', 'https://api.openai.com/v1')
        if not validate_api_url(api_base):
            return jsonify({'error': 'API URL 无效或不允许访问'}), 400

        # 创建批量任务
        generator = get_batch_generator()
        job = generator.create_job(
            items=items,
            api_config={
                'api_key': api_key,
                'api_base': api_base,
                'model_name': data.get('model_name', 'gpt-4o-mini'),
            },
            template_id=data.get('template_id', ''),
        )

        # 启动任务
        generator.start_job(job.job_id)

        logger.info(f"批量生成任务已创建: {job.job_id} ({len(items)} 项)")

        return jsonify({
            'success': True,
            'job_id': job.job_id,
            'total': len(items),
            'message': '批量任务已创建',
            'status_url': f'/api/batch/{job.job_id}'
        })

    except Exception as e:
        logger.error(f"创建批量任务失败: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/batch/<job_id>')
def get_batch_job(job_id):
    """获取批量任务状态"""
    from utils.batch import get_batch_generator

    job = get_batch_generator().get_job(job_id)
    if not job:
        return jsonify({'success': False, 'error': '任务不存在'}), 404

    return jsonify({
        'success': True,
        **job.to_dict()
    })


@app.route('/api/batch/<job_id>', methods=['DELETE'])
def cancel_batch_job(job_id):
    """取消批量任务"""
    from utils.batch import get_batch_generator

    success = get_batch_generator().cancel_job(job_id)
    if success:
        return jsonify({'success': True, 'message': '任务已取消'})
    return jsonify({
        'success': False,
        'error': '无法取消任务（可能已完成或不存在）'
    }), 400


@app.route('/api/batch')
def list_batch_jobs():
    """列出批量任务"""
    from utils.batch import get_batch_generator

    limit = min(int(request.args.get('limit', 20)), 100)
    jobs = get_batch_generator().list_jobs(limit=limit)

    return jsonify({
        'success': True,
        'jobs': jobs,
        'count': len(jobs)
    })


# ==================== PPT 编辑 API ====================

@app.route('/api/ppt/<path:filename>/info')
def get_ppt_info(filename):
    """获取 PPT 文件信息

    返回幻灯片列表、每页标题、内容类型等信息。
    """
    from ppt.editor import PPTEditor

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        # 安全检查
        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        editor = PPTEditor(filepath)
        return jsonify({
            'success': True,
            **editor.to_dict()
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/ppt/<path:filename>/slide/<int:index>', methods=['GET'])
def get_slide_info(filename, index):
    """获取指定幻灯片信息"""
    from ppt.editor import PPTEditor

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        editor = PPTEditor(filepath)
        info = editor.get_slide_info(index)

        if info is None:
            return jsonify({'success': False, 'error': '幻灯片不存在'}), 404

        return jsonify({
            'success': True,
            'slide': {
                'index': info.index,
                'title': info.title,
                'content_type': info.content_type,
                'shape_count': info.shape_count,
                'has_image': info.has_image,
                'text_preview': info.text_preview
            }
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/ppt/<path:filename>/slide/<int:index>/title', methods=['PUT'])
def update_slide_title(filename, index):
    """更新幻灯片标题

    请求体：{"title": "新标题"}
    """
    from ppt.editor import PPTEditor

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        data = request.json or {}
        new_title = data.get('title', '')

        if not new_title:
            return jsonify({'success': False, 'error': '标题不能为空'}), 400

        editor = PPTEditor(filepath)
        if editor.update_slide_title(index, new_title):
            editor.save()
            return jsonify({'success': True, 'message': '标题已更新'})
        else:
            return jsonify({'success': False, 'error': '更新失败'}), 400

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/ppt/<path:filename>/slide/<int:index>', methods=['DELETE'])
def delete_slide(filename, index):
    """删除幻灯片"""
    from ppt.editor import PPTEditor

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        editor = PPTEditor(filepath)
        if editor.delete_slide(index):
            editor.save()
            return jsonify({
                'success': True,
                'message': '幻灯片已删除',
                'slide_count': editor.get_slide_count()
            })
        else:
            return jsonify({'success': False, 'error': '删除失败'}), 400

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/ppt/<path:filename>/reorder', methods=['POST'])
def reorder_slides(filename):
    """重新排列幻灯片顺序

    请求体：{"order": [2, 0, 1, 3]}  // 新的索引顺序
    """
    from ppt.editor import PPTEditor

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        data = request.json or {}
        new_order = data.get('order', [])

        if not new_order or not isinstance(new_order, list):
            return jsonify({'success': False, 'error': '请提供有效的顺序列表'}), 400

        editor = PPTEditor(filepath)

        # 验证顺序
        slide_count = editor.get_slide_count()
        if len(new_order) != slide_count:
            return jsonify({
                'success': False,
                'error': f'顺序列表长度({len(new_order)})与幻灯片数量({slide_count})不匹配'
            }), 400

        if set(new_order) != set(range(slide_count)):
            return jsonify({'success': False, 'error': '顺序列表包含无效索引'}), 400

        # 执行重排
        from ppt.editor import reorder_slides as do_reorder
        if do_reorder(filepath, new_order):
            return jsonify({
                'success': True,
                'message': '幻灯片顺序已更新',
                'order': new_order
            })
        else:
            return jsonify({'success': False, 'error': '重排失败'}), 400

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/ppt/<path:filename>/slide/<int:index>/duplicate', methods=['POST'])
def duplicate_slide(filename, index):
    """复制幻灯片"""
    from ppt.editor import PPTEditor

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        editor = PPTEditor(filepath)
        new_index = editor.duplicate_slide(index)

        if new_index >= 0:
            editor.save()
            return jsonify({
                'success': True,
                'message': '幻灯片已复制',
                'new_index': new_index,
                'slide_count': editor.get_slide_count()
            })
        else:
            return jsonify({'success': False, 'error': '复制失败'}), 400

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ==================== 导出 API ====================

@app.route('/api/export/formats')
def get_export_formats():
    """获取可用的导出格式"""
    from utils.export import get_export_manager

    manager = get_export_manager()
    data = {
        'success': True,
        **manager.get_status()
    }
    # 格式列表变化不频繁，缓存1小时
    return cached_json_response(data, max_age=3600)


@app.route('/api/export/<path:filename>/pdf', methods=['POST'])
def export_to_pdf(filename):
    """将 PPT 导出为 PDF

    返回 PDF 文件下载链接。
    """
    from utils.export import get_export_manager, ExportError

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        # 安全检查
        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        manager = get_export_manager()
        if not manager.is_format_available('pdf'):
            return jsonify({
                'success': False,
                'error': 'PDF 导出不可用（需要安装 LibreOffice）'
            }), 503

        # 生成输出路径
        pdf_filename = Path(filename).stem + '.pdf'
        output_path = os.path.join(output_folder, pdf_filename)

        # 执行导出
        result_path = manager.export(filepath, output_path, 'pdf')

        return jsonify({
            'success': True,
            'message': 'PDF 导出成功',
            'filename': pdf_filename,
            'download_url': f'/api/download/{pdf_filename}'
        })

    except ExportError as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    except Exception as e:
        logger.error(f"PDF 导出失败: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/<path:filename>/images', methods=['POST'])
def export_to_images(filename):
    """将 PPT 导出为图片序列

    返回包含所有图片的 ZIP 文件下载链接。
    """
    from utils.export import get_export_manager, ExportError
    import zipfile

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        # 安全检查
        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        manager = get_export_manager()
        if not manager.is_format_available('images'):
            return jsonify({
                'success': False,
                'error': '图片导出不可用（需要安装 LibreOffice 和 pdf2image）'
            }), 503

        # 创建临时目录存放图片
        import tempfile
        with tempfile.TemporaryDirectory() as tmp_dir:
            # 执行导出
            manager.export(filepath, tmp_dir, 'images')

            # 打包为 ZIP
            zip_filename = Path(filename).stem + '_images.zip'
            zip_path = os.path.join(output_folder, zip_filename)

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for img_file in sorted(Path(tmp_dir).glob('*.png')):
                    zf.write(img_file, img_file.name)

        # 统计图片数量
        with zipfile.ZipFile(zip_path, 'r') as zf:
            image_count = len(zf.namelist())

        return jsonify({
            'success': True,
            'message': f'已导出 {image_count} 张图片',
            'filename': zip_filename,
            'image_count': image_count,
            'download_url': f'/api/download/{zip_filename}'
        })

    except ExportError as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    except Exception as e:
        logger.error(f"图片导出失败: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/<path:filename>/thumbnail', methods=['POST'])
def export_thumbnail(filename):
    """生成 PPT 缩略图

    返回缩略图下载链接。
    """
    from utils.export import get_export_manager, ExportError

    try:
        filepath = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
        output_folder = os.path.abspath(app.config['OUTPUT_FOLDER'])

        # 安全检查
        if not filepath.startswith(output_folder):
            return jsonify({'success': False, 'error': '非法的文件路径'}), 403

        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        manager = get_export_manager()
        if not manager.is_format_available('thumbnail'):
            return jsonify({
                'success': False,
                'error': '缩略图生成不可用'
            }), 503

        # 生成输出路径
        thumb_filename = Path(filename).stem + '_thumb.png'
        output_path = os.path.join(output_folder, thumb_filename)

        # 执行导出
        manager.export(filepath, output_path, 'thumbnail')

        return jsonify({
            'success': True,
            'message': '缩略图生成成功',
            'filename': thumb_filename,
            'download_url': f'/api/download/{thumb_filename}'
        })

    except ExportError as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    except Exception as e:
        logger.error(f"缩略图生成失败: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    logger.info(f"AI PPT 生成器启动 | 访问地址: http://localhost:{app_config.port}")
    app.run(debug=app_config.debug, host=app_config.host, port=app_config.port)
