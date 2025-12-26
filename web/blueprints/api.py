"""核心 API 蓝图 - 生成、下载、预览等基础功能"""

import os
from pathlib import Path
from datetime import datetime
from flask import Blueprint, request, jsonify, send_file, current_app
from werkzeug.utils import secure_filename

from core.ai_client import generate_ppt_plan, AIClientError, JSONParseError, test_api_connection
from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from ppt.template_manager import template_manager
from utils.file_parser import parse_file, get_text_summary, validate_file
from utils.logger import get_logger
from utils.rate_limit import rate_limit
from config.settings import (
    AIConfig, ImageConfig,
    DEFAULT_API_BASE_URL, DEFAULT_MODEL_NAME,
    RATE_LIMIT_PER_MINUTE, RATE_LIMIT_PER_HOUR,
    CACHE_CONFIG_TTL, CACHE_TEMPLATES_TTL, CACHE_STATS_TTL
)

from .common import (
    validate_request, sanitize_filename, cached_json_response,
    validate_generation_params, validate_api_url
)

logger = get_logger("web.api")

api_bp = Blueprint('api', __name__, url_prefix='/api')


@api_bp.route('/generate', methods=['POST'])
@rate_limit(requests_per_minute=RATE_LIMIT_PER_MINUTE, requests_per_hour=RATE_LIMIT_PER_HOUR)
@validate_request(['topic'])
def generate_ppt():
    """生成 PPT API"""
    try:
        data = request.json

        params, error = validate_generation_params(data)
        if error:
            return error

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

        ai_config = AIConfig(
            api_key=api_key,
            api_base_url=api_base,
            model_name=model_name
        )

        if unsplash_key:
            from utils.image_search import reset_searcher
            image_config = ImageConfig(unsplash_key=unsplash_key)
            reset_searcher(image_config)

        logger.info(f"生成 PPT: {topic} | 受众: {audience} | 模型: {ai_config.model_name} | 页数: {'自动' if auto_page_count else page_count}")

        import time
        start_time = time.time()

        plan_dict = generate_ppt_plan(
            topic, audience, page_count, description, auto_page_count, ai_config
        )
        plan = ppt_plan_from_dict(plan_dict)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_topic = sanitize_filename(topic, 30)
        filename = f"{safe_topic}_{timestamp}.pptx"
        output_path = os.path.join(current_app.config['OUTPUT_FOLDER'], filename)

        template_path = template_manager.get_template(template_id) if template_id else None

        build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)

        duration_ms = int((time.time() - start_time) * 1000)
        file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0

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


@api_bp.route('/download/<path:filename>')
def download_ppt(filename):
    """下载 PPT"""
    try:
        safe_filename = secure_filename(filename)
        if not safe_filename or safe_filename != filename.replace('/', '_').replace('\\', '_'):
            return jsonify({'error': '非法的文件名'}), 403

        output_folder = Path(current_app.config['OUTPUT_FOLDER']).resolve()
        filepath = (output_folder / safe_filename).resolve()

        if not filepath.is_relative_to(output_folder):
            return jsonify({'error': '非法的文件路径'}), 403

        if not filepath.exists():
            return jsonify({'error': f'文件不存在: {filename}'}), 404

        return send_file(
            filepath,
            as_attachment=True,
            download_name=safe_filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@api_bp.route('/preview', methods=['POST'])
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


@api_bp.route('/config')
def get_config():
    """获取配置信息"""
    data = {
        'ai_configured': bool(os.getenv('AI_API_KEY')),
        'image_search_available': bool(os.getenv('UNSPLASH_ACCESS_KEY'))
    }
    return cached_json_response(data, max_age=CACHE_CONFIG_TTL)


@api_bp.route('/test-connection', methods=['POST'])
def test_connection():
    """测试 API 连通性"""
    try:
        data = request.json or {}

        api_key = data.get('api_key', '')
        if not api_key:
            return jsonify({'success': False, 'message': '请输入 API Key'})

        ai_config = AIConfig(
            api_key=api_key,
            api_base_url=data.get('api_base', 'https://api.openai.com/v1'),
            model_name=data.get('model_name', 'gpt-4o-mini')
        )

        result = test_api_connection(ai_config)
        return jsonify(result)

    except Exception as e:
        return jsonify({'success': False, 'message': f'测试失败: {e}'})


@api_bp.route('/upload', methods=['POST'])
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
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
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
                    pass

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/templates')
def get_templates():
    """获取所有可用模板"""
    try:
        templates = template_manager.list_templates()
        data = {
            'success': True,
            'templates': templates,
            'count': len(templates)
        }
        return cached_json_response(data, max_age=CACHE_TEMPLATES_TTL)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'templates': []})


@api_bp.route('/history')
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


@api_bp.route('/history/stats')
def get_history_stats():
    """获取生成统计信息"""
    try:
        from utils.history import get_history
        stats = get_history().get_stats()
        data = {'success': True, **stats}
        return cached_json_response(data, max_age=CACHE_STATS_TTL)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/history/search')
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


@api_bp.route('/history/<int:record_id>')
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
