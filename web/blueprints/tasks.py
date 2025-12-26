"""异步任务蓝图 - 异步生成、SSE 进度推送"""

import os
from datetime import datetime
from flask import Blueprint, request, jsonify, current_app

from core.ai_client import generate_ppt_plan
from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from ppt.template_manager import template_manager
from utils.logger import get_logger
from utils.rate_limit import rate_limit
from config.settings import AIConfig, ImageConfig, RATE_LIMIT_PER_MINUTE, RATE_LIMIT_PER_HOUR

from .common import validate_request, sanitize_filename, validate_generation_params

logger = get_logger("web.tasks")

tasks_bp = Blueprint('tasks', __name__, url_prefix='/api')


@tasks_bp.route('/generate/async', methods=['POST'])
@rate_limit(requests_per_minute=RATE_LIMIT_PER_MINUTE, requests_per_hour=RATE_LIMIT_PER_HOUR)
@validate_request(['topic'])
def generate_ppt_async():
    """异步生成 PPT API"""
    from utils.async_tasks import get_task_manager

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

        task_manager = get_task_manager()
        task_id = task_manager.create_task()

        # 捕获当前 app 配置
        output_folder = current_app.config['OUTPUT_FOLDER']

        def generate_task(progress_callback=None):
            import time as _time
            start_time = _time.time()

            def set_progress(pct, msg):
                if progress_callback:
                    progress_callback(pct, msg)

            set_progress(5, "初始化配置...")

            ai_config = AIConfig(
                api_key=api_key,
                api_base_url=api_base,
                model_name=model_name
            )

            if unsplash_key:
                from utils.image_search import reset_searcher
                image_config = ImageConfig(unsplash_key=unsplash_key)
                reset_searcher(image_config)

            set_progress(10, "正在生成 PPT 结构...")

            plan_dict = generate_ppt_plan(
                topic, audience, page_count, description, auto_page_count, ai_config
            )
            plan = ppt_plan_from_dict(plan_dict)

            set_progress(50, f"生成了 {len(plan.slides)} 页幻灯片，正在构建...")

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_topic = sanitize_filename(topic, 30)
            filename = f"{safe_topic}_{timestamp}.pptx"
            output_path = os.path.join(output_folder, filename)

            template_path = template_manager.get_template(template_id) if template_id else None

            set_progress(60, "正在生成 PPT 文件...")

            build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)

            set_progress(90, "正在保存...")

            duration_ms = int((_time.time() - start_time) * 1000)
            file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0

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


@tasks_bp.route('/tasks/<task_id>')
def get_task_status(task_id):
    """获取任务状态"""
    from utils.async_tasks import get_task_manager

    task = get_task_manager().get_task(task_id)
    if not task:
        return jsonify({'success': False, 'error': '任务不存在'}), 404

    return jsonify({'success': True, **task.to_dict()})


@tasks_bp.route('/tasks/<task_id>', methods=['DELETE'])
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


@tasks_bp.route('/tasks')
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


@tasks_bp.route('/events/<task_id>')
def task_events(task_id):
    """任务进度 SSE 流"""
    from utils.sse import create_sse_response
    return create_sse_response(task_id)


@tasks_bp.route('/generate/stream', methods=['POST'])
@rate_limit(requests_per_minute=RATE_LIMIT_PER_MINUTE, requests_per_hour=RATE_LIMIT_PER_HOUR)
@validate_request(['topic'])
def generate_ppt_stream():
    """带 SSE 进度推送的异步生成"""
    from utils.async_tasks import get_task_manager
    from utils.sse import get_sse_manager

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

        task_manager = get_task_manager()
        sse_manager = get_sse_manager()
        task_id = task_manager.create_task()
        sse_channel = sse_manager.create_channel(task_id)

        output_folder = current_app.config['OUTPUT_FOLDER']

        def generate_with_sse(progress_callback=None):
            import time as _time
            start_time = _time.time()

            def set_progress(pct, msg):
                if progress_callback:
                    progress_callback(pct, msg)
                sse_channel.send_progress(pct, msg)

            set_progress(5, "初始化配置...")

            ai_config = AIConfig(
                api_key=api_key,
                api_base_url=api_base,
                model_name=model_name
            )

            if unsplash_key:
                from utils.image_search import reset_searcher
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
            output_path = os.path.join(output_folder, filename)

            template_path = template_manager.get_template(template_id) if template_id else None

            set_progress(60, "正在生成 PPT 文件...")
            build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)

            set_progress(90, "正在保存...")

            duration_ms = int((_time.time() - start_time) * 1000)
            file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0

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

            sse_channel.send_complete(result)
            return result

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
