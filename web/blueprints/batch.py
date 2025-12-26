"""批量生成蓝图"""

from flask import Blueprint, request, jsonify

from utils.logger import get_logger
from utils.rate_limit import rate_limit
from config.settings import (
    DEFAULT_API_BASE_URL, DEFAULT_MODEL_NAME,
    MAX_BATCH_ITEMS, BATCH_RATE_LIMIT_PER_MINUTE, BATCH_RATE_LIMIT_PER_HOUR
)

from .common import validate_api_url

logger = get_logger("web.batch")

batch_bp = Blueprint('batch', __name__, url_prefix='/api')


@batch_bp.route('/batch', methods=['POST'])
@rate_limit(requests_per_minute=BATCH_RATE_LIMIT_PER_MINUTE, requests_per_hour=BATCH_RATE_LIMIT_PER_HOUR)
def create_batch_job():
    """创建批量生成任务"""
    from utils.batch import get_batch_generator

    try:
        data = request.json or {}

        items = data.get('items', [])
        if not items:
            return jsonify({'error': '请提供要生成的项目列表'}), 400

        if len(items) > MAX_BATCH_ITEMS:
            return jsonify({'error': f'单次最多生成 {MAX_BATCH_ITEMS} 个 PPT'}), 400

        for i, item in enumerate(items):
            if not item.get('topic'):
                return jsonify({'error': f'项目 {i+1} 缺少主题'}), 400

        api_key = data.get('api_key', '')
        if not api_key:
            return jsonify({'error': '请先配置 AI API Key'}), 400

        api_base = data.get('api_base', DEFAULT_API_BASE_URL)
        if not validate_api_url(api_base):
            return jsonify({'error': 'API URL 无效或不允许访问'}), 400

        generator = get_batch_generator()
        job = generator.create_job(
            items=items,
            api_config={
                'api_key': api_key,
                'api_base': api_base,
                'model_name': data.get('model_name', DEFAULT_MODEL_NAME),
            },
            template_id=data.get('template_id', ''),
        )

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


@batch_bp.route('/batch/<job_id>')
def get_batch_job(job_id):
    """获取批量任务状态"""
    from utils.batch import get_batch_generator

    job = get_batch_generator().get_job(job_id)
    if not job:
        return jsonify({'success': False, 'error': '任务不存在'}), 404

    return jsonify({'success': True, **job.to_dict()})


@batch_bp.route('/batch/<job_id>', methods=['DELETE'])
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


@batch_bp.route('/batch')
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
