"""公共工具函数和装饰器"""

import os
from pathlib import Path
from functools import wraps
from urllib.parse import urlparse
from flask import request, jsonify, make_response, current_app
from werkzeug.utils import secure_filename

from config.settings import (
    DEFAULT_API_BASE_URL, DEFAULT_MODEL_NAME,
    MAX_TOPIC_LENGTH, MAX_AUDIENCE_LENGTH, MAX_DESCRIPTION_LENGTH,
    MAX_MODEL_NAME_LENGTH, MAX_PAGE_COUNT, MAX_FILENAME_LENGTH,
    FILE_CLEANUP_MAX_AGE_HOURS, FILE_CLEANUP_MAX_FILES
)


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


def sanitize_filename(filename: str, max_length: int = MAX_FILENAME_LENGTH) -> str:
    """安全处理文件名"""
    safe = "".join(c for c in filename if c.isalnum() or c in (' ', '-', '_', '.'))
    safe = safe.replace(' ', '_')
    return safe[:max_length]


def cached_json_response(data: dict, max_age: int = 3600, etag: str = None):
    """创建带缓存头的 JSON 响应"""
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
        if parsed.scheme not in ('http', 'https'):
            return False

        host = parsed.hostname or ''
        if not host:
            return False

        host_lower = host.lower()

        if host_lower.endswith('.'):
            return False
        if host_lower.endswith('.local') or host_lower.endswith('.internal'):
            return False

        blocked_hosts = ('localhost', '127.0.0.1', '0.0.0.0', '::1', '[::1]')
        if host_lower in blocked_hosts:
            return False

        import ipaddress
        try:
            ip_str = host.strip('[]')
            ip = ipaddress.ip_address(ip_str)

            if ip.is_private or ip.is_loopback or ip.is_link_local:
                return False
            if ip.is_multicast or ip.is_reserved or ip.is_unspecified:
                return False

            if isinstance(ip, ipaddress.IPv6Address) and ip.ipv4_mapped:
                mapped_v4 = ip.ipv4_mapped
                if mapped_v4.is_private or mapped_v4.is_loopback or mapped_v4.is_link_local:
                    return False

        except ValueError:
            if host.startswith('10.') or host.startswith('192.168.') or host.startswith('169.254.'):
                return False
            if host.startswith('172.'):
                try:
                    second_octet = int(host.split('.')[1])
                    if 16 <= second_octet <= 31:
                        return False
                except (IndexError, ValueError):
                    pass
            if host.isdigit():
                try:
                    decimal_ip = int(host)
                    if 0 <= decimal_ip <= 0xFFFFFFFF:
                        ip = ipaddress.ip_address(decimal_ip)
                        if ip.is_private or ip.is_loopback or ip.is_link_local:
                            return False
                except (ValueError, OverflowError):
                    pass

        return True
    except Exception:
        return False


def validate_generation_params(data: dict) -> tuple:
    """验证生成 PPT 的通用参数

    Returns:
        (params_dict, error_response) - 成功时 error_response 为 None
    """
    errors = []

    topic = data.get('topic', '').strip()
    if not topic:
        errors.append('主题不能为空')
    elif len(topic) > MAX_TOPIC_LENGTH:
        errors.append(f'主题不能超过 {MAX_TOPIC_LENGTH} 字符')

    audience = data.get('audience', '通用受众').strip() or '通用受众'
    if len(audience) > MAX_AUDIENCE_LENGTH:
        errors.append(f'受众描述不能超过 {MAX_AUDIENCE_LENGTH} 字符')

    try:
        page_count = int(data.get('page_count', 5))
        page_count = max(1, min(page_count, MAX_PAGE_COUNT))
    except (ValueError, TypeError):
        page_count = 5

    auto_download = bool(data.get('auto_download', False))
    description = (data.get('description') or '').strip()
    if len(description) > MAX_DESCRIPTION_LENGTH:
        description = description[:MAX_DESCRIPTION_LENGTH]

    auto_page_count = bool(data.get('auto_page_count', False))
    template_id = data.get('template_id', '')

    api_key = data.get('api_key', '')
    if not api_key:
        errors.append('请先配置 AI API Key')

    api_base = data.get('api_base', DEFAULT_API_BASE_URL)
    if api_base and not validate_api_url(api_base):
        errors.append('API URL 无效或不允许访问')

    model_name = data.get('model_name', DEFAULT_MODEL_NAME)
    if len(model_name) > MAX_MODEL_NAME_LENGTH:
        errors.append('模型名称过长')

    unsplash_key = data.get('unsplash_key', '')

    if errors:
        return None, (jsonify({'error': errors[0] if len(errors) == 1 else errors}), 400)

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


def validate_ppt_filepath(filename: str) -> tuple:
    """验证 PPT 文件路径安全性

    Returns:
        (filepath, error_response) - 如果验证通过，error_response 为 None
    """
    try:
        output_folder = Path(current_app.config['OUTPUT_FOLDER']).resolve()
        safe_filename = secure_filename(filename)

        if not safe_filename:
            return None, (jsonify({'success': False, 'error': '非法的件名'}), 403)

        filepath = (output_folder / safe_filename).resolve()

        if not filepath.is_relative_to(output_folder):
            return None, (jsonify({'success': False, 'error': '非法的文件路径'}), 403)

        if not filepath.exists():
            return None, (jsonify({'success': False, 'error': '文件不存在'}), 404)

        return str(filepath), None
    except Exception as e:
        return None, (jsonify({'success': False, 'error': str(e)}), 500)


def cleanup_old_files(folder: str, max_age_hours: int = FILE_CLEANUP_MAX_AGE_HOURS, max_files: int = FILE_CLEANUP_MAX_FILES):
    """清理旧文件，防止目录无限增长"""
    import time
    try:
        files = []
        for f in Path(folder).glob('*.pptx'):
            files.append((f, f.stat().st_mtime))

        files.sort(key=lambda x: x[1], reverse=True)

        now = time.time()
        for i, (filepath, mtime) in enumerate(files):
            age_hours = (now - mtime) / 3600
            if age_hours > max_age_hours or i >= max_files:
                try:
                    filepath.unlink()
                except OSError:
                    pass
    except OSError:
        pass
