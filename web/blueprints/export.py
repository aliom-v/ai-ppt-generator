"""导出蓝图 - PDF、图片导出"""

import os
from pathlib import Path
from flask import Blueprint, request, jsonify, current_app

from utils.logger import get_logger
from config.settings import CACHE_STATS_TTL

from .common import validate_ppt_filepath, cached_json_response

logger = get_logger("web.export")

export_bp = Blueprint('export', __name__, url_prefix='/api/export')


@export_bp.route('/formats')
def get_export_formats():
    """获取可用的导出格式"""
    from utils.export import get_export_manager

    manager = get_export_manager()
    data = {'success': True, **manager.get_status()}
    return cached_json_response(data, max_age=CACHE_STATS_TTL)


@export_bp.route('/<path:filename>/pdf', methods=['POST'])
def export_to_pdf(filename):
    """将 PPT 导出为 PDF"""
    from utils.export import get_export_manager, ExportError

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
        output_folder = Path(current_app.config['OUTPUT_FOLDER']).resolve()
        manager = get_export_manager()
        if not manager.is_format_available('pdf'):
            return jsonify({
                'success': False,
                'error': 'PDF 导出不可用（需要安装 LibreOffice）'
            }), 503

        pdf_filename = Path(filename).stem + '.pdf'
        output_path = str(output_folder / pdf_filename)

        manager.export(filepath, output_path, 'pdf')

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


@export_bp.route('/<path:filename>/images', methods=['POST'])
def export_to_images(filename):
    """将 PPT 导出为图片序列"""
    from utils.export import get_export_manager, ExportError
    import zipfile
    import tempfile

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
        output_folder = Path(current_app.config['OUTPUT_FOLDER']).resolve()
        manager = get_export_manager()
        if not manager.is_format_available('images'):
            return jsonify({
                'success': False,
                'error': '图片导出不可用（需要安装 LibreOffice 和 pdf2image）'
            }), 503

        with tempfile.TemporaryDirectory() as tmp_dir:
            manager.export(filepath, tmp_dir, 'images')

            zip_filename = Path(filename).stem + '_images.zip'
            zip_path = os.path.join(output_folder, zip_filename)

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for img_file in sorted(Path(tmp_dir).glob('*.png')):
                    zf.write(img_file, img_file.name)

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


@export_bp.route('/<path:filename>/thumbnail', methods=['POST'])
def export_thumbnail(filename):
    """生成 PPT 缩略图"""
    from utils.export import get_export_manager, ExportError

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
        output_folder = Path(current_app.config['OUTPUT_FOLDER']).resolve()
        manager = get_export_manager()
        if not manager.is_format_available('thumbnail'):
            return jsonify({
                'success': False,
                'error': '缩略图生成不可用'
            }), 503

        thumb_filename = Path(filename).stem + '_thumb.png'
        output_path = str(output_folder / thumb_filename)

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
