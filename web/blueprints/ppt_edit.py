"""PPT 编辑蓝图"""

from flask import Blueprint, request, jsonify

from utils.logger import get_logger

from .common import validate_ppt_filepath

logger = get_logger("web.ppt_edit")

ppt_edit_bp = Blueprint('ppt_edit', __name__, url_prefix='/api/ppt')


@ppt_edit_bp.route('/<path:filename>/info')
def get_ppt_info(filename):
    """获取 PPT 文件信息"""
    from ppt.editor import PPTEditor

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
        editor = PPTEditor(filepath)
        return jsonify({'success': True, **editor.to_dict()})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@ppt_edit_bp.route('/<path:filename>/slide/<int:index>', methods=['GET'])
def get_slide_info(filename, index):
    """获取指定幻灯片信息"""
    from ppt.editor import PPTEditor

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
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


@ppt_edit_bp.route('/<path:filename>/slide/<int:index>/title', methods=['PUT'])
def update_slide_title(filename, index):
    """更新幻灯片标题"""
    from ppt.editor import PPTEditor

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
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


@ppt_edit_bp.route('/<path:filename>/slide/<int:index>', methods=['DELETE'])
def delete_slide(filename, index):
    """删除幻灯片"""
    from ppt.editor import PPTEditor

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
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


@ppt_edit_bp.route('/<path:filename>/reorder', methods=['POST'])
def reorder_slides(filename):
    """重新排列幻灯片顺序"""
    from ppt.editor import PPTEditor, reorder_slides as do_reorder

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
        data = request.json or {}
        new_order = data.get('order', [])

        if not new_order or not isinstance(new_order, list):
            return jsonify({'success': False, 'error': '请提供有效的顺序列表'}), 400

        editor = PPTEditor(filepath)

        slide_count = editor.get_slide_count()
        if len(new_order) != slide_count:
            return jsonify({
                'success': False,
                'error': f'顺序列表长度({len(new_order)})与幻灯片数量({slide_count})不匹配'
            }), 400

        if set(new_order) != set(range(slide_count)):
            return jsonify({'success': False, 'error': '顺序列表包含无效索引'}), 400

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


@ppt_edit_bp.route('/<path:filename>/slide/<int:index>/duplicate', methods=['POST'])
def duplicate_slide(filename, index):
    """复制幻灯片"""
    from ppt.editor import PPTEditor

    filepath, error = validate_ppt_filepath(filename)
    if error:
        return error

    try:
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
