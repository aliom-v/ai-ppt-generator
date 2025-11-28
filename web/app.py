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
    except:
        pass

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from flask import Flask, render_template, request, jsonify, send_file, Response
from werkzeug.utils import secure_filename
import json
import hashlib
from datetime import datetime
from functools import wraps

from core.ai_client import generate_ppt_plan, AIClientError, JSONParseError, test_api_connection
from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from ppt.template_manager import template_manager
from utils.file_parser import parse_file, get_text_summary, validate_file
from config.settings import AIConfig, ImageConfig, AppConfig


# 应用配置
app_config = AppConfig.from_env()

app = Flask(__name__)
app.config['SECRET_KEY'] = app_config.secret_key
app.config['UPLOAD_FOLDER'] = app_config.upload_folder
app.config['OUTPUT_FOLDER'] = app_config.output_folder
app.config['MAX_CONTENT_LENGTH'] = app_config.max_upload_size
app.config['JSON_AS_ASCII'] = False

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
                except:
                    pass
    except:
        pass


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


@app.route('/')
def index():
    """首页"""
    return render_template('index.html')


@app.route('/api/generate', methods=['POST'])
@validate_request(['topic'])
def generate_ppt():
    """生成 PPT API"""
    try:
        data = request.json
        
        # 获取参数
        topic = data.get('topic', '').strip()
        audience = data.get('audience', '通用受众').strip() or '通用受众'
        
        # 验证页数范围
        try:
            page_count = int(data.get('page_count', 5))
            page_count = max(1, min(page_count, 100))  # 限制 1-100 页
        except (ValueError, TypeError):
            page_count = 5
        
        auto_download = bool(data.get('auto_download', False))
        description = (data.get('description') or '').strip()
        
        # 限制描述长度（防止过长的输入）
        if len(description) > 100000:
            description = description[:100000]
        
        auto_page_count = bool(data.get('auto_page_count', False))
        template_id = data.get('template_id', '')
        
        # 获取 API 配置
        api_key = data.get('api_key', '')
        if not api_key:
            return jsonify({'error': '请先配置 AI API Key'}), 400
        
        # 创建配置对象（不修改环境变量）
        ai_config = AIConfig(
            api_key=api_key,
            api_base_url=data.get('api_base', 'https://api.openai.com/v1'),
            model_name=data.get('model_name', 'gpt-4o-mini')
        )
        
        # 设置图片搜索配置（临时设置环境变量，后续可优化为传递配置对象）
        unsplash_key = data.get('unsplash_key', '')
        if unsplash_key:
            os.environ['UNSPLASH_ACCESS_KEY'] = unsplash_key
            # 重置全局搜索器以使用新的 API Key
            try:
                from utils.image_search import _searcher
                import utils.image_search as img_module
                img_module._searcher = None
            except:
                pass
        
        # 日志
        print(f"\n{'=' * 60}")
        print(f"📝 生成 PPT: {topic}")
        print(f"🎯 受众: {audience}")
        print(f"🤖 模型: {ai_config.model_name}")
        print(f"📄 页数: {'自动' if auto_page_count else page_count}")
        print(f"{'=' * 60}\n")
        
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
        
        return jsonify({
            'success': True,
            'filename': filename,
            'title': plan.title,
            'subtitle': plan.subtitle,
            'slide_count': len(plan.slides),
            'download_url': f'/api/download/{filename}'
        })
        
    except JSONParseError as e:
        return jsonify({'error': str(e), 'type': 'json_error'}), 400
    except AIClientError as e:
        return jsonify({'error': str(e), 'type': 'ai_error'}), 500
    except Exception as e:
        print(f"生成失败: {e}")
        import traceback
        traceback.print_exc()
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
    """获取配置信息"""
    return jsonify({
        'ai_configured': bool(os.getenv('AI_API_KEY')),
        'image_search_available': bool(os.getenv('UNSPLASH_ACCESS_KEY'))
    })


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
                except:
                    pass
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/templates')
def get_templates():
    """获取所有可用模板"""
    try:
        templates = template_manager.list_templates()
        return jsonify({
            'success': True,
            'templates': templates,
            'count': len(templates)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'templates': []})


@app.route('/health')
def health():
    """健康检查"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'service': 'ai-ppt-generator'
    }), 200


if __name__ == '__main__':
    print(f"\n{'=' * 60}")
    print("🚀 AI PPT 生成器 - Web 界面")
    print(f"{'=' * 60}")
    print(f"\n🌐 访问地址: http://localhost:{app_config.port}")
    print("💡 按 Ctrl+C 停止服务器\n")
    
    app.run(debug=app_config.debug, host=app_config.host, port=app_config.port)
