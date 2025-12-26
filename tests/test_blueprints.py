"""蓝图模块测试"""
import pytest


class TestBlueprintsCommon:
    """公共工具函数测试"""

    def test_sanitize_filename_basic(self):
        """测试基本文件名清理"""
        from web.blueprints.common import sanitize_filename
        assert sanitize_filename("test.pptx") == "test.pptx"
        assert sanitize_filename("测试文件.pptx") == "测试文件.pptx"

    def test_sanitize_filename_special_chars(self):
        """测试特殊字符清理"""
        from web.blueprints.common import sanitize_filename
        result = sanitize_filename("test<>file.pptx")
        assert "<" not in result
        assert ">" not in result

    def test_sanitize_filename_max_length(self):
        """测试文件名长度限制"""
        from web.blueprints.common import sanitize_filename
        long_name = "a" * 100 + ".pptx"
        result = sanitize_filename(long_name, max_length=50)
        assert len(result) <= 50

    def test_sanitize_filename_spaces(self):
        """测试空格替换"""
        from web.blueprints.common import sanitize_filename
        result = sanitize_filename("test file name.pptx")
        assert " " not in result
        assert "_" in result

    def test_validate_api_url_valid(self):
        """测试有效 API URL"""
        from web.blueprints.common import validate_api_url
        assert validate_api_url("https://api.openai.com/v1") is True
        assert validate_api_url("https://example.com/api") is True

    def test_validate_api_url_invalid_scheme(self):
        """测试无效协议"""
        from web.blueprints.common import validate_api_url
        assert validate_api_url("ftp://example.com") is False
        assert validate_api_url("file:///etc/passwd") is False

    def test_validate_api_url_localhost_blocked(self):
        """测试本地地址被阻止"""
        from web.blueprints.common import validate_api_url
        assert validate_api_url("http://localhost/api") is False
        assert validate_api_url("http://127.0.0.1/api") is False
        assert validate_api_url("http://0.0.0.0/api") is False

    def test_validate_api_url_private_ip_blocked(self):
        """测试内网地址被阻止"""
        from web.blueprints.common import validate_api_url
        assert validate_api_url("http://192.168.1.1/api") is False
        assert validate_api_url("http://10.0.0.1/api") is False
        assert validate_api_url("http://172.16.0.1/api") is False

    def test_validate_api_url_empty(self):
        """测试空 URL"""
        from web.blueprints.common import validate_api_url
        assert validate_api_url("") is False
        assert validate_api_url(None) is False

    def test_validate_generation_params_valid(self):
        """测试有效参数验证"""
        from flask import Flask
        from web.blueprints.common import validate_generation_params

        app = Flask(__name__)
        with app.app_context():
            data = {
                'topic': '测试主题',
                'audience': '通用受众',
                'page_count': 10,
                'api_key': 'test-key',
                'api_base': 'https://api.openai.com/v1'
            }
            params, error = validate_generation_params(data)
            assert error is None
            assert params['topic'] == '测试主题'
            assert params['page_count'] == 10

    def test_validate_generation_params_missing_topic(self):
        """测试缺少主题"""
        from flask import Flask
        from web.blueprints.common import validate_generation_params

        app = Flask(__name__)
        with app.app_context():
            data = {'api_key': 'test-key'}
            params, error = validate_generation_params(data)
            assert params is None
            assert error is not None

    def test_validate_generation_params_missing_api_key(self):
        """测试缺少 API Key"""
        from flask import Flask
        from web.blueprints.common import validate_generation_params

        app = Flask(__name__)
        with app.app_context():
            data = {'topic': '测试'}
            params, error = validate_generation_params(data)
            assert params is None
            assert error is not None

    def test_validate_generation_params_topic_too_long(self):
        """测试主题过长"""
        from flask import Flask
        from web.blueprints.common import validate_generation_params
        from config.settings import MAX_TOPIC_LENGTH

        app = Flask(__name__)
        with app.app_context():
            data = {
                'topic': 'x' * (MAX_TOPIC_LENGTH + 100),
                'api_key': 'test-key'
            }
            params, error = validate_generation_params(data)
            assert params is None
            assert error is not None

    def test_validate_generation_params_page_count_clamped(self):
        """测试页数被限制在有效范围"""
        from flask import Flask
        from web.blueprints.common import validate_generation_params
        from config.settings import MAX_PAGE_COUNT

        app = Flask(__name__)
        with app.app_context():
            data = {
                'topic': '测试',
                'api_key': 'test-key',
                'api_base': 'https://api.openai.com/v1',
                'page_count': 999
            }
            params, error = validate_generation_params(data)
            assert error is None
            assert params['page_count'] == MAX_PAGE_COUNT


class TestBlueprintsAPI:
    """API 蓝图测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_api_generate_endpoint_exists(self, client):
        """测试生成端点存在"""
        response = client.post('/api/generate', json={})
        # 应该返回 400（缺少参数）而不是 404
        assert response.status_code == 400

    def test_api_templates_endpoint(self, client):
        """测试模板端点"""
        response = client.get('/api/templates')
        assert response.status_code == 200
        data = response.get_json()
        assert 'templates' in data

    def test_api_config_endpoint(self, client):
        """测试配置端点"""
        response = client.get('/api/config')
        assert response.status_code == 200
        data = response.get_json()
        assert 'ai_configured' in data

    def test_api_history_endpoint(self, client):
        """测试历史记录端点"""
        response = client.get('/api/history')
        assert response.status_code == 200
        data = response.get_json()
        assert 'records' in data or 'success' in data


class TestBlueprintsTasks:
    """任务蓝图测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_tasks_list_endpoint(self, client):
        """测试任务列表端点"""
        response = client.get('/api/tasks')
        assert response.status_code == 200
        data = response.get_json()
        assert 'tasks' in data

    def test_tasks_async_generate_missing_params(self, client):
        """测试异步生成缺少参数"""
        response = client.post('/api/generate/async', json={})
        assert response.status_code == 400

    def test_tasks_stream_generate_missing_params(self, client):
        """测试流式生成缺少参数"""
        response = client.post('/api/generate/stream', json={})
        assert response.status_code == 400


class TestBlueprintsBatch:
    """批量生成蓝图测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_batch_list_endpoint(self, client):
        """测试批量任务列表端点"""
        response = client.get('/api/batch')
        assert response.status_code == 200
        data = response.get_json()
        assert 'jobs' in data

    def test_batch_create_missing_items(self, client):
        """测试创建批量任务缺少项目"""
        response = client.post('/api/batch', json={})
        assert response.status_code == 400
        data = response.get_json()
        assert 'error' in data

    def test_batch_create_missing_api_key(self, client):
        """测试创建批量任务缺少 API Key"""
        response = client.post('/api/batch', json={
            'items': [{'topic': '测试'}]
        })
        assert response.status_code == 400
        data = response.get_json()
        assert 'API Key' in data['error']

    def test_batch_create_too_many_items(self, client):
        """测试批量任务项目过多"""
        from config.settings import MAX_BATCH_ITEMS
        items = [{'topic': f'测试{i}'} for i in range(MAX_BATCH_ITEMS + 5)]
        response = client.post('/api/batch', json={
            'items': items,
            'api_key': 'test-key'
        })
        # 可能返回 400（项目过多）或 429（速率限制）
        assert response.status_code in [400, 429]
        if response.status_code == 400:
            data = response.get_json()
            assert str(MAX_BATCH_ITEMS) in data['error']


class TestBlueprintsExport:
    """导出蓝图测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_export_formats_endpoint(self, client):
        """测试导出格式端点"""
        response = client.get('/api/export/formats')
        assert response.status_code == 200
        data = response.get_json()
        assert 'success' in data

    def test_export_pdf_file_not_found(self, client):
        """测试导出 PDF 文件不存在"""
        response = client.post('/api/export/nonexistent.pptx/pdf')
        assert response.status_code == 404

    def test_export_images_file_not_found(self, client):
        """测试导出图片文件不存在"""
        response = client.post('/api/export/nonexistent.pptx/images')
        assert response.status_code == 404


class TestBlueprintsPPTEdit:
    """PPT 编辑蓝图测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_ppt_info_file_not_found(self, client):
        """测试获取 PPT 信息文件不存在"""
        response = client.get('/api/ppt/nonexistent.pptx/info')
        assert response.status_code == 404

    def test_ppt_slide_info_file_not_found(self, client):
        """测试获取幻灯片信息文件不存在"""
        response = client.get('/api/ppt/nonexistent.pptx/slide/0')
        assert response.status_code == 404

    def test_ppt_update_title_file_not_found(self, client):
        """测试更新标题文件不存在"""
        response = client.put('/api/ppt/nonexistent.pptx/slide/0/title', json={
            'title': '新标题'
        })
        assert response.status_code == 404
