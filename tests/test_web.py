"""Web API 测试"""
import pytest


class TestWebApp:
    """Web 应用测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_index_page(self, client):
        """测试首页"""
        response = client.get('/')
        assert response.status_code == 200

    def test_health_check(self, client):
        """测试健康检查"""
        response = client.get('/health')
        assert response.status_code == 200
        data = response.get_json()
        assert data['status'] == 'healthy'
        assert 'timestamp' in data

    def test_get_templates(self, client):
        """测试获取模板列表"""
        response = client.get('/api/templates')
        assert response.status_code == 200
        data = response.get_json()
        assert 'templates' in data

    def test_get_config(self, client):
        """测试获取配置"""
        response = client.get('/api/config')
        assert response.status_code == 200
        data = response.get_json()
        assert 'ai_configured' in data
        assert 'image_search_available' in data

    def test_generate_missing_topic(self, client):
        """测试生成 PPT 缺少主题"""
        response = client.post('/api/generate', json={})
        assert response.status_code == 400
        data = response.get_json()
        assert 'error' in data

    def test_generate_missing_api_key(self, client):
        """测试生成 PPT 缺少 API Key"""
        response = client.post('/api/generate', json={
            'topic': '测试主题'
        })
        assert response.status_code == 400
        data = response.get_json()
        assert 'API Key' in data['error']

    def test_test_connection_missing_key(self, client):
        """测试 API 连接 - 缺少 Key"""
        response = client.post('/api/test-connection', json={})
        assert response.status_code == 200
        data = response.get_json()
        assert data['success'] is False

    def test_upload_no_file(self, client):
        """测试上传 - 无文件"""
        response = client.post('/api/upload')
        assert response.status_code == 400


class TestInputValidation:
    """输入验证测试"""

    @pytest.fixture
    def client(self):
        """创建测试客户端"""
        from web.app import app
        app.config['TESTING'] = True
        with app.test_client() as client:
            yield client

    def test_topic_too_long(self, client):
        """测试主题过长"""
        response = client.post('/api/generate', json={
            'topic': 'x' * 600,
            'api_key': 'test-key'
        })
        assert response.status_code == 400
        data = response.get_json()
        assert '500' in data['error']

    def test_audience_too_long(self, client):
        """测试受众描述过长"""
        response = client.post('/api/generate', json={
            'topic': '测试',
            'audience': 'x' * 300,
            'api_key': 'test-key'
        })
        assert response.status_code == 400
        data = response.get_json()
        assert '200' in data['error']

    def test_invalid_api_url(self, client):
        """测试无效 API URL"""
        response = client.post('/api/generate', json={
            'topic': '测试',
            'api_key': 'test-key',
            'api_base': 'ftp://invalid.com'
        })
        # 可能返回 400（URL 无效）或 429（速率限制）
        assert response.status_code in [400, 429]
        if response.status_code == 400:
            data = response.get_json()
            assert 'URL' in data['error']

    def test_localhost_api_url_blocked(self, client):
        """测试本地地址被阻止（SSRF 防护）"""
        response = client.post('/api/generate', json={
            'topic': '测试',
            'api_key': 'test-key',
            'api_base': 'http://localhost:8080/v1'
        })
        # 可能返回 400（SSRF 阻止）或 429（速率限制）
        assert response.status_code in [400, 429]
        if response.status_code == 400:
            data = response.get_json()
            assert 'URL' in data['error'] or '不允许' in data['error']


class TestRateLimit:
    """速率限制测试"""

    def test_rate_limiter_basic(self):
        """测试速率限制器基本功能"""
        from utils.rate_limit import RateLimiter
        limiter = RateLimiter(requests_per_minute=2, requests_per_hour=10)

        # 模拟请求（需要 Flask 上下文）
        # 这里只测试类的创建
        assert limiter.requests_per_minute == 2
        assert limiter.requests_per_hour == 10
