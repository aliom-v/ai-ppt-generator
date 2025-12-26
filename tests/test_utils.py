"""工具模块测试"""
import pytest
import time


class TestRequestContext:
    """请求上下文测试"""

    def test_generate_request_id(self):
        """测试请求 ID 生成"""
        from utils.request_context import generate_request_id
        id1 = generate_request_id()
        id2 = generate_request_id()
        assert len(id1) == 8
        assert id1 != id2

    def test_set_get_request_id(self):
        """测试设置和获取请求 ID"""
        from utils.request_context import set_request_id, get_request_id, clear_request_id

        set_request_id("test123")
        assert get_request_id() == "test123"

        clear_request_id()
        assert get_request_id() is None


class TestAPIResponse:
    """API 响应测试"""

    @pytest.fixture
    def app_context(self):
        """创建 Flask 应用上下文"""
        from flask import Flask
        app = Flask(__name__)
        app.config['TESTING'] = True
        return app

    def test_success_response(self, app_context):
        """测试成功响应"""
        from utils.api_response import APIResponse
        from utils.request_context import set_request_id, clear_request_id

        with app_context.app_context():
            set_request_id("test123")
            try:
                response, status = APIResponse.success(data={"key": "value"})
                data = response.get_json()
                assert data["success"] is True
                assert data["data"]["key"] == "value"
                assert data["request_id"] == "test123"
                assert status == 200
            finally:
                clear_request_id()

    def test_error_response(self, app_context):
        """测试错误响应"""
        from utils.api_response import APIResponse
        from utils.request_context import set_request_id, clear_request_id

        with app_context.app_context():
            set_request_id("test456")
            try:
                response, status = APIResponse.error(
                    message="测试错误",
                    code="TEST_ERROR",
                    status_code=400
                )
                data = response.get_json()
                assert data["success"] is False
                assert data["error"]["message"] == "测试错误"
                assert data["error"]["code"] == "TEST_ERROR"
                assert status == 400
            finally:
                clear_request_id()

    def test_validation_error(self, app_context):
        """测试验证错误响应"""
        from utils.api_response import APIResponse

        with app_context.app_context():
            response, status = APIResponse.validation_error("字段无效", field="name")
            data = response.get_json()
            assert data["error"]["code"] == "VALIDATION_ERROR"
            assert data["error"]["details"]["field"] == "name"
            assert status == 400


class TestRateLimiter:
    """速率限制测试"""

    def test_rate_limiter_creation(self):
        """测试速率限制器创建"""
        from utils.rate_limit import RateLimiter
        limiter = RateLimiter(requests_per_minute=5, requests_per_hour=50)
        assert limiter.requests_per_minute == 5
        assert limiter.requests_per_hour == 50


class TestCache:
    """缓存测试"""

    def test_cache_set_get(self, temp_dir):
        """测试缓存存取"""
        from utils.cache import GenerationCache
        cache = GenerationCache(cache_dir=temp_dir)

        test_data = {"title": "测试", "slides": []}
        cache.set("测试主题", "通用受众", 5, test_data)

        result = cache.get("测试主题", "通用受众", 5)
        assert result is not None
        assert result["title"] == "测试"

    def test_cache_miss(self, temp_dir):
        """测试缓存未命中"""
        from utils.cache import GenerationCache
        cache = GenerationCache(cache_dir=temp_dir)

        result = cache.get("不存在的主题", "受众", 5)
        assert result is None

    def test_cache_clear(self, temp_dir):
        """测试清空缓存"""
        from utils.cache import GenerationCache
        cache = GenerationCache(cache_dir=temp_dir)

        cache.set("主题1", "受众", 5, {"title": "1"})
        cache.set("主题2", "受众", 5, {"title": "2"})

        count = cache.clear()
        assert count == 2

        assert cache.get("主题1", "受众", 5) is None


class TestImageSearch:
    """图片搜索测试"""

    def test_image_cache(self, temp_dir):
        """测试图片缓存"""
        from utils.image_search import ImageCache
        cache = ImageCache(cache_dir=temp_dir)

        cache.set("test_keyword", "/path/to/image.jpg")
        # 由于文件不存在，get 应该返回 None
        assert cache.get("test_keyword") is None

    def test_searcher_without_api_key(self):
        """测试无 API Key 时的搜索器"""
        from utils.image_search import ImageSearcher
        from config.settings import ImageConfig

        config = ImageConfig(unsplash_key="")
        searcher = ImageSearcher(config)

        # 无 API Key 时应返回空列表
        results = searcher.search_images("test")
        assert results == []


class TestSSRFProtection:
    """SSRF 防护测试"""

    def test_validate_api_url_valid(self):
        """测试有效 URL"""
        from web.blueprints.common import validate_api_url

        assert validate_api_url("https://api.openai.com/v1") is True
        assert validate_api_url("https://example.com/api") is True

    def test_validate_api_url_localhost(self):
        """测试本地地址拦截"""
        from web.blueprints.common import validate_api_url

        assert validate_api_url("http://localhost:8080") is False
        assert validate_api_url("http://127.0.0.1:8080") is False
        assert validate_api_url("http://0.0.0.0:8080") is False

    def test_validate_api_url_private_network(self):
        """测试内网地址拦截"""
        from web.blueprints.common import validate_api_url

        assert validate_api_url("http://192.168.1.1/api") is False
        assert validate_api_url("http://10.0.0.1/api") is False
        assert validate_api_url("http://172.16.0.1/api") is False
        assert validate_api_url("http://172.31.255.255/api") is False

        # 172.32.x.x 不是内网地址
        assert validate_api_url("http://172.32.0.1/api") is True

    def test_validate_api_url_invalid(self):
        """测试无效 URL"""
        from web.blueprints.common import validate_api_url

        assert validate_api_url("") is False
        assert validate_api_url("ftp://example.com") is False
        assert validate_api_url("http://example.local/api") is False


class TestScheduler:
    """调度器测试"""

    def test_scheduler_creation(self):
        """测试调度器创建"""
        from utils.scheduler import BackgroundScheduler
        scheduler = BackgroundScheduler()
        assert scheduler._running is False

    def test_add_task(self):
        """测试添加任务"""
        from utils.scheduler import BackgroundScheduler
        scheduler = BackgroundScheduler()

        def dummy_task():
            pass

        scheduler.add_task(dummy_task, interval_hours=1)
        assert len(scheduler._tasks) == 1
