"""新模块单元测试"""
import pytest
import json
import time
from unittest.mock import Mock, patch


class TestValidators:
    """验证器测试"""

    def test_required_validator_empty(self):
        """测试必填验证 - 空值"""
        from utils.validators import Validators
        error = Validators.required("", "topic")
        assert error is not None
        assert error.code == "REQUIRED"

    def test_required_validator_valid(self):
        """测试必填验证 - 有效值"""
        from utils.validators import Validators
        error = Validators.required("test", "topic")
        assert error is None

    def test_string_length_too_long(self):
        """测试字符串长度 - 过长"""
        from utils.validators import Validators
        error = Validators.string_length("x" * 600, "topic", max_length=500)
        assert error is not None
        assert error.code == "TOO_LONG"

    def test_string_length_valid(self):
        """测试字符串长度 - 有效"""
        from utils.validators import Validators
        error = Validators.string_length("valid topic", "topic", max_length=500)
        assert error is None

    def test_url_validator_invalid_scheme(self):
        """测试 URL 验证 - 无效协议"""
        from utils.validators import Validators
        error = Validators.url("ftp://example.com", "api_base")
        assert error is not None
        assert error.code == "INVALID_SCHEME"

    def test_url_validator_localhost_blocked(self):
        """测试 URL 验证 - 本地地址被阻止"""
        from utils.validators import Validators
        error = Validators.url("http://localhost:8080", "api_base")
        assert error is not None
        assert error.code == "BLOCKED_HOST"

    def test_url_validator_valid(self):
        """测试 URL 验证 - 有效"""
        from utils.validators import Validators
        error = Validators.url("https://api.openai.com/v1", "api_base")
        assert error is None

    def test_integer_range(self):
        """测试整数范围验证"""
        from utils.validators import Validators
        value, error = Validators.integer_range(5, "page_count", min_value=1, max_value=100)
        assert error is None
        assert value == 5

    def test_integer_range_clamped(self):
        """测试整数范围 - 超出范围被截断"""
        from utils.validators import Validators
        value, error = Validators.integer_range(200, "page_count", min_value=1, max_value=100)
        assert error is None
        assert value == 100

    def test_request_validator_chain(self):
        """测试请求验证器链式调用"""
        from utils.validators import RequestValidator

        data = {
            "topic": "测试主题",
            "page_count": "10",
            "api_key": "sk-test123456789",
        }

        validator = RequestValidator(data)
        validator.require("topic").string(max_length=500)
        validator.optional("page_count").integer(min_value=1, max_value=100, default=5)
        validator.require("api_key").api_key()

        assert validator.is_valid
        assert validator.data["topic"] == "测试主题"
        assert validator.data["page_count"] == 10


class TestRetry:
    """重试机制测试"""

    def test_retry_success_first_try(self):
        """测试重试 - 首次成功"""
        from utils.retry import retry

        call_count = 0

        @retry(max_attempts=3)
        def success_func():
            nonlocal call_count
            call_count += 1
            return "success"

        result = success_func()
        assert result == "success"
        assert call_count == 1

    def test_retry_success_after_failures(self):
        """测试重试 - 失败后成功"""
        from utils.retry import retry

        call_count = 0

        @retry(max_attempts=3, delay=0.01)
        def flaky_func():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise ValueError("temporary error")
            return "success"

        result = flaky_func()
        assert result == "success"
        assert call_count == 3

    def test_retry_all_failures(self):
        """测试重试 - 全部失败"""
        from utils.retry import retry, RetryError

        @retry(max_attempts=3, delay=0.01)
        def always_fail():
            raise ValueError("always fails")

        with pytest.raises(RetryError) as exc_info:
            always_fail()

        assert exc_info.value.attempts == 3

    def test_retry_context_manager(self):
        """测试重试上下文管理器"""
        from utils.retry import RetryContext

        results = []
        with RetryContext(max_attempts=3, delay=0.01) as ctx:
            for attempt in ctx:
                results.append(attempt)
                if attempt < 3:
                    ctx.record_failure(ValueError("temp"))
                else:
                    ctx.success()
                    break

        assert results == [1, 2, 3]


class TestSSE:
    """SSE 模块测试"""

    def test_sse_event_serialize(self):
        """测试 SSE 事件序列化"""
        from utils.sse import SSEEvent

        event = SSEEvent(
            event="progress",
            data={"progress": 50, "message": "处理中"},
            id="123",
        )
        serialized = event.serialize()

        assert "id: 123" in serialized
        assert "event: progress" in serialized
        assert "data:" in serialized
        assert serialized.endswith("\n\n")

    def test_sse_channel(self):
        """测试 SSE 通道"""
        from utils.sse import SSEChannel

        channel = SSEChannel("test-123")
        assert channel.is_active

        channel.send_progress(50, "处理中")
        channel.close()
        assert not channel.is_active

    def test_sse_manager(self):
        """测试 SSE 管理器"""
        from utils.sse import SSEManager

        manager = SSEManager(max_channels=10)
        channel = manager.create_channel("test-1")

        assert channel is not None
        assert manager.get_channel("test-1") is not None
        assert manager.get_channel("nonexistent") is None

        manager.send_to("test-1", "test", {"message": "hello"})
        stats = manager.get_stats()
        assert stats["total_channels"] == 1


class TestAsyncTasks:
    """异步任务测试"""

    def test_task_creation(self):
        """测试任务创建"""
        from utils.async_tasks import TaskManager, TaskStatus

        manager = TaskManager()
        task_id = manager.create_task()

        task = manager.get_task(task_id)
        assert task is not None
        assert task.status == TaskStatus.PENDING

    def test_task_update(self):
        """测试任务更新"""
        from utils.async_tasks import TaskManager, TaskStatus

        manager = TaskManager()
        task_id = manager.create_task()

        manager.update_task(task_id, status=TaskStatus.RUNNING, progress=50)

        task = manager.get_task(task_id)
        assert task.status == TaskStatus.RUNNING
        assert task.progress == 50

    def test_task_cancel(self):
        """测试任务取消"""
        from utils.async_tasks import TaskManager, TaskStatus

        manager = TaskManager()
        task_id = manager.create_task()

        assert manager.cancel_task(task_id) is True

        task = manager.get_task(task_id)
        assert task.status == TaskStatus.CANCELLED


class TestBatch:
    """批量生成测试"""

    def test_batch_job_creation(self):
        """测试批量任务创建"""
        from utils.batch import BatchGenerator, BatchStatus

        generator = BatchGenerator()
        job = generator.create_job(
            items=[
                {"topic": "主题1"},
                {"topic": "主题2"},
            ],
            api_config={"api_key": "test"},
        )

        assert job.total == 2
        assert job.status == BatchStatus.PENDING

    def test_batch_job_to_dict(self):
        """测试批量任务序列化"""
        from utils.batch import BatchGenerator

        generator = BatchGenerator()
        job = generator.create_job(
            items=[{"topic": "测试"}],
            api_config={"api_key": "test"},
        )

        data = job.to_dict()
        assert "job_id" in data
        assert "items" in data
        assert data["total"] == 1


class TestConfigValidator:
    """配置验证测试"""

    def test_validate_config_success(self):
        """测试配置验证 - 成功"""
        from dataclasses import dataclass
        from utils.config_validator import validate_config, config_field

        @dataclass
        class TestConfig:
            host: str = config_field(default="localhost")
            port: int = config_field(default=8080, min=1, max=65535)

        config = validate_config(TestConfig, {"port": 3000})
        assert config.host == "localhost"
        assert config.port == 3000

    def test_validate_config_failure(self):
        """测试配置验证 - 失败"""
        from dataclasses import dataclass
        from utils.config_validator import validate_config, config_field, ConfigValidationError

        @dataclass
        class TestConfig:
            name: str = config_field(required=True)

        with pytest.raises(ConfigValidationError):
            validate_config(TestConfig, {})


class TestStructuredLogging:
    """结构化日志测试"""

    def test_json_formatter(self):
        """测试 JSON 格式化器"""
        import logging
        from utils.structured_logging import JSONFormatter

        formatter = JSONFormatter()
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="Test message",
            args=(),
            exc_info=None,
        )

        output = formatter.format(record)
        data = json.loads(output)

        assert data["level"] == "INFO"
        assert data["logger"] == "test"
        assert data["message"] == "Test message"

    def test_structured_logger(self):
        """测试结构化日志器"""
        from utils.structured_logging import StructuredLogger

        logger = StructuredLogger("test")
        bound = logger.bind(user_id=123)

        # 不会抛出异常
        bound.info("测试消息", action="test")


class TestOpenAPI:
    """OpenAPI 文档测试"""

    def test_openapi_spec_generation(self):
        """测试 OpenAPI 规范生成"""
        from utils.openapi import OpenAPIGenerator, APIEndpoint, APIResponse

        generator = OpenAPIGenerator(title="Test API", version="1.0.0")
        generator.register_endpoint(APIEndpoint(
            path="/test",
            method="GET",
            summary="测试端点",
            responses=[APIResponse(200, "成功")],
        ))

        spec = generator.generate_spec()

        assert spec["openapi"] == "3.0.3"
        assert spec["info"]["title"] == "Test API"
        assert "/test" in spec["paths"]
        assert "get" in spec["paths"]["/test"]
