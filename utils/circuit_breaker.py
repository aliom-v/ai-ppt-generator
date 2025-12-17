"""熔断器模块

实现熔断器模式，防止 API 调用失败时的级联故障。
"""
import time
import threading
from enum import Enum
from typing import Any, Callable, Optional, TypeVar
from dataclasses import dataclass, field
from functools import wraps

from utils.logger import get_logger

logger = get_logger("circuit_breaker")

T = TypeVar("T")


class CircuitState(str, Enum):
    """熔断器状态"""
    CLOSED = "closed"      # 正常，允许请求通过
    OPEN = "open"          # 熔断，拒绝所有请求
    HALF_OPEN = "half_open"  # 半开，允许部分请求测试


class CircuitBreakerError(Exception):
    """熔断器错误"""

    def __init__(self, message: str, state: CircuitState):
        super().__init__(message)
        self.state = state


@dataclass
class CircuitStats:
    """熔断器统计"""
    total_calls: int = 0
    successful_calls: int = 0
    failed_calls: int = 0
    rejected_calls: int = 0
    last_failure_time: Optional[float] = None
    consecutive_failures: int = 0
    consecutive_successes: int = 0

    def record_success(self):
        self.total_calls += 1
        self.successful_calls += 1
        self.consecutive_successes += 1
        self.consecutive_failures = 0

    def record_failure(self):
        self.total_calls += 1
        self.failed_calls += 1
        self.consecutive_failures += 1
        self.consecutive_successes = 0
        self.last_failure_time = time.time()

    def record_rejected(self):
        self.rejected_calls += 1

    def to_dict(self):
        return {
            "total_calls": self.total_calls,
            "successful_calls": self.successful_calls,
            "failed_calls": self.failed_calls,
            "rejected_calls": self.rejected_calls,
            "failure_rate": round(self.failed_calls / max(self.total_calls, 1) * 100, 2),
            "consecutive_failures": self.consecutive_failures,
        }


class CircuitBreaker:
    """熔断器

    用于保护对外部服务的调用，防止级联故障。

    状态转换:
    - CLOSED -> OPEN: 连续失败次数超过阈值
    - OPEN -> HALF_OPEN: 超时时间过后
    - HALF_OPEN -> CLOSED: 测试请求成功
    - HALF_OPEN -> OPEN: 测试请求失败

    用法:
        breaker = CircuitBreaker("ai_api", failure_threshold=5)

        @breaker
        def call_ai_api():
            return requests.post(...)

        # 或者
        with breaker:
            result = call_ai_api()
    """

    def __init__(
        self,
        name: str,
        failure_threshold: int = 5,
        success_threshold: int = 2,
        timeout: float = 60.0,
        excluded_exceptions: tuple = (),
    ):
        """
        Args:
            name: 熔断器名称
            failure_threshold: 触发熔断的连续失败次数
            success_threshold: 半开状态下恢复需要的连续成功次数
            timeout: 熔断状态持续时间（秒）
            excluded_exceptions: 不计入失败的异常类型
        """
        self.name = name
        self.failure_threshold = failure_threshold
        self.success_threshold = success_threshold
        self.timeout = timeout
        self.excluded_exceptions = excluded_exceptions

        self._state = CircuitState.CLOSED
        self._stats = CircuitStats()
        self._lock = threading.RLock()
        self._opened_at: Optional[float] = None

    @property
    def state(self) -> CircuitState:
        """获取当前状态（自动检查是否应该转换到半开状态）"""
        with self._lock:
            if self._state == CircuitState.OPEN:
                if self._should_try_reset():
                    self._state = CircuitState.HALF_OPEN
                    logger.info(f"熔断器 {self.name}: OPEN -> HALF_OPEN")
            return self._state

    @property
    def stats(self) -> CircuitStats:
        return self._stats

    def _should_try_reset(self) -> bool:
        """是否应该尝试重置（超时后）"""
        if self._opened_at is None:
            return False
        return time.time() - self._opened_at >= self.timeout

    def _trip(self):
        """触发熔断"""
        self._state = CircuitState.OPEN
        self._opened_at = time.time()
        logger.warning(f"熔断器 {self.name}: CLOSED -> OPEN (连续失败 {self._stats.consecutive_failures} 次)")

    def _reset(self):
        """重置熔断器"""
        self._state = CircuitState.CLOSED
        self._opened_at = None
        self._stats.consecutive_failures = 0
        logger.info(f"熔断器 {self.name}: -> CLOSED (恢复正常)")

    def _handle_success(self):
        """处理成功调用"""
        with self._lock:
            self._stats.record_success()

            if self._state == CircuitState.HALF_OPEN:
                if self._stats.consecutive_successes >= self.success_threshold:
                    self._reset()

    def _handle_failure(self, exc: Exception):
        """处理失败调用"""
        # 检查是否是排除的异常
        if isinstance(exc, self.excluded_exceptions):
            return

        with self._lock:
            self._stats.record_failure()

            if self._state == CircuitState.CLOSED:
                if self._stats.consecutive_failures >= self.failure_threshold:
                    self._trip()
            elif self._state == CircuitState.HALF_OPEN:
                self._trip()

    def _can_execute(self) -> bool:
        """检查是否可以执行"""
        state = self.state  # 这会自动检查超时
        if state == CircuitState.OPEN:
            self._stats.record_rejected()
            return False
        return True

    def __call__(self, func: Callable[..., T]) -> Callable[..., T]:
        """装饰器用法"""
        @wraps(func)
        def wrapper(*args, **kwargs) -> T:
            if not self._can_execute():
                raise CircuitBreakerError(
                    f"熔断器 {self.name} 已打开，请稍后重试",
                    self._state,
                )

            try:
                result = func(*args, **kwargs)
                self._handle_success()
                return result
            except Exception as e:
                self._handle_failure(e)
                raise

        return wrapper

    def __enter__(self):
        """上下文管理器用法"""
        if not self._can_execute():
            raise CircuitBreakerError(
                f"熔断器 {self.name} 已打开，请稍后重试",
                self._state,
            )
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self._handle_success()
        elif exc_val is not None:
            self._handle_failure(exc_val)
        return False

    def execute(self, func: Callable[..., T], *args, **kwargs) -> T:
        """手动执行"""
        if not self._can_execute():
            raise CircuitBreakerError(
                f"熔断器 {self.name} 已打开，请稍后重试",
                self._state,
            )

        try:
            result = func(*args, **kwargs)
            self._handle_success()
            return result
        except Exception as e:
            self._handle_failure(e)
            raise

    def force_open(self):
        """强制打开熔断器"""
        with self._lock:
            self._state = CircuitState.OPEN
            self._opened_at = time.time()
            logger.warning(f"熔断器 {self.name}: 强制打开")

    def force_close(self):
        """强制关闭熔断器"""
        with self._lock:
            self._reset()
            logger.info(f"熔断器 {self.name}: 强制关闭")

    def get_status(self) -> dict:
        """获取状态信息"""
        return {
            "name": self.name,
            "state": self.state.value,
            "stats": self._stats.to_dict(),
            "failure_threshold": self.failure_threshold,
            "timeout": self.timeout,
        }


class CircuitBreakerRegistry:
    """熔断器注册表

    管理多个熔断器实例。
    """

    def __init__(self):
        self._breakers: dict[str, CircuitBreaker] = {}
        self._lock = threading.RLock()

    def get_or_create(
        self,
        name: str,
        **kwargs,
    ) -> CircuitBreaker:
        """获取或创建熔断器"""
        with self._lock:
            if name not in self._breakers:
                self._breakers[name] = CircuitBreaker(name, **kwargs)
            return self._breakers[name]

    def get(self, name: str) -> Optional[CircuitBreaker]:
        """获取熔断器"""
        return self._breakers.get(name)

    def get_all_status(self) -> list[dict]:
        """获取所有熔断器状态"""
        return [breaker.get_status() for breaker in self._breakers.values()]


# 全局注册表
_registry = CircuitBreakerRegistry()


def get_circuit_breaker(name: str, **kwargs) -> CircuitBreaker:
    """获取熔断器（全局注册表）"""
    return _registry.get_or_create(name, **kwargs)


def circuit_breaker(
    name: str,
    failure_threshold: int = 5,
    success_threshold: int = 2,
    timeout: float = 60.0,
):
    """熔断器装饰器

    用法:
        @circuit_breaker("ai_api", failure_threshold=3)
        def call_ai():
            ...
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        breaker = get_circuit_breaker(
            name,
            failure_threshold=failure_threshold,
            success_threshold=success_threshold,
            timeout=timeout,
        )
        return breaker(func)
    return decorator


# 预定义的熔断器
AI_API_BREAKER = "ai_api"
IMAGE_API_BREAKER = "image_api"


def get_ai_breaker() -> CircuitBreaker:
    """获取 AI API 熔断器"""
    return get_circuit_breaker(
        AI_API_BREAKER,
        failure_threshold=5,
        timeout=60.0,
    )


def get_image_breaker() -> CircuitBreaker:
    """获取图片 API 熔断器"""
    return get_circuit_breaker(
        IMAGE_API_BREAKER,
        failure_threshold=3,
        timeout=30.0,
    )
