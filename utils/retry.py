"""重试机制模块

提供可配置的重试装饰器，用于处理网络请求、API 调用等可能失败的操作。
"""
import time
import random
import functools
import logging
from typing import Callable, Tuple, Type, Union, Optional, Any

logger = logging.getLogger(__name__)


class RetryError(Exception):
    """重试失败异常"""

    def __init__(self, message: str, last_exception: Exception = None, attempts: int = 0):
        super().__init__(message)
        self.last_exception = last_exception
        self.attempts = attempts


def retry(
    max_attempts: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    max_delay: float = 60.0,
    jitter: bool = True,
    exceptions: Tuple[Type[Exception], ...] = (Exception,),
    on_retry: Optional[Callable[[Exception, int], None]] = None,
    should_retry: Optional[Callable[[Exception], bool]] = None,
):
    """重试装饰器

    当被装饰的函数抛出指定异常时，自动重试。

    Args:
        max_attempts: 最大尝试次数（包括首次）
        delay: 初始延迟时间（秒）
        backoff: 延迟增长倍数（指数退避）
        max_delay: 最大延迟时间
        jitter: 是否添加随机抖动（防止惊群效应）
        exceptions: 需要捕获并重试的异常类型
        on_retry: 重试时的回调函数 (exception, attempt_number)
        should_retry: 自定义判断是否应该重试的函数

    用法:
        @retry(max_attempts=3, delay=1.0)
        def fetch_data():
            response = requests.get("https://api.example.com/data")
            response.raise_for_status()
            return response.json()

        # 带自定义重试条件
        @retry(
            max_attempts=5,
            should_retry=lambda e: "rate limit" in str(e).lower()
        )
        def call_api():
            ...
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            last_exception = None
            current_delay = delay

            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e

                    # 检查是否应该重试
                    if should_retry and not should_retry(e):
                        raise

                    # 最后一次尝试不再重试
                    if attempt >= max_attempts:
                        break

                    # 计算延迟
                    wait_time = min(current_delay, max_delay)
                    if jitter:
                        wait_time = wait_time * (0.5 + random.random())

                    # 执行重试回调
                    if on_retry:
                        on_retry(e, attempt)
                    else:
                        logger.warning(
                            f"重试 {func.__name__} (尝试 {attempt}/{max_attempts}): {e}"
                        )

                    time.sleep(wait_time)
                    current_delay *= backoff

            raise RetryError(
                f"重试 {max_attempts} 次后仍然失败: {last_exception}",
                last_exception=last_exception,
                attempts=max_attempts,
            )

        return wrapper
    return decorator


def retry_async(
    max_attempts: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    max_delay: float = 60.0,
    jitter: bool = True,
    exceptions: Tuple[Type[Exception], ...] = (Exception,),
    on_retry: Optional[Callable[[Exception, int], None]] = None,
):
    """异步重试装饰器

    用于 async 函数的重试。

    用法:
        @retry_async(max_attempts=3)
        async def fetch_data():
            async with aiohttp.ClientSession() as session:
                async with session.get(url) as response:
                    return await response.json()
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        async def wrapper(*args, **kwargs) -> Any:
            import asyncio

            last_exception = None
            current_delay = delay

            for attempt in range(1, max_attempts + 1):
                try:
                    return await func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e

                    if attempt >= max_attempts:
                        break

                    wait_time = min(current_delay, max_delay)
                    if jitter:
                        wait_time = wait_time * (0.5 + random.random())

                    if on_retry:
                        on_retry(e, attempt)
                    else:
                        logger.warning(
                            f"异步重试 {func.__name__} (尝试 {attempt}/{max_attempts}): {e}"
                        )

                    await asyncio.sleep(wait_time)
                    current_delay *= backoff

            raise RetryError(
                f"重试 {max_attempts} 次后仍然失败: {last_exception}",
                last_exception=last_exception,
                attempts=max_attempts,
            )

        return wrapper
    return decorator


class RetryContext:
    """重试上下文管理器

    用于需要更细粒度控制的场景。

    用法:
        with RetryContext(max_attempts=3) as ctx:
            for attempt in ctx:
                try:
                    result = risky_operation()
                    break
                except NetworkError as e:
                    ctx.record_failure(e)
    """

    def __init__(
        self,
        max_attempts: int = 3,
        delay: float = 1.0,
        backoff: float = 2.0,
        max_delay: float = 60.0,
        jitter: bool = True,
    ):
        self.max_attempts = max_attempts
        self.delay = delay
        self.backoff = backoff
        self.max_delay = max_delay
        self.jitter = jitter

        self._attempt = 0
        self._last_exception: Optional[Exception] = None
        self._succeeded = False

    def __enter__(self) -> "RetryContext":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None and not self._succeeded:
            if self._attempt >= self.max_attempts:
                raise RetryError(
                    f"重试 {self.max_attempts} 次后仍然失败",
                    last_exception=self._last_exception or exc_val,
                    attempts=self._attempt,
                )
        return False

    def __iter__(self):
        return self

    def __next__(self) -> int:
        if self._succeeded:
            raise StopIteration

        self._attempt += 1

        if self._attempt > self.max_attempts:
            raise StopIteration

        if self._attempt > 1:
            # 计算延迟
            wait_time = self.delay * (self.backoff ** (self._attempt - 2))
            wait_time = min(wait_time, self.max_delay)
            if self.jitter:
                wait_time = wait_time * (0.5 + random.random())
            time.sleep(wait_time)

        return self._attempt

    def record_failure(self, exception: Exception):
        """记录失败"""
        self._last_exception = exception
        logger.warning(f"尝试 {self._attempt}/{self.max_attempts} 失败: {exception}")

    def success(self):
        """标记成功"""
        self._succeeded = True

    @property
    def attempt(self) -> int:
        """当前尝试次数"""
        return self._attempt

    @property
    def last_exception(self) -> Optional[Exception]:
        """最后一次异常"""
        return self._last_exception


# 预定义的重试配置
NETWORK_RETRY = {
    "max_attempts": 3,
    "delay": 1.0,
    "backoff": 2.0,
    "exceptions": (ConnectionError, TimeoutError),
}

API_RETRY = {
    "max_attempts": 5,
    "delay": 2.0,
    "backoff": 2.0,
    "max_delay": 30.0,
}

QUICK_RETRY = {
    "max_attempts": 2,
    "delay": 0.5,
    "backoff": 1.5,
}


def with_network_retry(func: Callable) -> Callable:
    """网络请求重试装饰器（使用预定义配置）"""
    return retry(**NETWORK_RETRY)(func)


def with_api_retry(func: Callable) -> Callable:
    """API 调用重试装饰器（使用预定义配置）"""
    return retry(**API_RETRY)(func)
