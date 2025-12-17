"""速率限制中间件"""
import time
from collections import defaultdict
from functools import wraps
from threading import Lock
from typing import Callable, Optional, Tuple

from flask import request, jsonify


class RateLimiter:
    """基于令牌桶算法的速率限制器"""

    def __init__(
        self,
        requests_per_minute: int = 10,
        requests_per_hour: int = 100,
        burst_size: int = 5
    ):
        """
        初始化速率限制器

        Args:
            requests_per_minute: 每分钟最大请求数
            requests_per_hour: 每小时最大请求数
            burst_size: 突发请求容量
        """
        self.requests_per_minute = requests_per_minute
        self.requests_per_hour = requests_per_hour
        self.burst_size = burst_size

        # 存储每个 IP 的请求记录
        self._minute_counts: dict = defaultdict(list)
        self._hour_counts: dict = defaultdict(list)
        self._lock = Lock()

    def _get_client_ip(self) -> str:
        """获取客户端 IP"""
        # 支持反向代理
        if request.headers.get('X-Forwarded-For'):
            return request.headers.get('X-Forwarded-For').split(',')[0].strip()
        if request.headers.get('X-Real-IP'):
            return request.headers.get('X-Real-IP')
        return request.remote_addr or '127.0.0.1'

    def _cleanup_old_requests(self, client_ip: str, current_time: float) -> None:
        """清理过期的请求记录"""
        minute_ago = current_time - 60
        hour_ago = current_time - 3600

        # 清理分钟级记录
        self._minute_counts[client_ip] = [
            t for t in self._minute_counts[client_ip] if t > minute_ago
        ]

        # 清理小时级记录
        self._hour_counts[client_ip] = [
            t for t in self._hour_counts[client_ip] if t > hour_ago
        ]

    def check_rate_limit(self) -> Tuple[bool, Optional[str], Optional[int]]:
        """
        检查是否超出速率限制

        Returns:
            (is_allowed, error_message, retry_after_seconds)
        """
        client_ip = self._get_client_ip()
        current_time = time.time()

        with self._lock:
            self._cleanup_old_requests(client_ip, current_time)

            minute_count = len(self._minute_counts[client_ip])
            hour_count = len(self._hour_counts[client_ip])

            # 检查分钟限制
            if minute_count >= self.requests_per_minute:
                oldest = min(self._minute_counts[client_ip]) if self._minute_counts[client_ip] else current_time
                retry_after = int(60 - (current_time - oldest)) + 1
                return False, f"请求过于频繁，请 {retry_after} 秒后重试", retry_after

            # 检查小时限制
            if hour_count >= self.requests_per_hour:
                oldest = min(self._hour_counts[client_ip]) if self._hour_counts[client_ip] else current_time
                retry_after = int(3600 - (current_time - oldest)) + 1
                return False, f"已达到每小时请求上限，请 {retry_after // 60} 分钟后重试", retry_after

            # 记录请求
            self._minute_counts[client_ip].append(current_time)
            self._hour_counts[client_ip].append(current_time)

            return True, None, None

    def get_remaining(self) -> dict:
        """获取剩余配额"""
        client_ip = self._get_client_ip()
        current_time = time.time()

        with self._lock:
            self._cleanup_old_requests(client_ip, current_time)

            return {
                'minute_remaining': max(0, self.requests_per_minute - len(self._minute_counts[client_ip])),
                'hour_remaining': max(0, self.requests_per_hour - len(self._hour_counts[client_ip])),
                'minute_limit': self.requests_per_minute,
                'hour_limit': self.requests_per_hour,
            }


# 全局速率限制器实例
_rate_limiter: Optional[RateLimiter] = None


def get_rate_limiter(
    requests_per_minute: int = 10,
    requests_per_hour: int = 100
) -> RateLimiter:
    """获取全局速率限制器"""
    global _rate_limiter
    if _rate_limiter is None:
        _rate_limiter = RateLimiter(
            requests_per_minute=requests_per_minute,
            requests_per_hour=requests_per_hour
        )
    return _rate_limiter


def rate_limit(
    requests_per_minute: int = 10,
    requests_per_hour: int = 100
) -> Callable:
    """
    速率限制装饰器

    用法:
        @app.route('/api/generate')
        @rate_limit(requests_per_minute=5, requests_per_hour=50)
        def generate():
            ...
    """
    limiter = get_rate_limiter(requests_per_minute, requests_per_hour)

    def decorator(f: Callable) -> Callable:
        @wraps(f)
        def wrapper(*args, **kwargs):
            is_allowed, error_msg, retry_after = limiter.check_rate_limit()

            if not is_allowed:
                response = jsonify({
                    'error': error_msg,
                    'type': 'rate_limit',
                    'retry_after': retry_after
                })
                response.status_code = 429
                response.headers['Retry-After'] = str(retry_after)
                return response

            # 添加速率限制头
            result = f(*args, **kwargs)

            # 如果返回的是 Response 对象，添加头信息
            if hasattr(result, 'headers'):
                remaining = limiter.get_remaining()
                result.headers['X-RateLimit-Limit-Minute'] = str(remaining['minute_limit'])
                result.headers['X-RateLimit-Remaining-Minute'] = str(remaining['minute_remaining'])
                result.headers['X-RateLimit-Limit-Hour'] = str(remaining['hour_limit'])
                result.headers['X-RateLimit-Remaining-Hour'] = str(remaining['hour_remaining'])

            return result
        return wrapper
    return decorator
