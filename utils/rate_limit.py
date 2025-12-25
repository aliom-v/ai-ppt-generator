"""速率限制中间件"""
import os
import time
from collections import defaultdict
from functools import wraps
from threading import Lock
from typing import Callable, Dict, Optional, Tuple

from flask import request, jsonify


class RateLimiter:
    """基于滑动窗口的速率限制器"""

    def __init__(
        self,
        requests_per_minute: int = 10,
        requests_per_hour: int = 100,
        burst_size: int = 5,
        trust_proxy: bool = False
    ):
        """
        初始化速率限制器

        Args:
            requests_per_minute: 每分钟最大请求数
            requests_per_hour: 每小时最大请求数
            burst_size: 突发请求容量
            trust_proxy: 是否信任代理头（仅在明确配置反向代理时启用）
        """
        self.requests_per_minute = requests_per_minute
        self.requests_per_hour = requests_per_hour
        self.burst_size = burst_size
        self._trust_proxy = trust_proxy or os.getenv('TRUST_PROXY', 'false').lower() == 'true'

        # 存储每个 IP 的请求记录
        self._minute_counts: Dict[str, list] = defaultdict(list)
        self._hour_counts: Dict[str, list] = defaultdict(list)
        self._lock = Lock()
        self._last_cleanup = time.time()
        self._cleanup_interval = 60  # 每分钟清理一次

    def _get_client_ip(self) -> str:
        """获取客户端 IP（安全处理代理头）"""
        # 仅在明确配置信任代理时才使用代理头，防止 IP 欺骗
        if self._trust_proxy:
            forwarded_for = request.headers.get('X-Forwarded-For')
            if forwarded_for:
                # 取第一个 IP（最接近客户端的）
                return forwarded_for.split(',')[0].strip()
            real_ip = request.headers.get('X-Real-IP')
            if real_ip:
                return real_ip.strip()
        return request.remote_addr or '127.0.0.1'

    def _cleanup_old_requests(self, client_ip: str, current_time: float) -> None:
        """清理过期的请求记录"""
        minute_ago = current_time - 60
        hour_ago = current_time - 3600

        # 清理分钟级记录
        if client_ip in self._minute_counts:
            self._minute_counts[client_ip] = [
                t for t in self._minute_counts[client_ip] if t > minute_ago
            ]
            if not self._minute_counts[client_ip]:
                del self._minute_counts[client_ip]

        # 清理小时级记录
        if client_ip in self._hour_counts:
            self._hour_counts[client_ip] = [
                t for t in self._hour_counts[client_ip] if t > hour_ago
            ]
            if not self._hour_counts[client_ip]:
                del self._hour_counts[client_ip]

    def _global_cleanup(self, current_time: float) -> None:
        """定期全局清理，防止内存泄漏"""
        if current_time - self._last_cleanup < self._cleanup_interval:
            return

        self._last_cleanup = current_time
        minute_ago = current_time - 60
        hour_ago = current_time - 3600

        # 清理所有过期的 IP 记录
        for ip in list(self._minute_counts.keys()):
            self._minute_counts[ip] = [t for t in self._minute_counts[ip] if t > minute_ago]
            if not self._minute_counts[ip]:
                del self._minute_counts[ip]

        for ip in list(self._hour_counts.keys()):
            self._hour_counts[ip] = [t for t in self._hour_counts[ip] if t > hour_ago]
            if not self._hour_counts[ip]:
                del self._hour_counts[ip]

    def check_rate_limit(self) -> Tuple[bool, Optional[str], Optional[int]]:
        """
        检查是否超出速率限制

        Returns:
            (is_allowed, error_message, retry_after_seconds)
        """
        client_ip = self._get_client_ip()
        current_time = time.time()

        with self._lock:
            # 定期全局清理
            self._global_cleanup(current_time)
            # 清理当前 IP 的过期记录
            self._cleanup_old_requests(client_ip, current_time)

            minute_count = len(self._minute_counts.get(client_ip, []))
            hour_count = len(self._hour_counts.get(client_ip, []))

            # 检查分钟限制
            if minute_count >= self.requests_per_minute:
                oldest = min(self._minute_counts[client_ip]) if self._minute_counts.get(client_ip) else current_time
                retry_after = int(60 - (current_time - oldest)) + 1
                return False, f"请求过于频繁，请 {retry_after} 秒后重试", retry_after

            # 检查小时限制
            if hour_count >= self.requests_per_hour:
                oldest = min(self._hour_counts[client_ip]) if self._hour_counts.get(client_ip) else current_time
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
                'minute_remaining': max(0, self.requests_per_minute - len(self._minute_counts.get(client_ip, []))),
                'hour_remaining': max(0, self.requests_per_hour - len(self._hour_counts.get(client_ip, []))),
                'minute_limit': self.requests_per_minute,
                'hour_limit': self.requests_per_hour,
            }


# 端点级别的速率限制器缓存
_endpoint_limiters: Dict[str, RateLimiter] = {}
_limiters_lock = Lock()


def get_endpoint_limiter(
    endpoint: str,
    requests_per_minute: int = 10,
    requests_per_hour: int = 100
) -> RateLimiter:
    """获取端点专用的速率限制器（支持不同端点不同配置）"""
    key = f"{endpoint}:{requests_per_minute}:{requests_per_hour}"

    with _limiters_lock:
        if key not in _endpoint_limiters:
            _endpoint_limiters[key] = RateLimiter(
                requests_per_minute=requests_per_minute,
                requests_per_hour=requests_per_hour
            )
        return _endpoint_limiters[key]


def rate_limit(
    requests_per_minute: int = 10,
    requests_per_hour: int = 100
) -> Callable:
    """
    速率限制装饰器（支持每个端点独立配置）

    用法:
        @app.route('/api/generate')
        @rate_limit(requests_per_minute=5, requests_per_hour=50)
        def generate():
            ...
    """
    def decorator(f: Callable) -> Callable:
        # 为每个端点创建独立的限流器
        endpoint_name = f.__name__

        @wraps(f)
        def wrapper(*args, **kwargs):
            limiter = get_endpoint_limiter(endpoint_name, requests_per_minute, requests_per_hour)
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


# 保留旧的全局限流器接口（向后兼容）
_rate_limiter: Optional[RateLimiter] = None


def get_rate_limiter(
    requests_per_minute: int = 10,
    requests_per_hour: int = 100
) -> RateLimiter:
    """获取全局速率限制器（向后兼容）"""
    global _rate_limiter
    if _rate_limiter is None:
        _rate_limiter = RateLimiter(
            requests_per_minute=requests_per_minute,
            requests_per_hour=requests_per_hour
        )
    return _rate_limiter
