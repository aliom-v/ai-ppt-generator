"""性能监控中间件 - 请求耗时统计和指标收集"""
import time
import threading
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional
from functools import wraps

from flask import Flask, request, g
from utils.logger import get_logger

logger = get_logger("metrics")


@dataclass
class RequestMetrics:
    """单个请求的指标"""
    path: str
    method: str
    status_code: int
    duration_ms: float
    timestamp: datetime


@dataclass
class EndpointStats:
    """端点统计信息"""
    total_requests: int = 0
    total_errors: int = 0
    total_duration_ms: float = 0
    min_duration_ms: float = float('inf')
    max_duration_ms: float = 0
    recent_durations: List[float] = field(default_factory=list)

    @property
    def avg_duration_ms(self) -> float:
        if self.total_requests == 0:
            return 0
        return self.total_duration_ms / self.total_requests

    @property
    def error_rate(self) -> float:
        if self.total_requests == 0:
            return 0
        return self.total_errors / self.total_requests

    @property
    def p95_duration_ms(self) -> float:
        """95 分位耗时"""
        if not self.recent_durations:
            return 0
        sorted_durations = sorted(self.recent_durations)
        idx = int(len(sorted_durations) * 0.95)
        return sorted_durations[min(idx, len(sorted_durations) - 1)]


class MetricsCollector:
    """指标收集器"""

    def __init__(self, max_recent_requests: int = 1000):
        self._stats: Dict[str, EndpointStats] = defaultdict(EndpointStats)
        self._recent_requests: List[RequestMetrics] = []
        self._max_recent = max_recent_requests
        self._lock = threading.Lock()
        self._start_time = datetime.now()

    def record(self, metrics: RequestMetrics):
        """记录请求指标"""
        with self._lock:
            key = f"{metrics.method}:{metrics.path}"
            stats = self._stats[key]

            stats.total_requests += 1
            stats.total_duration_ms += metrics.duration_ms
            stats.min_duration_ms = min(stats.min_duration_ms, metrics.duration_ms)
            stats.max_duration_ms = max(stats.max_duration_ms, metrics.duration_ms)

            stats.recent_durations.append(metrics.duration_ms)
            if len(stats.recent_durations) > 100:
                stats.recent_durations = stats.recent_durations[-100:]

            if metrics.status_code >= 400:
                stats.total_errors += 1

            self._recent_requests.append(metrics)
            if len(self._recent_requests) > self._max_recent:
                self._recent_requests = self._recent_requests[-self._max_recent:]

    def get_stats(self) -> Dict:
        """获取所有统计信息"""
        with self._lock:
            uptime = datetime.now() - self._start_time
            total_requests = sum(s.total_requests for s in self._stats.values())
            total_errors = sum(s.total_errors for s in self._stats.values())

            endpoints = {}
            for key, stats in self._stats.items():
                endpoints[key] = {
                    "total_requests": stats.total_requests,
                    "error_rate": round(stats.error_rate * 100, 2),
                    "avg_ms": round(stats.avg_duration_ms, 2),
                    "p95_ms": round(stats.p95_duration_ms, 2),
                }

            return {
                "uptime_seconds": int(uptime.total_seconds()),
                "total_requests": total_requests,
                "total_errors": total_errors,
                "error_rate": round(total_errors / max(total_requests, 1) * 100, 2),
                "endpoints": endpoints,
            }


_metrics_collector: Optional[MetricsCollector] = None


def get_metrics_collector() -> MetricsCollector:
    """获取全局指标收集器"""
    global _metrics_collector
    if _metrics_collector is None:
        _metrics_collector = MetricsCollector()
    return _metrics_collector


class PerformanceMiddleware:
    """Flask 性能监控中间件"""

    def __init__(self, app: Flask = None, slow_threshold_ms: float = 1000):
        self.slow_threshold = slow_threshold_ms
        self.collector = get_metrics_collector()
        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        app.before_request(self._before_request)
        app.after_request(self._after_request)

    def _before_request(self):
        g.start_time = time.time()

    def _after_request(self, response):
        if not hasattr(g, 'start_time'):
            return response

        duration_ms = (time.time() - g.start_time) * 1000

        metrics = RequestMetrics(
            path=request.path,
            method=request.method,
            status_code=response.status_code,
            duration_ms=duration_ms,
            timestamp=datetime.now()
        )
        self.collector.record(metrics)

        response.headers['X-Response-Time'] = f"{duration_ms:.2f}ms"

        if duration_ms >= self.slow_threshold:
            logger.warning(f"慢请求: {request.method} {request.path} - {duration_ms:.2f}ms")

        return response


def timed(name: str = None):
    """函数计时装饰器"""
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            start = time.time()
            try:
                return f(*args, **kwargs)
            finally:
                duration = (time.time() - start) * 1000
                func_name = name or f.__name__
                logger.debug(f"[计时] {func_name}: {duration:.2f}ms")
        return wrapper
    return decorator
