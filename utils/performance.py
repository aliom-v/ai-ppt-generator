"""
Performance monitoring and optimization utilities.
"""
import asyncio
import time
import psutil
import threading
from collections import defaultdict, deque
from contextlib import asynccontextmanager, contextmanager
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Callable, Any
from functools import wraps
import structlog

from .logger import get_logger

logger = get_logger(__name__)


@dataclass
class PerformanceMetrics:
    """Performance metrics data class."""
    operation: str
    duration: float
    timestamp: float
    success: bool
    metadata: Dict[str, Any] = field(default_factory=dict)


class PerformanceMonitor:
    """Performance monitoring and metrics collection."""

    def __init__(self, max_history: int = 1000):
        self.max_history = max_history
        self.metrics: Dict[str, deque] = defaultdict(lambda: deque(maxlen=max_history))
        self.active_operations: Dict[str, float] = {}
        self._lock = threading.Lock()

    def record_metric(self, metric: PerformanceMetrics):
        """Record a performance metric."""
        with self._lock:
            self.metrics[metric.operation].append(metric)

    def start_operation(self, operation: str):
        """Start tracking an operation."""
        with self._lock:
            self.active_operations[operation] = time.time()

    def end_operation(self, operation: str, success: bool = True, metadata: Optional[Dict[str, Any]] = None):
        """End tracking an operation and record the metric."""
        start_time = self.active_operations.pop(operation, None)
        if start_time is None:
            logger.warning(f"Operation {operation} was not started")
            return

        duration = time.time() - start_time
        metric = PerformanceMetrics(
            operation=operation,
            duration=duration,
            timestamp=time.time(),
            success=success,
            metadata=metadata or {}
        )
        self.record_metric(metric)

        # Log slow operations
        if duration > 5.0:  # Log operations taking more than 5 seconds
            logger.warning(
                "Slow operation detected",
                operation=operation,
                duration=duration,
                metadata=metadata
            )

    def get_stats(self, operation: str, window_seconds: Optional[float] = None) -> Dict[str, Any]:
        """Get performance statistics for an operation."""
        with self._lock:
            metrics = list(self.metrics[operation])

        if window_seconds:
            cutoff_time = time.time() - window_seconds
            metrics = [m for m in metrics if m.timestamp >= cutoff_time]

        if not metrics:
            return {"count": 0}

        durations = [m.duration for m in metrics]
        success_count = sum(1 for m in metrics if m.success)

        return {
            "count": len(metrics),
            "success_rate": success_count / len(metrics),
            "avg_duration": sum(durations) / len(durations),
            "min_duration": min(durations),
            "max_duration": max(durations),
            "p95_duration": sorted(durations)[int(len(durations) * 0.95)] if len(durations) > 20 else max(durations),
            "recent_timestamp": max(m.timestamp for m in metrics)
        }

    def get_all_stats(self) -> Dict[str, Dict[str, Any]]:
        """Get performance statistics for all operations."""
        return {op: self.get_stats(op) for op in self.metrics.keys()}

    def reset(self):
        """Reset all metrics."""
        with self._lock:
            self.metrics.clear()
            self.active_operations.clear()


# Global performance monitor instance
performance_monitor = PerformanceMonitor()


@contextmanager
def monitor_performance(operation: str, metadata: Optional[Dict[str, Any]] = None):
    """Context manager for monitoring performance."""
    performance_monitor.start_operation(operation)
    try:
        yield
        performance_monitor.end_operation(operation, success=True, metadata=metadata)
    except Exception as e:
        performance_monitor.end_operation(operation, success=False, metadata={**(metadata or {}), "error": str(e)})
        raise


@asynccontextmanager
async def monitor_async_performance(operation: str, metadata: Optional[Dict[str, Any]] = None):
    """Async context manager for monitoring performance."""
    performance_monitor.start_operation(operation)
    try:
        yield
        performance_monitor.end_operation(operation, success=True, metadata=metadata)
    except Exception as e:
        performance_monitor.end_operation(operation, success=False, metadata={**(metadata or {}), "error": str(e)})
        raise


def track_performance(operation: Optional[str] = None):
    """Decorator for tracking function performance."""
    def decorator(func: Callable) -> Callable:
        op_name = operation or f"{func.__module__}.{func.__name__}"

        if asyncio.iscoroutinefunction(func):
            @wraps(func)
            async def async_wrapper(*args, **kwargs):
                async with monitor_async_performance(op_name):
                    return await func(*args, **kwargs)
            return async_wrapper
        else:
            @wraps(func)
            def sync_wrapper(*args, **kwargs):
                with monitor_performance(op_name):
                    return func(*args, **kwargs)
            return sync_wrapper

    return decorator


class ResourceMonitor:
    """System resource monitoring."""

    def __init__(self):
        self.cpu_history = deque(maxlen=60)  # Last 60 seconds
        self.memory_history = deque(maxlen=60)
        self._running = False
        self._task = None

    async def start(self):
        """Start resource monitoring."""
        if self._running:
            return

        self._running = True
        self._task = asyncio.create_task(self._monitor_loop())

    async def stop(self):
        """Stop resource monitoring."""
        self._running = False
        if self._task:
            self._task.cancel()
            try:
                await self._task
            except asyncio.CancelledError:
                pass

    async def _monitor_loop(self):
        """Resource monitoring loop."""
        while self._running:
            try:
                cpu_percent = psutil.cpu_percent(interval=1)
                memory = psutil.virtual_memory()

                self.cpu_history.append(cpu_percent)
                self.memory_history.append(memory.percent)

                # Log warnings for high resource usage
                if cpu_percent > 80:
                    logger.warning("High CPU usage detected", cpu_percent=cpu_percent)

                if memory.percent > 85:
                    logger.warning("High memory usage detected", memory_percent=memory.percent)

            except Exception as e:
                logger.error("Resource monitoring error", error=str(e))

            await asyncio.sleep(1)

    def get_current_stats(self) -> Dict[str, Any]:
        """Get current resource statistics."""
        return {
            "cpu": {
                "current": psutil.cpu_percent(),
                "avg_1min": sum(self.cpu_history) / len(self.cpu_history) if self.cpu_history else 0,
                "max_1min": max(self.cpu_history) if self.cpu_history else 0
            },
            "memory": {
                "current": psutil.virtual_memory().percent,
                "avg_1min": sum(self.memory_history) / len(self.memory_history) if self.memory_history else 0,
                "max_1min": max(self.memory_history) if self.memory_history else 0,
                "available_gb": psutil.virtual_memory().available / (1024**3),
                "total_gb": psutil.virtual_memory().total / (1024**3)
            }
        }


# Global resource monitor instance
resource_monitor = ResourceMonitor()


class CircuitBreaker:
    """Circuit breaker pattern implementation for fault tolerance."""

    def __init__(
        self,
        failure_threshold: int = 5,
        recovery_timeout: float = 60.0,
        expected_exception: type = Exception
    ):
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self.expected_exception = expected_exception

        self.failure_count = 0
        self.last_failure_time = None
        self.state = "CLOSED"  # CLOSED, OPEN, HALF_OPEN

    def __call__(self, func: Callable) -> Callable:
        """Decorator for circuit breaker functionality."""

        if asyncio.iscoroutinefunction(func):
            @wraps(func)
            async def async_wrapper(*args, **kwargs):
                return await self._call_async(func, *args, **kwargs)
            return async_wrapper
        else:
            @wraps(func)
            def sync_wrapper(*args, **kwargs):
                return self._call_sync(func, *args, **kwargs)
            return sync_wrapper

    async def _call_async(self, func: Callable, *args, **kwargs):
        """Call async function with circuit breaker protection."""
        if not self._can_execute():
            raise Exception("Circuit breaker is OPEN")

        try:
            result = await func(*args, **kwargs)
            self._on_success()
            return result
        except self.expected_exception as e:
            self._on_failure()
            raise

    def _call_sync(self, func: Callable, *args, **kwargs):
        """Call sync function with circuit breaker protection."""
        if not self._can_execute():
            raise Exception("Circuit breaker is OPEN")

        try:
            result = func(*args, **kwargs)
            self._on_success()
            return result
        except self.expected_exception as e:
            self._on_failure()
            raise

    def _can_execute(self) -> bool:
        """Check if execution is allowed."""
        if self.state == "CLOSED":
            return True
        elif self.state == "OPEN":
            if time.time() - self.last_failure_time > self.recovery_timeout:
                self.state = "HALF_OPEN"
                return True
            return False
        else:  # HALF_OPEN
            return True

    def _on_success(self):
        """Handle successful execution."""
        self.failure_count = 0
        self.state = "CLOSED"

    def _on_failure(self):
        """Handle failed execution."""
        self.failure_count += 1
        self.last_failure_time = time.time()

        if self.failure_count >= self.failure_threshold:
            self.state = "OPEN"
            logger.warning(
                "Circuit breaker opened",
                failure_count=self.failure_count,
                threshold=self.failure_threshold
            )