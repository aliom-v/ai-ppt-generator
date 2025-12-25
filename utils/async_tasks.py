"""异步任务管理模块

支持后台生成 PPT，提供任务状态查询和结果获取。
"""
import threading
import uuid
import time
import traceback
from datetime import datetime
from enum import Enum
from typing import Dict, Any, Optional, Callable
from dataclasses import dataclass, field
from collections import OrderedDict

from utils.logger import get_logger

logger = get_logger("async_tasks")


class TaskStatus(str, Enum):
    """任务状态"""
    PENDING = "pending"      # 等待执行
    RUNNING = "running"      # 执行中
    SUCCESS = "success"      # 成功
    FAILED = "failed"        # 失败
    CANCELLED = "cancelled"  # 已取消


@dataclass
class TaskInfo:
    """任务信息"""
    task_id: str
    status: TaskStatus
    progress: int = 0  # 0-100
    message: str = ""
    result: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    created_at: float = field(default_factory=time.time)
    started_at: Optional[float] = None
    completed_at: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        data = {
            "task_id": self.task_id,
            "status": self.status.value,
            "progress": self.progress,
            "message": self.message,
            "created_at": datetime.fromtimestamp(self.created_at).isoformat(),
        }
        if self.started_at:
            data["started_at"] = datetime.fromtimestamp(self.started_at).isoformat()
        if self.completed_at:
            data["completed_at"] = datetime.fromtimestamp(self.completed_at).isoformat()
            data["duration_ms"] = int((self.completed_at - (self.started_at or self.created_at)) * 1000)
        if self.result:
            data["result"] = self.result
        if self.error:
            data["error"] = self.error
        return data


class TaskManager:
    """任务管理器

    线程安全的任务管理，支持创建、查询、取消任务。
    """

    def __init__(self, max_tasks: int = 100, max_concurrent: int = 3):
        """
        Args:
            max_tasks: 最大保留任务数（LRU 策略）
            max_concurrent: 最大并发任务数
        """
        self._tasks: OrderedDict[str, TaskInfo] = OrderedDict()
        self._lock = threading.RLock()
        self._max_tasks = max_tasks
        self._max_concurrent = max_concurrent
        self._running_count = 0
        self._semaphore = threading.Semaphore(max_concurrent)

    def create_task(self) -> str:
        """创建新任务

        Returns:
            任务 ID
        """
        task_id = str(uuid.uuid4())[:8]
        task = TaskInfo(task_id=task_id, status=TaskStatus.PENDING)

        with self._lock:
            # 清理超时的 RUNNING 任务
            self._cleanup_stale_tasks()

            # LRU 清理：删除最旧的已完成任务
            while len(self._tasks) >= self._max_tasks:
                oldest_id = next(iter(self._tasks))
                oldest = self._tasks[oldest_id]
                if oldest.status in (TaskStatus.SUCCESS, TaskStatus.FAILED, TaskStatus.CANCELLED):
                    del self._tasks[oldest_id]
                else:
                    break

            self._tasks[task_id] = task

        logger.info(f"创建任务: {task_id}")
        return task_id

    def _cleanup_stale_tasks(self, stale_timeout: int = 3600) -> int:
        """清理超时的 RUNNING 任务（防止内存泄漏）

        Args:
            stale_timeout: 超时时间（秒），默认 1 小时

        Returns:
            清理的任务数量
        """
        now = time.time()
        cleaned = 0
        for task_id, task in list(self._tasks.items()):
            if task.status == TaskStatus.RUNNING:
                started = task.started_at or task.created_at
                if (now - started) > stale_timeout:
                    task.status = TaskStatus.FAILED
                    task.error = "任务执行超时"
                    task.completed_at = now
                    logger.warning(f"任务超时已清理: {task_id}")
                    cleaned += 1
        return cleaned

    def get_task(self, task_id: str) -> Optional[TaskInfo]:
        """获取任务信息"""
        with self._lock:
            return self._tasks.get(task_id)

    def update_task(
        self,
        task_id: str,
        status: Optional[TaskStatus] = None,
        progress: Optional[int] = None,
        message: Optional[str] = None,
        result: Optional[Dict] = None,
        error: Optional[str] = None,
    ):
        """更新任务状态"""
        with self._lock:
            task = self._tasks.get(task_id)
            if not task:
                return

            if status:
                task.status = status
                if status == TaskStatus.RUNNING and not task.started_at:
                    task.started_at = time.time()
                elif status in (TaskStatus.SUCCESS, TaskStatus.FAILED, TaskStatus.CANCELLED):
                    task.completed_at = time.time()

            if progress is not None:
                task.progress = min(100, max(0, progress))
            if message is not None:
                task.message = message
            if result is not None:
                task.result = result
            if error is not None:
                task.error = error

    def cancel_task(self, task_id: str) -> bool:
        """取消任务

        只能取消等待中的任务。

        Returns:
            是否取消成功
        """
        with self._lock:
            task = self._tasks.get(task_id)
            if task and task.status == TaskStatus.PENDING:
                task.status = TaskStatus.CANCELLED
                task.completed_at = time.time()
                logger.info(f"任务已取消: {task_id}")
                return True
            return False

    def run_task(
        self,
        task_id: str,
        func: Callable[..., Dict[str, Any]],
        *args,
        progress_callback: Optional[Callable[[int, str], None]] = None,
        **kwargs,
    ):
        """在后台线程中运行任务

        Args:
            task_id: 任务 ID
            func: 要执行的函数，返回结果字典
            progress_callback: 进度回调 (progress, message)
        """
        def wrapper():
            # 等待信号量
            acquired = self._semaphore.acquire(timeout=300)  # 最多等待 5 分钟
            if not acquired:
                self.update_task(
                    task_id,
                    status=TaskStatus.FAILED,
                    error="任务队列已满，请稍后重试"
                )
                return

            try:
                with self._lock:
                    task = self._tasks.get(task_id)
                    if not task or task.status == TaskStatus.CANCELLED:
                        # 任务已取消，需要释放信号量
                        self._semaphore.release()
                        return
                    self._running_count += 1

                self.update_task(
                    task_id,
                    status=TaskStatus.RUNNING,
                    progress=0,
                    message="任务开始执行"
                )

                # 创建进度回调包装器
                def _progress(progress: int, message: str = ""):
                    self.update_task(task_id, progress=progress, message=message)
                    if progress_callback:
                        progress_callback(progress, message)

                # 执行任务
                result = func(*args, progress_callback=_progress, **kwargs)

                self.update_task(
                    task_id,
                    status=TaskStatus.SUCCESS,
                    progress=100,
                    message="任务完成",
                    result=result
                )
                logger.info(f"任务完成: {task_id}")

            except Exception as e:
                error_msg = str(e)
                logger.error(f"任务失败: {task_id} - {error_msg}")
                logger.debug(traceback.format_exc())
                self.update_task(
                    task_id,
                    status=TaskStatus.FAILED,
                    error=error_msg
                )
            finally:
                with self._lock:
                    if self._running_count > 0:
                        self._running_count -= 1
                self._semaphore.release()

        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()

    def get_stats(self) -> Dict[str, Any]:
        """获取任务统计"""
        with self._lock:
            status_counts = {}
            for task in self._tasks.values():
                status = task.status.value
                status_counts[status] = status_counts.get(status, 0) + 1

            return {
                "total_tasks": len(self._tasks),
                "running_count": self._running_count,
                "max_concurrent": self._max_concurrent,
                "status_counts": status_counts,
            }

    def list_tasks(self, limit: int = 20) -> list:
        """列出最近的任务"""
        with self._lock:
            tasks = list(self._tasks.values())[-limit:]
            return [t.to_dict() for t in reversed(tasks)]


# 全局任务管理器
_task_manager: Optional[TaskManager] = None


def get_task_manager() -> TaskManager:
    """获取全局任务管理器"""
    global _task_manager
    if _task_manager is None:
        _task_manager = TaskManager()
    return _task_manager
