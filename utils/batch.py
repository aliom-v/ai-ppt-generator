"""批量生成模块

支持批量生成多个 PPT，提供进度跟踪和结果汇总。
"""
import os
import time
import threading
import uuid
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from utils.logger import get_logger

logger = get_logger("batch")


class BatchStatus(str, Enum):
    """批量任务状态"""
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"
    PARTIAL = "partial"  # 部分成功


@dataclass
class BatchItem:
    """批量任务项"""
    index: int
    topic: str
    audience: str = "通用受众"
    page_count: int = 5
    description: str = ""
    template_id: str = ""
    status: str = "pending"
    result: Optional[Dict] = None
    error: Optional[str] = None
    started_at: Optional[float] = None
    completed_at: Optional[float] = None

    def to_dict(self) -> Dict:
        data = {
            "index": self.index,
            "topic": self.topic,
            "status": self.status,
        }
        if self.result:
            data["result"] = self.result
        if self.error:
            data["error"] = self.error
        if self.started_at:
            data["started_at"] = datetime.fromtimestamp(self.started_at).isoformat()
        if self.completed_at:
            data["completed_at"] = datetime.fromtimestamp(self.completed_at).isoformat()
            data["duration_ms"] = int((self.completed_at - self.started_at) * 1000)
        return data


@dataclass
class BatchJob:
    """批量任务"""
    job_id: str
    items: List[BatchItem]
    status: BatchStatus = BatchStatus.PENDING
    created_at: float = field(default_factory=time.time)
    started_at: Optional[float] = None
    completed_at: Optional[float] = None
    current_index: int = 0
    api_config: Optional[Dict] = None

    @property
    def total(self) -> int:
        return len(self.items)

    @property
    def completed_count(self) -> int:
        return sum(1 for item in self.items if item.status in ("success", "failed"))

    @property
    def success_count(self) -> int:
        return sum(1 for item in self.items if item.status == "success")

    @property
    def failed_count(self) -> int:
        return sum(1 for item in self.items if item.status == "failed")

    @property
    def progress(self) -> int:
        if self.total == 0:
            return 0
        return int(self.completed_count / self.total * 100)

    def to_dict(self) -> Dict:
        return {
            "job_id": self.job_id,
            "status": self.status.value,
            "total": self.total,
            "completed": self.completed_count,
            "success": self.success_count,
            "failed": self.failed_count,
            "progress": self.progress,
            "current_index": self.current_index,
            "created_at": datetime.fromtimestamp(self.created_at).isoformat(),
            "started_at": datetime.fromtimestamp(self.started_at).isoformat() if self.started_at else None,
            "completed_at": datetime.fromtimestamp(self.completed_at).isoformat() if self.completed_at else None,
            "items": [item.to_dict() for item in self.items],
        }


class BatchGenerator:
    """批量生成器

    用法:
        generator = BatchGenerator()

        # 创建批量任务
        job = generator.create_job([
            {"topic": "人工智能简介", "pages": 5},
            {"topic": "机器学习入门", "pages": 8},
        ], api_config={...})

        # 启动生成
        generator.start_job(job.job_id)

        # 查询状态
        status = generator.get_job(job.job_id)
    """

    def __init__(self, max_concurrent: int = 2, output_folder: str = "web/outputs"):
        self._jobs: Dict[str, BatchJob] = {}
        self._lock = threading.RLock()
        self._max_concurrent = max_concurrent
        self._output_folder = output_folder
        self._running_count = 0
        self._semaphore = threading.Semaphore(max_concurrent)

    def create_job(
        self,
        items: List[Dict],
        api_config: Dict,
        template_id: str = "",
    ) -> BatchJob:
        """创建批量任务

        Args:
            items: 任务项列表，每项包含 topic, audience, page_count 等
            api_config: API 配置
            template_id: 默认模板 ID
        """
        job_id = str(uuid.uuid4())[:8]

        batch_items = []
        for i, item in enumerate(items):
            batch_items.append(BatchItem(
                index=i,
                topic=item.get("topic", ""),
                audience=item.get("audience", "通用受众"),
                page_count=item.get("page_count", 5),
                description=item.get("description", ""),
                template_id=item.get("template_id", template_id),
            ))

        job = BatchJob(
            job_id=job_id,
            items=batch_items,
            api_config=api_config,
        )

        with self._lock:
            self._jobs[job_id] = job

        logger.info(f"创建批量任务: {job_id} ({len(items)} 项)")
        return job

    def start_job(
        self,
        job_id: str,
        progress_callback: Optional[Callable[[str, int, str], None]] = None,
    ):
        """启动批量任务

        Args:
            job_id: 任务 ID
            progress_callback: 进度回调 (job_id, progress, message)
        """
        job = self._jobs.get(job_id)
        if not job:
            raise ValueError(f"任务不存在: {job_id}")

        if job.status != BatchStatus.PENDING:
            raise ValueError(f"任务状态不正确: {job.status}")

        def worker():
            try:
                self._run_job(job, progress_callback)
            except Exception as e:
                logger.error(f"批量任务失败: {job_id} - {e}")
                job.status = BatchStatus.FAILED

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _run_job(
        self,
        job: BatchJob,
        progress_callback: Optional[Callable[[str, int, str], None]],
    ):
        """执行批量任务"""
        from config.settings import AIConfig
        from core.ai_client import generate_ppt_plan
        from core.ppt_plan import ppt_plan_from_dict
        from ppt.unified_builder import build_ppt_from_plan
        from ppt.template_manager import template_manager

        job.status = BatchStatus.RUNNING
        job.started_at = time.time()

        # 创建 AI 配置
        ai_config = AIConfig(
            api_key=job.api_config.get("api_key", ""),
            api_base_url=job.api_config.get("api_base", "https://api.openai.com/v1"),
            model_name=job.api_config.get("model_name", "gpt-4o-mini"),
        )

        # 确保输出目录存在
        Path(self._output_folder).mkdir(parents=True, exist_ok=True)

        for item in job.items:
            if job.status == BatchStatus.CANCELLED:
                break

            job.current_index = item.index
            item.status = "running"
            item.started_at = time.time()

            # 回调进度
            if progress_callback:
                progress_callback(
                    job.job_id,
                    job.progress,
                    f"正在生成: {item.topic} ({item.index + 1}/{job.total})"
                )

            try:
                # 获取信号量
                self._semaphore.acquire()

                # 生成 PPT 结构
                plan_dict = generate_ppt_plan(
                    item.topic,
                    item.audience,
                    item.page_count,
                    item.description,
                    config=ai_config,
                )
                plan = ppt_plan_from_dict(plan_dict)

                # 生成文件名
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                safe_topic = "".join(c for c in item.topic if c.isalnum() or c in " -_")[:30]
                filename = f"batch_{job.job_id}_{item.index}_{safe_topic}_{timestamp}.pptx"
                output_path = os.path.join(self._output_folder, filename)

                # 获取模板
                template_path = None
                if item.template_id:
                    template_path = template_manager.get_template(item.template_id)

                # 生成 PPT
                build_ppt_from_plan(plan, template_path, output_path)

                # 更新结果
                item.status = "success"
                item.result = {
                    "filename": filename,
                    "title": plan.title,
                    "slide_count": len(plan.slides),
                    "download_url": f"/api/download/{filename}",
                }

                logger.info(f"批量任务 {job.job_id} 项目 {item.index} 成功: {item.topic}")

            except Exception as e:
                item.status = "failed"
                item.error = str(e)
                logger.error(f"批量任务 {job.job_id} 项目 {item.index} 失败: {e}")

            finally:
                item.completed_at = time.time()
                self._semaphore.release()

        # 更新任务状态
        job.completed_at = time.time()
        if job.status == BatchStatus.CANCELLED:
            pass  # 保持取消状态
        elif job.failed_count == 0:
            job.status = BatchStatus.COMPLETED
        elif job.success_count == 0:
            job.status = BatchStatus.FAILED
        else:
            job.status = BatchStatus.PARTIAL

        # 最终回调
        if progress_callback:
            progress_callback(
                job.job_id,
                100,
                f"完成: {job.success_count}/{job.total} 成功"
            )

        logger.info(f"批量任务 {job.job_id} 完成: {job.success_count}/{job.total} 成功")

    def get_job(self, job_id: str) -> Optional[BatchJob]:
        """获取任务"""
        return self._jobs.get(job_id)

    def cancel_job(self, job_id: str) -> bool:
        """取消任务"""
        job = self._jobs.get(job_id)
        if not job:
            return False

        if job.status in (BatchStatus.PENDING, BatchStatus.RUNNING):
            job.status = BatchStatus.CANCELLED
            logger.info(f"批量任务已取消: {job_id}")
            return True

        return False

    def list_jobs(self, limit: int = 20) -> List[Dict]:
        """列出任务"""
        with self._lock:
            jobs = sorted(
                self._jobs.values(),
                key=lambda j: j.created_at,
                reverse=True,
            )[:limit]
            return [j.to_dict() for j in jobs]

    def cleanup_old_jobs(self, max_age_hours: int = 24):
        """清理旧任务"""
        now = time.time()
        max_age = max_age_hours * 3600

        with self._lock:
            to_delete = [
                job_id for job_id, job in self._jobs.items()
                if now - job.created_at > max_age
                and job.status in (BatchStatus.COMPLETED, BatchStatus.FAILED, BatchStatus.CANCELLED)
            ]
            for job_id in to_delete:
                del self._jobs[job_id]

        if to_delete:
            logger.info(f"清理了 {len(to_delete)} 个旧批量任务")


# 全局批量生成器
_batch_generator: Optional[BatchGenerator] = None


def get_batch_generator() -> BatchGenerator:
    """获取全局批量生成器"""
    global _batch_generator
    if _batch_generator is None:
        _batch_generator = BatchGenerator()
    return _batch_generator
