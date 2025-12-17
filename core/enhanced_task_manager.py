"""增强的任务管理器 - 支持异步并发和智能管理"""
import os
import uuid
import time
import json
import asyncio
import threading
from typing import Dict, Any, Optional, Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

from config.settings import settings
from core.ai_client import generate_ppt_plan
from core.ai_client_async import generate_ppt_plan_async
from ppt.unified_builder import build_ppt_from_plan
from utils.logger import get_logger
from utils.smart_cache import get_cache
from utils.metrics import record_metric, get_metrics_summary

logger = get_logger("task_manager")


class TaskManager:
    """增强的任务管理器 - 支持异步并发、智能调度"""

    def __init__(self, max_workers: int = 3, enable_async: bool = True):
        self.max_workers = max_workers
        self.enable_async = enable_async
        self.tasks: Dict[str, Dict] = {}
        self.task_results: Dict[str, Any] = {}
        self.task_futures: Dict[str, Any] = {}
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(max_workers=max_workers)
        self._cleanup_interval = 3600  # 1小时清理一次
        self._last_cleanup = time.time()

    def generate_task_id(self) -> str:
        """生成唯一任务ID"""
        return str(uuid.uuid4())[:8]

    def submit_task(self, task_config: Dict[str, Any]) -> str:
        """提交PPT生成任务"""
        task_id = self.generate_task_id()

        # 任务配置
        task = {
            "id": task_id,
            "status": "pending",
            "progress": 0,
            "message": "任务已提交",
            "created_at": datetime.now().isoformat(),
            "started_at": None,
            "completed_at": None,
            "config": task_config,
            "result": None,
            "error": None,
            "retries": 0,
            "max_retries": 3
        }

        with self._lock:
            self.tasks[task_id] = task

        # 提交到线程池
        future = self._executor.submit(self._execute_task, task_id)
        self.task_futures[task_id] = future

        logger.info(f"任务已提交: {task_id}")
        return task_id

    def _execute_task(self, task_id: str):
        """执行任务（支持异步优化）"""
        task = self.tasks.get(task_id)
        if not task:
            return

        try:
            # 更新状态
            task["status"] = "running"
            task["started_at"] = datetime.now().isoformat()
            task["message"] = "开始生成PPT..."
            task["progress"] = 10
            self._update_progress(task_id, 10, "开始生成PPT...")

            # 提取配置
            config = task["config"]
            topic = config["topic"]
            audience = config.get("audience", "通用受众")
            page_count = config.get("page_count", 10)
            description = config.get("description", "")
            auto_page_count = config.get("auto_page_count", False)
            template_id = config.get("template_id", "classic_blue")
            use_cache = config.get("use_cache", True)

            # 记录指标
            start_time = time.time()
            record_metric("ppt_generation_started", 1, tags={
                "template": template_id,
                "auto_count": str(auto_page_count)
            })

            # 步骤1: 生成PPT计划 (30%)
            self._update_progress(task_id, 20, "正在生成PPT结构...")

            try:
                if self.enable_async and page_count > 35:
                    # 大型PPT使用异步并发生成
                    plan = asyncio.run(generate_ppt_plan_async(
                        topic=topic,
                        audience=audience,
                        page_count=page_count,
                        description=description,
                        auto_page_count=auto_page_count,
                        config=config.get("ai_config"),
                        progress_callback=self._progress_callback(task_id),
                        use_cache=use_cache
                    ))
                else:
                    # 小型PPT使用同步生成
                    plan = generate_ppt_plan(
                        topic=topic,
                        audience=audience,
                        page_count=page_count,
                        description=description,
                        auto_page_count=auto_page_count,
                        config=config.get("ai_config"),
                        use_cache=use_cache
                    )

                self._update_progress(task_id, 50, "PPT结构生成完成，开始制作PPT...")

            except Exception as e:
                logger.error(f"生成PPT计划失败: {e}")
                raise Exception(f"生成PPT结构失败: {str(e)}")

            # 步骤2: 构建PPT文件 (40%)
            self._update_progress(task_id, 60, "正在制作PPT文件...")

            try:
                output_dir = config.get("output_dir", "web/outputs")
                auto_download_images = config.get("auto_download_images", False)

                ppt_path = build_ppt_from_plan(
                    plan=plan,
                    template_id=template_id,
                    output_dir=output_dir,
                    auto_download_images=auto_download_images
                )

                self._update_progress(task_id, 90, "PPT制作完成，正在保存...")

            except Exception as e:
                logger.error(f"构建PPT失败: {e}")
                raise Exception(f"制作PPT文件失败: {str(e)}")

            # 完成
            elapsed = time.time() - start_time

            task["status"] = "completed"
            task["progress"] = 100
            task["message"] = "PPT生成完成"
            task["completed_at"] = datetime.now().isoformat()
            task["result"] = {
                "ppt_path": ppt_path,
                "ppt_size": os.path.getsize(ppt_path) if os.path.exists(ppt_path) else 0,
                "slide_count": len(plan.get("slides", [])),
                "elapsed_time": elapsed,
                "title": plan.get("title", topic)
            }

            # 记录成功指标
            record_metric("ppt_generation_completed", 1, tags={
                "template": template_id,
                "success": "true"
            })
            record_metric("ppt_generation_duration", elapsed, tags={
                "template": template_id
            })

            logger.info(f"任务完成: {task_id}, 耗时: {elapsed:.1f}秒")

            # 预热缓存
            if use_cache:
                self._warm_related_cache(topic, audience, template_id)

        except Exception as e:
            # 错误处理
            task["status"] = "failed"
            task["error"] = str(e)
            task["message"] = f"生成失败: {str(e)}"
            task["completed_at"] = datetime.now().isoformat()

            # 记录失败指标
            record_metric("ppt_generation_completed", 1, tags={
                "success": "false",
                "error_type": type(e).__name__
            })

            logger.error(f"任务失败: {task_id}, 错误: {e}")

            # 重试逻辑
            if task["retries"] < task["max_retries"]:
                task["retries"] += 1
                task["status"] = "pending"
                task["message"] = f"准备重试 ({task['retries']}/{task['max_retries']})"

                # 延迟重试
                time.sleep(2 ** task["retries"])

                # 重新提交
                future = self._executor.submit(self._execute_task, task_id)
                self.task_futures[task_id] = future

        finally:
            # 定期清理
            self._maybe_cleanup()

    def _progress_callback(self, task_id: str):
        """生成进度回调函数"""
        def callback(current_batch: int, total_batches: int, message: str):
            # 计算进度 (20% - 50%)
            base_progress = 20
            batch_progress = 30 * current_batch / total_batches
            total_progress = base_progress + batch_progress

            self._update_progress(task_id, int(total_progress), message)

        return callback

    def _update_progress(self, task_id: str, progress: int, message: str):
        """更新任务进度"""
        with self._lock:
            task = self.tasks.get(task_id)
            if task:
                task["progress"] = progress
                task["message"] = message

    def _warm_related_cache(self, topic: str, audience: str, template_id: str):
        """预热相关缓存"""
        try:
            cache = get_cache()

            # 预热模板信息
            template_key = f"template_info:{template_id}"
            if not cache.get(template_key):
                from ppt.template_manager import get_template_manager
                template_info = get_template_manager().get_template_info(template_id)
                cache.set(template_key, template_info, ttl=3600)

        except Exception as e:
            logger.debug(f"预热缓存失败: {e}")

    def _maybe_cleanup(self):
        """定期清理旧任务"""
        now = time.time()
        if now - self._last_cleanup > self._cleanup_interval:
            self.cleanup_old_tasks()
            self._last_cleanup = now

    def get_task(self, task_id: str) -> Optional[Dict]:
        """获取任务状态"""
        with self._lock:
            task = self.tasks.get(task_id)
            if task:
                # 检查future状态
                future = self.task_futures.get(task_id)
                if future and future.done():
                    try:
                        future.result()  # 触发异常处理
                    except Exception as e:
                        if task["status"] == "running":
                            task["status"] = "failed"
                            task["error"] = str(e)
                            task["message"] = f"执行失败: {str(e)}"
            return task

    def cancel_task(self, task_id: str) -> bool:
        """取消任务"""
        with self._lock:
            future = self.task_futures.get(task_id)
            if future and not future.done():
                future.cancel()
                task = self.tasks.get(task_id)
                if task:
                    task["status"] = "cancelled"
                    task["message"] = "任务已取消"
                    task["completed_at"] = datetime.now().isoformat()
                logger.info(f"任务已取消: {task_id}")
                return True
            return False

    def list_tasks(self, status: Optional[str] = None, limit: int = 50) -> List[Dict]:
        """列出任务"""
        with self._lock:
            tasks = list(self.tasks.values())

            if status:
                tasks = [t for t in tasks if t["status"] == status]

            # 按创建时间倒序
            tasks.sort(key=lambda x: x["created_at"], reverse=True)

            # 隐藏敏感信息
            for task in tasks:
                if "config" in task and "ai_config" in task["config"]:
                    ai_config = task["config"]["ai_config"].copy()
                    if "api_key" in ai_config:
                        ai_config["api_key"] = "***"
                    task["config"]["ai_config"] = ai_config

            return tasks[:limit]

    def cleanup_old_tasks(self, max_age: int = 86400) -> int:
        """清理旧任务（默认24小时）"""
        with self._lock:
            cutoff_time = datetime.now().timestamp() - max_age
            old_tasks = [
                task_id for task_id, task in self.tasks.items()
                if datetime.fromisoformat(task["created_at"]).timestamp() < cutoff_time
                and task["status"] in ["completed", "failed", "cancelled"]
            ]

            for task_id in old_tasks:
                self.tasks.pop(task_id, None)
                self.task_futures.pop(task_id, None)

            if old_tasks:
                logger.info(f"清理了 {len(old_tasks)} 个旧任务")

            return len(old_tasks)

    def get_stats(self) -> Dict[str, Any]:
        """获取任务统计"""
        with self._lock:
            stats = {
                "total": len(self.tasks),
                "pending": 0,
                "running": 0,
                "completed": 0,
                "failed": 0,
                "cancelled": 0
            }

            for task in self.tasks.values():
                stats[task["status"]] += 1

            # 计算平均耗时
            completed_tasks = [t for t in self.tasks.values()
                             if t["status"] == "completed" and t["result"]]
            if completed_tasks:
                avg_time = sum(t["result"]["elapsed_time"] for t in completed_tasks) / len(completed_tasks)
                stats["avg_duration"] = round(avg_time, 2)

            # 添加缓存统计
            cache_stats = get_cache().get_stats()
            stats["cache"] = {
                "hit_rate": cache_stats["hit_rate"],
                "size": cache_stats["memory_cache"]["size"]
            }

            # 添加系统指标
            stats["metrics"] = get_metrics_summary()

            return stats

    def shutdown(self):
        """关闭任务管理器"""
        logger.info("正在关闭任务管理器...")

        # 取消所有运行中的任务
        with self._lock:
            for task_id, future in self.task_futures.items():
                if not future.done():
                    future.cancel()
                    task = self.tasks.get(task_id)
                    if task:
                        task["status"] = "cancelled"
                        task["message"] = "系统关闭，任务已取消"

        # 关闭线程池
        self._executor.shutdown(wait=True)
        logger.info("任务管理器已关闭")


# 全局实例
_task_manager: Optional[TaskManager] = None
_manager_lock = threading.Lock()


def get_task_manager() -> TaskManager:
    """获取全局任务管理器实例"""
    global _task_manager
    if _task_manager is None:
        with _manager_lock:
            if _task_manager is None:
                max_workers = int(os.getenv("MAX_WORKERS", "3"))
                enable_async = os.getenv("ENABLE_ASYNC", "true").lower() == "true"
                _task_manager = TaskManager(
                    max_workers=max_workers,
                    enable_async=enable_async
                )

                # 注册清理钩子
                import atexit
                atexit.register(_task_manager.shutdown)

    return _task_manager


# 便捷函数
def submit_ppt_task(config: Dict[str, Any]) -> str:
    """提交PPT生成任务"""
    return get_task_manager().submit_task(config)


def get_task_status(task_id: str) -> Optional[Dict]:
    """获取任务状态"""
    return get_task_manager().get_task(task_id)


def cancel_task(task_id: str) -> bool:
    """取消任务"""
    return get_task_manager().cancel_task(task_id)


def list_recent_tasks(limit: int = 20) -> List[Dict]:
    """列出最近的任务"""
    return get_task_manager().list_tasks(limit=limit)


def get_task_stats() -> Dict[str, Any]:
    """获取任务统计"""
    return get_task_manager().get_stats()