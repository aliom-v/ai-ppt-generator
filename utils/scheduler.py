"""后台任务调度器 - 定时清理缓存和临时文件"""
import threading
import time
from typing import Callable, Optional
from pathlib import Path

from utils.logger import get_logger

logger = get_logger("scheduler")


class BackgroundScheduler:
    """简单的后台任务调度器

    用于定期执行清理任务，不需要额外依赖。

    用法:
        scheduler = BackgroundScheduler()
        scheduler.add_task(cleanup_cache, interval_hours=1)
        scheduler.start()
    """

    def __init__(self):
        self._tasks: list = []
        self._running = False
        self._thread: Optional[threading.Thread] = None

    def add_task(
        self,
        func: Callable,
        interval_hours: float = 1,
        run_immediately: bool = True
    ):
        """添加定时任务

        Args:
            func: 要执行的函数
            interval_hours: 执行间隔（小时）
            run_immediately: 是否立即执行一次
        """
        self._tasks.append({
            'func': func,
            'interval': interval_hours * 3600,
            'last_run': 0 if run_immediately else time.time()
        })

    def start(self):
        """启动调度器"""
        if self._running:
            return

        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        logger.info("后台调度器已启动")

    def stop(self):
        """停止调度器"""
        self._running = False
        if self._thread:
            self._thread.join(timeout=5)
        logger.info("后台调度器已停止")

    def _run(self):
        """调度器主循环"""
        while self._running:
            current_time = time.time()

            for task in self._tasks:
                if current_time - task['last_run'] >= task['interval']:
                    try:
                        task['func']()
                        task['last_run'] = current_time
                    except Exception as e:
                        logger.error(f"定时任务执行失败: {e}")

            # 每分钟检查一次
            time.sleep(60)


def cleanup_expired_cache():
    """清理过期缓存"""
    try:
        from utils.cache import get_cache
        cache = get_cache()
        count = cache.cleanup_expired()
        if count > 0:
            logger.info(f"清理了 {count} 条过期缓存")
    except Exception as e:
        logger.error(f"清理缓存失败: {e}")


def cleanup_old_files(folder: str, max_age_hours: int = 24, max_files: int = 100):
    """清理旧文件

    Args:
        folder: 要清理的目录
        max_age_hours: 文件最大保留时间（小时）
        max_files: 最大保留文件数
    """
    try:
        folder_path = Path(folder)
        if not folder_path.exists():
            return

        files = []
        for pattern in ['*.pptx', '*.pdf', '*.tmp']:
            for f in folder_path.glob(pattern):
                try:
                    files.append((f, f.stat().st_mtime))
                except OSError:
                    continue

        # 按修改时间排序
        files.sort(key=lambda x: x[1], reverse=True)

        current_time = time.time()
        deleted = 0

        for i, (filepath, mtime) in enumerate(files):
            age_hours = (current_time - mtime) / 3600
            if age_hours > max_age_hours or i >= max_files:
                try:
                    filepath.unlink()
                    deleted += 1
                except OSError:
                    pass

        if deleted > 0:
            logger.info(f"清理了 {deleted} 个旧文件: {folder}")

    except Exception as e:
        logger.error(f"清理文件失败: {e}")


def cleanup_image_cache():
    """清理图片缓存"""
    cleanup_old_files("images/downloaded", max_age_hours=24 * 7, max_files=500)
    cleanup_old_files("images/cache", max_age_hours=24 * 7, max_files=500)


def cleanup_output_files():
    """清理输出文件"""
    cleanup_old_files("web/outputs", max_age_hours=24, max_files=100)


# 全局调度器实例
_scheduler: Optional[BackgroundScheduler] = None


def get_scheduler() -> BackgroundScheduler:
    """获取全局调度器"""
    global _scheduler
    if _scheduler is None:
        _scheduler = BackgroundScheduler()
    return _scheduler


def setup_cleanup_tasks():
    """设置清理任务"""
    scheduler = get_scheduler()

    # 每小时清理过期缓存
    scheduler.add_task(cleanup_expired_cache, interval_hours=1, run_immediately=True)

    # 每 6 小时清理图片缓存
    scheduler.add_task(cleanup_image_cache, interval_hours=6, run_immediately=False)

    # 每小时清理输出文件
    scheduler.add_task(cleanup_output_files, interval_hours=1, run_immediately=True)

    scheduler.start()
    return scheduler
