"""历史记录、任务队列和缓存模块"""
import os
import json
import hashlib
import threading
import time
from datetime import datetime
from typing import Optional, List, Dict, Any, Callable
from pathlib import Path
from dataclasses import dataclass, asdict, field
from enum import Enum
from queue import Queue, PriorityQueue
from concurrent.futures import ThreadPoolExecutor
import uuid

from utils.logger import get_logger

logger = get_logger("task_manager")


# ==================== 历史记录模块 ====================

@dataclass
class HistoryRecord:
    """PPT 生成历史记录"""
    id: str
    topic: str
    title: str
    subtitle: str
    slide_count: int
    template_id: str
    template_name: str
    filename: str
    filepath: str
    created_at: str
    file_size: int = 0
    audience: str = "通用受众"
    description: str = ""
    auto_download_images: bool = False
    generation_time: float = 0  # 生成耗时（秒）
    model_name: str = ""
    status: str = "completed"  # completed, failed, deleted

    def to_dict(self) -> Dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict) -> "HistoryRecord":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


class HistoryManager:
    """历史记录管理器"""

    def __init__(self, history_file: str = "web/data/history.json", max_records: int = 100):
        self.history_file = history_file
        self.max_records = max_records
        self._records: List[HistoryRecord] = []
        self._lock = threading.Lock()
        self._load()

    def _ensure_dir(self):
        """确保目录存在"""
        Path(os.path.dirname(self.history_file)).mkdir(parents=True, exist_ok=True)

    def _load(self):
        """加载历史记录"""
        self._ensure_dir()
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self._records = [HistoryRecord.from_dict(r) for r in data]
            except Exception as e:
                logger.error(f"加载历史记录失败: {e}")
                self._records = []

    def _save(self):
        """保存历史记录"""
        self._ensure_dir()
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump([r.to_dict() for r in self._records], f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存历史记录失败: {e}")

    def add(self, record: HistoryRecord) -> str:
        """添加记录"""
        with self._lock:
            # 生成 ID
            if not record.id:
                record.id = str(uuid.uuid4())[:8]

            # 添加到列表头部
            self._records.insert(0, record)

            # 限制数量
            if len(self._records) > self.max_records:
                self._records = self._records[:self.max_records]

            self._save()
            return record.id

    def get(self, record_id: str) -> Optional[HistoryRecord]:
        """获取记录"""
        with self._lock:
            for record in self._records:
                if record.id == record_id:
                    return record
            return None

    def list(self, limit: int = 20, offset: int = 0, status: str = None) -> List[HistoryRecord]:
        """列出记录"""
        with self._lock:
            records = self._records
            if status:
                records = [r for r in records if r.status == status]
            return records[offset:offset + limit]

    def delete(self, record_id: str, delete_file: bool = False) -> bool:
        """删除记录"""
        with self._lock:
            for i, record in enumerate(self._records):
                if record.id == record_id:
                    if delete_file and os.path.exists(record.filepath):
                        try:
                            os.remove(record.filepath)
                        except OSError as e:
                            logger.warning(f"删除文件失败 {record.filepath}: {e}")
                    self._records.pop(i)
                    self._save()
                    return True
            return False

    def clear(self, delete_files: bool = False) -> int:
        """清空历史"""
        with self._lock:
            count = len(self._records)
            if delete_files:
                for record in self._records:
                    if os.path.exists(record.filepath):
                        try:
                            os.remove(record.filepath)
                        except OSError as e:
                            logger.warning(f"删除文件失败 {record.filepath}: {e}")
            self._records = []
            self._save()
            return count

    def count(self) -> int:
        """记录数量"""
        return len(self._records)

    def search(self, keyword: str, limit: int = 20) -> List[HistoryRecord]:
        """搜索记录"""
        with self._lock:
            results = []
            keyword_lower = keyword.lower()
            for record in self._records:
                if (keyword_lower in record.topic.lower() or
                    keyword_lower in record.title.lower() or
                    keyword_lower in record.description.lower()):
                    results.append(record)
                    if len(results) >= limit:
                        break
            return results


# ==================== 任务队列模块 ====================

class TaskStatus(Enum):
    """任务状态"""
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class TaskItem:
    """任务项"""
    id: str
    topic: str
    audience: str
    page_count: int
    template_id: str
    description: str = ""
    auto_download: bool = False
    priority: int = 5  # 1-10，数字越小优先级越高
    status: TaskStatus = TaskStatus.PENDING
    progress: int = 0  # 0-100
    message: str = ""
    result: Dict = field(default_factory=dict)
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    started_at: str = ""
    completed_at: str = ""
    error: str = ""
    # API 配置
    api_key: str = ""
    api_base: str = ""
    model_name: str = ""

    def __lt__(self, other):
        """用于优先级队列比较"""
        return self.priority < other.priority

    def to_dict(self) -> Dict:
        data = asdict(self)
        data['status'] = self.status.value
        # 隐藏敏感信息
        data['api_key'] = '***' if self.api_key else ''
        return data


class TaskQueue:
    """任务队列管理器"""

    def __init__(self, max_workers: int = 2, queue_file: str = "web/data/task_queue.json"):
        self.max_workers = max_workers
        self.queue_file = queue_file
        self._tasks: Dict[str, TaskItem] = {}
        self._queue: PriorityQueue = PriorityQueue()
        self._lock = threading.Lock()
        self._executor: Optional[ThreadPoolExecutor] = None
        self._running = False
        self._callbacks: Dict[str, Callable] = {}
        self._load()

    def _ensure_dir(self):
        Path(os.path.dirname(self.queue_file)).mkdir(parents=True, exist_ok=True)

    def _load(self):
        """加载持久化的任务"""
        self._ensure_dir()
        if os.path.exists(self.queue_file):
            try:
                with open(self.queue_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for item_data in data:
                        item_data['status'] = TaskStatus(item_data['status'])
                        task = TaskItem(**{k: v for k, v in item_data.items() if k in TaskItem.__dataclass_fields__})
                        self._tasks[task.id] = task
                        if task.status == TaskStatus.PENDING:
                            self._queue.put((task.priority, task.id))
            except Exception as e:
                logger.error(f"加载任务队列失败: {e}")

    def _save(self):
        """保存任务队列"""
        self._ensure_dir()
        try:
            with open(self.queue_file, 'w', encoding='utf-8') as f:
                tasks_data = []
                for task in self._tasks.values():
                    task_dict = asdict(task)
                    task_dict['status'] = task.status.value
                    # 不保存敏感信息
                    task_dict['api_key'] = ''
                    tasks_data.append(task_dict)
                json.dump(tasks_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存任务队列失败: {e}")

    def add_task(self, task: TaskItem) -> str:
        """添加任务"""
        with self._lock:
            if not task.id:
                task.id = str(uuid.uuid4())[:8]

            self._tasks[task.id] = task
            self._queue.put((task.priority, task.id))
            self._save()
            return task.id

    def get_task(self, task_id: str) -> Optional[TaskItem]:
        """获取任务"""
        return self._tasks.get(task_id)

    def list_tasks(self, status: TaskStatus = None, limit: int = 50) -> List[TaskItem]:
        """列出任务"""
        tasks = list(self._tasks.values())
        if status:
            tasks = [t for t in tasks if t.status == status]
        # 按创建时间倒序
        tasks.sort(key=lambda x: x.created_at, reverse=True)
        return tasks[:limit]

    def cancel_task(self, task_id: str) -> bool:
        """取消任务"""
        with self._lock:
            if task_id in self._tasks:
                task = self._tasks[task_id]
                if task.status == TaskStatus.PENDING:
                    task.status = TaskStatus.CANCELLED
                    self._save()
                    return True
            return False

    def update_progress(self, task_id: str, progress: int, message: str = ""):
        """更新任务进度"""
        with self._lock:
            if task_id in self._tasks:
                task = self._tasks[task_id]
                task.progress = min(100, max(0, progress))
                if message:
                    task.message = message
                self._save()

    def complete_task(self, task_id: str, result: Dict):
        """完成任务"""
        with self._lock:
            if task_id in self._tasks:
                task = self._tasks[task_id]
                task.status = TaskStatus.COMPLETED
                task.progress = 100
                task.result = result
                task.completed_at = datetime.now().isoformat()
                self._save()

    def fail_task(self, task_id: str, error: str):
        """任务失败"""
        with self._lock:
            if task_id in self._tasks:
                task = self._tasks[task_id]
                task.status = TaskStatus.FAILED
                task.error = error
                task.completed_at = datetime.now().isoformat()
                self._save()

    def start_processing(self, processor: Callable[[TaskItem], Dict]):
        """启动任务处理"""
        if self._running:
            return

        self._running = True
        self._executor = ThreadPoolExecutor(max_workers=self.max_workers)

        def worker():
            while self._running:
                try:
                    if self._queue.empty():
                        time.sleep(1)
                        continue

                    priority, task_id = self._queue.get(timeout=1)
                    task = self._tasks.get(task_id)

                    if not task or task.status != TaskStatus.PENDING:
                        continue

                    # 开始处理
                    task.status = TaskStatus.RUNNING
                    task.started_at = datetime.now().isoformat()
                    self._save()

                    try:
                        result = processor(task)
                        self.complete_task(task_id, result)
                    except Exception as e:
                        self.fail_task(task_id, str(e))

                except Exception as e:
                    if self._running:
                        logger.error(f"任务处理错误: {e}")
                    time.sleep(1)

        # 启动工作线程
        for _ in range(self.max_workers):
            self._executor.submit(worker)

    def stop_processing(self):
        """停止任务处理"""
        self._running = False
        if self._executor:
            self._executor.shutdown(wait=False)

    def clear_completed(self) -> int:
        """清理已完成的任务"""
        with self._lock:
            to_remove = [
                tid for tid, task in self._tasks.items()
                if task.status in (TaskStatus.COMPLETED, TaskStatus.FAILED, TaskStatus.CANCELLED)
            ]
            for tid in to_remove:
                del self._tasks[tid]
            self._save()
            return len(to_remove)


# ==================== 缓存模块 ====================

@dataclass
class CacheEntry:
    """缓存条目"""
    key: str
    value: Any
    created_at: float
    ttl: int  # 生存时间（秒）
    hits: int = 0

    def is_expired(self) -> bool:
        return time.time() - self.created_at > self.ttl


class GenerationCache:
    """PPT 生成缓存 - 用于相似主题复用"""

    def __init__(self, cache_dir: str = "web/data/cache", max_size: int = 50, default_ttl: int = 86400):
        """
        Args:
            cache_dir: 缓存目录
            max_size: 最大缓存条目数
            default_ttl: 默认生存时间（秒），默认24小时
        """
        self.cache_dir = cache_dir
        self.max_size = max_size
        self.default_ttl = default_ttl
        self._cache: Dict[str, CacheEntry] = {}
        self._lock = threading.Lock()
        self._index_file = os.path.join(cache_dir, "cache_index.json")
        self._load()

    def _ensure_dir(self):
        Path(self.cache_dir).mkdir(parents=True, exist_ok=True)

    def _load(self):
        """加载缓存索引"""
        self._ensure_dir()
        if os.path.exists(self._index_file):
            try:
                with open(self._index_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for key, entry_data in data.items():
                        entry = CacheEntry(**entry_data)
                        if not entry.is_expired():
                            self._cache[key] = entry
            except Exception as e:
                logger.error(f"加载缓存失败: {e}")

    def _save(self):
        """保存缓存索引"""
        self._ensure_dir()
        try:
            data = {k: asdict(v) for k, v in self._cache.items() if not v.is_expired()}
            with open(self._index_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存缓存失败: {e}")

    def _make_key(self, topic: str, page_count: int, template_id: str) -> str:
        """生成缓存键"""
        # 标准化主题（小写、去空格）
        normalized_topic = topic.lower().strip()
        key_str = f"{normalized_topic}_{page_count}_{template_id}"
        return hashlib.md5(key_str.encode()).hexdigest()

    def _similarity_key(self, topic: str) -> str:
        """生成相似度键（用于模糊匹配）"""
        # 提取关键词
        words = set(topic.lower().split())
        # 移除常见停用词
        stop_words = {'的', '是', '在', '和', '了', 'a', 'an', 'the', 'is', 'are', 'in', 'on', 'for'}
        words = words - stop_words
        return '_'.join(sorted(words))

    def get(self, topic: str, page_count: int, template_id: str) -> Optional[Dict]:
        """获取缓存"""
        with self._lock:
            key = self._make_key(topic, page_count, template_id)

            if key in self._cache:
                entry = self._cache[key]
                if not entry.is_expired():
                    entry.hits += 1
                    # 从文件加载实际内容
                    cache_file = os.path.join(self.cache_dir, f"{key}.json")
                    if os.path.exists(cache_file):
                        try:
                            with open(cache_file, 'r', encoding='utf-8') as f:
                                return json.load(f)
                        except (IOError, json.JSONDecodeError) as e:
                            logger.warning(f"读取缓存文件失败 {cache_file}: {e}")
                else:
                    del self._cache[key]

            return None

    def get_similar(self, topic: str, page_count: int, tolerance: int = 2) -> Optional[Dict]:
        """获取相似主题的缓存"""
        with self._lock:
            sim_key = self._similarity_key(topic)

            for key, entry in self._cache.items():
                if entry.is_expired():
                    continue

                # 检查缓存内容
                cache_file = os.path.join(self.cache_dir, f"{key}.json")
                if os.path.exists(cache_file):
                    try:
                        with open(cache_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            cached_topic = data.get('_meta', {}).get('topic', '')
                            cached_pages = data.get('_meta', {}).get('page_count', 0)

                            # 检查相似度
                            if (self._similarity_key(cached_topic) == sim_key and
                                abs(cached_pages - page_count) <= tolerance):
                                entry.hits += 1
                                return data
                    except (IOError, json.JSONDecodeError) as e:
                        logger.debug(f"读取缓存文件失败 {cache_file}: {e}")
                        continue

            return None

    def set(self, topic: str, page_count: int, template_id: str, value: Dict, ttl: int = None):
        """设置缓存"""
        with self._lock:
            key = self._make_key(topic, page_count, template_id)

            # 添加元数据
            value['_meta'] = {
                'topic': topic,
                'page_count': page_count,
                'template_id': template_id,
                'cached_at': datetime.now().isoformat()
            }

            # 保存到文件
            cache_file = os.path.join(self.cache_dir, f"{key}.json")
            try:
                with open(cache_file, 'w', encoding='utf-8') as f:
                    json.dump(value, f, ensure_ascii=False, indent=2)
            except Exception as e:
                logger.error(f"写入缓存文件失败: {e}")
                return

            # 更新索引
            self._cache[key] = CacheEntry(
                key=key,
                value=cache_file,  # 存储文件路径
                created_at=time.time(),
                ttl=ttl or self.default_ttl
            )

            # 限制缓存大小
            self._evict_if_needed()
            self._save()

    def _evict_if_needed(self):
        """如果超出容量，驱逐旧条目"""
        if len(self._cache) <= self.max_size:
            return

        # 按访问次数和创建时间排序，驱逐最不常用的
        entries = sorted(
            self._cache.items(),
            key=lambda x: (x[1].hits, x[1].created_at)
        )

        # 删除超出的条目
        to_remove = len(self._cache) - self.max_size
        for i in range(to_remove):
            key = entries[i][0]
            # 删除缓存文件
            cache_file = os.path.join(self.cache_dir, f"{key}.json")
            if os.path.exists(cache_file):
                try:
                    os.remove(cache_file)
                except OSError as e:
                    logger.warning(f"删除缓存文件失败 {cache_file}: {e}")
            del self._cache[key]

    def invalidate(self, topic: str, page_count: int, template_id: str):
        """使缓存失效"""
        with self._lock:
            key = self._make_key(topic, page_count, template_id)
            if key in self._cache:
                # 删除缓存文件
                cache_file = os.path.join(self.cache_dir, f"{key}.json")
                if os.path.exists(cache_file):
                    try:
                        os.remove(cache_file)
                    except OSError as e:
                        logger.warning(f"删除缓存文件失败 {cache_file}: {e}")
                del self._cache[key]
                self._save()

    def clear(self) -> int:
        """清空缓存"""
        with self._lock:
            count = len(self._cache)

            # 删除所有缓存文件
            for key in self._cache:
                cache_file = os.path.join(self.cache_dir, f"{key}.json")
                if os.path.exists(cache_file):
                    try:
                        os.remove(cache_file)
                    except OSError as e:
                        logger.warning(f"删除缓存文件失败 {cache_file}: {e}")

            self._cache = {}
            self._save()
            return count

    def stats(self) -> Dict:
        """获取缓存统计"""
        with self._lock:
            total_hits = sum(e.hits for e in self._cache.values())
            expired = sum(1 for e in self._cache.values() if e.is_expired())
            return {
                'total_entries': len(self._cache),
                'active_entries': len(self._cache) - expired,
                'expired_entries': expired,
                'total_hits': total_hits,
                'max_size': self.max_size
            }


# ==================== 全局实例 ====================

_history_manager: Optional[HistoryManager] = None
_task_queue: Optional[TaskQueue] = None
_generation_cache: Optional[GenerationCache] = None
_global_lock = threading.Lock()


def get_history_manager() -> HistoryManager:
    """获取历史管理器实例（线程安全）"""
    global _history_manager
    if _history_manager is None:
        with _global_lock:
            if _history_manager is None:
                _history_manager = HistoryManager()
    return _history_manager


def get_task_queue() -> TaskQueue:
    """获取任务队列实例（线程安全）"""
    global _task_queue
    if _task_queue is None:
        with _global_lock:
            if _task_queue is None:
                _task_queue = TaskQueue()
    return _task_queue


def get_generation_cache() -> GenerationCache:
    """获取生成缓存实例（线程安全）"""
    global _generation_cache
    if _generation_cache is None:
        with _global_lock:
            if _generation_cache is None:
                _generation_cache = GenerationCache()
    return _generation_cache
