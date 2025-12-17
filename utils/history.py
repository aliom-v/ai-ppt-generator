"""生成历史记录模块 - SQLite 存储"""
import sqlite3
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any
from contextlib import contextmanager

from utils.logger import get_logger

logger = get_logger("history")


class GenerationHistory:
    """PPT 生成历史记录

    使用 SQLite 存储生成历史，支持查询和统计。
    """

    def __init__(self, db_path: str = "data/history.db"):
        self.db_path = db_path
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _get_conn(self):
        """获取数据库连接"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_db(self):
        """初始化数据库表"""
        with self._get_conn() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS generation_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    topic TEXT NOT NULL,
                    audience TEXT,
                    page_count INTEGER,
                    model_name TEXT,
                    template_id TEXT,
                    filename TEXT,
                    file_size INTEGER,
                    slide_count INTEGER,
                    duration_ms INTEGER,
                    status TEXT DEFAULT 'success',
                    error_message TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    request_id TEXT,
                    client_ip TEXT
                )
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_created_at ON generation_history(created_at)
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_status ON generation_history(status)
            """)

    def add(
        self,
        topic: str,
        audience: str = "",
        page_count: int = 0,
        model_name: str = "",
        template_id: str = "",
        filename: str = "",
        file_size: int = 0,
        slide_count: int = 0,
        duration_ms: int = 0,
        status: str = "success",
        error_message: str = "",
        request_id: str = "",
        client_ip: str = "",
    ) -> int:
        """添加历史记录

        Returns:
            记录 ID
        """
        with self._get_conn() as conn:
            cursor = conn.execute("""
                INSERT INTO generation_history
                (topic, audience, page_count, model_name, template_id, filename,
                 file_size, slide_count, duration_ms, status, error_message,
                 request_id, client_ip)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (topic, audience, page_count, model_name, template_id, filename,
                  file_size, slide_count, duration_ms, status, error_message,
                  request_id, client_ip))
            return cursor.lastrowid

    def get_recent(self, limit: int = 20, offset: int = 0) -> List[Dict]:
        """获取最近的历史记录"""
        with self._get_conn() as conn:
            rows = conn.execute("""
                SELECT * FROM generation_history
                ORDER BY created_at DESC
                LIMIT ? OFFSET ?
            """, (limit, offset)).fetchall()
            return [dict(row) for row in rows]

    def get_by_id(self, record_id: int) -> Optional[Dict]:
        """根据 ID 获取记录"""
        with self._get_conn() as conn:
            row = conn.execute("""
                SELECT * FROM generation_history WHERE id = ?
            """, (record_id,)).fetchone()
            return dict(row) if row else None

    def get_stats(self) -> Dict[str, Any]:
        """获取统计信息"""
        with self._get_conn() as conn:
            # 总数和成功率
            total = conn.execute(
                "SELECT COUNT(*) FROM generation_history"
            ).fetchone()[0]

            success = conn.execute(
                "SELECT COUNT(*) FROM generation_history WHERE status = 'success'"
            ).fetchone()[0]

            # 今日统计
            today = datetime.now().strftime("%Y-%m-%d")
            today_count = conn.execute("""
                SELECT COUNT(*) FROM generation_history
                WHERE DATE(created_at) = ?
            """, (today,)).fetchone()[0]

            # 平均耗时
            avg_duration = conn.execute("""
                SELECT AVG(duration_ms) FROM generation_history
                WHERE status = 'success' AND duration_ms > 0
            """).fetchone()[0] or 0

            # 热门主题
            popular_topics = conn.execute("""
                SELECT topic, COUNT(*) as count
                FROM generation_history
                GROUP BY topic
                ORDER BY count DESC
                LIMIT 5
            """).fetchall()

            return {
                "total_generations": total,
                "successful": success,
                "failed": total - success,
                "success_rate": round(success / max(total, 1) * 100, 2),
                "today_count": today_count,
                "avg_duration_ms": round(avg_duration, 2),
                "popular_topics": [
                    {"topic": row[0], "count": row[1]}
                    for row in popular_topics
                ],
            }

    def search(
        self,
        keyword: str = "",
        status: str = "",
        start_date: str = "",
        end_date: str = "",
        limit: int = 20,
    ) -> List[Dict]:
        """搜索历史记录"""
        conditions = []
        params = []

        if keyword:
            conditions.append("topic LIKE ?")
            params.append(f"%{keyword}%")

        if status:
            conditions.append("status = ?")
            params.append(status)

        if start_date:
            conditions.append("DATE(created_at) >= ?")
            params.append(start_date)

        if end_date:
            conditions.append("DATE(created_at) <= ?")
            params.append(end_date)

        where_clause = " AND ".join(conditions) if conditions else "1=1"

        with self._get_conn() as conn:
            rows = conn.execute(f"""
                SELECT * FROM generation_history
                WHERE {where_clause}
                ORDER BY created_at DESC
                LIMIT ?
            """, params + [limit]).fetchall()
            return [dict(row) for row in rows]

    def cleanup_old(self, days: int = 30) -> int:
        """清理旧记录

        Returns:
            删除的记录数
        """
        with self._get_conn() as conn:
            cursor = conn.execute("""
                DELETE FROM generation_history
                WHERE created_at < datetime('now', ?)
            """, (f"-{days} days",))
            count = cursor.rowcount
            if count > 0:
                logger.info(f"清理了 {count} 条历史记录（超过 {days} 天）")
            return count


# 全局实例
_history: Optional[GenerationHistory] = None


def get_history() -> GenerationHistory:
    """获取全局历史记录实例"""
    global _history
    if _history is None:
        _history = GenerationHistory()
    return _history
