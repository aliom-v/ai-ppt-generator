"""Server-Sent Events (SSE) 模块

提供轻量级的实时进度推送功能，无需额外依赖。
"""
import json
import queue
import threading
import time
from typing import Any, Dict, Generator, Optional
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class SSEEvent:
    """SSE 事件"""
    event: str  # 事件类型
    data: Any   # 事件数据
    id: Optional[str] = None  # 事件 ID
    retry: Optional[int] = None  # 重连间隔（毫秒）

    def serialize(self) -> str:
        """序列化为 SSE 格式"""
        lines = []

        if self.id:
            lines.append(f"id: {self.id}")

        if self.retry:
            lines.append(f"retry: {self.retry}")

        if self.event:
            lines.append(f"event: {self.event}")

        # 数据序列化
        if isinstance(self.data, (dict, list)):
            data_str = json.dumps(self.data, ensure_ascii=False)
        else:
            data_str = str(self.data)

        # 多行数据处理
        for line in data_str.split("\n"):
            lines.append(f"data: {line}")

        return "\n".join(lines) + "\n\n"


class SSEChannel:
    """SSE 通道

    管理单个客户端连接的事件队列。
    """

    def __init__(self, channel_id: str, timeout: float = 300.0):
        self.channel_id = channel_id
        self.timeout = timeout
        self._queue: queue.Queue = queue.Queue()
        self._active = True
        self._created_at = time.time()
        self._last_activity = time.time()

    def send(self, event: str, data: Any, event_id: str = None):
        """发送事件到通道"""
        if not self._active:
            return

        sse_event = SSEEvent(
            event=event,
            data=data,
            id=event_id or f"{self.channel_id}-{int(time.time() * 1000)}",
        )
        self._queue.put(sse_event)
        self._last_activity = time.time()

    def send_progress(self, progress: int, message: str = "", extra: Dict = None):
        """发送进度更新"""
        data = {
            "progress": progress,
            "message": message,
            "timestamp": datetime.now().isoformat(),
        }
        if extra:
            data.update(extra)
        self.send("progress", data)

    def send_complete(self, result: Dict):
        """发送完成事件"""
        self.send("complete", result)

    def send_error(self, error: str, code: str = "ERROR"):
        """发送错误事件"""
        self.send("error", {"error": error, "code": code})

    def close(self):
        """关闭通道"""
        self._active = False
        # 发送关闭事件
        self._queue.put(SSEEvent(event="close", data={"reason": "channel_closed"}))

    def stream(self) -> Generator[str, None, None]:
        """生成 SSE 事件流"""
        # 发送连接成功事件
        yield SSEEvent(
            event="connected",
            data={"channel_id": self.channel_id},
            retry=3000,  # 3 秒重连
        ).serialize()

        while self._active:
            try:
                event = self._queue.get(timeout=30)  # 30 秒超时
                yield event.serialize()

                if event.event == "close":
                    break

            except queue.Empty:
                # 发送心跳保持连接
                yield SSEEvent(event="heartbeat", data={"time": time.time()}).serialize()

                # 检查超时
                if time.time() - self._last_activity > self.timeout:
                    self.close()
                    break

    @property
    def is_active(self) -> bool:
        return self._active

    @property
    def age(self) -> float:
        """通道存活时间（秒）"""
        return time.time() - self._created_at


class SSEManager:
    """SSE 管理器

    管理所有 SSE 通道。

    用法:
        from flask import Response

        sse = SSEManager()

        @app.route('/events/<channel_id>')
        def stream(channel_id):
            channel = sse.create_channel(channel_id)
            return Response(
                channel.stream(),
                mimetype='text/event-stream',
                headers={
                    'Cache-Control': 'no-cache',
                    'Connection': 'keep-alive',
                }
            )

        # 在其他地方发送事件
        sse.send_to(channel_id, "update", {"status": "processing"})
    """

    def __init__(self, max_channels: int = 1000, cleanup_interval: int = 60):
        self._channels: Dict[str, SSEChannel] = {}
        self._lock = threading.RLock()
        self._max_channels = max_channels
        self._cleanup_interval = cleanup_interval
        self._start_cleanup_thread()

    def _start_cleanup_thread(self):
        """启动清理线程"""
        def cleanup():
            while True:
                time.sleep(self._cleanup_interval)
                self._cleanup_inactive_channels()

        thread = threading.Thread(target=cleanup, daemon=True)
        thread.start()

    def _cleanup_inactive_channels(self):
        """清理不活跃的通道"""
        with self._lock:
            inactive = [
                cid for cid, channel in self._channels.items()
                if not channel.is_active
            ]
            for cid in inactive:
                del self._channels[cid]

    def create_channel(self, channel_id: str, timeout: float = 300.0) -> SSEChannel:
        """创建或获取通道"""
        with self._lock:
            if channel_id in self._channels:
                channel = self._channels[channel_id]
                if channel.is_active:
                    return channel

            # 检查容量
            if len(self._channels) >= self._max_channels:
                self._cleanup_inactive_channels()
                if len(self._channels) >= self._max_channels:
                    # 删除最旧的通道
                    oldest_id = min(
                        self._channels.keys(),
                        key=lambda k: self._channels[k]._created_at
                    )
                    self._channels[oldest_id].close()
                    del self._channels[oldest_id]

            channel = SSEChannel(channel_id, timeout)
            self._channels[channel_id] = channel
            return channel

    def get_channel(self, channel_id: str) -> Optional[SSEChannel]:
        """获取通道"""
        with self._lock:
            channel = self._channels.get(channel_id)
            if channel and channel.is_active:
                return channel
            return None

    def send_to(self, channel_id: str, event: str, data: Any) -> bool:
        """发送事件到指定通道"""
        channel = self.get_channel(channel_id)
        if channel:
            channel.send(event, data)
            return True
        return False

    def send_progress(self, channel_id: str, progress: int, message: str = "") -> bool:
        """发送进度到指定通道"""
        channel = self.get_channel(channel_id)
        if channel:
            channel.send_progress(progress, message)
            return True
        return False

    def close_channel(self, channel_id: str):
        """关闭通道"""
        with self._lock:
            channel = self._channels.get(channel_id)
            if channel:
                channel.close()

    def get_stats(self) -> Dict[str, Any]:
        """获取统计信息"""
        with self._lock:
            active = sum(1 for c in self._channels.values() if c.is_active)
            return {
                "total_channels": len(self._channels),
                "active_channels": active,
                "max_channels": self._max_channels,
            }


# 全局 SSE 管理器
_sse_manager: Optional[SSEManager] = None


def get_sse_manager() -> SSEManager:
    """获取全局 SSE 管理器"""
    global _sse_manager
    if _sse_manager is None:
        _sse_manager = SSEManager()
    return _sse_manager


def create_sse_response(channel_id: str):
    """创建 SSE 响应（Flask 辅助函数）"""
    from flask import Response

    channel = get_sse_manager().create_channel(channel_id)

    return Response(
        channel.stream(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",  # 禁用 Nginx 缓冲
        },
    )
