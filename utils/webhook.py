"""Webhook 通知模块

支持任务完成后发送 Webhook 通知。
"""
import hashlib
import hmac
import json
import threading
import time
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional
from dataclasses import dataclass, field
from enum import Enum
from urllib.parse import urlparse

import requests

from utils.logger import get_logger
from utils.retry import retry

logger = get_logger("webhook")


class WebhookEvent(str, Enum):
    """Webhook 事件类型"""
    TASK_CREATED = "task.created"
    TASK_STARTED = "task.started"
    TASK_PROGRESS = "task.progress"
    TASK_COMPLETED = "task.completed"
    TASK_FAILED = "task.failed"
    BATCH_COMPLETED = "batch.completed"


@dataclass
class WebhookConfig:
    """Webhook 配置"""
    url: str
    secret: str = ""  # 用于签名验证
    events: List[str] = field(default_factory=lambda: ["task.completed", "task.failed"])
    headers: Dict[str, str] = field(default_factory=dict)
    timeout: int = 30
    retry_count: int = 3
    enabled: bool = True

    def to_dict(self) -> Dict:
        return {
            "url": self.url,
            "events": self.events,
            "enabled": self.enabled,
        }


@dataclass
class WebhookPayload:
    """Webhook 负载"""
    event: str
    timestamp: str
    data: Dict[str, Any]
    webhook_id: str = ""

    def to_dict(self) -> Dict:
        return {
            "event": self.event,
            "timestamp": self.timestamp,
            "data": self.data,
            "webhook_id": self.webhook_id,
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False)


@dataclass
class WebhookDelivery:
    """Webhook 投递记录"""
    webhook_id: str
    url: str
    event: str
    status: str  # pending, success, failed
    attempts: int = 0
    response_code: Optional[int] = None
    response_body: Optional[str] = None
    error: Optional[str] = None
    created_at: float = field(default_factory=time.time)
    delivered_at: Optional[float] = None

    def to_dict(self) -> Dict:
        return {
            "webhook_id": self.webhook_id,
            "url": self.url,
            "event": self.event,
            "status": self.status,
            "attempts": self.attempts,
            "response_code": self.response_code,
            "error": self.error,
            "created_at": datetime.fromtimestamp(self.created_at).isoformat(),
            "delivered_at": datetime.fromtimestamp(self.delivered_at).isoformat() if self.delivered_at else None,
        }


class WebhookSigner:
    """Webhook 签名器"""

    @staticmethod
    def sign(payload: str, secret: str) -> str:
        """生成 HMAC-SHA256 签名"""
        return hmac.new(
            secret.encode("utf-8"),
            payload.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()

    @staticmethod
    def verify(payload: str, signature: str, secret: str) -> bool:
        """验证签名"""
        expected = WebhookSigner.sign(payload, secret)
        return hmac.compare_digest(expected, signature)


class WebhookSender:
    """Webhook 发送器"""

    def __init__(self, config: WebhookConfig):
        self.config = config

    def send(self, payload: WebhookPayload) -> WebhookDelivery:
        """发送 Webhook"""
        import uuid
        webhook_id = str(uuid.uuid4())[:8]
        payload.webhook_id = webhook_id

        delivery = WebhookDelivery(
            webhook_id=webhook_id,
            url=self.config.url,
            event=payload.event,
            status="pending",
        )

        if not self.config.enabled:
            delivery.status = "skipped"
            return delivery

        # 检查事件是否在订阅列表中
        if payload.event not in self.config.events and "*" not in self.config.events:
            delivery.status = "filtered"
            return delivery

        # 构建请求头
        headers = {
            "Content-Type": "application/json",
            "User-Agent": "AI-PPT-Generator-Webhook/1.0",
            "X-Webhook-Event": payload.event,
            "X-Webhook-ID": webhook_id,
            **self.config.headers,
        }

        # 添加签名
        payload_json = payload.to_json()
        if self.config.secret:
            signature = WebhookSigner.sign(payload_json, self.config.secret)
            headers["X-Webhook-Signature"] = f"sha256={signature}"

        # 发送请求（带重试）
        last_error = None
        for attempt in range(self.config.retry_count):
            delivery.attempts = attempt + 1
            try:
                response = requests.post(
                    self.config.url,
                    data=payload_json,
                    headers=headers,
                    timeout=self.config.timeout,
                )
                delivery.response_code = response.status_code
                delivery.response_body = response.text[:500]  # 截断

                if 200 <= response.status_code < 300:
                    delivery.status = "success"
                    delivery.delivered_at = time.time()
                    logger.info(f"Webhook 发送成功: {webhook_id} -> {self.config.url}")
                    return delivery
                else:
                    last_error = f"HTTP {response.status_code}"

            except requests.Timeout:
                last_error = "请求超时"
            except requests.RequestException as e:
                last_error = str(e)

            # 重试前等待
            if attempt < self.config.retry_count - 1:
                time.sleep(2 ** attempt)  # 指数退避

        delivery.status = "failed"
        delivery.error = last_error
        logger.warning(f"Webhook 发送失败: {webhook_id} -> {self.config.url}: {last_error}")
        return delivery


class WebhookManager:
    """Webhook 管理器

    管理多个 Webhook 配置和发送。

    用法:
        manager = WebhookManager()

        # 注册 Webhook
        manager.register(WebhookConfig(
            url="https://example.com/webhook",
            events=["task.completed"],
            secret="my-secret",
        ))

        # 触发事件
        manager.trigger("task.completed", {
            "task_id": "123",
            "filename": "output.pptx",
        })
    """

    def __init__(self, async_send: bool = True):
        self._webhooks: Dict[str, WebhookConfig] = {}
        self._deliveries: List[WebhookDelivery] = []
        self._lock = threading.RLock()
        self._async_send = async_send
        self._max_deliveries = 1000

    def register(self, config: WebhookConfig, webhook_id: str = None) -> str:
        """注册 Webhook"""
        import uuid
        if not webhook_id:
            webhook_id = str(uuid.uuid4())[:8]

        # 验证 URL
        try:
            parsed = urlparse(config.url)
            if parsed.scheme not in ("http", "https"):
                raise ValueError("URL 必须是 HTTP 或 HTTPS")
        except Exception as e:
            raise ValueError(f"无效的 Webhook URL: {e}")

        with self._lock:
            self._webhooks[webhook_id] = config

        logger.info(f"注册 Webhook: {webhook_id} -> {config.url}")
        return webhook_id

    def unregister(self, webhook_id: str) -> bool:
        """取消注册 Webhook"""
        with self._lock:
            if webhook_id in self._webhooks:
                del self._webhooks[webhook_id]
                logger.info(f"取消注册 Webhook: {webhook_id}")
                return True
        return False

    def trigger(self, event: str, data: Dict[str, Any]):
        """触发事件"""
        payload = WebhookPayload(
            event=event,
            timestamp=datetime.utcnow().isoformat() + "Z",
            data=data,
        )

        with self._lock:
            webhooks = list(self._webhooks.values())

        for config in webhooks:
            if self._async_send:
                thread = threading.Thread(
                    target=self._send_webhook,
                    args=(config, payload),
                    daemon=True,
                )
                thread.start()
            else:
                self._send_webhook(config, payload)

    def _send_webhook(self, config: WebhookConfig, payload: WebhookPayload):
        """发送单个 Webhook"""
        sender = WebhookSender(config)
        delivery = sender.send(payload)

        with self._lock:
            self._deliveries.append(delivery)
            # 限制历史记录数量
            if len(self._deliveries) > self._max_deliveries:
                self._deliveries = self._deliveries[-self._max_deliveries:]

    def get_webhooks(self) -> List[Dict]:
        """获取所有 Webhook 配置"""
        with self._lock:
            return [
                {"id": wid, **config.to_dict()}
                for wid, config in self._webhooks.items()
            ]

    def get_deliveries(self, limit: int = 50) -> List[Dict]:
        """获取最近的投递记录"""
        with self._lock:
            return [d.to_dict() for d in self._deliveries[-limit:]]

    def get_stats(self) -> Dict:
        """获取统计信息"""
        with self._lock:
            total = len(self._deliveries)
            success = sum(1 for d in self._deliveries if d.status == "success")
            failed = sum(1 for d in self._deliveries if d.status == "failed")

            return {
                "webhook_count": len(self._webhooks),
                "total_deliveries": total,
                "successful": success,
                "failed": failed,
                "success_rate": round(success / max(total, 1) * 100, 2),
            }


# 全局 Webhook 管理器
_webhook_manager: Optional[WebhookManager] = None


def get_webhook_manager() -> WebhookManager:
    """获取全局 Webhook 管理器"""
    global _webhook_manager
    if _webhook_manager is None:
        _webhook_manager = WebhookManager()
    return _webhook_manager


def trigger_webhook(event: str, data: Dict[str, Any]):
    """触发 Webhook 事件（便捷函数）"""
    get_webhook_manager().trigger(event, data)


# 任务事件触发辅助函数
def notify_task_completed(task_id: str, result: Dict):
    """通知任务完成"""
    trigger_webhook(WebhookEvent.TASK_COMPLETED.value, {
        "task_id": task_id,
        **result,
    })


def notify_task_failed(task_id: str, error: str):
    """通知任务失败"""
    trigger_webhook(WebhookEvent.TASK_FAILED.value, {
        "task_id": task_id,
        "error": error,
    })


def notify_batch_completed(job_id: str, stats: Dict):
    """通知批量任务完成"""
    trigger_webhook(WebhookEvent.BATCH_COMPLETED.value, {
        "job_id": job_id,
        **stats,
    })
