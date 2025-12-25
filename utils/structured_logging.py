"""结构化日志模块

支持 JSON 格式日志输出，适合生产环境日志收集。
"""
import json
import logging
import sys
import traceback
from datetime import datetime, timezone
from typing import Any, Dict, Optional
from dataclasses import dataclass, asdict


@dataclass
class LogRecord:
    """结构化日志记录"""
    timestamp: str
    level: str
    logger: str
    message: str
    request_id: Optional[str] = None
    extra: Optional[Dict[str, Any]] = None
    exception: Optional[str] = None
    stack_trace: Optional[str] = None

    def to_dict(self) -> Dict:
        data = {
            "timestamp": self.timestamp,
            "level": self.level,
            "logger": self.logger,
            "message": self.message,
        }
        if self.request_id:
            data["request_id"] = self.request_id
        if self.extra:
            data.update(self.extra)
        if self.exception:
            data["exception"] = self.exception
        if self.stack_trace:
            data["stack_trace"] = self.stack_trace
        return data


class JSONFormatter(logging.Formatter):
    """JSON 格式化器

    将日志记录格式化为 JSON 格式。
    """

    def __init__(self, include_stack_trace: bool = True):
        super().__init__()
        self.include_stack_trace = include_stack_trace

    def format(self, record: logging.LogRecord) -> str:
        # 获取请求 ID
        request_id = getattr(record, "request_id", None)
        if not request_id:
            try:
                from utils.request_context import get_request_id
                request_id = get_request_id()
            except Exception:
                pass

        # 构建日志记录
        log_record = LogRecord(
            timestamp=datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            level=record.levelname,
            logger=record.name,
            message=record.getMessage(),
            request_id=request_id,
        )

        # 提取额外字段
        extra = {}
        for key, value in record.__dict__.items():
            if key not in (
                "name", "msg", "args", "created", "filename", "funcName",
                "levelname", "levelno", "lineno", "module", "msecs",
                "pathname", "process", "processName", "relativeCreated",
                "stack_info", "exc_info", "exc_text", "thread", "threadName",
                "request_id", "message",
            ):
                try:
                    json.dumps(value)  # 测试是否可序列化
                    extra[key] = value
                except (TypeError, ValueError):
                    extra[key] = str(value)

        if extra:
            log_record.extra = extra

        # 处理异常
        if record.exc_info:
            log_record.exception = str(record.exc_info[1])
            if self.include_stack_trace:
                log_record.stack_trace = "".join(
                    traceback.format_exception(*record.exc_info)
                )

        return json.dumps(log_record.to_dict(), ensure_ascii=False)


class PrettyFormatter(logging.Formatter):
    """美化格式化器

    用于开发环境，带颜色的可读格式。
    """

    COLORS = {
        "DEBUG": "\033[36m",     # 青色
        "INFO": "\033[32m",      # 绿色
        "WARNING": "\033[33m",   # 黄色
        "ERROR": "\033[31m",     # 红色
        "CRITICAL": "\033[35m",  # 紫色
    }
    RESET = "\033[0m"

    def __init__(self, use_colors: bool = True):
        super().__init__()
        self.use_colors = use_colors

    def format(self, record: logging.LogRecord) -> str:
        # 时间戳
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 请求 ID
        request_id = getattr(record, "request_id", "-")

        # 级别颜色
        level = record.levelname
        if self.use_colors:
            color = self.COLORS.get(level, "")
            level = f"{color}{level:8}{self.RESET}"
        else:
            level = f"{level:8}"

        # 格式化消息
        message = record.getMessage()

        # 组装
        line = f"{timestamp} | {level} | {record.name:15} | [{request_id}] {message}"

        # 异常信息
        if record.exc_info:
            line += "\n" + "".join(traceback.format_exception(*record.exc_info))

        return line


class StructuredLogger:
    """结构化日志器

    支持添加上下文字段的日志器包装。

    用法:
        logger = StructuredLogger("myapp")
        logger.info("用户登录", user_id=123, ip="192.168.1.1")
        # 输出: {"timestamp": "...", "level": "INFO", "message": "用户登录", "user_id": 123, "ip": "192.168.1.1"}
    """

    def __init__(self, name: str, base_logger: logging.Logger = None):
        self.logger = base_logger or logging.getLogger(name)
        self._context: Dict[str, Any] = {}

    def bind(self, **kwargs) -> "StructuredLogger":
        """绑定上下文字段，返回新的日志器"""
        new_logger = StructuredLogger(self.logger.name, self.logger)
        new_logger._context = {**self._context, **kwargs}
        return new_logger

    def _log(self, level: int, message: str, exc_info=None, **kwargs):
        """内部日志方法"""
        extra = {**self._context, **kwargs}
        self.logger.log(level, message, exc_info=exc_info, extra=extra)

    def debug(self, message: str, **kwargs):
        self._log(logging.DEBUG, message, **kwargs)

    def info(self, message: str, **kwargs):
        self._log(logging.INFO, message, **kwargs)

    def warning(self, message: str, **kwargs):
        self._log(logging.WARNING, message, **kwargs)

    def error(self, message: str, exc_info=None, **kwargs):
        self._log(logging.ERROR, message, exc_info=exc_info, **kwargs)

    def critical(self, message: str, exc_info=None, **kwargs):
        self._log(logging.CRITICAL, message, exc_info=exc_info, **kwargs)

    def exception(self, message: str, **kwargs):
        self._log(logging.ERROR, message, exc_info=True, **kwargs)


def setup_logging(
    level: str = "INFO",
    json_format: bool = False,
    log_file: Optional[str] = None,
    include_stack_trace: bool = True,
):
    """配置日志系统

    Args:
        level: 日志级别
        json_format: 是否使用 JSON 格式
        log_file: 日志文件路径
        include_stack_trace: 是否包含堆栈跟踪
    """
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, level.upper()))

    # 清除现有处理器
    root_logger.handlers.clear()

    # 选择格式化器
    if json_format:
        formatter = JSONFormatter(include_stack_trace=include_stack_trace)
    else:
        # 检查是否是终端
        use_colors = hasattr(sys.stdout, "isatty") and sys.stdout.isatty()
        formatter = PrettyFormatter(use_colors=use_colors)

    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # 文件处理器
    if log_file:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        # 文件总是用 JSON 格式
        file_handler.setFormatter(JSONFormatter(include_stack_trace=True))
        root_logger.addHandler(file_handler)

    return root_logger


def get_structured_logger(name: str) -> StructuredLogger:
    """获取结构化日志器"""
    return StructuredLogger(name)


# 日志上下文管理器
class LogContext:
    """日志上下文管理器

    在上下文中自动添加额外字段。

    用法:
        with LogContext(user_id=123, action="login"):
            logger.info("开始处理")  # 自动包含 user_id 和 action
    """

    _current: Dict[str, Any] = {}

    def __init__(self, **kwargs):
        self._kwargs = kwargs
        self._previous: Dict[str, Any] = {}

    def __enter__(self):
        self._previous = LogContext._current.copy()
        LogContext._current.update(self._kwargs)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        LogContext._current = self._previous
        return False

    @classmethod
    def get_current(cls) -> Dict[str, Any]:
        return cls._current.copy()


# 请求日志中间件
class RequestLoggingMiddleware:
    """请求日志中间件

    记录每个请求的详细信息。
    """

    def __init__(self, app=None, logger_name: str = "http"):
        self.logger = get_structured_logger(logger_name)
        if app:
            self.init_app(app)

    def init_app(self, app):
        from flask import request, g
        import time

        @app.before_request
        def log_request_start():
            g.request_start_time = time.time()

        @app.after_request
        def log_request_end(response):
            duration_ms = int((time.time() - g.get("request_start_time", time.time())) * 1000)

            # 不记录静态文件和健康检查
            if request.path.startswith("/static") or request.path == "/health":
                return response

            self.logger.info(
                f"{request.method} {request.path}",
                method=request.method,
                path=request.path,
                status=response.status_code,
                duration_ms=duration_ms,
                remote_addr=request.remote_addr,
                user_agent=request.user_agent.string[:100] if request.user_agent else None,
            )

            return response
