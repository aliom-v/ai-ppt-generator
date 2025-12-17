"""统一日志模块 - 支持请求 ID 追踪和文件日志"""
import logging
import sys
import os
from pathlib import Path
from logging.handlers import RotatingFileHandler
from typing import Optional


class RequestIdFilter(logging.Filter):
    """日志过滤器：自动添加请求 ID"""

    def filter(self, record):
        from utils.request_context import get_request_id
        record.request_id = get_request_id() or '-'
        return True


def setup_logger(
    name: str = "ai_ppt",
    level: int = logging.INFO,
    log_file: Optional[str] = None,
    max_bytes: int = 10 * 1024 * 1024,  # 10MB
    backup_count: int = 5
) -> logging.Logger:
    """配置并返回 logger

    Args:
        name: logger 名称
        level: 日志级别
        log_file: 日志文件路径（可选）
        max_bytes: 单个日志文件最大字节数
        backup_count: 保留的日志文件数量

    Returns:
        配置好的 logger 实例
    """
    logger = logging.getLogger(name)

    # 避免重复添加 handler
    if logger.handlers:
        return logger

    logger.setLevel(level)

    # 添加请求 ID 过滤器
    logger.addFilter(RequestIdFilter())

    # 控制台输出格式（带请求 ID）
    console_formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-5s | [%(request_id)s] %(message)s",
        datefmt="%H:%M:%S"
    )

    # 控制台 handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    # 文件 handler（如果指定了日志文件）
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)

        # 文件格式更详细
        file_formatter = logging.Formatter(
            fmt="%(asctime)s | %(levelname)-5s | [%(request_id)s] %(name)s:%(lineno)d | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

    return logger


# 根据环境变量配置日志
_log_level = logging.DEBUG if os.getenv("DEBUG", "").lower() == "true" else logging.INFO
_log_file = os.getenv("LOG_FILE", "logs/app.log") if os.getenv("LOG_TO_FILE", "").lower() == "true" else None

# 全局 logger 实例
logger = setup_logger(level=_log_level, log_file=_log_file)


def get_logger(name: str = None) -> logging.Logger:
    """获取 logger 实例

    Args:
        name: 子模块名称，如 "web"

    Returns:
        logger 实例
    """
    if name:
        child_logger = logging.getLogger(f"ai_ppt.{name}")
        # 继承父 logger 的配置
        if not child_logger.handlers:
            child_logger.parent = logger
        return child_logger
    return logger
