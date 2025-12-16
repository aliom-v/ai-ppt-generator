"""统一日志模块"""
import logging
import sys
from pathlib import Path


def setup_logger(name: str = "ai_ppt", level: int = logging.INFO) -> logging.Logger:
    """配置并返回 logger
    
    Args:
        name: logger 名称
        level: 日志级别
        
    Returns:
        配置好的 logger 实例
    """
    logger = logging.getLogger(name)
    
    # 避免重复添加 handler
    if logger.handlers:
        return logger
    
    logger.setLevel(level)
    
    # 控制台输出
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    
    # 格式化：简洁风格
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S"
    )
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger


# 全局 logger 实例
logger = setup_logger()


def get_logger(name: str = None) -> logging.Logger:
    """获取 logger 实例
    
    Args:
        name: 子模块名称，如 "ai_ppt.web"
        
    Returns:
        logger 实例
    """
    if name:
        return logging.getLogger(f"ai_ppt.{name}")
    return logger
