"""增强的健康检查模块"""
import os
import sys
import platform
import psutil
from datetime import datetime
from typing import Dict, Any
from pathlib import Path

from utils.logger import get_logger

logger = get_logger("health")

# 启动时间
_start_time = datetime.now()


def get_system_info() -> Dict[str, Any]:
    """获取系统信息"""
    try:
        cpu_percent = psutil.cpu_percent(interval=0.1)
        memory = psutil.virtual_memory()
        disk = psutil.disk_usage('/')
    except Exception:
        cpu_percent = 0
        memory = None
        disk = None

    return {
        "platform": platform.system(),
        "platform_version": platform.version(),
        "python_version": sys.version.split()[0],
        "cpu_percent": cpu_percent,
        "memory": {
            "total_gb": round(memory.total / (1024**3), 2) if memory else 0,
            "used_gb": round(memory.used / (1024**3), 2) if memory else 0,
            "percent": memory.percent if memory else 0,
        } if memory else None,
        "disk": {
            "total_gb": round(disk.total / (1024**3), 2) if disk else 0,
            "used_gb": round(disk.used / (1024**3), 2) if disk else 0,
            "percent": round(disk.percent, 1) if disk else 0,
        } if disk else None,
    }


def get_service_status() -> Dict[str, Any]:
    """获取服务状态"""
    uptime = datetime.now() - _start_time

    # 检查关键目录
    dirs_status = {}
    for dir_name in ["web/outputs", "web/uploads", "cache", "images/downloaded"]:
        path = Path(dir_name)
        dirs_status[dir_name] = {
            "exists": path.exists(),
            "writable": os.access(path, os.W_OK) if path.exists() else False,
        }

    # 检查配置
    config_status = {
        "ai_api_key": bool(os.getenv("AI_API_KEY")),
        "unsplash_key": bool(os.getenv("UNSPLASH_ACCESS_KEY")),
    }

    return {
        "uptime_seconds": int(uptime.total_seconds()),
        "uptime_human": str(uptime).split('.')[0],
        "start_time": _start_time.isoformat(),
        "directories": dirs_status,
        "config": config_status,
    }


def check_health() -> Dict[str, Any]:
    """完整健康检查"""
    from utils.metrics import get_metrics_collector

    checks = {
        "api": True,
        "cache_dir": Path("cache").exists(),
        "output_dir": Path("web/outputs").exists(),
    }

    all_healthy = all(checks.values())

    result = {
        "status": "healthy" if all_healthy else "degraded",
        "timestamp": datetime.now().isoformat(),
        "checks": checks,
    }

    return result


def get_detailed_health() -> Dict[str, Any]:
    """详细健康检查（包含系统信息和指标）"""
    from utils.metrics import get_metrics_collector

    health = check_health()
    metrics = get_metrics_collector().get_stats()

    return {
        **health,
        "service": get_service_status(),
        "system": get_system_info(),
        "metrics": metrics,
    }
