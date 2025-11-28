"""文件预览工具模块"""
import os
import sys
import subprocess


def open_with_default_app(path: str) -> None:
    """使用系统默认程序打开文件
    
    Args:
        path: 文件路径
        
    Raises:
        FileNotFoundError: 当文件不存在时
        RuntimeError: 当无法打开文件时
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"文件不存在: {path}")
    
    abs_path = os.path.abspath(path)
    
    try:
        if sys.platform == "win32":
            # Windows
            os.startfile(abs_path)
        elif sys.platform == "darwin":
            # macOS
            subprocess.run(["open", abs_path], check=True)
        else:
            # Linux
            subprocess.run(["xdg-open", abs_path], check=True)
    except Exception as e:
        raise RuntimeError(f"无法打开文件: {e}")
