"""Flask 蓝图模块"""

from .api import api_bp
from .tasks import tasks_bp
from .batch import batch_bp
from .ppt_edit import ppt_edit_bp
from .export import export_bp

__all__ = ['api_bp', 'tasks_bp', 'batch_bp', 'ppt_edit_bp', 'export_bp']
