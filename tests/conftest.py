"""测试配置和共享 fixtures"""
import os
import sys
import tempfile
from pathlib import Path

import pytest

# 添加项目根目录到路径
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture
def temp_dir():
    """创建临时目录"""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def sample_ppt_plan():
    """示例 PPT 规划数据"""
    return {
        "title": "测试演示文稿",
        "subtitle": "AI 自动生成",
        "slides": [
            {
                "title": "第一章：简介",
                "type": "bullets",
                "bullets": [
                    "要点一：这是第一个要点的详细说明",
                    "要点二：这是第二个要点的详细说明",
                    "要点三：这是第三个要点的详细说明",
                ]
            },
            {
                "title": "图文说明",
                "type": "image_with_text",
                "text": "这是一段图文说明的文字内容，用于测试图文混排页面的生成效果。",
                "image_keyword": "technology"
            },
            {
                "title": "时间线",
                "type": "timeline",
                "bullets": [
                    "阶段一：准备阶段",
                    "阶段二：执行阶段",
                    "阶段三：收尾阶段",
                ]
            },
            {
                "title": "感谢观看",
                "type": "ending",
                "subtitle": "欢迎提问"
            }
        ]
    }


@pytest.fixture
def mock_ai_config():
    """模拟 AI 配置"""
    from config.settings import AIConfig
    return AIConfig(
        api_key="test-api-key",
        api_base_url="https://api.openai.com/v1",
        model_name="gpt-4o-mini"
    )


@pytest.fixture
def mock_env_vars(monkeypatch):
    """设置测试环境变量"""
    monkeypatch.setenv("AI_API_KEY", "test-key")
    monkeypatch.setenv("AI_API_BASE", "https://api.openai.com/v1")
    monkeypatch.setenv("AI_MODEL_NAME", "gpt-4o-mini")
