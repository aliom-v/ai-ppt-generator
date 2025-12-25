"""PPT 构建器测试"""
import os
import pytest

from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from ppt.styles import FONT_CN, FONT_EN, ColorTheme


class TestPptBuilder:
    """PPT 构建器测试"""

    def test_get_default_fonts(self):
        """测试获取默认字体"""
        assert isinstance(FONT_CN, str)
        assert isinstance(FONT_EN, str)
        assert len(FONT_CN) > 0
        assert len(FONT_EN) > 0

    def test_color_theme(self):
        """测试颜色主题"""
        assert ColorTheme.PRIMARY is not None
        assert ColorTheme.SECONDARY is not None
        assert ColorTheme.ACCENT is not None
        assert ColorTheme.TEXT_DARK is not None

    def test_build_ppt_basic(self, sample_ppt_plan, temp_dir):
        """测试基本 PPT 生成"""
        plan = ppt_plan_from_dict(sample_ppt_plan)
        output_path = os.path.join(temp_dir, "test_output.pptx")

        build_ppt_from_plan(plan, None, output_path, auto_download_images=False)

        assert os.path.exists(output_path)
        assert os.path.getsize(output_path) > 0

    def test_build_ppt_with_empty_slides(self, temp_dir):
        """测试空幻灯片处理"""
        plan_dict = {
            "title": "测试",
            "subtitle": "副标题",
            "slides": []
        }
        plan = ppt_plan_from_dict(plan_dict)
        output_path = os.path.join(temp_dir, "test_empty.pptx")

        build_ppt_from_plan(plan, None, output_path)

        assert os.path.exists(output_path)

    def test_build_ppt_all_slide_types(self, temp_dir):
        """测试所有幻灯片类型"""
        plan_dict = {
            "title": "完整测试",
            "subtitle": "所有类型",
            "slides": [
                {"title": "要点页", "type": "bullets", "bullets": ["要点1", "要点2", "要点3"]},
                {"title": "图文页", "type": "image_with_text", "text": "说明文字", "image_keyword": "test"},
                {"title": "双栏页", "type": "two_column", "bullets": ["左1", "左2", "右1", "右2"]},
                {"title": "时间线", "type": "timeline", "bullets": ["阶段1", "阶段2", "阶段3"]},
                {"title": "对比页", "type": "comparison", "bullets": ["A1", "A2", "B1", "B2"]},
                {"title": "引用页", "type": "quote", "text": "这是一段引用", "subtitle": "作者"},
                {"title": "结束页", "type": "ending", "subtitle": "谢谢"},
            ]
        }
        plan = ppt_plan_from_dict(plan_dict)
        output_path = os.path.join(temp_dir, "test_all_types.pptx")

        build_ppt_from_plan(plan, None, output_path)

        assert os.path.exists(output_path)
        assert os.path.getsize(output_path) > 0

    def test_build_ppt_long_content(self, temp_dir):
        """测试长内容处理"""
        long_bullet = "这是一段很长的要点内容，" * 10
        plan_dict = {
            "title": "长内容测试",
            "subtitle": "测试",
            "slides": [
                {
                    "title": "长要点",
                    "type": "bullets",
                    "bullets": [long_bullet, long_bullet, long_bullet]
                }
            ]
        }
        plan = ppt_plan_from_dict(plan_dict)
        output_path = os.path.join(temp_dir, "test_long.pptx")

        build_ppt_from_plan(plan, None, output_path)

        assert os.path.exists(output_path)

    def test_build_ppt_special_characters(self, temp_dir):
        """测试特殊字符处理"""
        plan_dict = {
            "title": "特殊字符 <>&\"' 测试",
            "subtitle": "包含特殊符号",
            "slides": [
                {
                    "title": "特殊字符测试",
                    "type": "bullets",
                    "bullets": [
                        "包含 <html> 标签",
                        "包含 & 符号",
                        "包含 \"引号\" 和 '单引号'",
                    ]
                }
            ]
        }
        plan = ppt_plan_from_dict(plan_dict)
        output_path = os.path.join(temp_dir, "test_special.pptx")

        build_ppt_from_plan(plan, None, output_path)

        assert os.path.exists(output_path)
