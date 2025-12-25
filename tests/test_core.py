"""核心模块测试"""
import json
import pytest

from core.ppt_plan import PptPlan, Slide, ppt_plan_from_dict, ppt_plan_to_dict
from core.ai_common import clean_json_response, calculate_batches
from config.settings import AIConfig


class TestPptPlan:
    """PPT 规划数据模型测试"""

    def test_slide_creation(self):
        """测试 Slide 创建"""
        slide = Slide(
            title="测试标题",
            slide_type="bullets",
            bullets=["要点1", "要点2"]
        )
        assert slide.title == "测试标题"
        assert slide.slide_type == "bullets"
        assert len(slide.bullets) == 2

    def test_ppt_plan_from_dict(self, sample_ppt_plan):
        """测试从字典创建 PPT 规划"""
        plan = ppt_plan_from_dict(sample_ppt_plan)
        assert plan.title == "测试演示文稿"
        assert plan.subtitle == "AI 自动生成"
        assert len(plan.slides) == 4

    def test_ppt_plan_to_dict(self, sample_ppt_plan):
        """测试 PPT 规划转换为字典"""
        plan = ppt_plan_from_dict(sample_ppt_plan)
        result = ppt_plan_to_dict(plan)
        assert result["title"] == sample_ppt_plan["title"]
        assert len(result["slides"]) == len(sample_ppt_plan["slides"])


class TestAIClient:
    """AI 客户端测试"""

    def testclean_json_response_basic(self):
        """测试基本 JSON 清理"""
        content = '{"title": "测试"}'
        result = clean_json_response(content)
        assert result == '{"title": "测试"}'

    def testclean_json_response_with_markdown(self):
        """测试带 markdown 代码块的 JSON 清理"""
        content = '```json\n{"title": "测试"}\n```'
        result = clean_json_response(content)
        assert json.loads(result)["title"] == "测试"

    def testclean_json_response_with_chinese_quotes(self):
        """测试中文引号替换"""
        content = '{"title": "测试"}'
        result = clean_json_response(content)
        parsed = json.loads(result)
        assert parsed["title"] == "测试"

    def testclean_json_response_extract_json(self):
        """测试提取 JSON 部分"""
        content = 'Here is the JSON: {"title": "测试"} end of content'
        result = clean_json_response(content)
        parsed = json.loads(result)
        assert parsed["title"] == "测试"

    def testcalculate_batches_small(self):
        """测试小页数不分批"""
        assert calculate_batches(10) == [10]
        assert calculate_batches(35) == [35]

    def testcalculate_batches_medium(self):
        """测试中等页数分批"""
        batches = calculate_batches(50)
        assert len(batches) == 2
        assert sum(batches) == 50

    def testcalculate_batches_large(self):
        """测试大页数分批"""
        batches = calculate_batches(100)
        assert len(batches) == 3
        assert sum(batches) == 100

    def testcalculate_batches_max(self):
        """测试最大页数分批"""
        batches = calculate_batches(200)
        assert len(batches) == 4
        assert sum(batches) == 200


class TestAIConfig:
    """AI 配置测试"""

    def test_config_creation(self):
        """测试配置创建"""
        config = AIConfig(
            api_key="test-key",
            api_base_url="https://api.openai.com/v1",
            model_name="gpt-4o-mini"
        )
        assert config.api_key == "test-key"
        assert config.model_name == "gpt-4o-mini"

    def test_config_normalize_url(self):
        """测试 URL 规范化"""
        config = AIConfig(
            api_key="test-key",
            api_base_url="https://api.openai.com/v1/",
        )
        assert config.api_base_url == "https://api.openai.com/v1"

    def test_config_validate_missing_key(self):
        """测试缺少 API Key 验证"""
        config = AIConfig(api_key="", api_base_url="https://api.openai.com/v1")
        with pytest.raises(ValueError, match="API_KEY"):
            config.validate()

    def test_config_from_dict(self):
        """测试从字典创建配置"""
        data = {
            "api_key": "test-key",
            "api_base": "https://api.openai.com/v1",
            "model_name": "gpt-4"
        }
        config = AIConfig.from_dict(data)
        assert config.api_key == "test-key"
        assert config.model_name == "gpt-4"
