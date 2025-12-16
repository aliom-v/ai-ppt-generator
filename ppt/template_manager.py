"""PPT 模板管理器"""
import os
from typing import List, Dict, Optional
from pathlib import Path


class TemplateManager:
    """模板管理器 - 管理和提供 PPT 模板"""
    
    def __init__(self, templates_dir: str = "ppt/pptx_templates"):
        """初始化模板管理器
        
        Args:
            templates_dir: 模板文件夹路径
        """
        self.templates_dir = templates_dir
        self._ensure_templates_dir()
    
    def _ensure_templates_dir(self):
        """确保模板目录存在"""
        Path(self.templates_dir).mkdir(parents=True, exist_ok=True)
    
    def list_templates(self) -> List[Dict[str, str]]:
        """列出所有可用模板
        
        Returns:
            模板列表，每个模板包含 id, name, path, category, description
        """
        templates = []
        
        if not os.path.exists(self.templates_dir):
            return templates
        
        # 扫描模板文件夹
        for root, dirs, files in os.walk(self.templates_dir):
            for file in files:
                if file.endswith('.pptx') and not file.startswith('~'):
                    template_path = os.path.join(root, file)
                    template_info = self._get_template_info(template_path)
                    templates.append(template_info)
        
        return templates
    
    # 默认模板ID
    DEFAULT_TEMPLATE = "mckinsey_consulting"
    
    # 模板配置映射
    TEMPLATE_CONFIG = {
        "academic_latex": {
            "name": "学术答辩",
            "category": "学术",
            "description": "极简黑白，学术风"
        },
        "yc_pitch_deck": {
            "name": "融资路演",
            "category": "商务",
            "description": "大字大数，路演风"
        },
        "swiss_style": {
            "name": "瑞士设计",
            "category": "设计",
            "description": "蓝橙撞色，设计风"
        },
        "handdrawn_sketch": {
            "name": "手绘笔记",
            "category": "教育",
            "description": "手写涂鸦，笔记风"
        },
        "mckinsey_consulting": {
            "name": "咨询报告",
            "category": "商务",
            "description": "左文右图，咨询风"
        },
        "magazine_editorial": {
            "name": "杂志排版",
            "category": "创意",
            "description": "大图衬线，杂志风"
        },
        # 新增模板
        "dark_theme": {
            "name": "暗色主题",
            "category": "现代",
            "description": "深色背景，夜间模式"
        },
        "minimalist": {
            "name": "极简风格",
            "category": "设计",
            "description": "纯白简约，无装饰"
        },
        "chinese_style": {
            "name": "中国风",
            "category": "文化",
            "description": "红金配色，传统美学"
        },
        "gradient_blue": {
            "name": "渐变蓝",
            "category": "现代",
            "description": "蓝色渐变，科技感"
        },
        "tech_modern": {
            "name": "科技现代",
            "category": "科技",
            "description": "深色网格，未来感"
        },
        "warm_sunset": {
            "name": "暖色夕阳",
            "category": "创意",
            "description": "橙色渐变，温暖活力"
        }
    }
    
    def _get_template_info(self, template_path: str) -> Dict[str, str]:
        """获取模板信息"""
        filename = os.path.basename(template_path)
        name = os.path.splitext(filename)[0]
        
        # 从配置获取模板信息
        config = self.TEMPLATE_CONFIG.get(name, {})
        
        return {
            "id": name,
            "name": config.get("name", self._format_name(name)),
            "path": template_path,
            "category": config.get("category", "其他"),
            "description": config.get("description", "通用 PPT 模板"),
            "preview": self._get_preview_path(name)
        }
    
    def _format_name(self, name: str) -> str:
        """格式化模板名称"""
        return name.replace('_', ' ').replace('-', ' ').title()
    
    def _get_preview_path(self, template_id: str) -> str:
        """获取模板预览图路径"""
        preview_dir = os.path.join(self.templates_dir, "previews")
        preview_path = os.path.join(preview_dir, f"{template_id}.png")
        
        if os.path.exists(preview_path):
            return preview_path
        return ""
    
    def get_template(self, template_id: str) -> Optional[str]:
        """获取指定模板的路径
        
        Args:
            template_id: 模板 ID
            
        Returns:
            模板文件路径，如果不存在则返回 None
        """
        templates = self.list_templates()
        for template in templates:
            if template["id"] == template_id:
                return template["path"]
        return None
    
    def get_templates_by_category(self, category: str) -> List[Dict[str, str]]:
        """获取指定分类的模板
        
        Args:
            category: 分类名称
            
        Returns:
            模板列表
        """
        all_templates = self.list_templates()
        return [t for t in all_templates if t["category"] == category]
    
    def get_default_template(self) -> Optional[str]:
        """获取默认模板
        
        Returns:
            默认模板路径，如果没有则返回 None
        """
        # 优先返回配置的默认模板
        default_path = self.get_template(self.DEFAULT_TEMPLATE)
        if default_path:
            return default_path
        # 否则返回第一个模板
        templates = self.list_templates()
        if templates:
            return templates[0]["path"]
        return None
    
    def get_default_template_id(self) -> str:
        """获取默认模板ID"""
        return self.DEFAULT_TEMPLATE


# 全局模板管理器实例
template_manager = TemplateManager()


# 便捷函数
def list_templates() -> List[Dict[str, str]]:
    """列出所有模板"""
    return template_manager.list_templates()


def get_template(template_id: str) -> Optional[str]:
    """获取模板路径"""
    return template_manager.get_template(template_id)


def get_default_template() -> Optional[str]:
    """获取默认模板"""
    return template_manager.get_default_template()


if __name__ == "__main__":
    # 测试模板管理器
    print("=" * 60)
    print("PPT 模板管理器测试")
    print("=" * 60)
    
    manager = TemplateManager()
    templates = manager.list_templates()
    
    if templates:
        print(f"\n找到 {len(templates)} 个模板：\n")
        for template in templates:
            print(f"📄 {template['name']}")
            print(f"   分类：{template['category']}")
            print(f"   描述：{template['description']}")
            print(f"   路径：{template['path']}")
            print()
    else:
        print("\n⚠️  未找到模板文件")
        print(f"请将 .pptx 模板文件放到：{manager.templates_dir}")
