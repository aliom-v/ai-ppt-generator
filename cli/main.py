"""命令行入口模块"""
import sys
import os

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.ai_client import generate_ppt_plan
from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from utils.preview import open_with_default_app
from config.settings import settings


def main():
    """主函数：命令行交互入口"""
    print("=" * 50)
    print("AI PPT 生成器")
    print("=" * 50)
    print()
    
    try:
        # 1. 收集用户输入
        topic = input("请输入 PPT 主题: ").strip()
        if not topic:
            print("错误：主题不能为空")
            return
        
        audience = input("请输入目标受众（如：技术团队、管理层等）: ").strip()
        if not audience:
            audience = "通用受众"
        
        page_count_str = input("请输入内容页数量（不含封面，默认 5）: ").strip()
        page_count = int(page_count_str) if page_count_str else 5
        
        template_path_input = input(f"请输入模板路径（直接回车使用默认，默认：{settings.default_template}）: ").strip()
        template_path = template_path_input if template_path_input else None
        
        output_path_input = input(f"请输入输出文件名（默认：{settings.default_output}）: ").strip()
        output_path = output_path_input if output_path_input else settings.default_output
        
        print()
        print("正在调用 AI 生成 PPT 结构...")
        
        # 2. 调用 AI 生成 plan
        plan_dict = generate_ppt_plan(topic, audience, page_count)
        
        # 3. 转换为 PptPlan 对象
        plan = ppt_plan_from_dict(plan_dict)
        
        print(f"✓ 已生成 PPT 结构：{plan.title}")
        print(f"  - 副标题：{plan.subtitle}")
        print(f"  - 内容页数：{len(plan.slides)}")
        
        # 显示页面类型统计
        type_counts = {}
        for slide in plan.slides:
            slide_type = slide.slide_type
            type_counts[slide_type] = type_counts.get(slide_type, 0) + 1
        
        print(f"  - 页面类型：", end="")
        print(", ".join([f"{t}({c})" for t, c in type_counts.items()]))
        print()
        
        # 4. 询问是否需要添加图片
        image_slides = [s for s in plan.slides if s.slide_type == "image_with_text"]
        auto_download = False
        
        if image_slides:
            print(f"\n检测到 {len(image_slides)} 个图文页：")
            for i, slide in enumerate(image_slides, 1):
                print(f"  {i}. {slide.title}")
                if slide.image_keyword:
                    print(f"     建议图片：{slide.image_keyword}")
            
            print("\n图片选项：")
            print("  1. 自动联网搜索下载（需要配置 API Key）")
            print("  2. 手动输入本地图片路径")
            print("  3. 使用占位符（稍后手动添加）")
            
            choice = input("\n请选择 (1/2/3，直接回车选择3): ").strip()
            
            if choice == '1':
                # 自动下载
                api_key = os.getenv("UNSPLASH_ACCESS_KEY")
                if not api_key:
                    print("\n提示：未设置 UNSPLASH_ACCESS_KEY 环境变量")
                    print("获取免费 API Key: https://unsplash.com/developers")
                    print("设置方法：在 .env 文件中添加 UNSPLASH_ACCESS_KEY=your_key")
                    print("\n将使用占位符模式")
                else:
                    auto_download = True
                    print("✓ 将自动搜索并下载图片")
            
            elif choice == '2':
                # 手动输入路径
                for i, slide in enumerate(image_slides, 1):
                    img_path = input(f"  第 {i} 页图片路径（直接回车跳过）: ").strip()
                    if img_path and os.path.exists(img_path):
                        slide.image_path = img_path
                        print(f"    ✓ 已设置图片：{img_path}")
                    elif img_path:
                        print(f"    ✗ 文件不存在，将使用占位符")
            
            print()
        
        # 5. 生成 PPTX 文件
        print("正在生成 PPT 文件...")
        build_ppt_from_plan(plan, template_path, output_path, auto_download_images=auto_download)
        
        abs_output_path = os.path.abspath(output_path)
        print(f"✓ PPT 已生成：{abs_output_path}")
        print()
        
        # 6. 询问是否预览
        preview = input("是否立即打开预览？(y/n): ").strip().lower()
        if preview == 'y':
            print("正在打开 PPT...")
            open_with_default_app(output_path)
            print("✓ 已打开")
        
        print()
        print("完成！")
        
    except KeyboardInterrupt:
        print("\n\n操作已取消")
    except ValueError as e:
        print(f"\n错误：{e}")
    except Exception as e:
        print(f"\n发生错误：{e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
