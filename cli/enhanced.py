#!/usr/bin/env python3
"""AI PPT Generator - 命令行工具

增强版命令行工具，支持多种操作模式和丰富的选项。

用法:
    # 交互模式
    python -m cli.enhanced

    # 直接生成
    python -m cli.enhanced generate --topic "人工智能" --pages 10

    # 从文件生成
    python -m cli.enhanced generate --file content.txt --topic "项目总结"

    # 批量生成
    python -m cli.enhanced batch --input topics.json --output ./outputs

    # 列出模板
    python -m cli.enhanced templates

    # 查看历史
    python -m cli.enhanced history --limit 10
"""
import argparse
import json
import os
import sys
from pathlib import Path
from typing import List, Optional

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from config.settings import settings, AIConfig
from core.ai_client import generate_ppt_plan, test_api_connection
from core.ppt_plan import ppt_plan_from_dict
from ppt.unified_builder import build_ppt_from_plan
from ppt.template_manager import template_manager
from utils.preview import open_with_default_app


class Colors:
    """终端颜色"""
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    RESET = "\033[0m"
    BOLD = "\033[1m"


def colored(text: str, color: str) -> str:
    """添加颜色"""
    if not sys.stdout.isatty():
        return text
    return f"{color}{text}{Colors.RESET}"


def print_success(msg: str):
    print(colored(f"✓ {msg}", Colors.GREEN))


def print_error(msg: str):
    print(colored(f"✗ {msg}", Colors.RED))


def print_info(msg: str):
    print(colored(f"ℹ {msg}", Colors.BLUE))


def print_warning(msg: str):
    print(colored(f"⚠ {msg}", Colors.YELLOW))


def create_parser() -> argparse.ArgumentParser:
    """创建参数解析器"""
    parser = argparse.ArgumentParser(
        prog="ai-ppt",
        description="AI 驱动的 PPT 自动生成工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s generate --topic "人工智能简介" --pages 10
  %(prog)s generate --topic "项目报告" --file notes.txt --auto-images
  %(prog)s batch --input topics.json --output ./outputs
  %(prog)s templates --list
  %(prog)s history --stats
        """,
    )

    parser.add_argument(
        "--version", "-v",
        action="version",
        version="%(prog)s 1.0.0",
    )

    parser.add_argument(
        "--verbose",
        action="store_true",
        help="显示详细输出",
    )

    subparsers = parser.add_subparsers(dest="command", help="可用命令")

    # generate 命令
    gen_parser = subparsers.add_parser("generate", aliases=["gen", "g"], help="生成 PPT")
    gen_parser.add_argument("--topic", "-t", required=True, help="PPT 主题")
    gen_parser.add_argument("--audience", "-a", default="通用受众", help="目标受众")
    gen_parser.add_argument("--pages", "-p", type=int, default=5, help="页数 (1-100)")
    gen_parser.add_argument("--output", "-o", help="输出文件路径")
    gen_parser.add_argument("--template", help="模板 ID 或路径")
    gen_parser.add_argument("--file", "-f", help="从文件读取内容")
    gen_parser.add_argument("--auto-images", action="store_true", help="自动下载图片")
    gen_parser.add_argument("--preview", action="store_true", help="生成后打开预览")
    gen_parser.add_argument("--api-key", help="AI API Key (或设置环境变量 AI_API_KEY)")
    gen_parser.add_argument("--api-base", help="API Base URL")
    gen_parser.add_argument("--model", help="模型名称")

    # batch 命令
    batch_parser = subparsers.add_parser("batch", aliases=["b"], help="批量生成")
    batch_parser.add_argument("--input", "-i", required=True, help="输入 JSON 文件")
    batch_parser.add_argument("--output", "-o", required=True, help="输出目录")
    batch_parser.add_argument("--template", help="使用的模板")
    batch_parser.add_argument("--parallel", type=int, default=1, help="并行数")

    # templates 命令
    tpl_parser = subparsers.add_parser("templates", aliases=["tpl"], help="模板管理")
    tpl_parser.add_argument("--list", "-l", action="store_true", help="列出所有模板")
    tpl_parser.add_argument("--info", help="查看模板详情")

    # history 命令
    hist_parser = subparsers.add_parser("history", aliases=["hist", "h"], help="查看历史")
    hist_parser.add_argument("--limit", "-l", type=int, default=10, help="显示数量")
    hist_parser.add_argument("--stats", "-s", action="store_true", help="显示统计")
    hist_parser.add_argument("--search", help="搜索关键词")

    # test 命令
    test_parser = subparsers.add_parser("test", help="测试 API 连接")
    test_parser.add_argument("--api-key", help="API Key")
    test_parser.add_argument("--api-base", help="API Base URL")

    # interactive 命令
    subparsers.add_parser("interactive", aliases=["i"], help="交互模式")

    return parser


def cmd_generate(args):
    """生成 PPT"""
    print_info(f"开始生成 PPT: {args.topic}")

    # 获取 API 配置
    api_key = args.api_key or os.getenv("AI_API_KEY")
    if not api_key:
        print_error("请设置 API Key (--api-key 或环境变量 AI_API_KEY)")
        return 1

    ai_config = AIConfig(
        api_key=api_key,
        api_base_url=args.api_base or os.getenv("AI_API_BASE", "https://api.openai.com/v1"),
        model_name=args.model or os.getenv("AI_MODEL_NAME", "gpt-4o-mini"),
    )

    # 读取文件内容
    description = ""
    if args.file:
        try:
            from utils.file_parser import parse_file
            description = parse_file(args.file)
            print_info(f"已读取文件: {args.file} ({len(description)} 字)")
        except Exception as e:
            print_error(f"读取文件失败: {e}")
            return 1

    # 生成 PPT 结构
    print_info("正在调用 AI 生成结构...")
    try:
        plan_dict = generate_ppt_plan(
            args.topic,
            args.audience,
            args.pages,
            description,
            config=ai_config,
        )
        plan = ppt_plan_from_dict(plan_dict)
        print_success(f"已生成结构: {plan.title} ({len(plan.slides)} 页)")
    except Exception as e:
        print_error(f"生成结构失败: {e}")
        return 1

    # 确定输出路径
    output_path = args.output or f"{args.topic[:30].replace(' ', '_')}.pptx"

    # 获取模板
    template_path = None
    if args.template:
        template_path = template_manager.get_template(args.template)

    # 生成 PPT
    print_info("正在生成 PPT 文件...")
    try:
        build_ppt_from_plan(
            plan,
            template_path,
            output_path,
            auto_download_images=args.auto_images,
        )
        print_success(f"PPT 已生成: {os.path.abspath(output_path)}")
    except Exception as e:
        print_error(f"生成 PPT 失败: {e}")
        return 1

    # 预览
    if args.preview:
        print_info("正在打开预览...")
        open_with_default_app(output_path)

    return 0


def cmd_batch(args):
    """批量生成"""
    print_info(f"批量生成模式: {args.input}")

    # 读取输入文件
    try:
        with open(args.input, "r", encoding="utf-8") as f:
            tasks = json.load(f)
    except Exception as e:
        print_error(f"读取输入文件失败: {e}")
        return 1

    if not isinstance(tasks, list):
        print_error("输入文件必须是 JSON 数组")
        return 1

    # 创建输出目录
    Path(args.output).mkdir(parents=True, exist_ok=True)

    # 获取 API 配置
    api_key = os.getenv("AI_API_KEY")
    if not api_key:
        print_error("请设置环境变量 AI_API_KEY")
        return 1

    ai_config = AIConfig(api_key=api_key)

    # 处理每个任务
    success_count = 0
    for i, task in enumerate(tasks, 1):
        topic = task.get("topic", "")
        if not topic:
            print_warning(f"任务 {i}: 跳过（无主题）")
            continue

        print_info(f"[{i}/{len(tasks)}] 生成: {topic}")

        try:
            plan_dict = generate_ppt_plan(
                topic,
                task.get("audience", "通用受众"),
                task.get("pages", 5),
                task.get("description", ""),
                config=ai_config,
            )
            plan = ppt_plan_from_dict(plan_dict)

            output_file = Path(args.output) / f"{topic[:30].replace(' ', '_')}_{i}.pptx"
            template_path = template_manager.get_template(args.template) if args.template else None

            build_ppt_from_plan(plan, template_path, str(output_file))
            print_success(f"  -> {output_file.name}")
            success_count += 1

        except Exception as e:
            print_error(f"  失败: {e}")

    print_info(f"完成: {success_count}/{len(tasks)} 成功")
    return 0 if success_count == len(tasks) else 1


def cmd_templates(args):
    """模板管理"""
    templates = template_manager.list_templates()

    if args.info:
        for tpl in templates:
            if tpl["id"] == args.info or tpl["name"] == args.info:
                print(f"\n{colored('模板详情', Colors.BOLD)}")
                print(f"  ID: {tpl['id']}")
                print(f"  名称: {tpl['name']}")
                print(f"  路径: {tpl['path']}")
                return 0
        print_error(f"模板不存在: {args.info}")
        return 1

    # 列出模板
    print(f"\n{colored('可用模板', Colors.BOLD)} ({len(templates)} 个)\n")
    for tpl in templates:
        print(f"  {colored(tpl['id'], Colors.CYAN):20} {tpl['name']}")

    return 0


def cmd_history(args):
    """查看历史"""
    try:
        from utils.history import get_history
        history = get_history()

        if args.stats:
            stats = history.get_stats()
            print(f"\n{colored('生成统计', Colors.BOLD)}\n")
            print(f"  总生成数: {stats['total_generations']}")
            print(f"  成功数: {stats['successful']}")
            print(f"  失败数: {stats['failed']}")
            print(f"  成功率: {stats['success_rate']}%")
            print(f"  今日生成: {stats['today_count']}")
            print(f"  平均耗时: {stats['avg_duration_ms']:.0f}ms")

            if stats.get("popular_topics"):
                print(f"\n  热门主题:")
                for item in stats["popular_topics"]:
                    print(f"    - {item['topic']} ({item['count']}次)")
            return 0

        # 搜索或列出
        if args.search:
            records = history.search(keyword=args.search, limit=args.limit)
        else:
            records = history.get_recent(limit=args.limit)

        if not records:
            print_info("暂无历史记录")
            return 0

        print(f"\n{colored('生成历史', Colors.BOLD)} ({len(records)} 条)\n")
        for rec in records:
            status_color = Colors.GREEN if rec["status"] == "success" else Colors.RED
            status = colored("✓" if rec["status"] == "success" else "✗", status_color)
            print(f"  {status} [{rec['created_at']}] {rec['topic'][:30]}")
            print(f"      {rec['slide_count']}页 | {rec['duration_ms']}ms | {rec['filename']}")

    except Exception as e:
        print_error(f"获取历史失败: {e}")
        return 1

    return 0


def cmd_test(args):
    """测试 API 连接"""
    api_key = args.api_key or os.getenv("AI_API_KEY")
    if not api_key:
        print_error("请提供 API Key")
        return 1

    ai_config = AIConfig(
        api_key=api_key,
        api_base_url=args.api_base or os.getenv("AI_API_BASE", "https://api.openai.com/v1"),
    )

    print_info("测试 API 连接...")
    result = test_api_connection(ai_config)

    if result["success"]:
        print_success(result["message"])
        if result.get("model"):
            print_info(f"模型: {result['model']}")
    else:
        print_error(result["message"])
        return 1

    return 0


def cmd_interactive(args):
    """交互模式"""
    print(f"\n{colored('AI PPT Generator', Colors.BOLD)} - 交互模式")
    print(f"{colored('=' * 40, Colors.CYAN)}\n")

    # 检查 API Key
    api_key = os.getenv("AI_API_KEY")
    if not api_key:
        api_key = input("请输入 AI API Key: ").strip()
        if not api_key:
            print_error("API Key 不能为空")
            return 1

    # 收集输入
    topic = input("PPT 主题: ").strip()
    if not topic:
        print_error("主题不能为空")
        return 1

    audience = input("目标受众 (默认: 通用受众): ").strip() or "通用受众"

    pages_str = input("页数 (默认: 5): ").strip()
    pages = int(pages_str) if pages_str else 5

    output = input("输出文件名 (默认: output.pptx): ").strip() or "output.pptx"

    auto_images = input("自动下载图片? (y/n, 默认: n): ").strip().lower() == "y"

    # 模拟 args
    class Args:
        pass

    gen_args = Args()
    gen_args.topic = topic
    gen_args.audience = audience
    gen_args.pages = pages
    gen_args.output = output
    gen_args.template = None
    gen_args.file = None
    gen_args.auto_images = auto_images
    gen_args.preview = True
    gen_args.api_key = api_key
    gen_args.api_base = None
    gen_args.model = None

    return cmd_generate(gen_args)


def main():
    """主入口"""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        # 默认进入交互模式
        return cmd_interactive(args)

    commands = {
        "generate": cmd_generate,
        "gen": cmd_generate,
        "g": cmd_generate,
        "batch": cmd_batch,
        "b": cmd_batch,
        "templates": cmd_templates,
        "tpl": cmd_templates,
        "history": cmd_history,
        "hist": cmd_history,
        "h": cmd_history,
        "test": cmd_test,
        "interactive": cmd_interactive,
        "i": cmd_interactive,
    }

    handler = commands.get(args.command)
    if handler:
        try:
            return handler(args)
        except KeyboardInterrupt:
            print("\n\n操作已取消")
            return 130
        except Exception as e:
            print_error(f"发生错误: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
            return 1

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
