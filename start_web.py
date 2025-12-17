"""启动 Web 界面"""
import os
import sys
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent))

# 加载环境变量
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# 检查依赖
try:
    import flask
except ImportError:
    print("=" * 60)
    print("❌ 缺少依赖")
    print("=" * 60)
    print("\n请先安装依赖：")
    print("  pip install flask")
    print("\n或安装所有依赖：")
    print("  pip install -r requirements.txt")
    sys.exit(1)

# 执行启动检查
from utils.startup import initialize_app

if not initialize_app():
    print("\n❌ 启动检查失败，请修复上述问题后重试")
    sys.exit(1)

# 导入 Flask 应用
from web.app import app, app_config

if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("🚀 AI PPT 生成器 - Web 界面")
    print("=" * 60)
    print("\n✨ 功能特性：")
    print("  • AI 自动生成 PPT 结构")
    print("  • 多种页面类型支持")
    print("  • 自动搜索下载图片")
    print("  • 实时预览内容结构")
    print("\n🌐 访问地址：")
    print(f"  http://localhost:{app_config.port}")
    print(f"  http://127.0.0.1:{app_config.port}")
    print("\n💡 提示：")
    print("  • 按 Ctrl+C 停止服务器")
    print("  • 生成的 PPT 保存在 web/outputs/ 目录")
    print("  • 下载的图片保存在 images/downloaded/ 目录")
    print("\n" + "=" * 60 + "\n")

    app.run(
        debug=app_config.debug,
        host=app_config.host,
        port=app_config.port
    )
