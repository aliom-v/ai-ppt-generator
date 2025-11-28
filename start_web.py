"""启动 Web 界面"""
import os
import sys

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

# 检查配置
api_key = os.getenv('AI_API_KEY')
if not api_key:
    print("=" * 60)
    print("⚠️  警告：未配置 AI_API_KEY")
    print("=" * 60)
    print("\n请在 .env 文件中配置 AI API Key")
    print("参考 .env.example 文件")
    print("\n继续启动服务器...\n")

# 启动 Web 应用
from web.app import app

if __name__ == '__main__':
    print("=" * 60)
    print("🚀 AI PPT 生成器 - Web 界面")
    print("=" * 60)
    print("\n✨ 功能特性：")
    print("  • AI 自动生成 PPT 结构")
    print("  • 多种页面类型支持")
    print("  • 自动搜索下载图片")
    print("  • 实时预览内容结构")
    print("\n🌐 访问地址：")
    print("  http://localhost:5000")
    print("  http://127.0.0.1:5000")
    print("\n💡 提示：")
    print("  • 按 Ctrl+C 停止服务器")
    print("  • 生成的 PPT 保存在 web/outputs/ 目录")
    print("  • 下载的图片保存在 images/downloaded/ 目录")
    print("\n" + "=" * 60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
