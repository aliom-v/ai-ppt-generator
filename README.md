# AI PPT 生成器

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Flask](https://img.shields.io/badge/Flask-3.0+-orange.svg)](https://flask.palletsprojects.com/)

基于大模型的自动化 PPT 生成工具，支持自定义 API 端点和模板。

## 功能特性

- 🤖 调用大模型自动生成 PPT 结构和内容
- 📊 支持多种页面类型（要点页、图文页、时间线、对比页等）
- 🎨 6 个精美亮色模板可选
- 🖼️ 支持 Unsplash 图片搜索和自动下载
- 🌐 Web 图形界面
- 🔧 支持自定义 API 端点（兼容 OpenAI 格式）

## 项目结构

```
ai-ppt-generator/
├── cli/                    # 命令行工具
│   └── main.py
├── core/                   # 核心功能
│   ├── ai_client.py        # AI 客户端
│   ├── prompt_builder.py   # 提示词构建
│   ├── ppt_plan.py         # PPT 数据模型
│   └── types.py            # 类型定义
├── ppt/                    # PPT 生成引擎
│   ├── unified_builder.py  # PPT 构建器
│   ├── template_manager.py # 模板管理
│   ├── create_templates.py # 模板创建
│   └── pptx_templates/     # 模板文件
├── web/                    # Web 应用
│   ├── app.py              # Flask 应用
│   ├── templates/          # HTML 模板
│   └── outputs/            # 生成的 PPT
├── utils/                  # 工具函数
│   ├── image_search.py     # 图片搜索
│   └── file_parser.py      # 文件解析
├── images/                 # 图片目录
│   ├── downloaded/         # 下载的图片
│   └── cache/              # 图片缓存
├── .env.example            # 环境变量示例
├── requirements.txt        # 依赖列表
└── start_web.py            # Web 启动入口
```

## 环境要求

- Python 3.8+
- Windows / macOS / Linux

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置环境变量

```bash
# 复制环境变量模板
cp .env.example .env

# 编辑 .env 文件，填入你的 API Key
```

必须配置：
- `AI_API_KEY`: AI 模型的 API 密钥

可选配置：
- `AI_API_BASE`: API 端点（默认：https://api.openai.com/v1）
- `AI_MODEL_NAME`: 模型名称（默认：gpt-4o-mini）
- `UNSPLASH_ACCESS_KEY`: Unsplash API Key（用于自动配图）

### 3. 启动服务

```bash
python start_web.py
```

### 4. 访问

打开浏览器访问：http://localhost:5000

## 使用方式

### Web 界面（推荐）

1. 访问 http://localhost:5000
2. 填写 PPT 主题和页数
3. 选择模板样式
4. 点击生成，等待完成后下载

### 命令行

```bash
python cli/main.py
```

按提示输入主题、页数等信息，系统会自动生成 PPT。

## 页面类型

支持以下页面类型：

| 类型 | 说明 |
|------|------|
| bullets | 要点页，包含标题和要点列表 |
| image_with_text | 图文混排页，左侧图片 + 右侧文字 |
| two_column | 双栏布局页 |
| timeline | 时间线页，展示流程或历程 |
| comparison | 对比页，左右对比展示 |
| quote | 引用页，展示金句或观点 |
| ending | 结束页 |

## 模板样式

提供 6 个亮色风格模板：

1. 经典蓝 - 专业商务风格
2. 清新绿 - 自然环保风格
3. 活力橙 - 创意活泼风格
4. 优雅紫 - 高端优雅风格
5. 天空蓝 - 清新明亮风格
6. 薄荷绿 - 清爽简约风格

## 图片搜索功能

### 配置 Unsplash API

1. 访问 https://unsplash.com/developers 注册
2. 创建应用获取 Access Key
3. 在 `.env` 文件或 Web 界面中配置

### 使用方式

在 Web 界面勾选"自动搜索下载图片"选项，系统会根据内容自动搜索并下载配图。

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可

[MIT License](LICENSE)

## Star History

如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！
