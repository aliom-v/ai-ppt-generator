.PHONY: help install dev clean

help:
	@echo "AI PPT 生成器 - 可用命令："
	@echo ""
	@echo "  make install  - 安装依赖"
	@echo "  make dev      - 启动开发服务器"
	@echo "  make clean    - 清理临时文件"
	@echo ""

install:
	pip install -r requirements.txt

dev:
	python start_web.py

clean:
	find . -type d -name "__pycache__" -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type f -name "*.log" -delete
