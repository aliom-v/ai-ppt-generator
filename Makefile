# AI PPT Generator Makefile
.PHONY: help install install-dev run test lint format clean build docker-build docker-run docker-stop docs

# Default target
.DEFAULT_GOAL := help

# Variables
PYTHON := python3
PIP := pip3
DOCKER_COMPOSE := docker-compose
DOCKER := docker
PROJECT_NAME := ai-ppt-generator

help: ## Show this help message
	@echo "AI PPT Generator - Available Commands:"
	@echo ""
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-20s\033[0m %s\n", $$1, $$2}'

# Installation
install: ## Install production dependencies
	$(PIP) install -r requirements.txt
	$(PIP) install -e .

install-dev: ## Install development dependencies
	$(PIP) install -r requirements.txt
	$(PIP) install -r requirements-dev.txt
	$(PIP) install -e ".[dev]"
	pre-commit install

install-pre-commit: ## Install pre-commit hooks
	pre-commit install
	pre-commit install --hook-type commit-msg

# Development
run: ## Run the web application
	$(PYTHON) start_web.py

run-worker: ## Run the background worker
	$(PYTHON) -m core.task_manager

dev: ## Run in development mode with hot reload
	uvicorn web.app:app --reload --host 0.0.0.0 --port 8000

# Testing
test: ## Run all tests
	pytest tests/ -v --cov=core --cov=ppt --cov=utils --cov=web --cov=cli --cov-report=html --cov-report=term

test-unit: ## Run unit tests only
	pytest tests/unit/ -v

test-integration: ## Run integration tests only
	pytest tests/integration/ -v

test-watch: ## Run tests in watch mode
	ptw --runner "python -m pytest tests/ -v"

# Code Quality
lint: ## Run all linting tools
	flake8 core ppt utils web cli tests
	mypy core ppt utils web cli tests
	bandit -r core ppt utils web cli
	safety check
	pip-audit

format: ## Format code with black and isort
	black core ppt utils web cli tests
	isort core ppt utils web cli tests

format-check: ## Check if code is properly formatted
	black --check core ppt utils web cli tests
	isort --check-only core ppt utils web cli tests

pre-commit: ## Run all pre-commit hooks
	pre-commit run --all-files

# Security
security-scan: ## Run security scanning
	bandit -r core ppt utils web cli -f json -o bandit-report.json
	safety check --json --output safety-report.json
	pip-audit --format json --output audit-report.json

security-fix: ## Fix common security issues
	safety check --fix
	pip-audit --fix

# Database
db-init: ## Initialize database
	$(PYTHON) -c "from utils.database import init_db; init_db()"

db-migrate: ## Run database migrations
	alembic upgrade head

db-migration: ## Create new database migration
	alembic revision --autogenerate -m "$(MSG)"

db-reset: ## Reset database
	alembic downgrade base
	alembic upgrade head

# Docker
docker-build: ## Build Docker image
	$(DOCKER) build -t $(PROJECT_NAME):latest .

docker-build-dev: ## Build development Docker image
	$(DOCKER) build --target development -t $(PROJECT_NAME):dev .

docker-run: ## Run Docker container
	$(DOCKER) run -p 8000:8000 --env-file .env $(PROJECT_NAME):latest

docker-compose-up: ## Start services with docker-compose
	$(DOCKER_COMPOSE) up -d

docker-compose-down: ## Stop services with docker-compose
	$(DOCKER_COMPOSE) down

docker-compose-logs: ## Show docker-compose logs
	$(DOCKER_COMPOSE) logs -f

docker-clean: ## Clean Docker resources
	$(DOCKER) system prune -f
	$(DOCKER) volume prune -f

# Utilities
clean: ## Clean up temporary files
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	rm -rf build/
	rm -rf dist/
	rm -rf .coverage
	rm -rf htmlcov/
	rm -rf .pytest_cache/
	rm -rf .mypy_cache/
	rm -rf logs/*.log
	rm -rf data/temp/*

clean-all: clean ## Clean everything including Docker
	$(DOCKER_COMPOSE) down -v
	$(DOCKER) system prune -af
	$(DOCKER) volume prune -f

# Documentation
docs: ## Generate documentation
	cd docs && make html

docs-serve: ## Serve documentation
	cd docs/_build/html && $(PYTHON) -m http.server 8001

docs-clean: ## Clean documentation
	cd docs && make clean

# Performance
perf-test: ## Run performance tests
	$(PYTHON) -m pytest tests/performance/ -v

perf-profile: ## Profile application performance
	$(PYTHON) -m cProfile -o profile.stats start_web.py

# Monitoring
monitor: ## Start monitoring dashboard
	$(PYTHON) -m utils.monitoring

health-check: ## Check application health
	curl -f http://localhost:8000/health || exit 1

# Backup and Restore
backup: ## Backup data
	$(PYTHON) -m utils.backup create

restore: ## Restore data
	$(PYTHON) -m utils.backup restore $(FILE)

# Deployment
deploy-staging: ## Deploy to staging
	$(DOCKER_COMPOSE) -f docker-compose.staging.yml up -d

deploy-production: ## Deploy to production
	$(DOCKER_COMPOSE) -f docker-compose.production.yml up -d

# Version Management
version-patch: ## Bump patch version
	bump2version patch

version-minor: ## Bump minor version
	bump2version minor

version-major: ## Bump major version
	bump2version major

# Quick start
setup: install-dev db-init ## Complete setup for development
	@echo "Setup complete! Run 'make dev' to start the development server."

quick-test: format lint test-unit ## Quick quality check
	@echo "Quick test completed successfully!"

ci: format lint test security-scan ## Full CI pipeline
	@echo "CI pipeline completed successfully!"