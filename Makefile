SHELL := /bin/bash

.DEFAULT_GOAL := help

NPM ?= npm
PYTHON ?= python3
PIP ?= $(PYTHON) -m pip
PRE_COMMIT ?= pre-commit
PRE_COMMIT_HOME ?= /tmp/pre-commit-cache
PYTHONPATH_VALUE ?= python
COM_SMOKE_OUTPUT_DIR ?= artifacts/com-smoke
SNAPSHOT_WIDTH_PX ?= 1280

.PHONY: help install install-node install-py mcp mcp-start build start test test-py lint-py format-py format-check-py check precommit-install precommit-run com-smoke com-smoke-ooxml clean

help: ## Show available make targets
	@awk 'BEGIN {FS = ":.*##"; print "Available targets:"} /^[a-zA-Z0-9_.-]+:.*##/ {printf "  %-20s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

install-node: ## Install Node.js dependencies
	$(NPM) install

install-py: ## Install Python runtime and development dependencies
	$(PIP) install -r requirements.txt -r requirements-dev.txt

install: install-node install-py ## Install all project dependencies

mcp: ## Run MCP server in development mode
	$(NPM) run dev

mcp-start: build ## Build and run MCP server from dist/
	$(NPM) run start

build: ## Build TypeScript project
	$(NPM) run build

start: ## Start MCP server from dist/
	$(NPM) run start

test: ## Run TypeScript tests
	$(NPM) test

test-py: ## Run Python tests
	$(NPM) run test:py

lint-py: ## Run Ruff lint checks
	$(NPM) run lint:py

format-py: ## Format Python files with Black
	$(NPM) run format:py

format-check-py: ## Check Python formatting with Black
	$(NPM) run format:check:py

check: ## Run full validation suite
	$(NPM) run check

precommit-install: ## Install git pre-commit hooks
	PRE_COMMIT_HOME=$(PRE_COMMIT_HOME) $(PRE_COMMIT) install --install-hooks

precommit-run: ## Run all pre-commit hooks for the repository
	PRE_COMMIT_HOME=$(PRE_COMMIT_HOME) $(PRE_COMMIT) run --all-files

com-smoke: ## Run Windows COM smoke runner (licensed Windows + PowerPoint required)
	PYTHONPATH=$(PYTHONPATH_VALUE) $(PYTHON) scripts/windows_com_smoke.py --output-dir "$(COM_SMOKE_OUTPUT_DIR)" --snapshot-width-px $(SNAPSHOT_WIDTH_PX)

com-smoke-ooxml: ## Run smoke runner in OOXML debug mode (non-Windows script sanity check)
	PYTHONPATH=$(PYTHONPATH_VALUE) $(PYTHON) scripts/windows_com_smoke.py --allow-ooxml --skip-snapshot --output-dir "$(COM_SMOKE_OUTPUT_DIR)" --snapshot-width-px $(SNAPSHOT_WIDTH_PX)

clean: ## Remove generated caches and build artifacts
	rm -rf dist .pytest_cache .ruff_cache artifacts/com-smoke
