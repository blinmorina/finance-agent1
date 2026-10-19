.PHONY: install install-dev test run help

help:
	@echo "Targets:"
	@echo "  make install      Create venv and install dependencies"
	@echo "  make install-dev  install + pytest"
	@echo "  make test         Run unit tests"
	@echo "  make run          Interactive agent (uses .env)"

install:
	python3 -m venv .venv
	.venv/bin/pip install -U pip
	.venv/bin/pip install -r requirements.txt

install-dev: install
	.venv/bin/pip install pytest

test:
	.venv/bin/pytest tests/ -q

run:
	.venv/bin/python main.py
