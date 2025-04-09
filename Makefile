# Makefile para KeepAlive Manager

PYTHON = python
SRC = keep-alive-app.py keep-alive-app_electron.py

.PHONY: help lint format check run

help:
	@echo "Comandos disponíveis:"
	@echo "  make lint     - Verifica código com pylint e flake8"
	@echo "  make format   - Formata o código com black"
	@echo "  make check    - Verifica lint + estilo"
	@echo "  make run      - Executa a aplicação principal"

lint:
	@echo "Executando pylint..."
	@pylint $(SRC)
	@echo "Executando flake8..."
	@flake8

format:
	@echo "Formatando com black..."
	@black $(SRC)

check: lint

run:
	$(PYTHON) keep-alive-app_electron.py
