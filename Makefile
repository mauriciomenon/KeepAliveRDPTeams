.PHONY: lint format check run

lint:
	@echo "Executando pylint..."
	pylint *.py

format:
	@echo "Formatando com black..."
	black *.py

check:
	@echo "Executando pylint..."
	pylint *.py

run:
	python keep_alive_manager_final.py