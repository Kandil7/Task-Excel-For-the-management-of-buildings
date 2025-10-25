.PHONY: help install test test-coverage run run-legacy clean lint format

# Default target
help:
	@echo "Excel Building Management - Available Commands"
	@echo "=============================================="
	@echo ""
	@echo "Setup:"
	@echo "  make install        - Install dependencies"
	@echo ""
	@echo "Running:"
	@echo "  make run            - Run the modern version (v2)"
	@echo "  make run-legacy     - Run the legacy version"
	@echo ""
	@echo "Testing:"
	@echo "  make test           - Run all tests"
	@echo "  make test-coverage  - Run tests with coverage report"
	@echo ""
	@echo "Code Quality:"
	@echo "  make lint           - Check code style (requires flake8)"
	@echo "  make format         - Format code (requires black)"
	@echo ""
	@echo "Maintenance:"
	@echo "  make clean          - Remove generated files and caches"
	@echo ""

# Install dependencies
install:
	pip install -r requirements.txt

# Run the modern version
run:
	python excel_generate_v2.py

# Run the legacy version
run-legacy:
	python excel_generate.py

# Run tests
test:
	pytest tests/ -v

# Run tests with coverage
test-coverage:
	pytest tests/ --cov=. --cov-report=html --cov-report=term
	@echo ""
	@echo "Coverage report generated in htmlcov/index.html"

# Lint code (requires flake8)
lint:
	@which flake8 > /dev/null || (echo "flake8 not found. Install with: pip install flake8" && exit 1)
	flake8 excel_generate.py excel_generate_v2.py tests/ --max-line-length=100 --ignore=E501,W503

# Format code (requires black)
format:
	@which black > /dev/null || (echo "black not found. Install with: pip install black" && exit 1)
	black excel_generate.py excel_generate_v2.py tests/ --line-length=100

# Clean generated files
clean:
	rm -rf __pycache__ tests/__pycache__
	rm -rf .pytest_cache
	rm -rf htmlcov
	rm -rf .coverage
	rm -f *.xlsx
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -delete
