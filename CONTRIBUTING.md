# Contributing to Excel Building Management

Thank you for your interest in contributing to this project! We welcome contributions from everyone.

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment for all contributors.

## How to Contribute

### Reporting Bugs

If you find a bug, please open an issue with:
- A clear, descriptive title
- Steps to reproduce the issue
- Expected behavior vs actual behavior
- Your environment (OS, Python version, etc.)
- Relevant code snippets or error messages

### Suggesting Enhancements

We welcome suggestions for new features or improvements. Please open an issue with:
- A clear description of the enhancement
- Use cases and examples
- Why this enhancement would be useful

### Pull Requests

1. **Fork the repository** and create your branch from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes**:
   - Follow the existing code style
   - Add type hints to all functions
   - Write docstrings for new functions
   - Add tests for new functionality
   - Update documentation as needed

3. **Test your changes**:
   ```bash
   # Run tests
   pytest tests/ -v

   # Check test coverage
   pytest tests/ --cov=. --cov-report=html
   ```

4. **Commit your changes**:
   - Use clear, descriptive commit messages
   - Follow the conventional commits format:
     - `feat:` for new features
     - `fix:` for bug fixes
     - `docs:` for documentation changes
     - `test:` for adding tests
     - `refactor:` for code refactoring
     - `chore:` for maintenance tasks

   Example:
   ```bash
   git commit -m "feat: Add support for multi-currency rent tracking"
   ```

5. **Push to your fork**:
   ```bash
   git push origin feature/your-feature-name
   ```

6. **Open a Pull Request**:
   - Provide a clear description of the changes
   - Reference any related issues
   - Ensure all tests pass
   - Wait for review and address feedback

## Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/Kandil7/Task-Excel-For-the-management-of-buildings.git
   cd Task-Excel-For-the-management-of-buildings
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Create a config file for testing:
   ```bash
   cp config.example.json config.json
   ```

## Code Style Guidelines

### Python Code Style

- Follow [PEP 8](https://www.python.org/dev/peps/pep-0008/) style guidelines
- Use 4 spaces for indentation (no tabs)
- Maximum line length: 100 characters
- Use meaningful variable and function names

### Type Hints

All functions should include type hints:

```python
def calculate_total(amounts: List[float]) -> float:
    """
    Calculate the total of a list of amounts.

    Args:
        amounts: List of numeric amounts

    Returns:
        Sum of all amounts
    """
    return sum(amounts)
```

### Docstrings

Use Google-style docstrings:

```python
def function_name(param1: str, param2: int) -> bool:
    """
    Brief description of function.

    Longer description if needed, explaining what the function does
    and any important details.

    Args:
        param1: Description of param1
        param2: Description of param2

    Returns:
        Description of return value

    Raises:
        ValueError: When param2 is negative
    """
    pass
```

### Error Handling

Always include proper error handling:

```python
try:
    # Risky operation
    result = perform_operation()
except SpecificException as e:
    # Handle specific exception
    logger.error(f"Operation failed: {e}")
    raise
```

## Testing Guidelines

### Writing Tests

- Place tests in the `tests/` directory
- Name test files `test_*.py`
- Name test functions `test_*`
- Use descriptive test names that explain what is being tested
- Include both positive and negative test cases
- Test edge cases and error conditions

Example:

```python
def test_parse_valid_date():
    """Test parsing a valid date string"""
    result = parse_date("2024-01-15")
    expected = datetime(2024, 1, 15).date()
    assert result == expected

def test_parse_invalid_date():
    """Test parsing invalid date format"""
    result = parse_date("invalid-date")
    assert result is None
```

### Running Tests

Run all tests:
```bash
pytest tests/ -v
```

Run specific test file:
```bash
pytest tests/test_excel_generate_v2.py -v
```

Run with coverage:
```bash
pytest tests/ --cov=. --cov-report=html
```

## Documentation

- Update the README.md if you change functionality
- Add docstrings to all new functions and classes
- Update configuration examples if config structure changes
- Add examples for new features

## Questions?

If you have questions about contributing, feel free to:
- Open an issue with your question
- Review existing issues and pull requests
- Check the README.md for basic usage information

## License

By contributing to this project, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to make this project better!
