<<<<<<< HEAD
# automation
This repo contains all the code automated for quality checks
=======
# Cornerstone Automation

A well-structured Python project with modern development practices.

## Features

- Clean project structure
- Virtual environment management
- Dependency management with pip
- Testing setup with pytest
- Code formatting with black
- Linting with flake8
- Type checking with mypy

## Project Structure

```
cornerstone_automation/
├── src/
│   └── cornerstone_automation/
│       ├── __init__.py
│       └── main.py
├── tests/
│   ├── __init__.py
│   └── test_main.py
├── requirements.txt
├── requirements-dev.txt
├── setup.py
├── .gitignore
├── README.md
└── pyproject.toml
```

## Setup Instructions

1. **Create a virtual environment:**
   ```bash
   python -m venv venv
   ```

2. **Activate the virtual environment:**
   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - macOS/Linux:
     ```bash
     source venv/bin/activate
     ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install development dependencies:**
   ```bash
   pip install -r requirements-dev.txt
   ```

## Development

- **Run the application:**
  ```bash
  python src/cornerstone_automation/main.py
  ```

- **Run tests:**
  ```bash
  pytest
  ```

- **Format code:**
  ```bash
  black src/ tests/
  ```

- **Lint code:**
  ```bash
  flake8 src/ tests/
  ```

- **Type checking:**
  ```bash
  mypy src/
  ```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run tests and linting
5. Submit a pull request

## License

This project is licensed under the MIT License. 
>>>>>>> b69b19c (Adding all the automation code to github repo for the 1st time)
