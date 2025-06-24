# Contributing to CostSpirits

Thank you for your interest in contributing to CostSpirits! This document provides guidelines for contributing to the project.

## How to Contribute

### Reporting Issues

If you find a bug or have a feature request:

1. Check if the issue already exists in the [Issues](https://github.com/Harsh223/CostSpirits/issues) section
2. If not, create a new issue with:
   - Clear description of the problem or feature request
   - Steps to reproduce (for bugs)
   - Expected vs actual behavior
   - Screenshots if applicable
   - Your environment details (OS, Python version, etc.)

### Submitting Changes

1. **Fork the repository**
   ```bash
   git fork https://github.com/Harsh223/CostSpirits.git
   ```

2. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

3. **Make your changes**
   - Follow the existing code style
   - Add comments for complex logic
   - Update documentation if needed

4. **Test your changes**
   - Ensure the application runs without errors
   - Test the affected functionality thoroughly
   - Verify Excel export/import still works

5. **Commit your changes**
   ```bash
   git commit -am 'Add new feature: description'
   ```

6. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```

7. **Create a Pull Request**
   - Provide a clear description of your changes
   - Reference any related issues
   - Include screenshots for UI changes

## Development Guidelines

### Code Style

- Follow PEP 8 for Python code
- Use meaningful variable and function names
- Add docstrings for functions and classes
- Keep functions focused and concise

### Adding New Subsystems

To add a new subsystem type:

1. Add the subsystem to `AVAILABLE_SUBSYSTEMS` in `CostSpirits.py`
2. Add corresponding headers to `subsystem_headers.json`
3. Ensure headers include the core cost-related fields such as:
   - Mission
   - WBS Item
   - Lower/Higher Weight Range (lbs)
   - Lower/Higher D&D Cost Range
   - Lower/Higher Flight Unit Cost Range
   - Lower/Higher Total Cost Range
Please note that the above steps are a general guideline and may need to be adapted based on the specific requirements for the subsystem.
### Testing

Before submitting:

1. Test template generation for your changes
2. Test cost analysis functionality
3. Verify Excel export works correctly
4. Test with different subsystem combinations

## Project Structure

```
CostSpirits/
├── CostSpirits.py              # Main application
├── subsystem_headers.json      # Subsystem configuration
├── Inflation Table.xlsx        # Inflation data
├── requirements.txt            # Dependencies
├── README.md                   # Project documentation
├── LICENSE                     # License file
├── .gitignore                  # Git ignore rules
├── CHANGELOG.md               # Version history
└── CONTRIBUTING.md            # This file
```

## Questions?

If you have questions about contributing, feel free to:
- Open an issue for discussion
- Contact the maintainer
- Check existing issues and pull requests for similar questions

Thank you for contributing to CostSpirits!