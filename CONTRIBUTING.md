# Contributing to PDF Text Extractor

First off, thank you for considering contributing to PDF Text Extractor!

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates. When creating a bug report, include:

- **Clear title and description**
- **Steps to reproduce** the behavior
- **Expected vs actual behavior**
- **Screenshots** if applicable
- **Environment details** (OS, Python version, etc.)

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion, include:

- **Clear title and description** of the feature
- **Use cases** - explain why this would be useful
- **Possible implementation** if you have ideas

### Pull Requests

1. Fork the repository
2. Create a new branch (`git checkout -b feature/AmazingFeature`)
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
6. Push to the branch (`git push origin feature/AmazingFeature`)
7. Open a Pull Request

#### Pull Request Guidelines

- Follow the existing code style
- Update documentation as needed
- Add comments for complex logic
- Test on multiple platforms if possible
- Keep PRs focused on a single feature/fix

## Code Style

- Follow PEP 8 guidelines for Python code
- Use meaningful variable and function names
- Add docstrings for functions and classes
- Keep functions focused and modular

## Development Setup

```bash
# Clone your fork
git clone https://github.com/YOUR_USERNAME/PDF_File_Extractor.git
cd PDF_File_Extractor

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Testing

Before submitting a PR:

1. Test the application with various PDF files
2. Test with and without Tesseract installed
3. Test OCR with different languages
4. Verify Fast Mode works correctly
5. Check all UI elements function properly

## Questions?

Feel free to open an issue for any questions or concerns!

Thank you for contributing!
