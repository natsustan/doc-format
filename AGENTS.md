# AGENTS.md

## Commands

```bash
pip3 install -r requirements.txt             # Install dependencies
python3 format_doc.py doc.docx               # Run formatter
python3 format_doc.py doc.docx -c config.yaml  # With custom config
```

## Architecture

Single-file Python CLI tool for Word document formatting. Uses `python-docx` library.

- `format_doc.py` - Main script: loads config, formats paragraphs/tables/images, handles margins
- `config.yaml` - YAML config for fonts, spacing, margins, and heading styles

## Code Style

- Python 3, type hints in function signatures
- Chinese comments and user-facing messages
- Imports: stdlib first, then third-party (`yaml`, `docx`)
- Helper functions for font/spacing/alignment operations
- Use `Path` from pathlib for file paths
- Error handling: print error message and return non-zero exit code
