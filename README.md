# WordTool

A Python tool for processing Word documents, specifically designed to find and replace text patterns with sequential numbering.

## Features

### Text Pattern Enumeration

The main feature allows you to find and replace text patterns like `[REQ-XXX]` and `[SYS-XXX]` with sequential numbers (e.g., `[REQ-001]`, `[REQ-002]`, etc.).

#### Key Features:
- **Configurable prefixes**: Support for any prefix (REQ, SYS, BUG, FEATURE, etc.)
- **Independent numbering**: Each prefix maintains its own sequential counter
- **Comprehensive document processing**: Handles paragraphs, tables, headers, and footers
- **Flexible pattern matching**: Customizable regex patterns
- **Zero-padded numbers**: Numbers are formatted as 3-digit sequences (001, 002, etc.)

## Installation

```bash
# Install dependencies
pip install -r requirements.txt
# or using uv
uv sync
```

## Usage

### Command Line Interface

```bash
# Basic usage with default prefixes (REQ, SYS)
python src/enumerate.py input.docx output.docx

# Custom prefixes
python src/enumerate.py input.docx output.docx --prefixes REQ SYS BUG FEATURE

# Custom pattern (default is [PREFIX-XXX])
python src/enumerate.py input.docx output.docx --pattern '\[([A-Z]+)-XXX\]'
```

### Programmatic Usage

```python
from src.enumerate import find_and_replace_patterns

# Basic usage
replacements = find_and_replace_patterns('input.docx', 'output.docx')

# Custom prefixes
replacements = find_and_replace_patterns(
    'input.docx', 
    'output.docx',
    prefixes=['REQ', 'SYS', 'BUG', 'FEATURE']
)

# Custom pattern
replacements = find_and_replace_patterns(
    'input.docx', 
    'output.docx',
    pattern=r'\[([A-Z]+)-XXX\]'
)

# Print results
for prefix, numbers in replacements.items():
    if numbers:
        print(f"{prefix}: {len(numbers)} replacements")
        print(f"  Examples: {', '.join(numbers[:5])}")
```

## Examples

### Input Document Content:
```
This document contains several requirements:
- [REQ-XXX]: User authentication system
- [REQ-XXX]: Password validation
- [SYS-XXX]: Database connection
- [SYS-XXX]: Logging system
- [REQ-XXX]: Email notifications
```

### Output Document Content:
```
This document contains several requirements:
- [REQ-001]: User authentication system
- [REQ-002]: Password validation
- [SYS-001]: Database connection
- [SYS-002]: Logging system
- [REQ-003]: Email notifications
```

## API Reference

### `find_and_replace_patterns()`

Main function for processing Word documents.

**Parameters:**
- `input_file` (str): Path to input Word document
- `output_file` (str): Path to output Word document
- `prefixes` (List[str] | None): List of prefixes to process (default: ['REQ', 'SYS'])
- `pattern` (str): Regex pattern to match (default: r'\[([A-Z]+)-XXX\]')

**Returns:**
- `Dict[str, List[str]]`: Dictionary mapping prefixes to lists of generated numbers

## Requirements

- Python 3.12+
- bayoo-docx>=0.2.20
- click>=8.1.8

## Development

This project uses:
- **uv** for dependency management
- **flake8** for code formatting (max line length: 160)
- Single quotes preferred for Python strings
- Public elements at top of files, private at bottom

## License

Personal project by Boris Resnick (boris.resnick@gmail.com)
