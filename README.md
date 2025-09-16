# Markdown Consolidator

## Purpose
Automates consolidation and cleaning of markdown files across a structured directory.

## Setup
1. Install Python 3.7+
2. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

## Usage
```sh
python scripts/consolidate_markdown.py --parent_dir /path/to/PARENT --output_dir /path/to/output
```

## Notes
- Output files are named after each child folder.
- Unwanted lines matching `{number}------------------------------------------------` are removed.
