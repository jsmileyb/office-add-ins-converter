## Objective:

Develop a Python script to automate the consolidation and cleaning of markdown files across a structured directory.

## Functional Requirements

1. Input/Output Directories as Arguments:
   - The script must accept:
     - `--parent_dir`: Path to the top-level directory containing child folders.
     - `--output_dir`: Path for saving consolidated markdown files.
2. Directory Traversal:
   - For each child folder within the parent directory:
   - Locate all `.md` files within each child (only immediate children unless stated otherwise).
3. Content Consolidation & Cleaning:
   - Merge the contents of all markdown files within each child folder into a single markdown file.
   - Name each merged file after the corresponding child folder.
   - Remove any lines matching the pattern: `"{number}------------------------------------------------" (e.g., {2}------------------------------------------------)`.
4. Output:
   - Save each finalized, cleaned markdown file in the specified output directory.

## Project Structure & Supporting Files

- Include a `requirements.txt` specifying all required Python packages/libraries.
- Add a `.gitignore` (e.g., to ignore `__pycache__`, `.env`, etc.).
- Provide a `README.md that describes:
  - Project purpose
  - Setup instructions
  - Usage examples
  - Any relevant notes or caveats
- Use a modular directory structure for supporting code (e.g., `utils/`, `scripts/`).

## Testing Requirements

- Implement at least one test script to verify key functionalities:
  - Directory parsing and file detection
  - Markdown merging and cleaning/removal of unwanted lines
  - Output file creation and naming conventions
- Place tests in a designated `tests/` directory or as dedicated test scripts/files.
- Use Python’s `unittest` or `pytest` frameworks if appropriate.
- Ensure core functions are modular and importable for testing in isolation.

## Coding & Documentation Standards

Write idiomatic, well-documented Python code.

- Include function docstrings and inline comments as needed.
- Structure all imports at the top of each file per PEP8 guidelines.
- Follow standard Python project organization conventions for clarity and maintainability.
- List any third-party dependencies in `requirements.txt`.

## Deliverables

- Python script(s) as described above
- `requirements.txt`
- `.gitignore`
- `README.md`
- Supporting directory structure (e.g., `utils/`, `tests/`)
- At least one functional test script verifying the main features

## Example Invocation

```[python]
python consolidate_markdown.py --parent_dir /path/to/PARENT --output_dir /path/to/output

```
