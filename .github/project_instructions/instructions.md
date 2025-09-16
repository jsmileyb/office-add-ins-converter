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

## Output Directory Structure

When running the consolidation script, the finalized consolidated markdown files will be saved in the specified output directory.  
Each set of output files is stored in a uniquely named subfolder within the output directory, following the naming convention:

yyyyMMdd_mmss

```

- `yyyy` = four-digit year
- `MM`   = two-digit month
- `dd`   = two-digit day
- `mm`   = two-digit hour and
- `ss`   = two-digit minute and second based on the time of execution
```

**For example:**  
If the script is run on March 15, 2024 at 14:30:45, the output will be saved to: `output/20240315_143045/`

All consolidated markdown files generated during that run will be placed inside this timestamped folder, keeping outputs organized and preventing accidental overwrites.

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
