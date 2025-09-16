import os
import re
from datetime import datetime

def find_markdown_files(directory):
    """Return a list of .md files in the given directory (recursive)."""
    md_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.md'):
                md_files.append(os.path.join(root, file))
    return md_files

def clean_markdown_content(content):
    """Remove lines matching the pattern: {number}------------------------------------------------"""
    pattern = re.compile(r"^\{\d+\}-+\s*$")
    return '\n'.join(line for line in content.splitlines() if not pattern.match(line))

def consolidate_and_clean_markdown(parent_dir, output_dir):
    """Consolidate and clean markdown files from each child folder in parent_dir.
    Output is saved in a timestamped subfolder inside output_dir.
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    run_output_dir = os.path.join(output_dir, timestamp)
    os.makedirs(run_output_dir, exist_ok=True)
    print(f"Output directory for this run: {run_output_dir}")
    children = [child for child in os.listdir(parent_dir) if os.path.isdir(os.path.join(parent_dir, child))]
    total = len(children)
    for idx, child in enumerate(children, 1):
        percent = int((idx / total) * 100) if total else 100
        bar_len = 30
        filled_len = int(bar_len * percent // 100)
        bar = '=' * filled_len + '-' * (bar_len - filled_len)
        print(f"[{idx}/{total}] Processing folder: {child}")
        print(f"Progress: |{bar}| {percent}%", end='\r')
        child_path = os.path.join(parent_dir, child)
        md_files = find_markdown_files(child_path)
        merged_content = ''
        for md_file in md_files:
            print(f"    Merging: {os.path.basename(md_file)}")
            with open(md_file, 'r', encoding='utf-8') as f:
                merged_content += f.read() + '\n'
        cleaned_content = clean_markdown_content(merged_content)
        output_file = os.path.join(run_output_dir, f"{child}.md")
        with open(output_file, 'w', encoding='utf-8') as out_f:
            out_f.write(cleaned_content)
        print(f"    Saved: {output_file}")
    print()  # Newline after progress bar
    print(f"All done! {total} folders processed. Outputs in: {run_output_dir}")
