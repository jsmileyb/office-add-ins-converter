import os
import re


def find_markdown_files(directory):
    """Return a list of .md files in the given directory (non-recursive)."""
    return [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.md') and os.path.isfile(os.path.join(directory, f))]


def clean_markdown_content(content):
    """Remove lines matching the pattern: {number}------------------------------------------------"""
    pattern = re.compile(r"^\{\d+\}-+\s*$")
    return '\n'.join(line for line in content.splitlines() if not pattern.match(line))


def consolidate_and_clean_markdown(parent_dir, output_dir):
    """Consolidate and clean markdown files from each child folder in parent_dir."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    for child in os.listdir(parent_dir):
        child_path = os.path.join(parent_dir, child)
        if os.path.isdir(child_path):
            md_files = find_markdown_files(child_path)
            merged_content = ''
            for md_file in md_files:
                with open(md_file, 'r', encoding='utf-8') as f:
                    merged_content += f.read() + '\n'
            cleaned_content = clean_markdown_content(merged_content)
            output_file = os.path.join(output_dir, f"{child}.md")
            with open(output_file, 'w', encoding='utf-8') as out_f:
                out_f.write(cleaned_content)
