import argparse
from .utils.markdown_utils import consolidate_and_clean_markdown

def main():
    parser = argparse.ArgumentParser(description="Consolidate and clean markdown files.")
    parser.add_argument('--parent_dir', required=True, help='Path to the parent directory')
    parser.add_argument('--output_dir', required=True, help='Path to the output directory')
    args = parser.parse_args()

    consolidate_and_clean_markdown(args.parent_dir, args.output_dir)

if __name__ == "__main__":
    main()
