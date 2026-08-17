import argparse
from pathlib import Path

from mcp_app import convert_document_file


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert a DOCX or text-based file to Markdown or YAML")
    parser.add_argument("document_path", help="Path to the document file")
    parser.add_argument("--format", choices=["markdown", "yaml"], default="markdown", help="Output format")
    args = parser.parse_args()

    if not Path(args.document_path).exists():
        raise FileNotFoundError(f"Document not found: {args.document_path}")

    output = convert_document_file(args.document_path, args.format)
    print(output)


if __name__ == '__main__':
    main()
