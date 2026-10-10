"""Interactive XLSX-to-PDF conversion example."""

import os
import sys
from pathlib import Path


# Allow this example to run directly from the repository checkout.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose.cells_foss import Workbook


def convert_xlsx_to_pdf(source_path, target_path):
    """Convert an XLSX workbook to PDF and return the output path."""
    source = Path(source_path).expanduser()
    target = Path(target_path).expanduser()

    if not source.is_file():
        raise FileNotFoundError(f"Source XLSX file does not exist: {source}")

    target.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook(str(source))
    workbook.save_as_pdf(str(target))
    return target


def main():
    """Prompt for file paths and perform the conversion."""
    source_path = input("Enter the source XLSX file path: ").strip().strip('"')
    target_path = input("Enter the target PDF file path: ").strip().strip('"')

    if not source_path or not target_path:
        print("Both source and target file paths are required.", file=sys.stderr)
        return 1

    try:
        output_path = convert_xlsx_to_pdf(source_path, target_path)
    except Exception as exc:
        print(f"Conversion failed: {exc}", file=sys.stderr)
        return 1

    print(f"PDF file created successfully: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
