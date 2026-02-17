"""Bulk DOCX to PDF converter using Microsoft Word COM automation."""

import argparse
import sys
from pathlib import Path

import win32com.client


def convert_docx_to_pdf(input_dir: Path, output_dir: Path) -> None:
    docx_files = sorted(input_dir.glob("*.docx"))
    if not docx_files:
        print(f"No .docx files found in {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        total = len(docx_files)
        success = 0
        failed = []

        for i, docx_path in enumerate(docx_files, 1):
            pdf_path = output_dir / (docx_path.stem + ".pdf")
            print(f"[{i}/{total}] {docx_path.name} -> {pdf_path.name}", end=" ")

            try:
                doc = word.Documents.Open(str(docx_path.resolve()))
                doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
                print("OK")
                success += 1
            except Exception as e:
                print(f"FAILED: {e}")
                failed.append(docx_path.name)

        print(f"\nDone: {success}/{total} converted successfully.")
        if failed:
            print("Failed files:")
            for name in failed:
                print(f"  - {name}")
    finally:
        word.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(description="Bulk convert DOCX files to PDF.")
    parser.add_argument("input_dir", help="Directory containing .docx files")
    parser.add_argument(
        "-o", "--output-dir",
        help="Output directory for PDFs (defaults to input_dir)",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"Error: {input_dir} is not a directory", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else input_dir

    convert_docx_to_pdf(input_dir, output_dir)


if __name__ == "__main__":
    main()
