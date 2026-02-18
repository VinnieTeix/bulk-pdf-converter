"""Bulk DOCX/ODT/XLSX/XLS to PDF converter using Microsoft Office COM automation."""

import argparse
import sys
from pathlib import Path

import win32com.client

WORD_EXTENSIONS = {".docx", ".odt"}
EXCEL_EXTENSIONS = {".xlsx", ".xls"}
ALL_EXTENSIONS = WORD_EXTENSIONS | EXCEL_EXTENSIONS


def convert_to_pdf(input_dir: Path, output_dir: Path) -> None:
    files = sorted(
        f for f in input_dir.iterdir() if f.suffix.lower() in ALL_EXTENSIONS
    )
    if not files:
        print(f"No supported files found in {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)

    word_files = [f for f in files if f.suffix.lower() in WORD_EXTENSIONS]
    excel_files = [f for f in files if f.suffix.lower() in EXCEL_EXTENSIONS]

    total = len(files)
    success = 0
    failed = []
    counter = 0

    if word_files:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            for src_path in word_files:
                counter += 1
                pdf_path = output_dir / (src_path.stem + ".pdf")
                print(f"[{counter}/{total}] {src_path.name} -> {pdf_path.name}", end=" ")
                try:
                    doc = word.Documents.Open(str(src_path.resolve()))
                    doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)  # 17 = wdFormatPDF
                    doc.Close()
                    print("OK")
                    success += 1
                except Exception as e:
                    print(f"FAILED: {e}")
                    failed.append(src_path.name)
        finally:
            word.Quit()

    if excel_files:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            for src_path in excel_files:
                counter += 1
                pdf_path = output_dir / (src_path.stem + ".pdf")
                print(f"[{counter}/{total}] {src_path.name} -> {pdf_path.name}", end=" ")
                try:
                    wb = excel.Workbooks.Open(str(src_path.resolve()))
                    wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))  # 0 = xlTypePDF
                    wb.Close(False)
                    print("OK")
                    success += 1
                except Exception as e:
                    print(f"FAILED: {e}")
                    failed.append(src_path.name)
        finally:
            excel.Quit()

    print(f"\nDone: {success}/{total} converted successfully.")
    if failed:
        print("Failed files:")
        for name in failed:
            print(f"  - {name}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Bulk convert DOCX/ODT/XLSX/XLS files to PDF.")
    parser.add_argument("input_dir", help="Directory containing files to convert")
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

    convert_to_pdf(input_dir, output_dir)


if __name__ == "__main__":
    main()
