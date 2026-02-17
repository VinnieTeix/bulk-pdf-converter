# MassDocx2PDF

Bulk convert DOCX files to PDF using Microsoft Word's COM automation on Windows.

## Prerequisites

- **Windows** with **Microsoft Word** installed
- **Python 3.7+**

## Installation

```bash
pip install pywin32
```

## Usage

```bash
# Convert all .docx files in a directory (PDFs saved alongside originals)
python convert.py "C:\path\to\docx\files"

# Convert to a separate output directory
python convert.py "C:\path\to\docx\files" -o "C:\path\to\output"
```

### Options

| Argument | Description |
|---|---|
| `input_dir` | Directory containing `.docx` files |
| `-o`, `--output-dir` | Output directory for PDFs (defaults to `input_dir`) |

## Example

```
> python convert.py ./documents -o ./pdfs

[1/3] report.docx -> report.pdf OK
[2/3] invoice.docx -> invoice.pdf OK
[3/3] letter.docx -> letter.pdf OK

Done: 3/3 converted successfully.
```
