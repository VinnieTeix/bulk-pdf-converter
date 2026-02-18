# MassDocx2PDF

Bulk convert DOCX, ODT, XLSX, and XLS files to PDF using Microsoft Office COM automation on Windows.

## Supported Formats

| Format | Application |
|---|---|
| `.docx`, `.odt` | Microsoft Word |
| `.xlsx`, `.xls` | Microsoft Excel |

## Prerequisites

- **Windows** with **Microsoft Word** and/or **Microsoft Excel** installed
- **Python 3.7+**

## Installation

```bash
pip install pywin32
```

## Usage

```bash
# Convert all supported files in a directory (PDFs saved alongside originals)
python convert.py "C:\path\to\files"

# Convert to a separate output directory
python convert.py "C:\path\to\files" -o "C:\path\to\output"
```

### Options

| Argument | Description |
|---|---|
| `input_dir` | Directory containing files to convert |
| `-o`, `--output-dir` | Output directory for PDFs (defaults to `input_dir`) |

## Example

```
> python convert.py ./documents -o ./pdfs

[1/4] report.docx -> report.pdf OK
[2/4] letter.odt -> letter.pdf OK
[3/4] budget.xlsx -> budget.pdf OK
[4/4] data.xls -> data.pdf OK

Done: 4/4 converted successfully.
```
