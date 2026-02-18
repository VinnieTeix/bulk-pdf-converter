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

Place files in the `input/` folder (subfolders supported), then run:

```bash
python convert.py
```

PDFs are saved to `output/`, mirroring the subfolder structure of `input/`.

### Custom directories

```bash
python convert.py "C:\path\to\source" -o "C:\path\to\output"
```

### Options

| Argument | Description |
|---|---|
| `input_dir` | Source directory (default: `input`) |
| `-o`, `--output-dir` | Output directory (default: `output`) |

## Example

```
input/
├── ClientA/
│   ├── report.docx
│   └── budget.xlsx
└── ClientB/
    └── letter.odt
```

```
> python convert.py

[1/3] ClientA\report.docx -> ClientA\report.pdf OK
[2/3] ClientB\letter.odt -> ClientB\letter.pdf OK
[3/3] ClientA\budget.xlsx -> ClientA\budget.pdf OK

Done: 3/3 converted successfully.
```

```
output/
├── ClientA/
│   ├── report.pdf
│   └── budget.pdf
└── ClientB/
    └── letter.pdf
```
