# Batch DOC/DOCX → PDF Converter

A Python utility for bulk-converting Microsoft Word documents (`.doc` / `.docx`) to PDF. Supports recursive folder scanning, mirrored output structure, and optional merge of all PDFs into a single file.

Inspired by [batch-excel-xlsx-csv-to-pdf-converter](https://github.com/SakibAhmedShuva/batch-excel-xlsx-csv-to-pdf-converter).

---

## Features

- **Batch conversion** — process an entire folder of Word documents in one run
- **Recursive search** — scans all subfolders and mirrors the folder structure in output
- **Merge to single PDF** — optionally combine all converted PDFs into one file
- **Dual backend** — auto-detects LibreOffice; falls back to pure-Python `mammoth` + `weasyprint` if unavailable
- **Cross-platform** — Windows, macOS, Linux
- **Graceful error handling** — per-file errors are logged; conversion continues for remaining files
- **Timestamped logging** — clear pass/fail summary at the end

---

## Requirements

### Option A — LibreOffice (recommended, best fidelity)

Install LibreOffice from [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/).

| Platform | Default install path |
|----------|----------------------|
| Windows  | `C:\Program Files\LibreOffice\program\soffice.exe` |
| macOS    | `/Applications/LibreOffice.app/Contents/MacOS/soffice` |
| Linux    | `/usr/bin/soffice` |

If LibreOffice is installed to a non-standard location, set the path in the `_LO_CANDIDATES` list at the top of the script.

### Option B — Pure Python fallback

```bash
pip install mammoth weasyprint pypdf
```

No system dependencies required. Fidelity may vary for complex documents.

---

## Installation

```bash
git clone https://github.com/SakibAhmedShuva/batch-doc-to-pdf-converter.git
cd batch-doc-to-pdf-converter
```

No additional `pip install` needed if LibreOffice is present. For the Python fallback:

```bash
pip install mammoth weasyprint pypdf
```

---

## Usage

### 1. Configure paths

Open `batch_docx_to_pdf.py` and edit the configuration block at the top:

```python
INPUT_DIR      = r"C:\Users\You\Documents\my_docs"   # folder containing .doc / .docx files
OUTPUT_PDF_DIR = r"C:\Users\You\Documents\output"     # where PDFs will be saved
```

### 2. Run

```bash
python batch_docx_to_pdf.py
```

### 3. Output

```
[2026-05-23 16:45:20] [INFO]    Batch DOCX → PDF Converter
[2026-05-23 16:45:20] [INFO]    Backend  : libreoffice
[2026-05-23 16:45:20] [INFO]    Found 22 file(s) to process.
[2026-05-23 16:45:21] [OK]  ✓  Saved: output\CD 01.pdf  (84.2 KB)
[2026-05-23 16:45:22] [OK]  ✓  Saved: output\CD 02.pdf  (91.7 KB)
...
[2026-05-23 16:46:10] [INFO]    Converted : 22 | Skipped : 0 | Failed : 0
```

---

## Configuration Reference

All options are in the configuration block at the top of `batch_docx_to_pdf.py`:

| Option | Default | Description |
|--------|---------|-------------|
| `INPUT_DIR` | *(set by user)* | Folder containing Word documents |
| `OUTPUT_PDF_DIR` | *(set by user)* | Destination folder for PDFs |
| `RECURSIVE_SEARCH` | `True` | Scan subfolders recursively |
| `CREATE_SUBFOLDERS_IN_OUTPUT` | `True` | Mirror input folder structure in output |
| `FILE_EXTENSIONS_TO_CONVERT` | `[".docx", ".doc"]` | File types to include |
| `OVERWRITE_EXISTING` | `True` | Overwrite PDFs that already exist |
| `MERGE_ALL_INTO_ONE_PDF` | `False` | Combine all output PDFs into one file |
| `MERGED_PDF_FILENAME` | `"merged_output.pdf"` | Name of the combined PDF |

---

## Merge All to Single PDF

Set `MERGE_ALL_INTO_ONE_PDF = True` in the config block. After all individual conversions complete, a single `merged_output.pdf` will be written to `OUTPUT_PDF_DIR` containing every converted document in sorted order.

```python
MERGE_ALL_INTO_ONE_PDF = True
MERGED_PDF_FILENAME    = "merged_output.pdf"
```

Requires `pypdf`:

```bash
pip install pypdf
```

---

## Backend Detection Order

The script automatically selects the best available backend:

```
LibreOffice (soffice)  →  mammoth + weasyprint  →  error with install instructions
```

LibreOffice is searched in ~15 common install locations across Windows, macOS, and Linux before falling back to the Python stack. To force a specific path, add it first in `_LO_CANDIDATES`:

```python
_LO_CANDIDATES = [
    r"D:\Program Files\LibreOffice\program\soffice.exe",  # custom path
    ...
]
```

---

## Notes

- `.doc` (legacy Word 97–2003) files are fully supported via both backends
- Conversion runs through a temp directory to avoid permission conflicts when input and output folders are the same
- LibreOffice is launched in `--headless` mode; no GUI window will appear
- Each file has a 120-second timeout; hung conversions are skipped with an error log

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Author

**Sakib Ahmed Shuva**  
[github.com/SakibAhmedShuva](https://github.com/SakibAhmedShuva)
