# PDF & Word Keyword Scanner

A desktop GUI tool that scans every PDF (`.pdf`) and Word document (`.docx`) in a folder for one or more keywords, and shows the results in a sortable table.

## Features

- **Directory picker** — browse to any folder; defaults to the script's own folder on launch
- **Multiple keywords** — comma-separated; unquoted keywords are full Python regex patterns (e.g. `PO-\d+`, `overdue|outstanding`, `inv(oice|oicing)`)
- **Exact / whole-word matching** — wrap a keyword in double quotes (e.g. `"PO"`) to match it as a literal whole word; regex metacharacters are escaped automatically, so `"U.S.A."` matches dots literally
- **Pre-scan regex validation** — invalid patterns are caught and reported before scanning begins
- **Mixed file types** — scans `.pdf` and `.docx` files together in one pass
- **Word table support** — searches inside Word table cells as well as body text
- **Per-keyword results** — one Yes/No column per keyword, plus an "Any Match" summary column
- **Case sensitivity toggle** — off by default; applies to both quoted and unquoted keywords
- **Recursive scan toggle** — optionally include subfolders
- **Background scanning** — runs on a separate thread so the GUI doesn't freeze on large folders, with live progress in the status bar
- **Error handling** — corrupt, encrypted, or unreadable files are marked `Error` instead of crashing the scan
- **CSV export** — save results for use elsewhere
- **Auto-installing dependencies** — installs `pypdf` and `python-docx` automatically on first run if missing

## Requirements

- Python 3.8+
- `tkinter` (included with most standard Python installs; on Linux you may need `sudo apt install python3-tk` if missing)

No manual `pip install` is required — the script checks for `pypdf` and `python-docx` on startup and installs them automatically if they're not already present.

## Usage

```bash
python pdf_keyword_scanner.pyw
```

1. On launch, the directory field defaults to the folder where the script resides.
2. Click **Browse...** to change directory if needed.
3. Type one or more keywords into the **Keyword(s)** field, separated by commas.
   - **Unquoted** → full Python regex pattern (case-insensitive by default)
     - `invoice` — simple substring (same as before)
     - `inv(oice|oicing)` — alternation
     - `PO-\d+` — matches "PO-2024", "PO-555", etc.
     - `overdue|outstanding` — either word anywhere in the text
   - **`"quoted"`** → exact whole-word literal match (regex metacharacters are escaped automatically)
     - `"PO"` matches the standalone word "PO" but not inside "REPORT" or "PORT"
     - `"U.S.A."` matches the literal string "U.S.A." (dots are not treated as regex)
   - Mix both freely: `inv(oice|oicing), "PO", overdue|outstanding`
   - Quoted keywords may contain commas (e.g. `"item, version 2"`)
4. (Optional) Tick **Case sensitive** — applies to both regex and quoted keywords.
5. (Optional) Tick **Include subfolders**.
6. Click **Scan** (or press Enter). Invalid regex patterns are flagged immediately before the scan starts.
7. Review results in the table. Exact-match columns show their header in quotes.
8. (Optional) Click **Export CSV** to save results.

## Output columns

| Column | Description |
|---|---|
| Filename | Relative path of the file within the scanned folder |
| Type | `PDF` or `Word` |
| *(one column per keyword)* | `Yes` / `No` / `Error` |
| Any Match | `Yes` if at least one keyword was found, `No` otherwise |

## Notes & limitations

- **"Whole word" boundaries**: exact matches use word-boundary detection, so punctuation/hyphens count as boundaries (e.g. `"PO"` matches inside "PO-2024-555" since the hyphen separates it).
- **Password-protected / encrypted PDFs** are not automatically decrypted — these will show as `Error`.
- **Scanned/image-only PDFs** (no embedded text layer) will not yield matches — this tool does not perform OCR.
- **Legacy `.doc` files** (old binary Word format) are not supported — only `.docx`. Convert to `.docx` first if needed.
- Matching is a simple substring search (not whole-word or regex).

## Files

- `pdf_keyword_scanner.py` — the application (single file, no other project files needed)
