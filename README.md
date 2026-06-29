# PDF & Word Keyword Scanner

A desktop GUI tool that scans every PDF (`.pdf`) and Word document (`.docx`) in a folder for one or more keywords, and shows the results in a sortable table.

## Features

- **Directory picker** — browse to any folder
- **Multiple keywords** — comma-separated (e.g. `invoice, overdue, PO-2024`)
- **Exact / whole-word matching** — wrap a keyword in double quotes (e.g. `"PO"`) to match it as a whole word only, instead of as a substring. Useful for short keywords that would otherwise false-match inside longer words (e.g. `"PO"` won't match inside "REPORT" or "PORT", but unquoted `po` would).
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
python3 pdf_keyword_scanner.py
```

1. Click **Browse...** and select the folder containing your PDFs / Word docs.
2. Type one or more keywords into the **Keyword(s)** field, separated by commas.
   - Unquoted keyword → substring match, e.g. `invoice` matches "invoice", "Invoicing", "reinvoice".
   - `"quoted keyword"` → exact whole-word match, e.g. `"PO"` matches the standalone word "PO" but not "REPORT" or "PORT".
   - Example: `invoice, "PO", overdue` searches for "invoice" anywhere, the exact word "PO", and "overdue" anywhere.
   - Quoted keywords may contain commas (e.g. `"item, version 2"`) and spaces.
3. (Optional) Tick **Case sensitive** or **Include subfolders**.
4. Click **Scan** (or press Enter in the keyword field).
5. Review results in the table — one row per file, one column per keyword. Exact-match columns show their header in quotes.
6. (Optional) Click **Export CSV** to save the results.

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
