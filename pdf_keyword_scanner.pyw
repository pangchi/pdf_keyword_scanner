#!/usr/bin/env python3
"""
PDF & Word Keyword Scanner
---------------------------
- Pick a directory
- Enter one or more keywords (comma-separated)
- Scans every PDF (.pdf) and Word document (.docx) in the directory
  (non-recursive by default, optional recursive)
- Shows a table: filename, file type, per-keyword found/not found, and overall match
"""

import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox


def _ensure_package(import_name, pip_name=None):
    """Install a package automatically if it isn't already available."""
    pip_name = pip_name or import_name
    try:
        __import__(import_name)
        return
    except ImportError:
        pass

    print(f"{pip_name} not found - installing automatically...")
    cmd = [sys.executable, "-m", "pip", "install", pip_name]
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError:
        # Some systems (e.g. Debian/Ubuntu-managed Python) require this flag
        subprocess.check_call(cmd + ["--break-system-packages"])

    print(f"{pip_name} installed successfully.")


_ensure_package("pypdf")
_ensure_package("docx", pip_name="python-docx")

from pypdf import PdfReader
from pypdf.errors import PdfReadError
from docx import Document as DocxDocument


class PDFKeywordScanner(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF & Word Keyword Scanner")
        self.geometry("900x550")
        self.minsize(700, 400)

        self.directory = tk.StringVar()
        self.keywords_raw = tk.StringVar()
        self.case_sensitive = tk.BooleanVar(value=False)
        self.recursive = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="Ready.")

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        # Directory row
        ttk.Label(top, text="Directory:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.directory).grid(
            row=0, column=1, sticky="ew", padx=4
        )
        ttk.Button(top, text="Browse...", command=self.browse_directory).grid(
            row=0, column=2, padx=4
        )
        top.columnconfigure(1, weight=1)

        # Keywords row
        ttk.Label(top, text="Keyword(s):").grid(row=1, column=0, sticky="w")
        kw_entry = ttk.Entry(top, textvariable=self.keywords_raw)
        kw_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=(6, 0))
        kw_entry.bind("<Return>", lambda e: self.start_scan())
        ttk.Label(top, text='(comma-separated; use "exact" for whole-word match)').grid(
            row=1, column=2, sticky="w", pady=(6, 0)
        )

        # Options row
        opts = ttk.Frame(self)
        opts.pack(fill="x", padx=8)
        ttk.Checkbutton(
            opts, text="Case sensitive", variable=self.case_sensitive
        ).pack(side="left")
        ttk.Checkbutton(
            opts, text="Include subfolders", variable=self.recursive
        ).pack(side="left", padx=12)
        self.scan_btn = ttk.Button(opts, text="Scan", command=self.start_scan)
        self.scan_btn.pack(side="right")
        ttk.Button(opts, text="Export CSV", command=self.export_csv).pack(
            side="right", padx=8
        )

        # Results table
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=8, pady=8)

        self.tree = ttk.Treeview(table_frame, show="headings")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Status bar
        status = ttk.Label(
            self, textvariable=self.status_text, anchor="w", relief="sunken"
        )
        status.pack(fill="x", side="bottom")

        self._last_results = []  # cache for CSV export
        self._last_keywords = []

    # ------------------------------------------------------------- actions
    def browse_directory(self):
        path = filedialog.askdirectory(title="Select directory containing PDFs / Word docs")
        if path:
            self.directory.set(path)

    def get_keywords(self):
        """
        Parse the keyword field into a list of (text, is_exact) tuples.

        Keywords wrapped in double quotes (e.g. "PO") are matched as whole
        words only. Unquoted keywords (e.g. invoice) are matched as a plain
        substring anywhere in the text. Quoted keywords may contain commas
        and escaped quotes (use "" for a literal " inside a quoted keyword).
        """
        raw = self.keywords_raw.get()
        if not raw.strip():
            return []

        keywords = []
        pos = 0
        n = len(raw)
        while pos < n:
            # Skip separators and whitespace between keywords
            while pos < n and raw[pos] in " \t,":
                pos += 1
            if pos >= n:
                break

            if raw[pos] == '"':
                # Quoted -> exact (whole-word) match. Read until the closing
                # quote, treating "" as an escaped literal quote character.
                end = pos + 1
                buf = []
                while end < n:
                    if raw[end] == '"':
                        if end + 1 < n and raw[end + 1] == '"':
                            buf.append('"')
                            end += 2
                            continue
                        break
                    buf.append(raw[end])
                    end += 1
                text = "".join(buf).strip()
                if text:
                    keywords.append((text, True))
                pos = end + 1
            else:
                # Unquoted -> substring match. Read until the next comma.
                end = raw.find(",", pos)
                if end == -1:
                    end = n
                text = raw[pos:end].strip()
                if text:
                    keywords.append((text, False))
                pos = end

        return keywords

    def start_scan(self):
        directory = self.directory.get().strip()
        keywords = self.get_keywords()

        if not directory or not os.path.isdir(directory):
            messagebox.showerror("Error", "Please select a valid directory.")
            return
        if not keywords:
            messagebox.showerror("Error", "Please enter at least one keyword.")
            return

        # Deduplicate by keyword text (case-sensitive as typed) so quoted and
        # unquoted versions of the same word don't collide on the same
        # results column. First occurrence wins.
        seen = set()
        deduped = []
        for text, is_exact in keywords:
            if text in seen:
                continue
            seen.add(text)
            deduped.append((text, is_exact))
        keywords = deduped

        self.scan_btn.config(state="disabled")
        self.status_text.set("Scanning...")
        self._clear_table()

        # Run in background thread so the GUI doesn't freeze on large folders
        thread = threading.Thread(
            target=self._scan_worker, args=(directory, keywords), daemon=True
        )
        thread.start()

    SUPPORTED_EXTENSIONS = (".pdf", ".docx")

    def _find_documents(self, directory):
        if self.recursive.get():
            paths = []
            for root, _, files in os.walk(directory):
                for f in files:
                    if f.lower().endswith(self.SUPPORTED_EXTENSIONS):
                        paths.append(os.path.join(root, f))
            return sorted(paths)
        else:
            return sorted(
                os.path.join(directory, f)
                for f in os.listdir(directory)
                if f.lower().endswith(self.SUPPORTED_EXTENSIONS)
            )

    def _scan_worker(self, directory, keywords):
        doc_paths = self._find_documents(directory)
        results = []

        if not doc_paths:
            self.after(
                0,
                self._scan_finished,
                [],
                keywords,
                "No PDF or Word files found.",
            )
            return

        case_sensitive = self.case_sensitive.get()

        # Pre-compile a whole-word regex for each "exact" keyword.
        # Unquoted keywords are matched as plain substrings (faster, and
        # behaves the same as before this feature was added).
        flags = 0 if case_sensitive else re.IGNORECASE
        compiled = []
        for text, is_exact in keywords:
            if is_exact:
                pattern = re.compile(r"\b" + re.escape(text) + r"\b", flags)
                compiled.append((text, is_exact, pattern))
            else:
                needle = text if case_sensitive else text.lower()
                compiled.append((text, is_exact, needle))

        for i, path in enumerate(doc_paths, start=1):
            display_name = os.path.relpath(path, directory)
            ext = os.path.splitext(path)[1].lower()
            file_type = "PDF" if ext == ".pdf" else "Word"
            self.after(
                0,
                self.status_text.set,
                f"Scanning {i}/{len(doc_paths)}: {display_name}",
            )

            row = {"filename": display_name, "type": file_type, "error": ""}
            try:
                text = self._extract_text(path, ext)
                text_for_substring = text if case_sensitive else text.lower()
                for kw_text, is_exact, matcher in compiled:
                    if is_exact:
                        found = bool(matcher.search(text))
                    else:
                        found = matcher in text_for_substring
                    row[kw_text] = "Yes" if found else "No"
            except Exception as e:
                for kw_text, _, _ in compiled:
                    row[kw_text] = "Error"
                row["error"] = str(e)

            results.append(row)

        self.after(0, self._scan_finished, results, keywords, None)

    @staticmethod
    def _extract_text(path, ext):
        if ext == ".pdf":
            return PDFKeywordScanner._extract_text_pdf(path)
        elif ext == ".docx":
            return PDFKeywordScanner._extract_text_docx(path)
        else:
            raise ValueError(f"Unsupported file type: {ext}")

    @staticmethod
    def _extract_text_pdf(path):
        reader = PdfReader(path)
        text_parts = []
        for page in reader.pages:
            try:
                text_parts.append(page.extract_text() or "")
            except Exception:
                continue
        return "\n".join(text_parts)

    @staticmethod
    def _extract_text_docx(path):
        doc = DocxDocument(path)
        text_parts = [p.text for p in doc.paragraphs]

        # Also search inside tables, since body text alone misses table content
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text_parts.append(cell.text)

        return "\n".join(text_parts)

    def _scan_finished(self, results, keywords, message):
        self._last_results = results
        self._last_keywords = keywords
        self._populate_table(results, keywords)
        self.scan_btn.config(state="normal")
        if message:
            self.status_text.set(message)
        else:
            n_errors = sum(1 for r in results if r.get("error"))
            msg = f"Done. Scanned {len(results)} file(s)."
            if n_errors:
                msg += f" {n_errors} file(s) had errors (see Error column tooltip via export)."
            self.status_text.set(msg)

    # --------------------------------------------------------------- table
    def _clear_table(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    def _populate_table(self, results, keywords):
        self._clear_table()
        kw_ids = [kw_text for kw_text, _ in keywords]
        columns = ["filename", "type"] + kw_ids + ["any_match"]
        self.tree["columns"] = columns

        self.tree.heading("filename", text="Filename")
        self.tree.column("filename", width=300, anchor="w")
        self.tree.heading("type", text="Type")
        self.tree.column("type", width=60, anchor="center")
        for kw_text, is_exact in keywords:
            header = f'"{kw_text}"' if is_exact else kw_text
            self.tree.heading(kw_text, text=header)
            self.tree.column(kw_text, width=100, anchor="center")
        self.tree.heading("any_match", text="Any Match")
        self.tree.column("any_match", width=90, anchor="center")

        for row in results:
            values = [row["filename"], row.get("type", "")]
            any_yes = False
            for kw_id in kw_ids:
                v = row.get(kw_id, "")
                if v == "Yes":
                    any_yes = True
                values.append(v)
            values.append("Yes" if any_yes else "No")
            self.tree.insert("", "end", values=values)

    # -------------------------------------------------------------- export
    def export_csv(self):
        if not self._last_results:
            messagebox.showinfo("Nothing to export", "Run a scan first.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save results as CSV",
        )
        if not path:
            return

        import csv

        keywords = self._last_keywords
        kw_ids = [kw_text for kw_text, _ in keywords]
        kw_headers = [f'"{t}"' if exact else t for t, exact in keywords]
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Filename", "Type"] + kw_headers + ["Any Match", "Error"])
            for row in self._last_results:
                any_yes = "Yes" if any(row.get(k) == "Yes" for k in kw_ids) else "No"
                writer.writerow(
                    [row["filename"], row.get("type", "")]
                    + [row.get(k, "") for k in kw_ids]
                    + [any_yes, row.get("error", "")]
                )

        messagebox.showinfo("Exported", f"Results saved to:\n{path}")


if __name__ == "__main__":
    app = PDFKeywordScanner()
    app.mainloop()
