# -*- coding: utf-8 -*-
"""
SYS/SUB (Annotation) → Excel Match (Subsystem Only, try right then left) → Fix Subsystem Description (Optional)

CSV headers:
  - pdf filename
  - pdf sysno
  - pdf subno
  - pdf description            (after update if any; else original)
  - status                     (match | mismatch) — based on BEFORE-update comparison
  - excel sysno                (from Excel System No. if present; else derived from Excel Subsystem No.)
  - excel description
  - update file path           (only when updated; else 'None')

Usage:
  pip install pymupdf pandas openpyxl
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import pandas as pd

# =========================================================
# Normalization helpers
# =========================================================

def norm_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = " ".join(s.split())
    return s.upper()

def norm_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("-\n", "")
    s = s.replace("\r", "\n").replace("\n", " ")
    s = " ".join(s.split())
    return s.strip()

def _normalize_code_piece(piece: str) -> str:
    if not piece:
        return ""
    piece = piece.replace("⁄", "/")
    piece = " ".join(piece.split())
    return piece.strip()

# ---- NEW: derive System No. when Excel has no explicit System No. column
_DERIVE_TAIL_RE = re.compile(r"([/_-])\d{1,3}$")

def derive_sysno_from_sub(sub: str) -> str:
    """
    Derive a System No. from a Subsystem No. by removing only the *final*
    numeric suffix (e.g., '-01', '-02') **if and only if** there are at least
    TWO numeric tokens in the whole code.

    Examples:
      '772-P-001-02'   -> '772-P-001'     (two numeric tokens: 001, 02)
      '775-E-026-01'   -> '775-E-026'     (two numeric tokens: 026, 01)
      '775-E-026'      -> '775-E-026'     (one numeric token only -> do not strip)
      '772-P-001/02'   -> '772-P-001'     (handles slash)
      '772-P-001_02'   -> '772-P-001'     (handles underscore)
    """
    if not sub:
        return ""
    s = str(sub).strip()

    # Normalize common separators to '-' for tokenization (keep original text intact for return)
    sep_normalized = s.replace("/", "-").replace("_", "-")
    tokens = sep_normalized.split("-")

    # Count numeric tokens (1–3 digits is common, but allow more just in case)
    numeric_idxs = [i for i, t in enumerate(tokens) if re.fullmatch(r"\d+", t)]
    if len(numeric_idxs) < 2:
        # Only one (or zero) numeric token in the whole code → DO NOT STRIP
        return s

    # If the last token is purely digits, drop it; else keep as-is
    if re.fullmatch(r"\d+", tokens[-1]):
        # Rebuild with original separators replaced back to '-'
        base_tokens = tokens[:-1]
        derived = "-".join(base_tokens)

        # Try to reconstruct into the original separator scheme (optional).
        # Since CSV display is fine with '-', we can return derived as-is.
        return derived

    return s
# =========================================================
# SUBJECT patterns (annotation) — prefer SUBSYSTEM DESCRIPTION for updates
# =========================================================

SUBJECT_SYS_SUB_RE = re.compile(
    r"""SYSTEM\s*[\/⁄]?\s*SUB\s*-?\s*SYSTEM\s*NO\b""",
    re.IGNORECASE,
)

# Strict SUBSYSTEM DESCRIPTION (preferred)
SUBJECT_SUBSYS_DESC_RE = re.compile(r"""SUBSYSTEM\s*DESCRIPTION\b""", re.IGNORECASE)

# SYSTEM DESCRIPTION (fallback read if SUBSYSTEM not found)
SUBJECT_SYS_DESC_RE = re.compile(r"""SYSTEM\s*DESCRIPTION\b""", re.IGNORECASE)

# =========================================================
# Split "A / B" robustly
# =========================================================

SYS_SUBSYS_SPLIT_RE = re.compile(
    r"""^\s*
    (?P<sys>[A-Za-z0-9][A-Za-z0-9._\-\s]*[A-Za-z0-9])
    \s*[\/⁄]\s*
    (?P<sub>[A-Za-z0-9][A-Za-z0-9._\-\s]*[A-Za-z0-9])
    \s*$""",
    re.VERBOSE,
)

def split_sys_sub(raw_value: str) -> Tuple[str, str, str]:
    """
    Returns (raw_clean, sys_left, subsys_right) if 'A / B' exists; else ('','','').
    """
    if not raw_value:
        return "", "", ""
    val = " ".join(str(raw_value).replace("⁄", "/").split())

    m = SYS_SUBSYS_SPLIT_RE.match(val)
    if m:
        sys_part = _normalize_code_piece(m.group("sys"))
        sub_part = _normalize_code_piece(m.group("sub"))
        return f"{sys_part} / {sub_part}", sys_part, sub_part

    if "/" in val:
        left, right = val.split("/", 1)
        left = _normalize_code_piece(left)
        right = _normalize_code_piece(right)
        if left and right:
            return f"{left} / {right}", left, right

    return "", "", ""

def expand_subsys_if_suffix(sys_left: str, subsys_right: str) -> str:
    if subsys_right and re.fullmatch(r"\d{1,3}", subsys_right):
        return f"{sys_left}-{subsys_right}"
    return subsys_right

# =========================================================
# Annotation readers / writers
# =========================================================

def _iter_page_annots(page: fitz.Page):
    """
    Safe iterator over annotations of a page.
    """
    try:
        annot = page.first_annot
    except Exception:
        annot = None
    if annot is None:
        try:
            it = page.annots()
            if it is not None:
                for a in it:
                    yield a
            return
        except Exception:
            return
    while annot:
        yield annot
        annot = annot.next

def _get_annot_subject_and_content(annot) -> Tuple[str, str]:
    """
    Extract subject and content from an annotation (Typewriter / FreeText).
    """
    info = {}
    try:
        info = annot.info or {}
    except Exception:
        info = {}
    subject = info.get("subject") or info.get("title") or ""
    content = info.get("content") or ""
    subject = " ".join(str(subject).split()).strip()
    content = "\n".join([ln.rstrip() for ln in str(content).splitlines()]).strip()
    return subject, content

def extract_sys_sub_by_subject(pdf_path: Path, max_pages: int = 600) -> Tuple[bool, str, str, str]:
    """
    Search annotations for Subject ≈ SYSTEM / SUBSYSTEM NO and read the value.
    """
    with fitz.open(pdf_path) as doc:
        end = min(len(doc), max_pages)
        any_label = False
        for i in range(end):
            page = doc[i]
            for annot in _iter_page_annots(page):
                try:
                    subject, content = _get_annot_subject_and_content(annot)
                except Exception:
                    continue
                if not subject:
                    continue
                if SUBJECT_SYS_SUB_RE.search(subject):
                    any_label = True
                    raw_clean, sys_left, subsys_right = split_sys_sub(content)
                    if raw_clean and sys_left and subsys_right:
                        subsys_right = expand_subsys_if_suffix(sys_left, subsys_right)
                        return True, sys_left, subsys_right, raw_clean
                    else:
                        return True, "", "", content.strip()
        return (any_label, "", "", "") if any_label else (False, "", "", "")

def find_all_desc_annots(doc: fitz.Document, max_pages: int = 600, prefer_subsystem: bool = True) -> List[Tuple[int, object]]:
    """
    Return list of (page_index, annot) matching description subjects.
    Prefer 'SUBSYSTEM DESCRIPTION'; if none, allow 'SYSTEM DESCRIPTION' as fallback.
    """
    end = min(len(doc), max_pages)
    hits_sub: List[Tuple[int, object]] = []
    hits_sys: List[Tuple[int, object]] = []
    for i in range(end):
        page = doc[i]
        for annot in _iter_page_annots(page):
            try:
                subject, _content = _get_annot_subject_and_content(annot)
            except Exception:
                continue
            if subject:
                if SUBJECT_SUBSYS_DESC_RE.search(subject):
                    hits_sub.append((i, annot))
                elif SUBJECT_SYS_DESC_RE.search(subject):
                    hits_sys.append((i, annot))
    if prefer_subsystem and hits_sub:
        return hits_sub
    return hits_sub if hits_sub else hits_sys

def extract_preferred_desc_by_subject(pdf_path: Path, max_pages: int = 600, prefer_subsystem: bool = True) -> str:
    with fitz.open(pdf_path) as doc:
        hits = find_all_desc_annots(doc, max_pages=max_pages, prefer_subsystem=prefer_subsystem)
        if hits:
            _pi, annot = hits[0]
            _subj, content = _get_annot_subject_and_content(annot)
            return " ".join(content.split()).strip()
    return ""

def update_all_preferred_desc_annotations(pdf_path: Path, new_text: str, out_path: Path, max_pages: int = 600, prefer_subsystem: bool = True) -> bool:
    """
    Update ALL description annotations (prefer 'SUBSYSTEM DESCRIPTION'; fallback to 'SYSTEM DESCRIPTION' only if none).
    Save to out_path. Returns True if at least one updated & saved.
    """
    changed = False
    with fitz.open(pdf_path) as doc:
        hits = find_all_desc_annots(doc, max_pages=max_pages, prefer_subsystem=prefer_subsystem)
        if not hits:
            return False

        for (_pi, annot) in hits:
            try:
                if hasattr(annot, "set_info"):
                    annot.set_info({"content": new_text})
                    try:
                        annot.set_info({"richtext": new_text})
                    except Exception:
                        pass
                else:
                    try:
                        if hasattr(annot, "set_contents"):
                            annot.set_contents(new_text)
                    except Exception:
                        pass
                    try:
                        if hasattr(annot, "setRichText"):
                            annot.setRichText(new_text)  # type: ignore
                    except Exception:
                        pass
                try:
                    annot.update()
                except Exception:
                    pass
                changed = True
            except Exception:
                continue

        if changed:
            out_path.parent.mkdir(parents=True, exist_ok=True)
            try:
                doc.save(out_path, deflate=True)
            except Exception:
                doc.save(out_path)
    return changed

# =========================================================
# Fallback (line-based) for description (if annotation missing)
# =========================================================

def lines_from_page(page: fitz.Page) -> List[str]:
    text = page.get_text("text")
    text = text.replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n")]
    return [ln for ln in lines if ln]

def extract_system_description_linebased(pdf_path: Path, max_pages: int = 600) -> str:
    flags = re.IGNORECASE
    sys_desc_variants = [
        r"SYSTEM\s*DESCRIPTION",
        r"SUBSYSTEM\s*DESCRIPTION",
        r"SYSTEM\s*/\s*SUBSYSTEM\s*DESCRIPTION",
    ]
    sys_desc_pats = [
        re.compile(rf"(?:{v})\s*[:\-–—]?\s*(.+)$", flags) for v in sys_desc_variants
    ] + [re.compile(v, flags) for v in sys_desc_variants]

    def extract_labeled_value_from_lines(lines: List[str]) -> Optional[str]:
        for i, line in enumerate(lines):
            for pat in sys_desc_pats:
                m = pat.search(line)
                if not m:
                    continue
                if m.lastindex:
                    val = m.group(m.lastindex).strip()
                    if val:
                        return val
                tail = line[m.end():].strip(" :\t-—")
                if tail:
                    return tail
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    j += 1
                if j < len(lines):
                    nxt = lines[j].strip(" :\t-—")
                    if nxt:
                        return nxt
        return None

    with fitz.open(pdf_path) as doc:
        end = min(len(doc), max_pages)
        for i in range(end):
            lines = lines_from_page(doc[i])
            val = extract_labeled_value_from_lines(lines)
            if val:
                return val.strip()
    return ""

def extract_preferred_description(pdf_path: Path, max_pages: int = 600, prefer_subsystem: bool = True) -> str:
    v = extract_preferred_desc_by_subject(pdf_path, max_pages=max_pages, prefer_subsystem=prefer_subsystem)
    if v:
        return v
    return extract_system_description_linebased(pdf_path, max_pages=max_pages)

# =========================================================
# Excel matching (Subsystem Only + fuzzy headers + prefer both-match)
# =========================================================

def _normalize_header(h: str) -> str:
    """Uppercase and remove non-alnum to make fuzzy compares."""
    return re.sub(r"[^A-Z0-9]+", "", str(h).upper())

EXCEL_HEADER_ALIAS = {
    "sub_no":  {"SUBSYSTEMNO", "SUBSYSTEMNUMBER", "SUBNO", "SUBSYSTEM", "SUBSYSNO"},
    "sub_desc":{"SUBSYSTEMDESCRIPTION", "SUBSYSDESCRIPTION", "DESCRIPTION", "SUBDESC"},
    "sys_no":  {"SYSTEMNO", "SYSTEMNUMBER", "SYSNO", "SYSTEM"},
}

def _find_excel_columns(df: pd.DataFrame) -> Tuple[str, str, Optional[str]]:
    """
    Return the actual column names for:
      (subsystem_no_col, subsystem_desc_col, system_no_col_or_None)
    using fuzzy header matching.
    Raises ValueError if required columns not found.
    """
    norm_map = {_normalize_header(c): c for c in df.columns}

    # Subsystem No.
    sub_no_col = None
    for key in EXCEL_HEADER_ALIAS["sub_no"]:
        if key in norm_map:
            sub_no_col = norm_map[key]; break
    if sub_no_col is None:
        for c in df.columns:
            if c.strip().lower() in {"subsystem no.", "subsystem no", "subsystem no"}:
                sub_no_col = c; break
    if sub_no_col is None:
        raise ValueError("Excel missing required 'Subsystem No.' column.")

    # Subsystem Description
    sub_desc_col = None
    for key in EXCEL_HEADER_ALIAS["sub_desc"]:
        if key in norm_map:
            sub_desc_col = norm_map[key]; break
    if sub_desc_col is None:
        for c in df.columns:
            if c.strip().lower() in {"subsystem description", "sub-system description"}:
                sub_desc_col = c; break
    if sub_desc_col is None:
        raise ValueError("Excel missing required 'Subsystem Description' column.")

    # Optional System No.
    sys_no_col = None
    for key in EXCEL_HEADER_ALIAS["sys_no"]:
        if key in norm_map:
            sys_no_col = norm_map[key]; break
    if sys_no_col is None:
        for c in df.columns:
            if c.strip().lower() in {"system no.", "system no", "system number"}:
                sys_no_col = c; break

    return sub_no_col, sub_desc_col, sys_no_col

class ExcelIndex:
    def __init__(self, df: pd.DataFrame):
        sub_no_col, sub_desc_col, sys_no_col = _find_excel_columns(df)
        self.sub_no_col = sub_no_col
        self.sub_desc_col = sub_desc_col
        self.sys_no_col = sys_no_col  # may be None

        self.df = df
        self.df["_N_SUB"] = self.df[self.sub_no_col].apply(norm_key)
        if self.sys_no_col:
            self.df["_N_SYS"] = self.df[self.sys_no_col].apply(norm_key)
        else:
            self.df["_N_SYS"] = ""

        # Index by Subsystem No. (many rows can share same)
        self.map_sub: Dict[str, List[int]] = {}
        for idx, r in self.df.iterrows():
            k = r["_N_SUB"]
            self.map_sub.setdefault(k, []).append(idx)

    def find_row_best(self, sys_left: str, subsys_right: str) -> Tuple[Optional[int], str]:
        """
        Try RIGHT (subsys_right) first; if not found, try LEFT (sys_left).
        If multiple rows share a Subsystem No. and System No. column exists,
        prefer the row whose System No. equals sys_left.
        """
        k_r = norm_key(subsys_right)
        k_l = norm_key(sys_left)

        def choose_by_sys(candidates: List[int]) -> Optional[int]:
            if not candidates:
                return None
            if self.sys_no_col and k_l:
                for idx in candidates:
                    if norm_key(self.df.loc[idx, self.sys_no_col]) == k_l:
                        return idx
            return candidates[0]

        # 1) Subsystem match using RIGHT token
        cand_r = self.map_sub.get(k_r, [])
        if cand_r:
            return choose_by_sys(cand_r), ""

        # 2) Fallback: use LEFT token
        cand_l = self.map_sub.get(k_l, [])
        if cand_l:
            return choose_by_sys(cand_l), ""

        return None, "Subsystem No. not in Excel (tried right and left from PDF)"

def load_excel_index(xlsx: Path) -> ExcelIndex:
    df = pd.read_excel(xlsx, engine="openpyxl")
    return ExcelIndex(df)

# =========================================================
# Processing
# =========================================================

def process_folder(pdf_folder: Path,
                   excel_file: Path,
                   max_pages: int,
                   update_pdfs: bool,
                   output_folder: Optional[Path],
                   log_fn=print) -> pd.DataFrame:

    xidx = load_excel_index(excel_file)
    pdfs = sorted([p for p in pdf_folder.glob("*.pdf") if p.is_file()])
    rows = []

    if not pdfs:
        log_fn(f"[WARN] No PDFs found in: {pdf_folder}")
        return pd.DataFrame(columns=[
            "pdf filename", "pdf sysno", "pdf subno", "pdf description",
            "status", "excel sysno", "excel description", "update file path"
        ])

    for i, pdf in enumerate(pdfs, 1):
        out_file = "None"
        try:
            # 1) Read SYS/SUB from annotation
            label_found, sys_left, sub_right, raw_value = extract_sys_sub_by_subject(pdf, max_pages=max_pages)

            # 2) Read PDF Description (prefer SUBSYSTEM DESCRIPTION)
            pdf_desc_before = extract_preferred_description(pdf, max_pages=max_pages, prefer_subsystem=True)

            pdf_sysno = sys_left if sys_left else ""
            pdf_subno = sub_right if sub_right else ""

            excel_sysno = ""
            excel_desc = ""
            status = "mismatch"  # 'match' only if BEFORE equals Excel
            pdf_desc_after = pdf_desc_before

            if label_found and (pdf_subno or pdf_sysno):
                # Excel match (try RIGHT then LEFT)
                row_idx, _reason = xidx.find_row_best(pdf_sysno, pdf_subno)
                if row_idx is not None:
                    # Pull Excel description
                    excel_desc = str(xidx.df.loc[row_idx, xidx.sub_desc_col] or "").strip()
                    # Always patch excel_sysno: from Excel System No. if exists; else derive from Excel Subsystem No.
                    if xidx.sys_no_col:
                        excel_sysno = str(xidx.df.loc[row_idx, xidx.sys_no_col] or "").strip()
                    else:
                        # derive from the matched Excel Subsystem No.
                        excel_sub_val = str(xidx.df.loc[row_idx, xidx.sub_no_col] or "").strip()
                        excel_sysno = derive_sysno_from_sub(excel_sub_val)

                    # Compare BEFORE update
                    if norm_text(pdf_desc_before) == norm_text(excel_desc):
                        status = "match"
                        pdf_desc_after = pdf_desc_before
                        out_file = "None"
                        log_fn(f"[{i}/{len(pdfs)}] {pdf.name} | MATCH (no change)")
                    else:
                        # mismatch → update if enabled
                        if update_pdfs and output_folder:
                            updated_name = pdf.with_suffix("").name + ".pdf"
                            out_path = output_folder / updated_name
                            changed = update_all_preferred_desc_annotations(
                                pdf, excel_desc, out_path, max_pages=max_pages, prefer_subsystem=True
                            )
                            if changed:
                                out_file = str(out_path)
                                pdf_desc_after = extract_preferred_description(out_path, max_pages=max_pages, prefer_subsystem=True)
                                log_fn(f"[{i}/{len(pdfs)}] {pdf.name} | MISMATCH → UPDATED → {out_file}")
                            else:
                                out_file = "None"
                                log_fn(f"[{i}/{len(pdfs)}] {pdf.name} | MISMATCH → UPDATE FAILED")
                        else:
                            log_fn(f"[{i}/{len(pdfs)}] {pdf.name} | MISMATCH (dry-run)")
                else:
                    # Not found in Excel
                    log_fn(f"[{i}/{len(pdfs)}] {pdf.name} | Subsystem not in Excel (tried right & left) → treated as mismatch")
            else:
                log_fn(f"[{i}/{len(pdfs)}] {pdf.name} | SYS/SUB annotation missing or invalid → treated as mismatch")

            # Write CSV row
            rows.append({
                "pdf filename": pdf.name,
                "pdf sysno": pdf_sysno,
                "pdf subno": pdf_subno,
                "pdf description": pdf_desc_after,
                "status": status,                    # 'match' only if BEFORE equal; else 'mismatch'
                "excel sysno": excel_sysno,          # ALWAYS populated after Excel row match
                "excel description": excel_desc,
                "update file path": out_file,        # actual path only if updated; else 'None'
            })

        except Exception as e:
            log_fn(f"[ERR] {pdf.name}: {e}")
            rows.append({
                "pdf filename": pdf.name,
                "pdf sysno": "",
                "pdf subno": "",
                "pdf description": "",
                "status": "mismatch",
                "excel sysno": "",
                "excel description": "",
                "update file path": "None",
            })

    return pd.DataFrame(rows)

# =========================================================
# GUI
# =========================================================

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Annotation → Excel (Subsystem) → Fix SUBSYSTEM DESCRIPTION")
        self.geometry("1140x760")
        self.resizable(True, True)

        self.pdf_folder_var = tk.StringVar()
        self.excel_file_var = tk.StringVar()
        self.max_pages_var = tk.IntVar(value=600)
        self.save_csv_var = tk.BooleanVar(value=True)
        self.csv_path_var = tk.StringVar()
        # Default ON so mismatches are fixed
        self.update_pdfs_var = tk.BooleanVar(value=True)
        self.output_folder_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 8, 'pady': 6}
        frm = ttk.Frame(self); frm.pack(fill="x", **pad)

        # PDF Folder
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        ttk.Label(row, text="PDF Folder:").pack(side="left", padx=(0,6))
        ttk.Entry(row, textvariable=self.pdf_folder_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_pdf_folder).pack(side="left")

        # Excel File
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        ttk.Label(row, text="Excel File:").pack(side="left", padx=(0,6))
        ttk.Entry(row, textvariable=self.excel_file_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_excel_file).pack(side="left")

        # Max pages
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        ttk.Label(row, text="Max pages to scan per PDF:").pack(side="left", padx=(0,6))
        ttk.Spinbox(row, from_=1, to=50, textvariable=self.max_pages_var, width=6).pack(side="left", padx=6)

        # Save CSV
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        ttk.Checkbutton(row, text="Save results to CSV", variable=self.save_csv_var).pack(side="left")
        ttk.Entry(row, textvariable=self.csv_path_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Save As…", command=self.pick_csv).pack(side="left")

        # Update PDFs
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        ttk.Checkbutton(row, text="Update PDF when description mismatches Excel", variable=self.update_pdfs_var).pack(side="left")
        ttk.Label(row, text="Output Folder:").pack(side="left", padx=(12,6))
        ttk.Entry(row, textvariable=self.output_folder_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_output_folder).pack(side="left")

        # Run + Close
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        self.run_btn = ttk.Button(row, text="Run", command=self.on_run)
        self.run_btn.pack(side="left")
        ttk.Button(row, text="Close", command=self.destroy).pack(side="left", padx=6)

        # Log
        self.log_box = ScrolledText(self, height=24, wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)

    def log(self, msg: str):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def pick_pdf_folder(self):
        d = filedialog.askdirectory(title="Select PDF Folder")
        if d:
            self.pdf_folder_var.set(d)

    def pick_excel_file(self):
        f = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if f:
            self.excel_file_var.set(f)
            p = Path(f)
            if not self.csv_path_var.get():
                self.csv_path_var.set(str(p.with_name(p.stem + "_annots_match_report.csv")))

    def pick_output_folder(self):
        d = filedialog.askdirectory(title="Select Output (Updated PDFs) Folder")
        if d:
            self.output_folder_var.set(d)

    def pick_csv(self):
        f = filedialog.asksaveasfilename(
            title="Save Results CSV As",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if f:
            self.csv_path_var.set(f)

    def _validate(self) -> Optional[str]:
        pdf_folder = Path(self.pdf_folder_var.get().strip())
        excel_file = Path(self.excel_file_var.get().strip())
        max_pages = int(self.max_pages_var.get())
        update_pdfs = self.update_pdfs_var.get()
        output_folder_text = self.output_folder_var.get().strip()

        if not pdf_folder.exists() or not pdf_folder.is_dir():
            return "Please pick a valid PDF folder."
        if not excel_file.exists() or not excel_file.is_file():
            return "Please pick a valid Excel file."
        if max_pages < 1:
            return "Max pages must be at least 1."
        if self.save_csv_var.get() and not self.csv_path_var.get().strip():
            return "Please choose a CSV path or uncheck 'Save results to CSV'."
        if update_pdfs:
            if not output_folder_text:
                return "Please choose an Output Folder to save updated PDFs."
            out = Path(output_folder_text)
            if not out.exists():
                try:
                    out.mkdir(parents=True, exist_ok=True)
                except Exception:
                    return "Cannot create the Output Folder. Please choose a valid path."
        return None

    def on_run(self):
        err = self._validate()
        if err:
            messagebox.showerror("Input Error", err)
            return

        pdf_folder = Path(self.pdf_folder_var.get().strip())
        excel_file = Path(self.excel_file_var.get().strip())
        max_pages = int(self.max_pages_var.get())
        save_csv = self.save_csv_var.get()
        csv_path = Path(self.csv_path_var.get().strip()) if save_csv else None
        update_pdfs = self.update_pdfs_var.get()
        output_folder = Path(self.output_folder_var.get().strip()) if self.output_folder_var.get().strip() else None

        self.run_btn.config(state="disabled")
        self.log_box.delete("1.0", "end")
        self.log("Starting… Read SYS/SUB (annotation) → Match Excel (try right then left) → Fix SUBSYSTEM DESCRIPTION if mismatched.\n")

        try:
            df = process_folder(pdf_folder, excel_file, max_pages, update_pdfs, output_folder, log_fn=self.log)
            if save_csv and not df.empty:
                csv_path.parent.mkdir(parents=True, exist_ok=True)
                df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                self.log(f"\n🧾 Results saved to CSV: {csv_path}")
            self.log("\nDone.")
            messagebox.showinfo("Done", "Completed. Check the log and CSV (if saved).")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state="normal")

# =========================================================
# Entrypoint
# =========================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()