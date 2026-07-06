# -*- coding: utf-8 -*-
"""
Excel-driven PDF page replacement (grouped per PDF so multiple pages are all applied correctly).

Template columns (case-insensitive; flexible names allowed):
  Required:
    - PDF / PDF_Name / PDF_Path / Input_PDF       -> target PDF (name or full path)
    - Page / Page_No / Page_Number                -> 1-based page number to replace
    - Latest / Latest_Page_Path / Replacement_PDF -> FULL PATH to SINGLE-PAGE PDF (exact path in Excel)
  Optional:
    - Output / Output_PDF_Path                    -> if given:
         * if it's a folder, output = folder/<original_filename>.pdf
         * if it's a .pdf file path, output = that exact file path
      If not given, the script asks you to choose a single Output folder.

Behavior:
- Groups rows by (target PDF, computed output path).
- Opens each target PDF once, applies all its page replacements in memory, saves once.
- Uses page.show_pdf_page(...) to redraw target page content (keeps page count intact).
- Writes <template_name>_results.xlsx next to the template with: Result, Message, Saved_Output_Path.
"""

import os
import sys
import tempfile
import shutil
from typing import Optional, Dict, Tuple, List

import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import fitz  # PyMuPDF


# ----------------------------
# Tk helpers
# ----------------------------
def tk_select_file(title: str, filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))) -> Optional[str]:
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return path or None


def tk_select_folder(title: str) -> Optional[str]:
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path or None


# ----------------------------
# Column mapping / normalization
# ----------------------------
def normalize_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Map user headers to canonical keys:
      'pdf'    -> path or name
      'page'   -> 1-based page number
      'latest' -> replacement single-page pdf path (EXACT path from Excel)
      'output' -> optional output (folder or file)
    """
    canonical = {
        'pdf':  None,
        'page': None,
        'latest': None,
        'output': None,  # optional
    }

    aliases = {
        'pdf':    {'pdf', 'pdf_name', 'pdf file', 'pdf_path', 'input_pdf', 'source_pdf', 'file', 'filename'},
        'page':   {'page', 'page_no', 'page_number', 'pageno', 'pagenum'},
        'latest': {'latest', 'latest_page', 'latest_page_path', 'latest_path', 'replacement', 'replacement_pdf', 'new_page', 'new_pdf'},
        'output': {'output', 'output_path', 'output_pdf', 'output_pdf_path', 'save_as', 'dest', 'destination'},
    }

    lower_map = {c.lower().strip(): c for c in df.columns}
    for canon, candidates in aliases.items():
        for c in candidates:
            if c in lower_map:
                canonical[canon] = lower_map[c]
                break

    if canonical['pdf'] is None or canonical['page'] is None or canonical['latest'] is None:
        missing = [k for k in ('pdf', 'page', 'latest') if canonical[k] is None]
        raise ValueError(f"Template is missing required column(s): {', '.join(missing)}")

    return canonical


# ----------------------------
# Utility helpers
# ----------------------------
def is_pdf_file(path: str) -> bool:
    return bool(path) and path.lower().endswith('.pdf')


def resolve_target_pdf(pdf_cell_value: str, base_folder: Optional[str]) -> Optional[str]:
    """
    Resolve the target PDF path.
      - If cell value is an existing absolute path, return it.
      - Else, if base_folder provided, try base_folder/<cell> and also search recursively.
      - Else, return None.
    """
    if not pdf_cell_value:
        return None

    candidate = os.path.normpath(str(pdf_cell_value).strip().strip('"').strip("'"))
    if os.path.isabs(candidate) and os.path.isfile(candidate) and is_pdf_file(candidate):
        return candidate

    # If it's just a file name, try base folder directly
    if base_folder:
        direct = os.path.join(base_folder, os.path.basename(candidate))
        if os.path.isfile(direct) and is_pdf_file(direct):
            return os.path.normpath(direct)

        # Recursive search as fallback
        needle = os.path.basename(candidate).lower()
        for root_dir, _, files in os.walk(base_folder):
            for f in files:
                if f.lower() == needle and f.lower().endswith('.pdf'):
                    return os.path.normpath(os.path.join(root_dir, f))

    return None


def resolve_latest_exact(latest_cell_value: str) -> Optional[str]:
    """
    Latest must be an exact, absolute, existing .pdf path.
    No relative paths, no searching.
    """
    if latest_cell_value is None:
        return None

    v = str(latest_cell_value).strip().strip('"').strip("'")
    if not v or v.lower() == 'nan':
        return None

    v_norm = os.path.normpath(v)
    if not os.path.isabs(v_norm):
        return None
    if not os.path.isfile(v_norm):
        return None
    if not is_pdf_file(v_norm):
        return None

    return v_norm


def resolve_output_path(row_output_value: Optional[str],
                        fallback_output_folder: Optional[str],
                        original_pdf_path: str) -> str:
    """
    Decide where to save:
      - If row has Output_PDF_Path:
           * if it's a folder -> save as folder/<original_filename>.pdf
           * if it ends with .pdf -> save exactly there
      - Else if fallback_output_folder selected -> save there using original filename
      - Else -> save next to original with suffix _replaced.pdf
    """
    filename_only = os.path.basename(original_pdf_path)

    if row_output_value:
        out = os.path.normpath(str(row_output_value).strip().strip('"').strip("'"))
        if os.path.isdir(out):
            return os.path.join(out, filename_only)
        # If it's intended as a file path
        if out.lower().endswith('.pdf'):
            parent = os.path.dirname(out)
            if parent and not os.path.isdir(parent):
                os.makedirs(parent, exist_ok=True)
            return out
        # If neither folder nor .pdf file: treat as folder name to be created
        os.makedirs(out, exist_ok=True)
        return os.path.join(out, filename_only)

    if fallback_output_folder:
        os.makedirs(fallback_output_folder, exist_ok=True)
        return os.path.join(fallback_output_folder, filename_only)

    # Last resort: save next to original with suffix
    folder = os.path.dirname(original_pdf_path)
    base = os.path.splitext(filename_only)[0]
    return os.path.join(folder, f"{base}_replaced.pdf")


# ----------------------------
# Grouped replacement (apply many changes to the same PDF in memory)
# ----------------------------
def apply_group_replacements(target_pdf: str,
                             ops: List[Tuple[int, str, int]],
                             output_pdf: str) -> Tuple[bool, str, List[Tuple[int, str]]]:
    """
    Apply multiple page replacements to one PDF and save once.

    Parameters:
      target_pdf: path to the input PDF
      ops: list of (page_1based, latest_pdf_path, row_index)
      output_pdf: destination path

    Returns:
      (ok, message, per_op_results)
      where per_op_results = list of (row_index, message_per_row)
    """
    per_op_results: List[Tuple[int, str]] = []

    if not os.path.isfile(target_pdf):
        msg = f"Target not found: {target_pdf}"
        for _, _, row_idx in ops:
            per_op_results.append((row_idx, msg))
        return False, msg, per_op_results

    # Open target once
    try:
        doc = fitz.open(target_pdf)
    except Exception as e:
        msg = f"Failed to open target: {e}"
        for _, _, row_idx in ops:
            per_op_results.append((row_idx, msg))
        return False, msg, per_op_results

    # Sort ops by page (optional, keeps predictable order)
    ops_sorted = sorted(ops, key=lambda t: int(float(t[0])))

    try:
        # Apply each op
        for page_1based, latest_pdf, row_idx in ops_sorted:
            # Validate page number
            try:
                p = int(float(page_1based)) - 1
            except Exception:
                per_op_results.append((row_idx, f"Invalid page number: {page_1based}"))
                continue
            if p < 0 or p >= doc.page_count:
                per_op_results.append((row_idx, f"Page out of range (has {doc.page_count}, req {page_1based})"))
                continue

            # Validate latest page PDF
            if not (os.path.isabs(latest_pdf) and os.path.isfile(latest_pdf) and is_pdf_file(latest_pdf)):
                per_op_results.append((row_idx, f"Latest page not found: {latest_pdf}"))
                continue

            try:
                rep = fitz.open(latest_pdf)
            except Exception as e:
                per_op_results.append((row_idx, f"Cannot open latest page: {e}"))
                continue

            try:
                if rep.page_count != 1:
                    per_op_results.append((row_idx, f"Latest page must be single-page (got {rep.page_count})."))
                    rep.close()
                    continue

                page = doc.load_page(p)

                # Try to clean contents; if not available, just draw over
                try:
                    page.clean_contents()
                except Exception:
                    pass

                # Draw the replacement page into the full rect
                page.show_pdf_page(page.rect, rep, 0)
                per_op_results.append((row_idx, "OK"))

            except Exception as e:
                per_op_results.append((row_idx, f"Failed to replace page: {e}"))
            finally:
                try:
                    rep.close()
                except Exception:
                    pass

        # Save atomically in destination dir
        dest_dir = os.path.dirname(os.path.abspath(output_pdf)) or "."
        os.makedirs(dest_dir, exist_ok=True)
        tmp_fd, tmp_path = tempfile.mkstemp(dir=dest_dir, suffix=".pdf")
        os.close(tmp_fd)
        doc.save(tmp_path)

    except Exception as e:
        try:
            doc.close()
        except Exception:
            pass
        msg = f"Group save error: {e}"
        return False, msg, per_op_results

    # Close doc before moving
    try:
        doc.close()
    except Exception:
        pass

    # Move temp to final
    try:
        if os.path.exists(output_pdf):
            try:
                os.remove(output_pdf)
            except PermissionError:
                try:
                    os.replace(output_pdf, output_pdf + ".old")
                except Exception:
                    pass
        os.replace(tmp_path, output_pdf)
    except OSError:
        try:
            shutil.move(tmp_path, output_pdf)
        except Exception as ex2:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            return False, f"Failed to move output to destination: {ex2}", per_op_results

    return True, f"Saved: {output_pdf}", per_op_results


# ----------------------------
# Main
# ----------------------------
def main():
    # Pick template
    xlsx_path = tk_select_file("Select Excel TEMPLATE (.xlsx)")
    if not xlsx_path:
        print("No template selected. Exiting.")
        return

    try:
        df = pd.read_excel(xlsx_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel: {e}")
        return

    try:
        colmap = normalize_columns(df)
    except ValueError as ve:
        messagebox.showerror("Template Error", str(ve))
        return

    template_dir = os.path.dirname(xlsx_path)

    # If PDF column contains only names/relative, ask for base folder
    needs_base_target = False
    pdf_col = colmap['pdf']
    for v in df[pdf_col].dropna().astype(str):
        v_norm = os.path.normpath(v.strip().strip('"').strip("'"))
        if not (os.path.isabs(v_norm) and os.path.isfile(v_norm)):
            needs_base_target = True
            break

    base_folder_target = None
    if needs_base_target:
        base_folder_target = tk_select_folder("Select BASE folder where TARGET PDFs are located")
        if not base_folder_target:
            messagebox.showwarning("Cancelled", "No base folder selected for target PDFs. Exiting.")
            return

    # If no Output column, ask for a single output folder
    output_col = colmap.get('output')
    fallback_output_folder = None
    if output_col is None:
        fallback_output_folder = tk_select_folder("Select OUTPUT folder (all updated PDFs will be saved here)")
        if not fallback_output_folder:
            # Default next to template in 'Output'
            fallback_output_folder = os.path.join(template_dir, "Output")
            os.makedirs(fallback_output_folder, exist_ok=True)
            messagebox.showinfo("Info", f"No output folder chosen. Using: {fallback_output_folder}")

    # Build operations list from rows
    # Each row -> resolve target path, latest (exact), output path; store if valid enough to group
    ops_rows: List[Tuple[int, str, str, str]] = []  # (row_index, target_pdf, latest_pdf, output_pdf)
    per_row_msg = {}  # row_idx -> message if early failure

    for idx, row in df.iterrows():
        pdf_cell = str(row.get(pdf_col, "")).strip()
        latest_cell = str(row.get(colmap['latest'], "")).strip()
        page_cell = row.get(colmap['page'], None)
        row_output_cell_raw = row.get(output_col, None) if output_col else None
        row_output_cell = None if (row_output_cell_raw is None or (isinstance(row_output_cell_raw, float) and pd.isna(row_output_cell_raw))) else str(row_output_cell_raw)

        if not pdf_cell or not latest_cell or pd.isna(page_cell):
            per_row_msg[idx] = ("SKIPPED", "Missing required cell(s).", None)
            continue

        target_pdf = resolve_target_pdf(pdf_cell, base_folder_target)
        if not target_pdf:
            per_row_msg[idx] = ("ERROR", f"Cannot find target PDF for '{pdf_cell}'.", None)
            continue

        latest_pdf = resolve_latest_exact(latest_cell)
        if not latest_pdf:
            per_row_msg[idx] = ("ERROR", f"Latest page path invalid or not found: {latest_cell}", None)
            continue

        output_pdf = resolve_output_path(row_output_value=row_output_cell,
                                         fallback_output_folder=fallback_output_folder,
                                         original_pdf_path=target_pdf)

        try:
            page_1based = int(float(page_cell))
        except Exception:
            per_row_msg[idx] = ("ERROR", f"Page is not a number: {page_cell}", None)
            continue

        ops_rows.append((idx, target_pdf, latest_pdf, output_pdf, page_1based))

    # Group by (target_pdf, output_pdf)
    from collections import defaultdict
    groups = defaultdict(list)  # key -> list of (page_1based, latest_pdf, row_idx)

    for idx, target_pdf, latest_pdf, output_pdf, page_1based in ops_rows:
        key = (os.path.abspath(target_pdf), os.path.abspath(output_pdf))
        groups[key].append((page_1based, latest_pdf, idx))

    # Execute groups
    results: List[Tuple[str, str, Optional[str]]] = [("SKIPPED", "", None)] * len(df)  # prefill
    # Fill pre-known failures
    for idx, (r, m, p) in per_row_msg.items():
        results[idx] = (r, m, p)

    # Process each group
    for (target_pdf, output_pdf), ops in groups.items():
        ok, group_msg, per_op = apply_group_replacements(target_pdf, ops, output_pdf)

        # Map per-op messages back
        per_op_map = {row_idx: msg for (row_idx, msg) in per_op}

        if ok:
            # Mark each row: OK if its per-op msg == "OK", else ERROR with msg
            for _page_1based, _latest_pdf, row_idx in ops:
                row_msg = per_op_map.get(row_idx, "Unknown")
                if row_msg == "OK":
                    results[row_idx] = ("OK", f"Saved: {output_pdf}", output_pdf)
                else:
                    results[row_idx] = ("ERROR", row_msg, None)
        else:
            # Group save failed -> all rows in this group fail with group message
            for _page_1based, _latest_pdf, row_idx in ops:
                # If an individual op already had an error message, keep it; else set group failure
                row_msg = per_op_map.get(row_idx, group_msg) if per_op_map else group_msg
                results[row_idx] = ("ERROR", row_msg, None)

    # Write results next to template
    out_df = df.copy()
    # If results shorter (e.g., empty df), guard
    if len(out_df) != len(results):
        # Resize to match rows; shouldn't happen, but safe
        while len(results) < len(out_df):
            results.append(("SKIPPED", "No operation", None))

    out_df["Result"] = [r[0] for r in results]
    out_df["Message"] = [r[1] for r in results]
    out_df["Saved_Output_Path"] = [r[2] for r in results]

    result_path = os.path.join(
        template_dir,
        f"{os.path.splitext(os.path.basename(xlsx_path))[0]}_results.xlsx"
    )
    try:
        out_df.to_excel(result_path, index=False, engine="openpyxl")
        print(f"\nDone. Results saved to: {result_path}")
        # Small summary
        ok_count = sum(1 for r in results if r[0] == "OK")
        err_count = sum(1 for r in results if r[0] == "ERROR")
        skip_count = sum(1 for r in results if r[0] == "SKIPPED")
        summary = f"OK: {ok_count} | ERROR: {err_count} | SKIPPED: {skip_count}\nResults: {result_path}"
        messagebox.showinfo("Completed", f"Processing finished.\n{summary}")
    except Exception as e:
        print(f"Failed to write results Excel: {e}")
        messagebox.showerror("Write Error", f"Failed to write results Excel:\n{e}")


if __name__ == "__main__":
    main()