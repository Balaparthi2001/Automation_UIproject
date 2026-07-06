# -*- coding: utf-8 -*-
"""
Batch: Add text to PDFs ONLY when a target text is NOT present.
- Batch (folder → folder) only.
- Separate inputs:
    * Target text (presence check)
    * Add text (what to insert if target is absent)
- Include / Exclude page lists + "All pages" toggle
- Relative or Absolute bbox
- Font / size / color / alignment controls
- No border / no background, as requested
- Clear per-file console summary

Requires:
    pip install pymupdf
"""

from pathlib import Path
from typing import List, Optional, Tuple, Literal, Union

import fitz  # PyMuPDF
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter import StringVar, BooleanVar, DoubleVar


# -------------------------------
# Types & helpers
# -------------------------------
Color = Tuple[float, float, float]   # (r,g,b) in 0..1
BBox  = Tuple[float, float, float, float]  # (x,y,w,h)


def _resolve_pages(
    page_count: int,
    include_pages: Optional[List[int]] = None,
    exclude_pages: Optional[List[int]] = None,
) -> List[int]:
    """
    Build list of 0-based page indices to edit.
    include_pages / exclude_pages are 1-based in the UI.
    Priority: if include_pages is given, exclude_pages is ignored.
    """
    all_pages = list(range(page_count))
    if include_pages:
        inc_0 = sorted({p - 1 for p in include_pages if 1 <= p <= page_count})
        return inc_0
    if exclude_pages:
        exc_0 = {p - 1 for p in exclude_pages if 1 <= p <= page_count}
        return [i for i in all_pages if i not in exc_0]
    return all_pages


def _compute_rect(
    page: fitz.Page,
    bbox: BBox,
    bbox_mode: Literal["absolute", "relative"] = "relative",
    pad: float = 0.0
) -> fitz.Rect:
    """
    Convert (x,y,w,h) to page coordinates.
      - relative: x,y,w,h are 0..1 (percent of page size).
      - absolute: x,y,w,h are points (1/72 inch).
    """
    px = page.rect
    if bbox_mode == "relative":
        x = px.x0 + bbox[0] * px.width
        y = px.y0 + bbox[1] * px.height
        w = bbox[2] * px.width
        h = bbox[3] * px.height
    else:
        x, y, w, h = bbox

    rect = fitz.Rect(x, y, x + max(0, w), y + max(0, h)) & page.rect
    if pad > 0:
        rect = fitz.Rect(rect.x0 + pad, rect.y0 + pad, rect.x1 - pad, rect.y1 - pad)
    return rect


def _alignment_value(align: Literal["left", "center", "right", "justify"]) -> int:
    return {"left": 0, "center": 1, "right": 2, "justify": 3}[align]


def _normalize_text(s: str) -> str:
    """Lowercase, collapse whitespace, remove hyphen-newline joins for robust contains-check."""
    if not s:
        return ""
    s = s.replace("-\n", "")
    s = s.replace("\r", "\n").replace("\n", " ")
    s = " ".join(s.split())
    return s.lower().strip()


def _text_present(page: fitz.Page, clip: fitz.Rect, target: str) -> bool:
    """Check if target text appears in the clip (text layer only)."""
    extracted = page.get_text("text", clip=clip)
    return _normalize_text(target) in _normalize_text(extracted)


# -------------------------------
# Core: single PDF (called by batch)
# -------------------------------
def add_text_to_pdf(
    input_pdf_path: Union[str, Path],
    output_pdf_path: Union[str, Path],
    target_text: str,   # presence check
    add_text: str,      # what to insert if target is absent
    bbox: BBox,
    *,
    bbox_mode: Literal["absolute", "relative"] = "relative",
    include_pages: Optional[List[int]] = None,
    exclude_pages: Optional[List[int]] = None,
    font_name: str = "helv",
    font_size: float = 10.0,
    font_color: Color = (0, 0, 0),
    align: Literal["left", "center", "right", "justify"] = "left",
    inner_padding: float = 2.0,
    overprint: bool = True,
    check_scope: Literal["bbox", "page"] = "bbox",
) -> Tuple[List[int], List[int]]:
    """
    For each selected page:
      - If target_text (or add_text if target_text empty) is present in scope => skip
      - Else insert add_text inside bbox
    Returns (added_pages, skipped_pages) as 1-based lists.
    """
    in_path = Path(input_pdf_path)
    out_path = Path(output_pdf_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if not in_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {in_path}")

    # If target_text is blank, use add_text as the presence target
    presence_token = target_text.strip() if target_text.strip() else add_text.strip()

    doc = fitz.open(in_path)
    try:
        target_idx = _resolve_pages(doc.page_count, include_pages, exclude_pages)
        added_pages, skipped_pages = [], []

        for pidx in target_idx:
            page = doc[pidx]
            text_rect = _compute_rect(page, bbox, bbox_mode=bbox_mode, pad=inner_padding)

            # presence check
            check_rect = text_rect if check_scope == "bbox" else page.rect
            if presence_token and _text_present(page, check_rect, presence_token):
                skipped_pages.append(pidx + 1)
                continue

            # insert add_text
            page.insert_textbox(
                text_rect,
                add_text,
                fontname=font_name,
                fontsize=font_size,
                color=font_color,
                align=_alignment_value(align),
                overlay=overprint,
            )
            added_pages.append(pidx + 1)

        # save
        doc.save(out_path, deflate=True)
        return added_pages, skipped_pages
    finally:
        doc.close()


# -------------------------------
# Batch: a folder of PDFs
# -------------------------------
def batch_add_text_to_pdfs(
    input_folder: Union[str, Path],
    output_folder: Union[str, Path],
    target_text: str,
    add_text: str,
    bbox: BBox,
    *,
    bbox_mode: Literal["absolute", "relative"] = "relative",
    include_pages: Optional[List[int]] = None,
    exclude_pages: Optional[List[int]] = None,
    apply_all_pages: bool = False,
    font_name: str = "helv",
    font_size: float = 10.0,
    font_color: Color = (0, 0, 0),
    align: Literal["left", "center", "right", "justify"] = "left",
    inner_padding: float = 2.0,
    suffix: str = "_with_note",
    overprint: bool = True,
    check_scope: Literal["bbox", "page"] = "bbox",
) -> None:
    """
    Process every PDF in 'input_folder' and write outputs to 'output_folder'
    using the 'target_text' for presence check and 'add_text' for insertion.
    """
    in_dir = Path(input_folder)
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    pdfs = sorted([p for p in in_dir.glob("*.pdf") if p.is_file()])
    if not pdfs:
        print(f"[INFO] No PDFs found in: {in_dir}")
        return

    for pdf in pdfs:
        out_pdf = out_dir / f"{pdf.stem}{suffix}.pdf"
        print(f"\nProcessing: {pdf.name}")
        inc = None if apply_all_pages else include_pages
        exc = None if apply_all_pages else exclude_pages

        added, skipped = add_text_to_pdf(
            input_pdf_path=pdf,
            output_pdf_path=out_pdf,
            target_text=target_text,
            add_text=add_text,
            bbox=bbox,
            bbox_mode=bbox_mode,
            include_pages=inc,
            exclude_pages=exc,
            font_name=font_name,
            font_size=font_size,
            font_color=font_color,
            align=align,
            inner_padding=inner_padding,
            overprint=overprint,
            check_scope=check_scope,
        )

        if added:
            print(f"  Added on pages:  {added}")
        if skipped:
            print(f"  Skipped (already present): {skipped}")
        if not added and not skipped:
            print("  No target pages (due to include/exclude)")

    print(f"\nDone. Edited PDFs saved to: {out_dir}")


# -------------------------------
# GUI (Batch only)
# -------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Batch: Add Text When Target Absent")
        self.geometry("980x760")
        self.resizable(True, True)

        # Folders
        self.input_folder  = StringVar()
        self.output_folder = StringVar()
        self.suffix        = StringVar(value="_with_note")

        # Texts (separate)
        self.target_text   = StringVar(value="")  # used for presence check (optional)
        self.add_text      = StringVar(value="All Pipe attached supports shall be welded before galvanising.")

        # BBox
        self.bbox_mode_var = StringVar(value="relative")
        self.bbox_x        = DoubleVar(value=0.08)
        self.bbox_y        = DoubleVar(value=0.80)
        self.bbox_w        = DoubleVar(value=0.84)
        self.bbox_h        = DoubleVar(value=0.12)
        self.inner_padding = DoubleVar(value=2.0)

        # Pages
        self.apply_all_pages = BooleanVar(value=True)
        self.include_pages   = StringVar(value="")
        self.exclude_pages   = StringVar(value="")

        # Style
        self.font_name = StringVar(value="helv")
        self.font_size = DoubleVar(value=10.0)
        self.font_color_hex = StringVar(value="#000000")
        self.align_var  = StringVar(value="left")
        self.overprint  = BooleanVar(value=True)

        # Presence check scope
        self.check_scope = StringVar(value="bbox")  # "bbox" or "page"

        self._build_ui()

    # ---- utils ----
    @staticmethod
    def hex_to_rgb01(hex_color: str) -> Tuple[float, float, float]:
        s = hex_color.lstrip('#')
        if len(s) == 3: s = ''.join([c*2 for c in s])
        if len(s) != 6: return (0, 0, 0)
        r = int(s[0:2], 16) / 255.0
        g = int(s[2:4], 16) / 255.0
        b = int(s[4:6], 16) / 255.0
        return (r, g, b)

    def pick_color(self):
        c = colorchooser.askcolor(initialcolor=self.font_color_hex.get(), title="Choose font color")
        if c and c[1]:
            self.font_color_hex.set(c[1])

    def _parse_pages(self, s: str) -> Optional[List[int]]:
        s = (s or "").strip()
        if not s: return None
        out = []
        for part in s.split(","):
            part = part.strip()
            if not part: continue
            try:
                n = int(part)
                if n >= 1:
                    out.append(n)
            except ValueError:
                pass
        return out or None

    def _log(self, msg: str):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    # ---- UI ----
    def _build_ui(self):
        pad = {'padx': 8, 'pady': 6}

        # Folders
        f_frame = ttk.LabelFrame(self, text="Folders")
        f_frame.pack(fill="x", **pad)
        row = ttk.Frame(f_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Input Folder:").pack(side="left")
        ttk.Entry(row, textvariable=self.input_folder).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_input_folder).pack(side="left")

        row = ttk.Frame(f_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Output Folder:").pack(side="left")
        ttk.Entry(row, textvariable=self.output_folder).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_output_folder).pack(side="left")
        ttk.Label(row, text="File suffix:").pack(side="left", padx=(12, 4))
        ttk.Entry(row, textvariable=self.suffix, width=16).pack(side="left")

        # Texts
        t_frame = ttk.LabelFrame(self, text="Texts")
        t_frame.pack(fill="x", **pad)
        row = ttk.Frame(t_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Target text (presence check):").pack(side="left")
        ttk.Entry(row, textvariable=self.target_text).pack(side="left", fill="x", expand=True, padx=6)

        row = ttk.Frame(t_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Add text (to insert if target is absent):").pack(side="left")
        ttk.Entry(row, textvariable=self.add_text).pack(side="left", fill="x", expand=True, padx=6)

        # BBox
        b_frame = ttk.LabelFrame(self, text="BBox")
        b_frame.pack(fill="x", **pad)
        row = ttk.Frame(b_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Mode:").pack(side="left")
        ttk.Combobox(row, values=["relative", "absolute"], textvariable=self.bbox_mode_var, width=12, state="readonly").pack(side="left", padx=6)
        ttk.Label(row, text="x:").pack(side="left"); ttk.Entry(row, textvariable=self.bbox_x, width=8).pack(side="left")
        ttk.Label(row, text="y:").pack(side="left"); ttk.Entry(row, textvariable=self.bbox_y, width=8).pack(side="left")
        ttk.Label(row, text="w:").pack(side="left"); ttk.Entry(row, textvariable=self.bbox_w, width=8).pack(side="left")
        ttk.Label(row, text="h:").pack(side="left"); ttk.Entry(row, textvariable=self.bbox_h, width=8).pack(side="left")
        ttk.Label(row, text="Inner padding (pt):").pack(side="left"); ttk.Entry(row, textvariable=self.inner_padding, width=8).pack(side="left")

        # Pages
        p_frame = ttk.LabelFrame(self, text="Pages")
        p_frame.pack(fill="x", **pad)
        row = ttk.Frame(p_frame); row.pack(fill="x", **pad)
        ttk.Checkbutton(row, text="Apply to ALL pages", variable=self.apply_all_pages).pack(side="left")

        row = ttk.Frame(p_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Include (1-based, e.g., 1,2,5):").pack(side="left")
        ttk.Entry(row, textvariable=self.include_pages, width=28).pack(side="left", padx=6)
        ttk.Label(row, text="Except / Exclude:").pack(side="left")
        ttk.Entry(row, textvariable=self.exclude_pages, width=28).pack(side="left", padx=6)

        # Style
        s_frame = ttk.LabelFrame(self, text="Style")
        s_frame.pack(fill="x", **pad)
        row = ttk.Frame(s_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Font:").pack(side="left")
        ttk.Entry(row, textvariable=self.font_name, width=14).pack(side="left", padx=6)
        ttk.Label(row, text="Size:").pack(side="left")
        ttk.Entry(row, textvariable=self.font_size, width=8).pack(side="left", padx=6)
        ttk.Label(row, text="Font color:").pack(side="left")
        ttk.Entry(row, textvariable=self.font_color_hex, width=12).pack(side="left")
        ttk.Button(row, text="Pick…", command=self.pick_color).pack(side="left", padx=6)
        ttk.Label(row, text="Align:").pack(side="left", padx=(16, 4))
        ttk.Combobox(row, values=["left", "center", "right", "justify"], textvariable=self.align_var, width=10, state="readonly").pack(side="left", padx=6)
        ttk.Checkbutton(row, text="Overlay (draw on top)", variable=self.overprint).pack(side="left", padx=(16, 0))

        # presence scope
        pr_frame = ttk.LabelFrame(self, text="Presence Check Scope")
        pr_frame.pack(fill="x", **pad)
        row = ttk.Frame(pr_frame); row.pack(fill="x", **pad)
        ttk.Label(row, text="Check in:").pack(side="left")
        ttk.Combobox(row, values=["bbox", "page"], textvariable=self.check_scope, width=10, state="readonly").pack(side="left", padx=6)
        ttk.Label(row, text="(Adds only if target is absent)").pack(side="left", padx=(8, 0))

        # run
        run_row = ttk.Frame(self); run_row.pack(fill="x", **pad)
        ttk.Button(run_row, text="Run (Batch)", command=self.on_run).pack(side="left")
        ttk.Button(run_row, text="Close", command=self.destroy).pack(side="left", padx=6)

        # log
        self.log_box = tk.Text(self, height=12, wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)

    # folder pickers
    def pick_input_folder(self):
        d = filedialog.askdirectory(title="Select Input Folder (PDFs)")
        if d:
            self.input_folder.set(d)

    def pick_output_folder(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self.output_folder.set(d)

    # run
    def on_run(self):
        try:
            in_dir  = self.input_folder.get().strip()
            out_dir = self.output_folder.get().strip()
            if not in_dir:
                messagebox.showerror("Input Error", "Please choose the input folder.")
                return
            if not out_dir:
                messagebox.showerror("Input Error", "Please choose the output folder.")
                return

            target_text = self.target_text.get().strip()   # may be blank
            add_text    = self.add_text.get().strip()
            if not add_text:
                messagebox.showerror("Input Error", "Please enter the 'Add text' to insert.")
                return

            bbox_mode = self.bbox_mode_var.get()
            bbox = (
                float(self.bbox_x.get()),
                float(self.bbox_y.get()),
                float(self.bbox_w.get()),
                float(self.bbox_h.get())
            )
            inner_padding = float(self.inner_padding.get())

            apply_all = bool(self.apply_all_pages.get())
            inc = self._parse_pages(self.include_pages.get())
            exc = self._parse_pages(self.exclude_pages.get())

            font_name = self.font_name.get().strip() or "helv"
            font_size = float(self.font_size.get())
            font_color = self.hex_to_rgb01(self.font_color_hex.get())
            align = self.align_var.get()
            overprint = bool(self.overprint.get())
            scope = self.check_scope.get()
            suffix = self.suffix.get().strip() or "_with_note"

            self._log(f"Batch: {in_dir} → {out_dir}")
            self._log(f"  Target text: '{target_text or '(empty → will use Add text)'}'")
            self._log(f"  Add text   : '{add_text}'")

            batch_add_text_to_pdfs(
                input_folder=in_dir,
                output_folder=out_dir,
                target_text=target_text,
                add_text=add_text,
                bbox=bbox,
                bbox_mode=bbox_mode,
                include_pages=inc,
                exclude_pages=exc,
                apply_all_pages=apply_all,
                font_name=font_name,
                font_size=font_size,
                font_color=font_color,
                align=align,
                inner_padding=inner_padding,
                suffix=suffix,
                overprint=overprint,
                check_scope=scope,
            )
            self._log("Done ✔  (See console for per-file page details)")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(f"[ERR] {e}")


# Entrypoint
if __name__ == "__main__":
    app = App()
    app.mainloop()