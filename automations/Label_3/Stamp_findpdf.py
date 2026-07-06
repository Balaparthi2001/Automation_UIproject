# -*- coding: utf-8 -*-

import os
import sys
import contextlib
from typing import List, Dict, Tuple, Optional

import numpy as np
import cv2
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd

# PDF bookmarks
from PyPDF2 import PdfReader  # pypdf2/PyPDF2 compatible import


# =========================
# Utilities
# =========================

@contextlib.contextmanager
def suppress_stderr():
    """Silence noisy library warnings printed to stderr."""
    with open(os.devnull, "w") as devnull:
        old_stderr = sys.stderr
        sys.stderr = devnull
        try:
            yield
        finally:
            sys.stderr = old_stderr


def pil_to_bgr(img_pil: Image.Image) -> np.ndarray:
    """Convert PIL Image to OpenCV BGR numpy array."""
    rgb = np.array(img_pil.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def render_page_to_image(pdf_path: str, page_number: int, dpi: int, poppler_path: Optional[str]):
    """Render a single page (0-based index) of a PDF to a PIL Image using pdf2image."""
    images = convert_from_path(
        pdf_path,
        first_page=page_number + 1,
        last_page=page_number + 1,
        dpi=dpi,
        poppler_path=poppler_path
    )
    return images[0] if images else None


def draw_labeled_box(img: np.ndarray, box: Tuple[int, int, int, int], label: str, color: Tuple[int, int, int], thickness: int = 3):
    x, y, w, h = box
    cv2.rectangle(img, (x, y), (x + w, y + h), color, thickness)
    if label:
        font = cv2.FONT_HERSHEY_SIMPLEX
        scale = 0.6
        t_size, _ = cv2.getTextSize(label, font, scale, 2)
        # Background for label
        cv2.rectangle(img, (x, max(0, y - t_size[1] - 6)), (x + t_size[0] + 6, y), color, -1)
        cv2.putText(img, label, (x + 3, y - 6), font, scale, (255, 255, 255), 2, cv2.LINE_AA)


# =========================
# CONFIG — EDIT THESE (defaults; can be overridden by folder pickers)
# =========================

# Default/fallback Poppler bin path (Windows). On macOS/Linux, set to None if poppler is in PATH.
POPPLER_PATH = r"D:\Bala\poppler-25.12.0\Library\bin"  # set to None if poppler is in PATH

# Rendering settings
DPI = 300

# Red color detection (HSV)
RED_HSV_LOWER1 = np.array([0,   70, 70], dtype=np.uint8)
RED_HSV_UPPER1 = np.array([10, 255,255], dtype=np.uint8)
RED_HSV_LOWER2 = np.array([170, 70, 70], dtype=np.uint8)
RED_HSV_UPPER2 = np.array([180,255,255], dtype=np.uint8)

# Morphology to connect dashed/zig-zag red borders
DASH_CONNECT_KERNEL = (3, 3)
DASH_DILATE_ITERS = 2
DASH_CLOSE_ITERS  = 1

# Candidate filtering
MIN_BOX_AREA_RATIO = 0.01
MAX_BOX_AREA_RATIO = 0.8
MIN_ASPECT = 0.2
MAX_ASPECT = 10.0

# Border-ness test
BORDER_BAND_PCT = 0.05
MIN_BORDER_RED_RATIO = 0.70

# Empty-inside test
INSIDE_MARGIN_PCT = 0.06
NONWHITE_GRAY_THR = 240
EMPTY_NONWHITE_RATIO_THR = 0.005
CANNY_LOWER = 80
CANNY_UPPER = 200
EMPTY_EDGE_DENSITY_THR = 0.002

# Decide per page: True if ANY empty box exists on the page
PER_PAGE_TRUE_IF_ANY_EMPTY = True

# Output file name
OUTPUT_EXCEL_NAME = "Detected_Empty_Red_Boxes_By_Page.xlsx"

# Bookmark target filter
TARGET_L1_CONTAINS = "Engg ISO"   # set to None to disable L1 filtering
TARGET_L2_CONTAINS = None         # set to a string to require L2 contains it (or None to ignore)


# =========================
# Folder selection helpers
# =========================

def choose_folder(prompt: str, initial: Optional[str] = None) -> str:
    """
    Ask the user to pick a folder via GUI (tkinter) or fallback to console input.
    Returns a valid existing folder path.
    """
    # 1) Try tkinter dialog
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)  # bring dialog to front

        selected = filedialog.askdirectory(
            title=prompt,
            initialdir=(initial if (initial and os.path.isdir(initial)) else os.path.expanduser("~"))
        )
        root.destroy()

        if selected and os.path.isdir(selected):
            return selected
        print(f"[Info] No folder selected for: {prompt}.")
    except Exception as _:
        print(f"[Info] GUI not available for: {prompt}. Falling back to console input.")

    # 2) Fallback to console input
    while True:
        path = input(f"Enter folder path for '{prompt}': ").strip('"').strip()
        if os.path.isdir(path):
            return path
        print("  -> Invalid folder. Please try again.")


# =========================
# Red box detection
# =========================

def get_red_mask(img_bgr: np.ndarray) -> np.ndarray:
    """Return binary mask of red regions."""
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    mask1 = cv2.inRange(hsv, RED_HSV_LOWER1, RED_HSV_UPPER1)
    mask2 = cv2.inRange(hsv, RED_HSV_LOWER2, RED_HSV_UPPER2)
    mask = cv2.bitwise_or(mask1, mask2)
    if DASH_DILATE_ITERS > 0:
        mask = cv2.dilate(mask, np.ones(DASH_CONNECT_KERNEL, np.uint8), iterations=DASH_DILATE_ITERS)
    if DASH_CLOSE_ITERS > 0:
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones(DASH_CONNECT_KERNEL, np.uint8), iterations=DASH_CLOSE_ITERS)
    return mask


def is_box_like(red_mask: np.ndarray, bbox: Tuple[int, int, int, int]) -> bool:
    """Check if red pixels in bbox are concentrated along the border band (like a hollow rectangle)."""
    x, y, w, h = bbox
    roi = red_mask[y:y+h, x:x+w]
    if roi.size == 0:
        return False
    total_red = int(cv2.countNonZero(roi))
    if total_red == 0:
        return False
    band = max(1, int(min(w, h) * BORDER_BAND_PCT))
    border_mask = np.zeros_like(roi, dtype=np.uint8)
    cv2.rectangle(border_mask, (0, 0), (w-1, h-1), 255, thickness=band)
    red_in_border = int(cv2.countNonZero(cv2.bitwise_and(roi, border_mask)))
    border_ratio = red_in_border / float(total_red)
    return border_ratio >= MIN_BORDER_RED_RATIO


def empty_inside(img_bgr: np.ndarray, bbox: Tuple[int, int, int, int]) -> Tuple[bool, float, float]:
    x, y, w, h = bbox
    if w <= 0 or h <= 0:
        return (False, 1.0, 1.0)
    margin = max(2, int(min(w, h) * INSIDE_MARGIN_PCT))
    ix = x + margin
    iy = y + margin
    iw = max(1, w - 2 * margin)
    ih = max(1, h - 2 * margin)
    if iw < 10 or ih < 10:
        return (False, 1.0, 1.0)
    roi = img_bgr[iy:iy+ih, ix:ix+iw]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    nonwhite = cv2.threshold(gray, NONWHITE_GRAY_THR, 255, cv2.THRESH_BINARY_INV)[1]
    nonwhite_ratio = float(cv2.countNonZero(nonwhite)) / float(iw * ih)
    edges = cv2.Canny(gray, CANNY_LOWER, CANNY_UPPER)
    edge_density = float(cv2.countNonZero(edges)) / float(iw * ih)
    is_empty = (nonwhite_ratio < EMPTY_NONWHITE_RATIO_THR) and (edge_density < EMPTY_EDGE_DENSITY_THR)
    return (is_empty, nonwhite_ratio, edge_density)


def detect_red_border_boxes(img_bgr: np.ndarray) -> List[Dict]:
    H, W = img_bgr.shape[:2]
    page_area = float(H * W)
    mask = get_red_mask(img_bgr)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    detections: List[Dict] = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        area_ratio = area / page_area
        if not (MIN_BOX_AREA_RATIO <= area_ratio <= MAX_BOX_AREA_RATIO):
            continue
        aspect = (w / float(h)) if h > 0 else 0
        if not (MIN_ASPECT <= aspect <= MAX_ASPECT):
            continue
        if not is_box_like(mask, (x, y, w, h)):
            continue
        is_empty, _nr, _ed = empty_inside(img_bgr, (x, y, w, h))
        detections.append({'box': (x, y, w, h), 'is_empty': bool(is_empty)})
    return detections


# =============================
# PDF outline → per-page L1/L2 labeling
# =============================

def get_outline_nodes(reader: PdfReader) -> List[Tuple[int, str, int]]:
    """Return a flat list of (level, title, page_index) from the PDF outlines."""
    try:
        try:
            outlines = reader.outline
        except Exception:
            outlines = reader.outlines
    except Exception:
        outlines = []
    nodes: List[Tuple[int, str, int]] = []

    def get_page_num_from_dest(dest) -> Optional[int]:
        try:
            return reader.get_destination_page_number(dest)
        except Exception:
            try:
                page_obj = getattr(dest, "page", None)
                if page_obj is not None:
                    for i, p in enumerate(reader.pages):
                        if p == page_obj:
                            return i
            except Exception:
                pass
        return None

    def walk(items, level: int):
        for item in items:
            if isinstance(item, list):
                walk(item, level + 1)
            else:
                title = ""
                try:
                    title = (getattr(item, "title", "") or str(getattr(item, "name", ""))).strip()
                except Exception:
                    title = ""
                page_idx = get_page_num_from_dest(item)
                if page_idx is not None:
                    nodes.append((level, title, page_idx))
                try:
                    kids = getattr(item, "children", None)
                    if kids:
                        walk(kids, level + 1)
                except Exception:
                    pass

    try:
        if outlines:
            walk(outlines, level=1)  # Level 1 = top-level bookmarks
    except Exception:
        pass

    nodes.sort(key=lambda t: (t[0], t[2]))
    return nodes


def build_page_l1_l2_map(reader: PdfReader) -> Dict[int, Dict[str, Optional[str]]]:
    """
    Build a mapping: page_index -> {'L1': str|None, 'L2': str|None}
    Pages inherit the current L1 (until next L1) and current L2 (until next L2 or next L1).
    """
    num_pages = len(reader.pages)
    mapping = {i: {'L1': None, 'L2': None} for i in range(num_pages)}

    nodes = get_outline_nodes(reader)
    l1_nodes = [(title, page) for lvl, title, page in nodes if lvl == 1]
    l2_nodes = [(title, page, None) for lvl, title, page in nodes if lvl == 2]

    l1_bounds: List[Tuple[int, int, int, str]] = []  # (l1_index, start, end, title)
    if l1_nodes:
        # L1 ranges
        for i, (l1_title, l1_start) in enumerate(l1_nodes):
            l1_end = (l1_nodes[i+1][1] - 1) if (i + 1 < len(l1_nodes)) else (num_pages - 1)
            l1_bounds.append((i, l1_start, l1_end, l1_title))

        # Attach L2 to parent L1 and create L2 ranges within L1
        l2_list: List[Tuple[int, str, int, int]] = []
        for (l2_title, l2_page, _) in l2_nodes:
            parent_idx = None
            for (i, s, e, _t) in l1_bounds:
                if s <= l2_page <= e:
                    parent_idx = i
                    break
            if parent_idx is None:
                continue
            l2_list.append((parent_idx, l2_title, l2_page, -1))

        from collections import defaultdict
        l2_by_l1 = defaultdict(list)
        for item in l2_list:
            l2_by_l1[item[0]].append(item)

        finalized_l2_ranges: List[Tuple[int, str, int, int]] = []
        for (i, s, e, _t) in l1_bounds:
            group = sorted(l2_by_l1.get(i, []), key=lambda x: x[2])
            for j, (_pi, l2_title, l2_start, _end_) in enumerate(group):
                l2_end = group[j+1][2] - 1 if (j + 1 < len(group)) else e
                finalized_l2_ranges.append((i, l2_title, l2_start, l2_end))

        # Fill mapping
        for (i, s, e, l1_title) in l1_bounds:
            for p in range(max(0, s), min(num_pages - 1, e) + 1):
                mapping[p]['L1'] = l1_title

        for (_pi, l2_title, l2_start, l2_end) in finalized_l2_ranges:
            for p in range(max(0, l2_start), min(num_pages - 1, l2_end) + 1):
                mapping[p]['L2'] = l2_title

    return mapping


def page_matches_target(l1: Optional[str], l2: Optional[str]) -> bool:
    def contains(needle: Optional[str], hay: Optional[str]) -> bool:
        if needle is None:
            return True
        if hay is None:
            return False
        return needle.lower() in hay.lower()
    return contains(TARGET_L1_CONTAINS, l1) and contains(TARGET_L2_CONTAINS, l2)


# =========================
# Main
# =========================

if __name__ == "__main__":
    # --- Ask user to choose folders (source PDFs and output) ---
    print("Please select the **Source PDFs folder**…")
    SRC_FOLDER = choose_folder("Select the Source PDFs folder", initial=os.path.expanduser("~"))

    print("Please select the **Output folder**…")
    OUT_FOLDER = choose_folder("Select the Output folder", initial=os.path.expanduser("~"))

    # Prepare debug folder (as in your current logic)
    DEBUG_IMG_FOLDER = os.path.join(OUT_FOLDER, "debug")
    os.makedirs(OUT_FOLDER, exist_ok=True)
    os.makedirs(DEBUG_IMG_FOLDER, exist_ok=True)

    rows_for_excel: List[Dict] = []

    for filename in os.listdir(SRC_FOLDER):
        if not filename.lower().endswith(".pdf"):
            continue

        pdf_file = os.path.join(SRC_FOLDER, filename)
        print(f"\nProcessing: {filename}")

        # Build per-page L1/L2 map for this PDF
        try:
            reader = PdfReader(pdf_file)
            page_lmap = build_page_l1_l2_map(reader)
            num_pages = len(reader.pages)
        except Exception as e:
            print(f"  Unable to read outlines/pages: {e}")
            page_lmap = {}
            num_pages = 0

        processed_any = False

        try:
            page_idx = 0
            while True:
                if num_pages and page_idx >= num_pages:
                    break

                # Decide whether this page is within target L1/L2
                l1 = page_lmap.get(page_idx, {}).get('L1')
                l2 = page_lmap.get(page_idx, {}).get('L2')

                if not page_matches_target(l1, l2):
                    page_idx += 1
                    continue

                # Render page
                try:
                    with suppress_stderr():
                        pil_img = render_page_to_image(pdf_file, page_idx, DPI, POPPLER_PATH)
                    if pil_img is None:
                        break
                except Exception:
                    break

                img_bgr = pil_to_bgr(pil_img)
                detections = detect_red_border_boxes(img_bgr)

                # Page-level EmptyBox flag
                if PER_PAGE_TRUE_IF_ANY_EMPTY:
                    empty_box_flag = any(d['is_empty'] for d in detections)
                else:
                    empty_box_flag = (len(detections) > 0) and all(d['is_empty'] for d in detections)

                # Draw debug overlay (save in selected output)
                dbg = img_bgr.copy()
                for k, det in enumerate(detections, start=1):
                    x, y, w, h = det['box']
                    label = f"Box{k} - {'EMPTY' if det['is_empty'] else 'NOT EMPTY'}"
                    color = (255, 0, 0) if det['is_empty'] else (0, 0, 255)
                    draw_labeled_box(dbg, (x, y, w, h), label, color, thickness=3)
                dbg_name = f"{os.path.splitext(filename)[0]}_p{page_idx+1}.png"
                cv2.imwrite(os.path.join(DEBUG_IMG_FOLDER, dbg_name), dbg)

                # Add to Excel rows
                rows_for_excel.append({
                    "Filename": filename,
                    "Page_Number": page_idx + 1,  # 1-based
                    "L1": l1,
                    "L2": l2,
                    "EmptyBox": bool(empty_box_flag)
                })

                print(f"  Page {page_idx + 1}: boxes={len(detections)} | EmptyBox={empty_box_flag} | L1={l1} | L2={l2}")
                processed_any = True
                page_idx += 1

                if not num_pages and pil_img is None:
                    break

            if not processed_any:
                print("  (No pages matched the target L1/L2; nothing processed.)")

        except Exception as e:
            print(f"  ERROR: {e}")

    # Save Excel in the selected output folder
    output_excel = os.path.join(OUT_FOLDER, OUTPUT_EXCEL_NAME)
    df = pd.DataFrame(rows_for_excel, columns=["Filename", "Page_Number", "L1", "L2", "EmptyBox"])
    df.to_excel(output_excel, index=False, engine="openpyxl")

    print(f"\nDetection complete. Results saved to {output_excel}")
    print(f"Debug images saved in {DEBUG_IMG_FOLDER}")
