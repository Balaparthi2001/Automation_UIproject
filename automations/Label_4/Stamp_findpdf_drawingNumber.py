# -*- coding: utf-8 -*-

import os
import sys
import re
import contextlib
import logging
import traceback
from datetime import datetime
import subprocess
from typing import List, Dict, Tuple, Optional

import numpy as np
import cv2
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd

# PDF bookmarks
from PyPDF2 import PdfReader

# OCR
import pytesseract
from pytesseract import Output

# --- GUI for picking input/output folders and simple messages ---
import tkinter as tk
from tkinter import filedialog, messagebox


# =========================
# Runtime / embedded tools
# =========================
def runtime_dir() -> str:
    """
    Base dir where bundled resources live.
    - PyInstaller onefile: sys._MEIPASS (temp extraction)
    - PyInstaller onedir: folder of the executable
    - Running from source: folder of this script
    """
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", None)
        if base and os.path.isdir(base):
            return base
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_tools() -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Return (poppler_bin, tesseract_exe, tessdata_prefix).

    Tries these layouts:
      Poppler:
        - ./poppler/bin
        - ./poppler-25.12.0/Library/bin
        - env POPPLER_PATH
      Tesseract:
        - ./tesseract/tesseract.exe
        - ./Tesseract-OCR/tesseract.exe
    """
    base = runtime_dir()

    # Poppler candidates (dir must contain pdftoppm.exe)
    poppler_candidates = [
        os.path.join(base, "poppler", "bin"),
        os.path.join(base, "poppler-25.12.0", "Library", "bin"),
        os.environ.get("POPPLER_PATH"),  # if set
    ]
    poppler_bin: Optional[str] = None
    for cand in poppler_candidates:
        if cand and os.path.isdir(cand) and os.path.isfile(os.path.join(cand, "pdftoppm.exe")):
            poppler_bin = cand
            break

    # Tesseract candidates (exe file)
    cand_tess1 = os.path.join(base, "tesseract", "tesseract.exe")
    cand_tess2 = os.path.join(base, "Tesseract-OCR", "tesseract.exe")
    if os.path.isfile(cand_tess1):
        tesseract_exe = cand_tess1
        tessdata_prefix = os.path.join(base, "tesseract")
    elif os.path.isfile(cand_tess2):
        tesseract_exe = cand_tess2
        tessdata_prefix = os.path.join(base, "Tesseract-OCR")
    else:
        tesseract_exe, tessdata_prefix = None, None

    return poppler_bin, tesseract_exe, tessdata_prefix


# =========================
# Logging
# =========================
def setup_logger() -> logging.Logger:
    """
    Create a log file under %TEMP%\Stamp_Project_Logs (or cwd fallback).
    """
    try:
        base = os.environ.get("TEMP", None) or (os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd())
        log_dir = os.path.join(base, "Stamp_Project_Logs")
        os.makedirs(log_dir, exist_ok=True)
    except Exception:
        log_dir = os.getcwd()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(log_dir, f"run_{ts}.log")

    logger = logging.getLogger("stamp")
    logger.setLevel(logging.DEBUG)
    if not logger.handlers:
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    logger.info("==== Run started ====")
    logger.info("Python: %s", sys.version)
    logger.info("Frozen: %s", getattr(sys, "frozen", False))
    logger.info("cwd: %s", os.getcwd())
    logger.info("log: %s", log_path)
    return logger


def list_embed_tree(logger: logging.Logger, base: str):
    try:
        logger.info("Listing embedded tree under: %s", base)
        for root, _dirs, files in os.walk(base):
            rel = os.path.relpath(root, base)
            logger.debug("DIR %s", rel)
            for f in files:
                logger.debug("  - %s", os.path.join(rel, f))
    except Exception as e:
        logger.warning("list_embed_tree failed: %s", e)


def try_pdftoppm(logger: logging.Logger, poppler_dir: Optional[str]) -> None:
    """
    Try calling 'pdftoppm -v' from the selected poppler_dir (if provided).
    """
    try:
        env = os.environ.copy()
        if poppler_dir:
            env["PATH"] = poppler_dir + os.pathsep + env.get("PATH", "")
        cmd = ["pdftoppm", "-v"]
        logger.info("Testing pdftoppm: PATH startswith=%s", env.get("PATH", "")[:160])
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
        out = p.communicate()[0].decode("utf-8", errors="replace")
        logger.info("pdftoppm -v exit=%s; output:\n%s", p.returncode, out)
    except Exception as e:
        logger.error("pdftoppm test failed: %s", e)


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
    rgb = np.array(img_pil.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def render_page_to_image(pdf_path: str, page_number: int, dpi: int, poppler_path: Optional[str]):
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
        cv2.rectangle(img, (x, max(0, y - t_size[1] - 6)), (x + t_size[0] + 6, y), color, -1)
        cv2.putText(img, label, (x + 3, y - 6), font, scale, (255, 255, 255), 2, cv2.LINE_AA)


# =========================
# CONFIG — HARD-CODED OCR & POPPLER PATHS
# =========================
# >>> EDIT THESE TWO PATHS ONLY <<<
TESSERACT_PATH: Optional[str] = r"C:\Users\450574\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
POPPLER_PATH:   Optional[str] = r"D:\Bala\poppler-25.12.0\Library\bin"  # must contain pdftoppm.exe

# General settings
DPI = 300
OUTPUT_EXCEL_NAME = "Detected_Empty_Red_Boxes_By_Page.xlsx"

# Bookmark filter (L1 filter stays as-is)
TARGET_L1_CONTAINS = "engg iso"
TARGET_L2_CONTAINS = None

# ============== DN & SHT ROI (bottom-right narrow strip) ==============
DN_ROI_X1_FRAC, DN_ROI_X2_FRAC = 0.81, 0.99
DN_ROI_Y1_FRAC, DN_ROI_Y2_FRAC = 0.94, 0.97

def _clamp_roi():
    global DN_ROI_X1_FRAC, DN_ROI_X2_FRAC, DN_ROI_Y1_FRAC, DN_ROI_Y2_FRAC
    DN_ROI_X1_FRAC = float(min(max(DN_ROI_X1_FRAC, 0.0), 1.0))
    DN_ROI_X2_FRAC = float(min(max(DN_ROI_X2_FRAC, 0.0), 1.0))
    DN_ROI_Y1_FRAC = float(min(max(DN_ROI_Y1_FRAC, 0.0), 1.0))
    DN_ROI_Y2_FRAC = float(min(max(DN_ROI_Y2_FRAC, 0.0), 1.0))
    if DN_ROI_X1_FRAC >= DN_ROI_X2_FRAC: DN_ROI_X1_FRAC, DN_ROI_X2_FRAC = 0.80, 0.99
    if DN_ROI_Y1_FRAC >= DN_ROI_Y2_FRAC: DN_ROI_Y1_FRAC, DN_ROI_Y2_FRAC = 0.94, 0.97
_clamp_roi()

# OCR / regex
OCR_CONFIG_LINE   = r'--oem 3 --psm 7 -l eng -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_/.'   # single line
OCR_CONFIG_BLOCK  = r'--oem 3 --psm 6 -l eng -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_/.'   # block
OCR_CONFIG_SPARSE = r'--oem 3 --psm 11 -l eng -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_/.'  # sparse

# DN patterns
DN_REGEX_STRICT   = re.compile(r"^[A-Z]{2,3}\s*[-_/]\s*\d{6,7}$", re.IGNORECASE)
DN_REGEX_FALLBACK = re.compile(r"[A-Z]{2,4}\s*[-_/]\s*\d{5,7}", re.IGNORECASE)

# SHT labels & tight 3-digit pattern
SHT_LABEL_KEYWORDS = [
    "SHT. NO", "SHT NO", "SHEET NO", "SHEET. NO", "SHT NO.", "SHEET.NO.", "SHT"
]
SHT_REGEX = re.compile(r"\d{3}")


# =========================
# Red box detection
# =========================
RED_HSV_LOWER1 = np.array([0,   70, 70], dtype=np.uint8)
RED_HSV_UPPER1 = np.array([10, 255,255], dtype=np.uint8)
RED_HSV_LOWER2 = np.array([170, 70, 70], dtype=np.uint8)
RED_HSV_UPPER2 = np.array([180,255,255], dtype=np.uint8)

DASH_CONNECT_KERNEL = (3, 3)
DASH_DILATE_ITERS = 2
DASH_CLOSE_ITERS  = 1

MIN_BOX_AREA_RATIO = 0.01
MAX_BOX_AREA_RATIO = 0.8
MIN_ASPECT = 0.2
MAX_ASPECT = 10.0

BORDER_BAND_PCT = 0.05
MIN_BORDER_RED_RATIO = 0.70

INSIDE_MARGIN_PCT = 0.06
NONWHITE_GRAY_THR = 240
EMPTY_NONWHITE_RATIO_THR = 0.005
CANNY_LOWER = 80
CANNY_UPPER = 200
EMPTY_EDGE_DENSITY_THR = 0.002
PER_PAGE_TRUE_IF_ANY_EMPTY = True


def get_red_mask(img_bgr: np.ndarray) -> np.ndarray:
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


# =========================
# PDF outline mapping
# =========================
def get_outline_nodes(reader: PdfReader) -> List[Tuple[int, str, int]]:
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
            walk(outlines, level=1)
    except Exception:
        pass

    nodes.sort(key=lambda t: (t[0], t[2]))
    return nodes


def build_page_l1_l2_map(reader: PdfReader) -> Dict[int, Dict[str, Optional[str]]]:
    """
    L1: filled as ranges (standard).
    L2: set ONLY on the exact page where the L2 bookmark starts (no carry-forward).
    """
    num_pages = len(reader.pages)
    mapping = {i: {'L1': None, 'L2': None} for i in range(num_pages)}
    nodes = get_outline_nodes(reader)

    l1_nodes = [(lvl, title, page) for (lvl, title, page) in nodes if lvl == 1]
    l2_nodes = [(lvl, title, page) for (lvl, title, page) in nodes if lvl == 2]

    # --- L1 as ranges ---
    if l1_nodes:
        l1_titles_pages = [(t, p) for (_lvl, t, p) in l1_nodes]
        l1_bounds: List[Tuple[int, int, str]] = []
        for i, (l1_title, l1_start) in enumerate(l1_titles_pages):
            l1_end = (l1_titles_pages[i+1][1] - 1) if (i + 1 < len(l1_titles_pages)) else (num_pages - 1)
            l1_bounds.append((l1_start, l1_end, l1_title))
        for (s, e, l1_title) in l1_bounds:
            for p in range(max(0, s), min(num_pages - 1, e) + 1):
                mapping[p]['L1'] = l1_title

    # --- L2 point-only (NO range fill) ---
    for (_lvl, l2_title, l2_page) in l2_nodes:
        if 0 <= l2_page < num_pages:
            mapping[l2_page]['L2'] = l2_title

    return mapping


def page_matches_target(l1: Optional[str], l2: Optional[str]) -> bool:
    def contains(needle: Optional[str], hay: Optional[str]) -> bool:
        if needle is None: return True
        if hay is None: return False
        return needle.lower() in hay.lower()
    return contains(TARGET_L1_CONTAINS, l1) and contains(TARGET_L2_CONTAINS, l2)


# =========================
# DN + SHT extraction — ROI ONLY
# =========================
def _roi_from_page(img_bgr: np.ndarray) -> Tuple[np.ndarray, int, int, int, int]:
    H, W = img_bgr.shape[:2]
    rx1 = int(DN_ROI_X1_FRAC * W); rx2 = int(DN_ROI_X2_FRAC * W)
    ry1 = int(DN_ROI_Y1_FRAC * H); ry2 = int(DN_ROI_Y2_FRAC * H)
    rx1 = max(0, min(W-1, rx1)); rx2 = max(1, min(W,   rx2))
    ry1 = max(0, min(H-1, ry1)); ry2 = max(1, min(H,   ry2))
    return img_bgr[ry1:ry2, rx1:rx2].copy(), rx1, ry1, rx2, ry2


def _preprocess_for_ocr(roi_bgr: np.ndarray) -> np.ndarray:
    """Upscale + binarize for bold CAD text; remove thin graphics if any."""
    gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    scale = 2.0 if max(gray.shape) < 1200 else 1.5
    gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY, 31, 15)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
    th = cv2.morphologyEx(th, cv2.MORPH_OPEN, kernel, iterations=1)
    return th


def _extract_words(tsv_df: pd.DataFrame) -> pd.DataFrame:
    df = tsv_df.copy()
    if df is None or df.empty or "text" not in df.columns:
        return pd.DataFrame()
    df = df[df["text"].notna()].copy()
    if df.empty: return df
    df["left"]   = pd.to_numeric(df["left"], errors="coerce").fillna(0).astype(int)
    df["top"]    = pd.to_numeric(df["top"], errors="coerce").fillna(0).astype(int)
    df["width"]  = pd.to_numeric(df["width"], errors="coerce").fillna(0).astype(int)
    df["height"] = pd.to_numeric(df["height"], errors="coerce").fillna(0).astype(int)
    df["right"]  = df["left"] + df["width"]
    df["bottom"] = df["top"]  + df["height"]
    df["conf"]   = pd.to_numeric(df.get("conf", 0), errors="coerce").fillna(0)
    df["text_u"] = df["text"].astype(str).str.upper().str.replace("—","-", regex=False).str.replace("–","-", regex=False)
    df["text_u_nospace"] = df["text_u"].str.replace(" ", "", regex=False)
    return df


def _find_keyword_bbox(words_df: pd.DataFrame, keywords: List[str]) -> Optional[Tuple[int,int,int,int]]:
    """Find any word that looks like the label."""
    if words_df.empty: return None
    kws = [k.upper().replace(".", "").replace(" ", "") for k in keywords]
    for _, r in words_df.iterrows():
        t = str(r["text_u"]).upper().replace(".", "").replace(" ", "")
        if any(kw in t for kw in kws):
            return (int(r["left"]), int(r["top"]), int(r["width"]), int(r["height"]))
    return None


def _pick_dn(words_df: pd.DataFrame) -> Tuple[Optional[str], Optional[Tuple[int,int,int,int]]]:
    if words_df.empty: return (None, None)
    best, best_bb, best_score = None, None, -1
    for _, r in words_df.iterrows():
        tok = str(r["text_u"])
        tok_ns = str(r["text_u_nospace"])
        strict_ok = DN_REGEX_STRICT.match(tok_ns) is not None
        fb = DN_REGEX_FALLBACK.search(tok)
        if not strict_ok and not fb:
            continue
        lvl = 2 if strict_ok else 1
        cand = tok_ns if strict_ok else fb.group(0).replace(" ", "")
        score = lvl*100 + len(cand) + 0.01*float(r["conf"])
        if score > best_score:
            best, best_bb, best_score = cand, (int(r["left"]), int(r["top"]), int(r["width"]), int(r["height"])), score
    return best, best_bb


def _first3_digits(token: str) -> Optional[str]:
    """Normalize O->0 and return first 3 digits if present."""
    if not token: return None
    s = token.upper().replace("O","0")
    m = re.search(r"\d{3,}", s)
    return m.group(0)[:3] if m else None


def _pick_sht_from_label(words_df: pd.DataFrame, label_bb: Tuple[int,int,int,int]) -> Tuple[Optional[str], Optional[Tuple[int,int,int,int]]]:
    """Pick SHT value in a tight window directly below the SHT label (first 3 digits only)."""
    if label_bb is None or words_df.empty: return (None, None)
    lx, ly, lw, lh = label_bb
    sx1 = int(lx - 0.10 * lw)
    sx2 = int(lx + 1.20 * lw)
    sy1 = int(ly + 0.60 * lh)
    sy2 = int(ly + 2.00 * lh)

    cand, cand_bb, best_y = None, None, 1e9
    for _, r in words_df.iterrows():
        x1, y1, x2, y2 = int(r["left"]), int(r["top"]), int(r["right"]), int(r["bottom"])
        if (y1 >= sy1 and y2 <= sy2 and x1 <= sx2 and x2 >= sx1):
            token = str(r["text_u"]).strip()
            d3 = _first3_digits(token)
            if d3 and y1 < best_y:
                cand, cand_bb, best_y = d3.zfill(3), (int(r["left"]), int(r["top"]), int(r["width"]), int(r["height"])), y1
    return cand, cand_bb


def _pick_sht_right_of_dn(words_df: pd.DataFrame, dn_bb: Tuple[int,int,int,int]) -> Tuple[Optional[str], Optional[Tuple[int,int,int,int]]]:
    """Fallback: pick first 3-digit number immediately to the right of DN box (ignore REV)."""
    if dn_bb is None or words_df.empty: return (None, None)
    dx, dy, dw, dh = dn_bb
    # Tight band near DN row
    sx1 = int(dx + 0.85 * dw)
    sy1 = int(dy - 0.35 * dh)
    sy2 = int(dy + 0.80 * dh)

    best, best_bb, best_dist = None, None, 1e9
    dn_cx = dx + dw//2
    for _, r in words_df.iterrows():
        token = str(r["text_u"]).strip()
        d3 = _first3_digits(token)
        if not d3:
            continue
        x1, y1, x2, y2 = int(r["left"]), int(r["top"]), int(r["right"]), int(r["bottom"])
        if (x1 >= sx1) and (sy1 <= y1 <= sy2):
            cx = (x1 + x2) // 2
            dist = abs(cx - dn_cx)
            if dist < best_dist:
                best, best_bb, best_dist = d3.zfill(3), (int(r["left"]), int(r["top"]), int(r["width"]), int(r["height"])), dist
    return best, best_bb


def extract_dn_and_sht_from_page(img_bgr: np.ndarray) -> Tuple[Optional[str], Optional[str], Optional[Tuple[int,int,int,int]], Optional[Tuple[int,int,int,int]]]:
    """
    ROI-only extraction:
      - Find DN token by regex.
      - Find SHT value (prefer below SHT label; else to the right of DN).
      - Returns (dn_text, sht_text, dn_bbox_abs, sht_bbox_abs).
    """
    roi_bgr, rx1, ry1, rx2, ry2 = _roi_from_page(img_bgr)

    proc = _preprocess_for_ocr(roi_bgr)
    tsv = pytesseract.image_to_data(proc, output_type=Output.DATAFRAME, config=OCR_CONFIG_BLOCK)
    words = _extract_words(tsv)

    # 1) DN
    dn_text, dn_bb_ocr = _pick_dn(words)

    # 2) SHT from SHT label (if visible inside ROI)
    sht_label_bb = _find_keyword_bbox(words, SHT_LABEL_KEYWORDS)

    sht_text, sht_bb_ocr = (None, None)
    if sht_label_bb:
        sht_text, sht_bb_ocr = _pick_sht_from_label(words, sht_label_bb)

    # 3) If not found, pick to the right of DN
    if (sht_text is None) and dn_bb_ocr:
        sht_text, sht_bb_ocr = _pick_sht_right_of_dn(words, dn_bb_ocr)

    # map OCR bbox -> ROI -> page
    def _to_abs(bb_ocr: Optional[Tuple[int,int,int,int]]) -> Optional[Tuple[int,int,int,int]]:
        if bb_ocr is None: return None
        x, y, w, h = bb_ocr
        scale_x = proc.shape[1] / max(1, roi_bgr.shape[1])
        scale_y = proc.shape[0] / max(1, roi_bgr.shape[0])
        x0 = int(x / max(1e-6, scale_x)); y0 = int(y / max(1e-6, scale_y))
        w0 = int(w / max(1e-6, scale_x)); h0 = int(h / max(1e-6, scale_y))
        return (rx1 + x0, ry1 + y0, w0, h0)

    dn_bb_abs  = _to_abs(dn_bb_ocr)
    sht_bb_abs = _to_abs(sht_bb_ocr)

    # normalize sht_text: zero-pad to 3 digits
    if sht_text and re.fullmatch(r"\d{1,3}", sht_text):
        sht_text = sht_text.zfill(3)

    return dn_text, sht_text, dn_bb_abs, sht_bb_abs


# =========================
# Simple GUI helpers (folder selection + messages)
# =========================
def _new_tk_root() -> tk.Tk:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root

def pick_folder(title: str, initialdir: Optional[str] = None) -> Optional[str]:
    root = _new_tk_root()
    sel = filedialog.askdirectory(title=title, initialdir=initialdir or os.path.expanduser("~"))
    root.destroy()
    return sel if sel and os.path.isdir(sel) else None

def show_info(msg: str):
    try:
        root = _new_tk_root()
        messagebox.showinfo("Info", msg, parent=root)
        root.destroy()
    except Exception:
        pass

def show_error(msg: str):
    try:
        root = _new_tk_root()
        messagebox.showerror("Error", msg, parent=root)
        root.destroy()
    except Exception:
        pass

def choose_folder(prompt: str, initial: Optional[str] = None) -> str:
    """GUI first; console fallback."""
    try:
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        sel = filedialog.askdirectory(title=prompt, initialdir=initial or os.path.expanduser("~"))
        root.destroy()
        if sel and os.path.isdir(sel):
            return sel
        print(f"[Info] No folder selected for: {prompt}.")
    except Exception:
        print(f"[Info] GUI not available for: {prompt}. Falling back to console input.")
    while True:
        p = input(f"Enter folder path for '{prompt}': ").strip('"').strip()
        if os.path.isdir(p):
            return p
        print("  -> Invalid folder. Please try again.")


# =========================
# Main
# =========================
def main():
    logger = setup_logger()
    logger.info("Starting main()")
    logger.info("runtime_dir=%s", runtime_dir())

    # --- Resolve embedded paths (if any) ---
    embed_poppler, embed_tesseract_exe, embed_tess_prefix = load_tools()
    logger.info("embed_poppler=%s", embed_poppler)
    logger.info("embed_tesseract_exe=%s ; embed_tess_prefix=%s", embed_tesseract_exe, embed_tess_prefix)

    def _valid_poppler_dir(p: Optional[str]) -> bool:
        ok = bool(p and os.path.isdir(p) and os.path.isfile(os.path.join(p, "pdftoppm.exe")))
        logger.debug("check poppler_dir=%s -> %s", p, ok)
        return ok

    def _valid_tess_exe(p: Optional[str]) -> bool:
        ok = bool(p and os.path.isfile(p))
        logger.debug("check tess_exe=%s -> %s", p, ok)
        return ok

    # ---------------- Poppler: hardcoded > embedded > env > PATH ----------------
    poppler_candidates = []
    if _valid_poppler_dir(POPPLER_PATH):                 # 1) hardcoded
        poppler_candidates.append(("hardcoded", POPPLER_PATH))
    if _valid_poppler_dir(embed_poppler):                # 2) embedded (load_tools)
        poppler_candidates.append(("embedded", embed_poppler))
    env_poppler = os.environ.get("POPPLER_PATH")
    if _valid_poppler_dir(env_poppler):                  # 3) env var
        poppler_candidates.append(("env", env_poppler))
    logger.info("poppler_candidates=%s", poppler_candidates)

    if poppler_candidates:
        src, chosen = poppler_candidates[0]
        chosen = os.path.abspath(chosen)
        print(f"[OK] Using Poppler ({src}) at: {chosen}")
        logger.info("[OK] Using Poppler (%s) at: %s", src, chosen)
        os.environ["PATH"] = chosen + os.pathsep + os.environ.get("PATH", "")
        poppler_for_pdf2image = chosen
    else:
        print("[WARN] Poppler dir not found in hardcoded/embedded/env. Will rely on system PATH for pdftoppm.exe")
        logger.warning("Poppler dir not found; relying on system PATH")
        poppler_for_pdf2image = None  # allow PATH fallback

    # ---------------- Tesseract: hardcoded > embedded > env ----------------
    tess_candidates = []
    if _valid_tess_exe(TESSERACT_PATH):                  # 1) hardcoded
        tess_candidates.append(("hardcoded", TESSERACT_PATH))
    if _valid_tess_exe(embed_tesseract_exe):             # 2) embedded (load_tools)
        tess_candidates.append(("embedded", embed_tesseract_exe))
    env_tess = os.environ.get("TESSERACT_PATH")
    if _valid_tess_exe(env_tess):                        # 3) env var
        tess_candidates.append(("env", env_tess))
    logger.info("tess_candidates=%s", tess_candidates)

    if tess_candidates:
        src, chosen_tess = tess_candidates[0]
        chosen_tess = os.path.abspath(chosen_tess)
        pytesseract.pytesseract.tesseract_cmd = chosen_tess
        if embed_tess_prefix:
            os.environ["TESSDATA_PREFIX"] = embed_tess_prefix  # harmless if None
            logger.info("TESSDATA_PREFIX=%s", embed_tess_prefix)
        print(f"[OK] Using Tesseract ({src}) at: {chosen_tess}")
        logger.info("[OK] Using Tesseract (%s) at: %s", src, chosen_tess)
    else:
        print("[ERR] Tesseract not found (hardcoded/embedded/env).")
        logger.error("Tesseract not found; exiting.")
        show_error("Tesseract not found. Please set TESSERACT_PATH correctly.")
        sys.exit(2)

    # --- Write embedded tree listing (helps confirm packaging) ---
    try:
        list_embed_tree(logger, runtime_dir())
    except Exception as e:
        logger.warning("embed tree listing failed: %s", e)

    # --- Test pdftoppm availability ---
    try_pdftoppm(logger, poppler_for_pdf2image)

    # --- Pick folders (with GUI first, fallback console) ---
    try:
        FOLDER_PATH = pick_folder("Select INPUT folder containing PDF files")
        if not FOLDER_PATH:
            FOLDER_PATH = choose_folder("Select INPUT folder containing PDF files")
    except Exception:
        FOLDER_PATH = choose_folder("Select INPUT folder containing PDF files")
    if not FOLDER_PATH:
        show_error("No input folder selected. Exiting.")
        logger.error("No input folder selected; exiting.")
        sys.exit(1)

    try:
        DES_FOLDER_PATH = pick_folder("Select OUTPUT folder (Cancel to use '<input>\\output')", initialdir=FOLDER_PATH)
    except Exception:
        DES_FOLDER_PATH = None
    if not DES_FOLDER_PATH:
        DES_FOLDER_PATH = os.path.join(FOLDER_PATH, "output")
    os.makedirs(DES_FOLDER_PATH, exist_ok=True)
    logger.info("INPUT=%s ; OUTPUT=%s", FOLDER_PATH, DES_FOLDER_PATH)

    rows_for_excel: List[Dict] = []

    # Collect PDFs
    pdf_names = [f for f in os.listdir(FOLDER_PATH) if f.lower().endswith(".pdf")]
    logger.info("Found %d PDFs", len(pdf_names))
    if not pdf_names:
        show_info("No PDF files found in the selected input folder.")
        logger.info("No PDFs found; exiting.")
        sys.exit(0)

    # -------------- Processing loop --------------
    for filename in pdf_names:
        pdf_file = os.path.join(FOLDER_PATH, filename)
        logger.info("Processing: %s", filename)
        print(f"\nProcessing: {filename}")

        # Build per-page L1/L2 map (L2 point-only)
        try:
            reader = PdfReader(pdf_file)
            page_lmap = build_page_l1_l2_map(reader)
            num_pages = len(reader.pages)
            logger.info("PDF pages=%s", num_pages)
        except Exception as e:
            logger.error("PdfReader failed: %s", e)
            print(f"  Unable to read outlines/pages: {e}")
            page_lmap = {}
            num_pages = 0

        processed_any = False

        try:
            page_idx = 0
            while True:
                if num_pages and page_idx >= num_pages:
                    break

                l1 = page_lmap.get(page_idx, {}).get('L1')
                l2 = page_lmap.get(page_idx, {}).get('L2')

                # Filter: process only pages where L1 contains "engg iso"
                if not page_matches_target(l1, l2):
                    page_idx += 1
                    continue

                # Render page
                try:
                    with suppress_stderr():
                        pil_img = render_page_to_image(pdf_file, page_idx, DPI, poppler_for_pdf2image)
                    if pil_img is None:
                        logger.warning("render_page_to_image returned None (page %s)", page_idx+1)
                        break
                except Exception as e:
                    logger.error("render_page_to_image error (page %s): %s", page_idx+1, e)
                    print(f"  Render error: {e}")
                    break

                img_bgr = pil_to_bgr(pil_img)

                # --- Red box detection ---
                detections = detect_red_border_boxes(img_bgr)
                empty_box_flag = any(d['is_empty'] for d in detections) if PER_PAGE_TRUE_IF_ANY_EMPTY \
                                 else ((len(detections) > 0) and all(d['is_empty'] for d in detections))

                # --- DN + SHT extraction (ROI ONLY) ---
                dn_text, sht_text, dn_bb, sht_bb = extract_dn_and_sht_from_page(img_bgr)

                # Compose final DN value: "DN-SHT"
                dn_final = None
                if dn_text and sht_text:
                    dn_final = f"{dn_text}-{sht_text}"
                elif dn_text:
                    dn_final = dn_text
                elif sht_text:
                    dn_final = f"-{sht_text}"

                # L2 output = bookmark if present; else DN token (no carry-forward)
                l2_out = l2 if l2 else (dn_final or None)

                rows_for_excel.append({
                    "Filename": filename,
                    "Page_Number": page_idx + 1,
                    "L1": l1,
                    "L2": l2_out,
                    "EmptyBox": bool(empty_box_flag),
                    "DN": dn_final
                })

                logger.debug("Page %s: boxes=%s, EmptyBox=%s, L1=%s, L2_out=%s, DN=%s",
                             page_idx+1, len(detections), empty_box_flag, l1, l2_out, dn_final)
                print(f"  Page {page_idx + 1}: boxes={len(detections)} | EmptyBox={empty_box_flag} | L1={l1} | L2_out={l2_out} | DN={dn_final}")
                processed_any = True
                page_idx += 1

                if not num_pages and pil_img is None:
                    break

            if not processed_any:
                logger.info("No pages matched the target L1/L2 for this file.")
                print("  (No pages matched the target L1/L2; nothing processed.)")

        except Exception as e:
            logger.error("Unhandled error in file loop: %s\n%s", e, traceback.format_exc())
            print(f"  ERROR: {e}")

    # Save Excel
    try:
        df = pd.DataFrame(rows_for_excel, columns=["Filename", "Page_Number", "L1", "L2", "EmptyBox", "DN"])
        output_excel = os.path.join(DES_FOLDER_PATH, OUTPUT_EXCEL_NAME)
        df.to_excel(output_excel, index=False, engine="openpyxl")
        logger.info("Excel saved: %s (rows=%s)", output_excel, len(df))
        print(f"\nDetection complete. Results saved to {output_excel}")
        show_info(f"Detection complete.\n\nResults saved to:\n{output_excel}")
    except Exception as e:
        logger.error("Saving Excel failed: %s\n%s", e, traceback.format_exc())
        show_error(f"Saving Excel failed: {e}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Catch-any crash to log as well
        try:
            logger = setup_logger()
            logger.error("Fatal crash: %s\n%s", e, traceback.format_exc())
        finally:
            raise
