#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import ttk, filedialog, messagebox
import os
import re
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdfminer.high_level import extract_text
import webbrowser
import threading
from pathlib import Path
import pandas as pd
import io  # for in-memory overlay PDF

# ------------------------
# Global state (as original)
# ------------------------
packing_slips_paths = []
upc_to_sku_map = {}        # { '012345678905': 'TSD641-GREY', ... }
selected_upc_excel = None

# ------------------------
# Helpers
# ------------------------
def downloads_dir():
    return str(Path.home() / "Downloads")

def human_basename(p):
    return os.path.basename(p)

def calculate_total_pages(paths):
    total_pages = 0
    for path in paths:
        try:
            with open(path, 'rb') as f:
                reader = PdfReader(f)
                total_pages += len(reader.pages)
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not open {path}: {e}")
    return total_pages

def update_page_count_display():
    packing_slips_page_count = calculate_total_pages(packing_slips_paths)
    treeview_packing_slips.heading('#0', text=f'Packing Slip Files - Page Count: {packing_slips_page_count}')

# ------------------------
# UPC/SKU logic
# ------------------------
# NOTE:
#  - SKUs can contain decimals (e.g., OTS25113-COFFEE-7.5)
#  - Some PDF text extractors occasionally glue the next line label ("Status")
#    to the SKU token (e.g., "OT24700-BROWNStatus"). We defensively strip
#    any trailing "STATUS" token when extracting.
SKU_REGEX = re.compile(r'\b(OT|OTA|OTB|OTC|OTS|TSD)\d+[A-Za-z0-9]*(?:-[A-Za-z0-9.]+)*\b')

SKU_SORT_PARTS = re.compile(r'^(OTA|OTB|OTC|OTS|OT|TSD|AAA)(\d+)([A-Za-z0-9.-]*)$')

PREFIX_ORDER = {
    "AAA": 0,  # placeholder / missing
    "OT": 1,
    "OTA": 2,
    "OTB": 3,
    "OTC": 4,
    "OTS": 5,
    "TSD": 6,
}

NUMERIC_RUN = re.compile(r'\b(\d{11,15})\b')

def normalize_to_upca(num_str):
    '''
    Return a list of plausible UPC-A (12-digit) candidates derived from a numeric run.

    Handles common PDF-extraction quirks where line numbers or other digits get concatenated
    to the front of the UPC (e.g., '1' + '686162007974' -> '1686162007974').

    Rules:
      - 11 digits  -> pad left with '0'
      - 12 digits  -> as-is
      - 13 digits starting with '0' (EAN-13 with leading 0) -> drop leading 0
      - 13-15 digits -> generate all 12-digit windows (most robust)
    '''
    s = (num_str or "").strip()
    s = re.sub(r"\D", "", s)

    if len(s) == 11:
        return ["0" + s]
    if len(s) == 12:
        return [s]
    if len(s) == 13 and s.startswith("0"):
        return [s[1:]]

    # If pdf text concatenated line numbers or other digits, slide a 12-digit window
    if 13 <= len(s) <= 15:
        cands = []
        seen = set()
        for i in range(0, len(s) - 12 + 1):
            w = s[i:i+12]
            if w not in seen:
                seen.add(w)
                cands.append(w)
        return cands

    return []

def build_upc_map_from_excel(xlsx_path):
    upc_to_sku = {}
    try:
        df = pd.read_excel(xlsx_path, header=0, usecols=[0, 1])
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel mapping: {e}")

    for _, row in df.iterrows():
        upc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        sku = canonical_sku(str(row.iloc[1]).strip()) if pd.notna(row.iloc[1]) else ""
        if not upc or not sku:
            continue
        digits = re.sub(r'\D+', '', upc)
        cands = normalize_to_upca(digits)
        if cands:
            upc12 = cands[0]
            if len(upc12) == 12:
                upc_to_sku[upc12] = sku
    return upc_to_sku

def extract_page_text(pdf_path, page_num):
    try:
        return extract_text(pdf_path, page_numbers=[page_num]) or ""
    except Exception:
        return ""

def _normalize_text_for_sku(text: str) -> str:
    # Normalize unicode hyphens/soft-hyphens + NBSP to stabilize parsing
    return (
        (text or "")
        .replace("\xa0", " ")
        # Some PDFs embed odd "replacement" glyphs between wrapped fields.
        # On QVC slips we've observed a stray box char between GREY and M36.
        # Normalize it to a hyphen so our SKU regexes can see a continuous token.
        .replace("￾", "-")
        .replace("\ufffd", "-")
        .replace("\ufeff", "")
        .replace("\u00ad", "-")
        .replace("\u2010", "-")
        .replace("\u2011", "-")
        .replace("\u2012", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )

def _prepare_text_for_sku_scans(text: str) -> str:
    """Normalize page text and repair common extraction artifacts before SKU scanning."""
    t = _normalize_text_for_sku(text)

    # Strip glued "Status" label from the end of SKU tokens (case-insensitive)
    t = re.sub(
        r'(?i)(\b(?:OT|OTA|OTB|OTC|OTS|TSD)\d+[A-Za-z0-9.-]*?)STATUS\b',
        r'\1',
        t,
    )

    # Join SKUs split across *lines* where the first line ends with a hyphen.
    # Examples: OT21710- + TURQUOISE  -> OT21710-TURQUOISE;  OTB2400-GREY- + M36 -> OTB2400-GREY-M36
    def _join_split_sku(m):
        left = m.group(1)  # includes trailing '-'
        right = (m.group(2) or "").strip()
        return left + right

    t = re.sub(
        r'(\b(?:OT|OTA|OTB|OTC|OTS|TSD)\d+(?:-[A-Za-z0-9.]+)*-)\s*(?:\r?\n|\r)+\s*([A-Za-z0-9.]+)\b',
        _join_split_sku,
        t,
        flags=re.MULTILINE,
    )

    return t


def find_sku_in_text(text):
    text = _prepare_text_for_sku_scans(text)

    # Allow decimals in the final segment (e.g., "-7.5").
    split_sku_regex = re.compile(r'(OT|OTA|OTB|OTC|OTS|TSD)(\d+)\W+([A-Za-z0-9.-]+)')
    m = split_sku_regex.search(text)
    if m:
        prefix, number, suffix = m.groups()
        candidate = f"{prefix}{number}-{suffix}"
        # Guard against partial matches that end with a hyphen when the SKU is
        # actually wrapped to the next line/field (e.g., "OTB2400-GREY-" + "M36").
        if not candidate.endswith("-"):
            return candidate

    m2 = SKU_REGEX.search(text)
    if m2:
        return m2.group(0)

    return None

def find_upc_mapped_sku_in_text(text, upc_to_sku):
    if not upc_to_sku:
        return (None, None)

    digits_only = re.sub(r'\D+', '', text or "")
    for upc12, sku in upc_to_sku.items():
        if upc12 and upc12 in digits_only:
            return (sku, upc12)
    return (None, None)


def find_all_upc_mapped_skus_in_text(text: str, upc_to_sku: dict) -> set:
    """
    Robust UPC→SKU matching used for Pick List (and can be used elsewhere).

    This intentionally mirrors the logic that works for pages missing SKUs:
      - Strip ALL non-digits from the page text into one long digits_only string
      - Look for any mapped UPC (12-digit) as a substring anywhere in that digits_only string

    Returns a set of mapped SKUs (each SKU at most once per page).
    """
    if not text or not upc_to_sku:
        return set()

    digits_only = re.sub(r'\D+', '', text or '')
    if len(digits_only) < 12:
        return set()

    # Fast path: slide a 12-digit window and check membership in the mapping
    mapped = set()
    seen_upc = set()
    for i in range(0, len(digits_only) - 12 + 1):
        upc12 = digits_only[i:i+12]
        if upc12 in seen_upc:
            continue
        sku = upc_to_sku.get(upc12)
        if sku:
            mapped.add(canonical_sku(sku))
            seen_upc.add(upc12)

    return mapped


def find_any_upc_candidate(text):
    for num in NUMERIC_RUN.findall(text or ""):
        for cand in normalize_to_upca(num):
            if len(cand) == 12:
                return cand
    return None

def extract_skus_with_sources(input_pdf, upc_to_sku):
    results = []
    reader = PdfReader(input_pdf)
    for i in range(len(reader.pages)):
        text = extract_page_text(input_pdf, i)

        direct_sku = find_sku_in_text(text)
        if direct_sku:
            results.append({"page": i, "sku": direct_sku, "source": "SKU", "upc": None})
            continue

        mapped_sku, upc12 = find_upc_mapped_sku_in_text(text, upc_to_sku)
        if mapped_sku:
            results.append({"page": i, "sku": mapped_sku, "source": "UPC", "upc": upc12})
            continue

        cand = find_any_upc_candidate(text)
        if cand:
            results.append({"page": i, "sku": "AAA0000", "source": "UNMAPPED_UPC", "upc": cand})
        else:
            results.append({"page": i, "sku": "AAA0000", "source": "MISSING", "upc": None})
    return results

def custom_sku_sort_key(sku_str):
    m = SKU_SORT_PARTS.match(sku_str)
    if not m:
        prefix, number, suffix = "AAA", "0000", ""
    else:
        prefix, number, suffix = m.groups()
    order = PREFIX_ORDER.get(prefix, 0)
    try:
        num_val = int(number)
    except ValueError:
        num_val = 0
    return (order, num_val, suffix)

# ------------------------
# PDF building
# ------------------------
def merge_pdfs(paths, output_pdf):
    merger = PdfMerger()
    for p in paths:
        merger.append(p)
    with open(output_pdf, 'wb') as f:
        merger.write(f)
    merger.close()

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib import colors  # NEW for table grid rendering

# ------------------------
# NEW: Excel-like table drawing helpers
# ------------------------
def _draw_grid_table(
    c, x, y_top, col_widths, headers, rows,
    row_h=14,
    header_font_name="Helvetica-Bold",
    body_font_name="Helvetica",
    font_size=9,
    header_fill=colors.lightgrey,
    grid_color=colors.black,
    padding_x=3, padding_y=3,
    bold_body_cols=None,
):
    """
    Draw an Excel-like table with fixed column widths.
    - Headers are drawn in header_font_name (bold by default)
    - Body columns listed in bold_body_cols are drawn bold
    Returns y_bottom after drawing.
    """
    if bold_body_cols is None:
        bold_body_cols = set()
    else:
        bold_body_cols = set(bold_body_cols)

    table_w = sum(col_widths)
    n_cols = len(col_widths)

    # Header background
    c.setFillColor(header_fill)
    c.rect(x, y_top - row_h, table_w, row_h, stroke=0, fill=RenderPDFTrue)
    c.setFillColor(colors.black)

    # Header text (BOLD)
    c.setFont(header_font_name, font_size)
    cx = x
    for j in range(n_cols):
        c.drawString(cx + padding_x, y_top - row_h + padding_y, str(headers[j]))
        cx += col_widths[j]

    # Body text
    y = y_top - row_h
    for r in rows:
        y -= row_h
        cx = x
        for j in range(n_cols):
            val = "" if j >= len(r) or r[j] is None else str(r[j])

            # Bold selected body columns
            if j in bold_body_cols:
                font_name = "Helvetica-Bold"
            else:
                font_name = body_font_name

            # Dynamic fit for PAGE #'s column (index 2) so it never spills outside the grid.
            # Shrinks font size as needed; if still too long, truncates with ellipsis.
            if j == 2:
                avail = max(1, col_widths[j] - (2 * padding_x))
                fitted_text, fitted_size = _fit_text_to_width(
                    c, val, font_name, max_size=font_size, min_size=6, max_width=avail
                )
                c.setFont(font_name, fitted_size)
                c.drawString(cx + padding_x, y + padding_y, fitted_text)
            else:
                c.setFont(font_name, font_size)
                c.drawString(cx + padding_x, y + padding_y, val)

            cx += col_widths[j]

    # Grid lines
    total_rows = 1 + len(rows)
    y_bottom = y_top - total_rows * row_h

    c.setStrokeColor(grid_color)

    # Horizontal
    for i in range(total_rows + 1):
        yy = y_top - i * row_h
        c.line(x, yy, x + table_w, yy)

    # Vertical
    cx = x
    c.line(cx, y_top, cx, y_bottom)
    for w in col_widths:
        cx += w
        c.line(cx, y_top, cx, y_bottom)

    return y_bottom


def _fit_text_to_width(c, text, font_name, max_size, min_size, max_width):
    """Return (fitted_text, fitted_font_size) so that text fits within max_width.
    If it still doesn't fit at min_size, truncate with ellipsis.
    """
    if text is None:
        return ("", max_size)
    s = str(text)
    size = max_size
    while size > min_size and c.stringWidth(s, font_name, size) > max_width:
        size -= 0.5
    if c.stringWidth(s, font_name, size) <= max_width:
        return (s, size)

    ell = "…"
    if max_width <= c.stringWidth(ell, font_name, min_size):
        return (ell, min_size)

    lo, hi = 0, len(s)
    best = ell
    while lo <= hi:
        mid = (lo + hi) // 2
        candidate = s[:mid].rstrip() + ell
        if c.stringWidth(candidate, font_name, min_size) <= max_width:
            best = candidate
            lo = mid + 1
        else:
            hi = mid - 1
    return (best, min_size)

# reportlab's Canvas.rect fill parameter expects 1/0. Use constant to avoid mistakes.
RenderPDFTrue = 1


def canonical_sku(s: str) -> str:
    """Canonicalize SKU strings so extraction quirks don't split Pick List rows.

    Handles:
      - Case-only differences (e.g., "TSD21308-Olive" vs "TSD21308-OLIVE")
      - pdf text extraction occasionally glues the next-line label onto the token
        (e.g., "OT24700-BROWNStatus" -> "OT24700-BROWN")
      - Normalizes various hyphen characters to '-'
    """
    if s is None:
        return ""
    s = _normalize_text_for_sku(str(s)).strip()
    if not s:
        return ""

    # Collapse whitespace around hyphens so split-line SKUs don't become distinct keys.
    # e.g. "OT21710- TURQUOISE" -> "OT21710-TURQUOISE"
    s = re.sub(r"\s*-\s*", "-", s)

    # Strip ONLY a trailing STATUS label if it was glued onto the SKU token.
    # We do this at canonicalization time to ensure *all* extraction paths
    # (regex matches, split matches, mapped SKUs, etc.) converge.
    s = re.sub(r'(?i)STATUS\s*$', '', s).strip()

    return s.upper()

def compress_page_ranges(nums):
    """Return a compact page list string like '12–15, 18, 20–22'."""
    if not nums:
        return ""
    nums = sorted({int(n) for n in nums if n is not None})
    ranges = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        ranges.append((start, prev))
        start = prev = n
    ranges.append((start, prev))

    parts = []
    for a, b in ranges:
        if a == b:
            parts.append(str(a))
        else:
            parts.append(f"{a}–{b}")  # en-dash
    return ", ".join(parts)


def add_page_numbers_to_pdf(input_pdf_path: str, output_pdf_path: str):
    """
    Overlay 1-indexed page numbers at the bottom-left corner of EVERY page.
    The number shown equals the final output PDF page index.
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        mediabox = page.mediabox
        w = float(mediabox.width)
        h = float(mediabox.height)

        packet = io.BytesIO()
        c = rl_canvas.Canvas(packet, pagesize=(w, h))
        c.setFont("Helvetica", 9)
        # Bottom-left: a little in from the edge
        c.drawString(28, 20, str(i + 1))
        c.save()
        packet.seek(0)

        overlay = PdfReader(packet).pages[0]
        page.merge_page(overlay)
        writer.add_page(page)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)


import math

def _summary_page_count(mapped_rows_len, unmatched_rows_len):
    """
    Deterministic page count for the Summary, matching generate_summary_pdf exactly.
    """
    page_w, page_h = letter
    top = page_h - 50
    bottom = 50
    row_h = 14

    # Title (26) + section label (18)
    y_after = top - 26 - 18
    available_h = y_after - bottom
    fit = max(1, int(available_h // row_h) - 1)  # minus header row

    def pages_for_rows(n):
        if n <= 0:
            return 0
        import math
        return math.ceil(n / fit)

    # If no unmatched rows, we still emit the "Unmatched Pages: None" page.
    if unmatched_rows_len <= 0:
        base = 1
    else:
        base = pages_for_rows(unmatched_rows_len)

    mapped_pages = pages_for_rows(mapped_rows_len)

    return base + mapped_pages

def _picklist_page_count(total_pick_rows, rows_per_col):
    per_page_capacity = 2 * rows_per_col
    return max(1, math.ceil(total_pick_rows / per_page_capacity))


# ------------------------
# UPDATED: Summary PDF as grid tables
# ------------------------
def generate_summary_pdf(path, mapped_rows, unmatched_rows):
    """
    mapped_rows:    list of (UPC, SKU, Final Page #)
    unmatched_rows: list of (Reason, UPC_or_None, Final Page #)
    """
    c = rl_canvas.Canvas(path, pagesize=letter)
    page_w, page_h = letter

    left = 50
    right = 50
    top = page_h - 50
    bottom = 50

    table_w = page_w - left - right
    row_h = 14

    def draw_page_title(y):
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, y, "Pages Sorted via UPC Mapping (Summary)")
        return y - 26

    def rows_fit_per_page():
        y_after = top - 26 - 18  # title + section label
        available_h = y_after - bottom
        n = int(available_h // row_h) - 1  # minus header row
        return max(1, n)

    fit = rows_fit_per_page()

    def draw_section(section_title, headers, rows, col_widths, bold_body_cols):
        remaining = list(rows)
        first = True
        while remaining:
            if not first:
                c.showPage()
            first = False

            y = top
            y = draw_page_title(y)

            c.setFont("Helvetica-Bold", 12)
            c.drawString(left, y, section_title)
            y -= 18

            chunk = remaining[:fit]
            remaining = remaining[fit:]

            _draw_grid_table(
                c, left, y, col_widths, headers, chunk,
                row_h=row_h, font_size=9,
                header_font_name="Helvetica-Bold",
                body_font_name="Helvetica",
                bold_body_cols=bold_body_cols
            )

    # ---- Unmatched Pages ----
    if unmatched_rows:
        headers = ["Reason", "UPC (if any)", "Final Page #"]
        rows = [(reason, upc if upc else "", page_no) for (reason, upc, page_no) in unmatched_rows]
        col_widths = [table_w * 0.45, table_w * 0.33, table_w * 0.22]

        draw_section(
            "Unmatched Pages",
            headers,
            rows,
            col_widths,
            bold_body_cols={2}
        )
    else:
        # Still produce a clean first page even if none
        y = top
        y = draw_page_title(y)

        c.setFont("Helvetica-Bold", 12)
        c.drawString(left, y, "Unmatched Pages")
        y -= 18
        c.setFont("Helvetica", 10)
        c.drawString(left, y, "None")

    # ---- Mapped UPC Pages ----
    if mapped_rows:
        c.showPage()
        headers = ["UPC", "SKU", "Final Page #"]
        rows = [(upc, sku, page_no) for (upc, sku, page_no) in mapped_rows]
        col_widths = [table_w * 0.30, table_w * 0.52, table_w * 0.18]

        draw_section(
            "Mapped UPC Pages",
            headers,
            rows,
            col_widths,
            bold_body_cols={2}
        )

    c.save()

def is_macys_page(text: str) -> bool:
    t = _normalize_text_for_sku(text)
    return MACYS_HEADER_RE.search(t) is not None

def extract_macys_qty_by_sku(text: str, upc_to_sku: dict) -> dict:
    totals = {}
    if not text:
        return totals

    norm = _prepare_text_for_sku_scans(text)

    for m in MACYS_ROW_RE.finditer(norm):
        upc_raw = m.group(1)
        mid = (m.group(2) or "").strip()
        qty_ordered_str = m.group(4)

        try:
            qty_ordered = int(qty_ordered_str)
        except Exception:
            qty_ordered = 1
        if qty_ordered <= 0:
            qty_ordered = 1

        upc_digits = re.sub(r"\D+", "", upc_raw)
        upc12 = None
        for cand in normalize_to_upca(upc_digits):
            if len(cand) == 12:
                upc12 = cand
                break

        sku = None

        if "SKU:" in mid.upper():
            idx = mid.upper().find("SKU:")
            sku_slice = mid[idx: idx + 140]
            sku = find_sku_in_text(sku_slice) or find_sku_in_text(mid)

        if not sku:
            sku = find_sku_in_text(mid)

        if not sku and upc12 and upc_to_sku:
            sku = upc_to_sku.get(upc12)

        if not sku or sku == "AAA0000":
            continue

        totals[sku] = totals.get(sku, 0) + qty_ordered

    return totals


def extract_direct_skus_from_page_text(text: str) -> set:
    """
    Extract ONLY SKUs that are explicitly present in the page text.
    This intentionally does NOT use UPC→SKU mapping.

    IMPORTANT: We avoid "split token" regexes that can produce partial SKUs like
    'OTB2400-GREY-' when the size (e.g., M36) is wrapped to the next line.
    """
    if not text:
        return set()

    norm = _prepare_text_for_sku_scans(text)
    skus: set[str] = set()

    # Prefer "SKU:" labeled occurrences
    for m in re.finditer(r'(?i)SKU:\s*', norm):
        start = m.end()
        slice_ = norm[start:start + 140]
        sku = find_sku_in_text(slice_)
        if sku:
            skus.add(canonical_sku(sku))

    # Also capture any SKU-like tokens present anywhere on the page
    for m in SKU_REGEX.finditer(norm):
        sku = m.group(0)
        if sku and sku != 'AAA0000':
            skus.add(canonical_sku(sku))

    skus.discard('AAA0000')
    return skus

def extract_mapped_skus_from_page_text(text: str, upc_to_sku: dict) -> set:
    """
    Extract mapped SKUs by finding UPC candidates in the text and mapping them via upc_to_sku.
    """
    if not text or not upc_to_sku:
        return set()

    norm = _prepare_text_for_sku_scans(text)

    mapped = set()

    # Find all 11-13 digit runs; normalize them to UPC-A 12 digits; map any that exist
    for raw in NUMERIC_RUN.findall(norm):
        for cand in normalize_to_upca(raw):
            if len(cand) == 12:
                sku = upc_to_sku.get(cand)
                if sku:
                    mapped.add(sku)

    mapped.discard("AAA0000")
    return mapped


def _mapped_sku_near_index(norm_text: str, center_idx: int, upc_to_sku: dict, window: int = 220) -> str | None:
    """Look for a UPC near a given index and return the mapped SKU if found."""
    if not upc_to_sku:
        return None
    start = max(0, center_idx - window)
    end = min(len(norm_text), center_idx + window)
    snippet = norm_text[start:end]

    # Search numeric runs in the local snippet and map the first one that resolves
    for raw in NUMERIC_RUN.findall(snippet):
        for cand in normalize_to_upca(raw):
            if len(cand) == 12:
                sku = upc_to_sku.get(cand)
                if sku and sku != "AAA0000":
                    return sku
    return None

def extract_picklist_sku_occurrences_with_upc_override(text: str, upc_to_sku: dict) -> list[str]:
    """
    Pick List extraction with UPC override:
      - Extract each SKU occurrence present on the page.
      - If a UPC near that SKU matches the UPC→SKU Excel mapping, REPLACE the SKU with the mapped SKU.
      - If no SKU occurrences exist at all, return [] (caller can fall back to UPC-only mapping).
    """
    if not text:
        return []

    norm = _prepare_text_for_sku_scans(text)
    results: list[str] = []

    # 1) Prefer labeled 'SKU:' occurrences (most reliable positions)
    for m in re.finditer(r'(?i)SKU:\s*', norm):
        start = m.end()
        slice_ = norm[start:start + 160]
        sku = find_sku_in_text(slice_)
        if not sku:
            continue
        override = _mapped_sku_near_index(norm, start, upc_to_sku)
        results.append(canonical_sku(override or sku))

    # 2) Capture other SKU-like tokens (covers pdfminer line-break weirdness)
    #    We keep positions so we can look for nearby UPCs.
    for m in SKU_REGEX.finditer(norm):
        sku = m.group(0)
        if sku == "AAA0000":
            continue
        override = _mapped_sku_near_index(norm, m.start(), upc_to_sku)
        results.append(canonical_sku(override or sku))

    # Remove placeholder
    results = [x for x in results if x and x != "AAA0000"]
    return results

def remap_macys_qty_by_upc_near_sku(text: str, sku_qty: dict, upc_to_sku: dict) -> dict:
    """If a UPC near a Macy's SKU maps to a different SKU, move the quantity to the mapped SKU."""
    if not text or not sku_qty or not upc_to_sku:
        return sku_qty or {}

    norm = _prepare_text_for_sku_scans(text)
    out: dict[str, int] = {}
    for sku, qty in sku_qty.items():
        try:
            idx = norm.find(sku)
        except Exception:
            idx = -1
        mapped = _mapped_sku_near_index(norm, idx if idx != -1 else 0, upc_to_sku) if idx != -1 else None
        final_sku = mapped or sku
        out[final_sku] = out.get(final_sku, 0) + int(qty)
    return out

def extract_skus_for_picklist_from_page_text(text: str, upc_to_sku: dict) -> set:
    """
    Pick List rule:
      - If ANY explicit SKU(s) exist on the page, use ONLY those.
      - Otherwise (no SKU found), fall back to UPC→SKU mapping.
    """
    direct = extract_direct_skus_from_page_text(text)
    if direct:
        return direct
    return extract_mapped_skus_from_page_text(text, upc_to_sku)

# ------------------------
# UPDATED: Pick List PDF as a 2-column grid table (same page overflow)
# ------------------------
# ------------------------
# UPDATED: Pick List PDF as a 2-column grid table (same page overflow)
# ------------------------
def generate_pick_list_pdf(path, pick_rows):
    """
    pick_rows: list of (sku, qty, pages_str) already sorted in final order.

    Renders as an Excel-like grid table in TWO COLUMNS per page:
      [SKU | QTY | PAGE #'s]  [SKU | QTY | PAGE #'s]
    Only adds a new page when both columns fill.

    - Bold headers
    - Bold QTY values
    - If multiple pages, title becomes: Pick List Page N
    """
    c = rl_canvas.Canvas(path, pagesize=letter)
    page_w, page_h = letter

    left = 50
    right = 50
    top = page_h - 50
    bottom = 50
    gutter = 20

    usable_w = page_w - left - right
    col_block_w = (usable_w - gutter) / 2.0

    sku_w = col_block_w * 0.56
    qty_w = col_block_w * 0.14
    pages_w = col_block_w * 0.30
    col_widths = [sku_w, qty_w, pages_w]

    title_h = 22
    section_gap = 8
    row_h = 14

    def compute_rows_per_col():
        y = top
        y -= title_h
        y -= section_gap
        available_h = y - bottom
        n = int(available_h // row_h) - 1  # minus header row
        return max(1, n)

    rows_per_col = compute_rows_per_col()
    total_pages = _picklist_page_count(len(pick_rows), rows_per_col)

    def draw_title(page_num):
        y = top
        c.setFont("Helvetica-Bold", 14)
        if total_pages > 1:
            c.drawString(left, y, f"Pick List Page {page_num}")
        else:
            c.drawString(left, y, "Pick List")
        y -= title_h
        return y - section_gap

    idx = 0
    total = len(pick_rows)
    page_num = 1

    while True:
        y_table_top = draw_title(page_num)

        left_rows = pick_rows[idx: idx + rows_per_col]
        idx += len(left_rows)

        right_rows = pick_rows[idx: idx + rows_per_col]
        idx += len(right_rows)

        left_table_rows = [(sku, qty, pages_str) for (sku, qty, pages_str) in left_rows]
        right_table_rows = [(sku, qty, pages_str) for (sku, qty, pages_str) in right_rows]

        _draw_grid_table(
            c,
            x=left,
            y_top=y_table_top,
            col_widths=col_widths,
            headers=["SKU", "QTY", "PAGE #'s"],
            rows=left_table_rows,
            row_h=row_h,
            font_size=9,
            header_font_name="Helvetica-Bold",
            body_font_name="Helvetica",
            bold_body_cols={1}  # QTY bold
        )

        if right_table_rows:
            _draw_grid_table(
                c,
                x=left + col_block_w + gutter,
                y_top=y_table_top,
                col_widths=col_widths,
                headers=["SKU", "QTY", "PAGE #'s"],
                rows=right_table_rows,
                row_h=row_h,
                font_size=9,
                header_font_name="Helvetica-Bold",
                body_font_name="Helvetica",
                padding_x=2,
                bold_body_cols={1}
            )

        if idx >= total:
            break

        c.showPage()
        page_num += 1

    c.save()

def add_sku_overlay_to_page(page, sku_text):
    mediabox = page.mediabox
    width = float(mediabox.width)
    height = float(mediabox.height)

    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("Helvetica-Bold", 18)
    label = f"SKU: {sku_text}"
    c.drawCentredString(width / 2, height / 2, label)
    c.save()
    packet.seek(0)

    overlay_reader = PdfReader(packet)
    overlay_page = overlay_reader.pages[0]
    page.merge_page(overlay_page)
    return page

# ------------------------
# IMPORTANT FIX: Pick List uses PyPDF2 text fallback if pdfminer text is incomplete
# ------------------------
def _get_text_for_picklist(reader: PdfReader, input_pdf: str, page_index: int) -> str:
    """
    Use pdfminer first (original behavior). If it looks incomplete, fall back to PyPDF2.
    This is ONLY used for Pick List parsing.
    """
    t1 = extract_page_text(input_pdf, page_index) or ""
    t1 = _normalize_text_for_sku(t1)

    # If pdfminer output is tiny or missing obvious structure, use PyPDF2 as fallback
    if len(t1.strip()) < 150 or ("LINE" not in t1 and "SKU" not in t1):
        try:
            t2 = reader.pages[page_index].extract_text() or ""
        except Exception:
            t2 = ""
        t2 = _normalize_text_for_sku(t2)
        if len(t2.strip()) > len(t1.strip()):
            return t2

    return t1

def create_sorted_pdf_with_summary(input_pdf, page_records, output_pdf):
    reader = PdfReader(input_pdf)
    sorted_info = sorted(page_records, key=lambda rec: custom_sku_sort_key(rec["sku"]))

    # -------------------------
    # Build Pick List quantities + page tracking (body positions)
    # -------------------------
    # Rule (current script behavior):
    #   1) If ANY mapped UPC(s) appear on the page, use those mapped SKU(s) for Pick List and ignore SKU text.
    #   2) Else, if explicit SKU(s) exist, use those.
    #   3) Else, fall back to the sorting-resolved SKU (if any).
    #
    # We also track *which body pages* each SKU appeared on, so we can print PAGE #'s
    # that correspond to FINAL output page numbers on the packing slip pages.
    pick_qty: dict[str, int] = {}
    sku_to_body_positions: dict[str, set[int]] = {}

    for body_pos, rec in enumerate(sorted_info):
        text = _get_text_for_picklist(reader, input_pdf, rec["page"])

        used_skus = set()

        mapped_skus = find_all_upc_mapped_skus_in_text(text, upc_to_sku_map)
        if mapped_skus:
            used_skus = set(mapped_skus)
        else:
            direct_skus = extract_direct_skus_from_page_text(text)
            if direct_skus:
                used_skus = set(direct_skus)
            else:
                sku = rec.get("sku")
                if sku and sku != "AAA0000":
                    used_skus = {canonical_sku(sku)}

        # Canonicalize SKUs so case differences don't split Pick List rows
        used_skus = {canonical_sku(s) for s in used_skus if canonical_sku(s)}

        for sku in used_skus:
            pick_qty[sku] = pick_qty.get(sku, 0) + 1
            sku_to_body_positions.setdefault(sku, set()).add(body_pos)

    # -----------------------------------------------
    # Write sorted body (with overlay for UPC SKUs)
    # -----------------------------------------------
    writer = PdfWriter()
    for rec in sorted_info:
        page = reader.pages[rec["page"]]
        if rec["source"] == "UPC" and rec["sku"] and rec["sku"] != "AAA0000":
            page = add_sku_overlay_to_page(page, rec["sku"])
        writer.add_page(page)

    # ---------------------------------------------------------
    # Compute Summary/Pick List page counts for Final Page #
    # ---------------------------------------------------------
    tmp_upc_positions = []
    tmp_unmatched_positions = []

    for rec in sorted_info:
        if rec["source"] == "UPC":
            tmp_upc_positions.append((rec["upc"], rec["sku"], None))
        elif rec["source"] == "UNMAPPED_UPC":
            tmp_unmatched_positions.append(("UPC not in mapping", rec["upc"], None))
        elif rec["source"] == "MISSING":
            tmp_unmatched_positions.append(("No UPC/SKU found", None, None))

    summary_pages = _summary_page_count(len(tmp_upc_positions), len(tmp_unmatched_positions))

    # Pick List pages: mirror generate_pick_list_pdf's rows-per-column math (only vertical matters)
    page_w, page_h = letter
    top = page_h - 50
    bottom = 50
    title_h = 22
    section_gap = 8
    row_h = 14

    available_h = (top - title_h - section_gap) - bottom
    rows_per_col = max(1, int(available_h // row_h) - 1)  # minus header row
    pick_pages = _picklist_page_count(len(pick_qty), rows_per_col)

    # Body starts AFTER summary + pick list (1-indexed)
    body_start_page = summary_pages + pick_pages + 1

    # Assign Final Page #'s (packing slips only) with correct offset
    upc_positions = []
    unmatched_positions = []

    for body_pos, rec in enumerate(sorted_info):
        final_page_number = body_start_page + body_pos

        if rec["source"] == "UPC":
            upc_positions.append((rec["upc"], rec["sku"], final_page_number))
        elif rec["source"] == "UNMAPPED_UPC":
            unmatched_positions.append(("UPC not in mapping", rec["upc"], final_page_number))
        elif rec["source"] == "MISSING":
            unmatched_positions.append(("No UPC/SKU found", None, final_page_number))

    # Build Pick List rows with PAGE #'s (final output page numbers for packing slips)
    pick_rows = []
    for sku, qty in pick_qty.items():
        body_positions = sku_to_body_positions.get(sku, set())
        final_pages = [body_start_page + pos for pos in body_positions]
        pages_str = compress_page_ranges(final_pages)
        pick_rows.append((sku, qty, pages_str))

    pick_rows = sorted(pick_rows, key=lambda x: custom_sku_sort_key(x[0]))

    # -------------------------
    # Write temp PDFs + merge
    # -------------------------
    out_dir = os.path.dirname(output_pdf)
    tmp_body = os.path.join(out_dir, "_tmp_sorted_body.pdf")
    with open(tmp_body, "wb") as f:
        writer.write(f)

    tmp_summary = os.path.join(out_dir, "_tmp_upc_summary.pdf")
    tmp_pick = os.path.join(out_dir, "_tmp_pick_list.pdf")
    tmp_merged = os.path.join(out_dir, "_tmp_merged.pdf")

    generate_summary_pdf(tmp_summary, upc_positions, unmatched_positions)
    generate_pick_list_pdf(tmp_pick, pick_rows)

    merger = PdfMerger()
    merger.append(tmp_summary)
    merger.append(tmp_pick)
    merger.append(tmp_body)

    with open(tmp_merged, "wb") as f_out:
        merger.write(f_out)
    merger.close()

    # Final step: page-number every page in the final output
    add_page_numbers_to_pdf(tmp_merged, output_pdf)

    for p in (tmp_summary, tmp_pick, tmp_body, tmp_merged):
        try:
            os.remove(p)
        except OSError:
            pass

def select_packing_slips():
    files = filedialog.askopenfilenames(
        title="Select Packing Slip PDFs",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not files:
        return
    for f in files:
        packing_slips_paths.append(f)
        treeview_packing_slips.insert('', 'end', text=human_basename(f), values=(f,))
    update_page_count_display()

def on_drop_packing_slips_treeview(event):
    raw = event.data
    paths = []
    buf = []
    brace = False
    for ch in raw:
        if ch == "{":
            brace = True
            buf = []
        elif ch == "}":
            brace = False
            paths.append("".join(buf))
            buf = []
        elif ch == " " and not brace:
            if buf:
                paths.append("".join(buf))
                buf = []
        else:
            buf.append(ch)
    if buf:
        paths.append("".join(buf))

    for p in paths:
        if p.lower().endswith(".pdf"):
            packing_slips_paths.append(p)
            treeview_packing_slips.insert('', 'end', text=human_basename(p), values=(p,))
    update_page_count_display()

def remove_selected_items():
    selected_items = treeview_packing_slips.selection()
    paths_to_remove = []
    for item in selected_items:
        values = treeview_packing_slips.item(item, 'values')
        if values:
            paths_to_remove.append(values[0])
        treeview_packing_slips.delete(item)
    for path in paths_to_remove:
        if path in packing_slips_paths:
            packing_slips_paths.remove(path)
    update_page_count_display()

def on_select_upc_excel():
    global upc_to_sku_map, selected_upc_excel
    xlsx = filedialog.askopenfilename(
        title="Select UPC→SKU Excel (Col A: UPC, Col B: SKU; headers in A1,B1)",
        filetypes=[("Excel", "*.xlsx *.xls")]
    )
    if not xlsx:
        return
    try:
        upc_to_sku_map = build_upc_map_from_excel(xlsx)
        selected_upc_excel = xlsx
        messagebox.showinfo(
            "UPC→SKU mapping loaded",
            f"Loaded {len(upc_to_sku_map)} UPC→SKU rows from:\n{human_basename(xlsx)}"
        )
    except Exception as e:
        messagebox.showerror("Excel Read Error", str(e))

def process_files():
    if not packing_slips_paths:
        messagebox.showwarning("No Files", "Please add at least one PDF.")
        return

    if not upc_to_sku_map:
        if not messagebox.askyesno(
            "No UPC Mapping Selected",
            "You haven't selected an Excel UPC→SKU mapping file.\n\n"
            "Without it, pages lacking SKUs will still be grouped at the front, "
            "but we won't be able to map UPCs to SKUs.\n\n"
            "Continue anyway?"
        ):
            return

    def worker():
        try:
            process_button.config(state=tk.DISABLED, text="Processing…")

            out_dir = downloads_dir()
            os.makedirs(out_dir, exist_ok=True)
            combined_path = os.path.join(out_dir, "CombinedPackingSlips.pdf")
            sorted_path = os.path.join(out_dir, "SortedPackingSlips.pdf")

            merge_pdfs(packing_slips_paths, combined_path)
            page_records = extract_skus_with_sources(combined_path, upc_to_sku_map)
            create_sorted_pdf_with_summary(combined_path, page_records, sorted_path)

            def done():
                process_button.config(state=tk.NORMAL, text="Process Files")
                messagebox.showinfo(
                    "Success",
                    f"Finished!\n\nCombined: {combined_path}\nSorted: {sorted_path}\n\n"
                    "The sorted file will open now."
                )
                try:
                    webbrowser.open('file://' + os.path.realpath(sorted_path))
                except Exception:
                    pass
            root.after(0, done)

        except Exception as e:
            def fail():
                process_button.config(state=tk.NORMAL, text="Process Files")
                messagebox.showerror("Processing Error", str(e))
            root.after(0, fail)

    threading.Thread(target=worker, daemon=True).start()

# ------------------------
# Build UI (match original look & placements)
# ------------------------
root = TkinterDnD.Tk()
root.title("Order Processing Utility v2.0")

top_btn_frame = tk.Frame(root)
top_btn_frame.pack(fill=tk.X, padx=5, pady=5)

select_packing_slips_button = tk.Button(top_btn_frame, text="Select Packing Slip Files", command=select_packing_slips)
select_packing_slips_button.pack(side=tk.LEFT, padx=5, pady=5)

select_upc_button = tk.Button(top_btn_frame, text="Select UPC→SKU Excel File", command=on_select_upc_excel)
select_upc_button.pack(side=tk.LEFT, padx=5, pady=5)

treeview_packing_slips = ttk.Treeview(root, columns=('Full Path',), displaycolumns=(), selectmode='extended')
treeview_packing_slips.pack(fill=tk.BOTH, expand=True)
treeview_packing_slips.heading('#0', text='Packing Slip Files - Page Count: 0', anchor='w')

treeview_packing_slips.drop_target_register(DND_FILES)
treeview_packing_slips.dnd_bind('<<Drop>>', on_drop_packing_slips_treeview)

button_frame = tk.Frame(root)
button_frame.pack(fill=tk.X, pady=0)

remove_button = tk.Button(button_frame, text="Remove Selected Files", command=remove_selected_items, fg="red")
remove_button.pack(side=tk.LEFT, padx=5, pady=5)

process_button = tk.Button(button_frame, text="Process Files", command=process_files, fg="green")
process_button.pack(side=tk.RIGHT, padx=5, pady=5)

width = 450
height = 600
root.geometry(f'{width}x{height}')

banner_label = tk.Label(root, text="2025, Created by Sean Mali", bg="gray", fg="white", height=1)
banner_label.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

root.mainloop()
