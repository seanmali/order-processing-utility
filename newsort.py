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
SKU_REGEX = re.compile(r'(OT|OTA|OTB|TSD)\d+[A-Za-z-]*')
SKU_SORT_PARTS = re.compile(r'^(OT|OTA|OTB|TSD|AAA)(\d+)([A-Za-z-]*)$')
NUMERIC_RUN = re.compile(r'\b(\d{11,13})\b')
PREFIX_ORDER = {"AAA": 0, "OT": 1, "OTA": 2, "OTB": 3, "TSD": 4}

def normalize_to_upca(num_str):
    s = num_str.strip()
    if len(s) == 12:
        return [s]
    if len(s) == 11:
        return ["0" + s]
    if len(s) == 13 and s.startswith("0"):
        return [s[1:]]
    return []

def build_upc_map_from_excel(xlsx_path):
    upc_to_sku = {}
    try:
        df = pd.read_excel(xlsx_path, header=0, usecols=[0, 1])
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel mapping: {e}")

    for _, row in df.iterrows():
        upc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        sku = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
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

def find_sku_in_text(text):
    m = SKU_REGEX.search(text)
    return m.group(0) if m else None

def find_upc_mapped_sku_in_text(text, upc_to_sku):
    if not upc_to_sku:
        return (None, None)

    seen = set()
    for num in NUMERIC_RUN.findall(text):
        for cand in normalize_to_upca(num):
            if cand and cand not in seen:
                seen.add(cand)
                sku = upc_to_sku.get(cand)
                if sku:
                    return (sku, cand)
    return (None, None)

def find_any_upc_candidate(text):
    for num in NUMERIC_RUN.findall(text):
        for cand in normalize_to_upca(num):
            if len(cand) == 12:
                return cand
    return None

def extract_skus_with_sources(input_pdf, upc_to_sku):
    """
    Returns list of dicts per page:
      {
        "page": i,
        "sku": <sku string>,
        "source": "SKU"|"UPC"|"UNMAPPED_UPC"|"MISSING",
        "upc": <upc12 or None>
      }
    """
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

def generate_summary_pdf(path, mapped_rows, unmatched_rows):
    """
    mapped_rows:    list of (UPC, SKU, Final Page #)
    unmatched_rows: list of (Reason, UPC_or_None, Final Page #)
    """
    c = rl_canvas.Canvas(path, pagesize=letter)
    width, height = letter
    left = 50
    top = height - 50

    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, top, "Pages Sorted via UPC Mapping (Summary)")
    y = top - 24
    c.setFont("Helvetica", 10)

    # Section 1
    if mapped_rows:
        c.drawString(left, y, f"{'UPC':<16}  {'SKU':<28}  {'Final Page #'}")
        y -= 12
        c.line(left, y, width - 50, y)
        y -= 16
        for (upc, sku, page_no) in mapped_rows:
            c.drawString(left, y, f"{upc:<16}  {sku:<28}  {page_no}")
            y -= 12
            if y < 60:
                c.showPage()
                y = top
                c.setFont("Helvetica", 10)
        y -= 12

    # Section 2
    if unmatched_rows:
        if y < 90:
            c.showPage()
            y = top
            c.setFont("Helvetica", 10)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(left, y, "Unmatched Pages")
        y -= 16
        c.setFont("Helvetica", 10)
        c.drawString(left, y, f"{'Reason':<24}  {'UPC (if any)':<16}  {'Final Page #'}")
        y -= 12
        c.line(left, y, width - 50, y)
        y -= 16
        for (reason, upc_or_none, page_no) in unmatched_rows:
            upc_display = upc_or_none if upc_or_none else ""
            c.drawString(left, y, f"{reason:<24}  {upc_display:<16}  {page_no}")
            y -= 12
            if y < 60:
                c.showPage()
                y = top
                c.setFont("Helvetica", 10)

    c.showPage()
    c.save()

def create_sorted_pdf_with_summary(input_pdf, page_records, output_pdf):
    reader = PdfReader(input_pdf)
    sorted_info = sorted(page_records, key=lambda rec: custom_sku_sort_key(rec["sku"]))

    writer = PdfWriter()
    for rec in sorted_info:
        writer.add_page(reader.pages[rec["page"]])

    upc_positions = []
    unmatched_positions = []
    for idx, rec in enumerate(sorted_info):
        final_page_number = idx + 2  # 1-based, with summary occupying page 1
        if rec["source"] == "UPC":
            upc_positions.append((rec["upc"], rec["sku"], final_page_number))
        elif rec["source"] == "UNMAPPED_UPC":
            unmatched_positions.append(("UPC not in mapping", rec["upc"], final_page_number))
        elif rec["source"] == "MISSING":
            unmatched_positions.append(("No UPC/SKU found", None, final_page_number))

    out_dir = os.path.dirname(output_pdf)
    tmp_body = os.path.join(out_dir, "_tmp_sorted_body.pdf")
    with open(tmp_body, "wb") as f:
        writer.write(f)

    if upc_positions or unmatched_positions:
        tmp_summary = os.path.join(out_dir, "_tmp_upc_summary.pdf")
        generate_summary_pdf(tmp_summary, upc_positions, unmatched_positions)
        merger = PdfMerger()
        merger.append(tmp_summary)
        merger.append(tmp_body)
        with open(output_pdf, "wb") as f_out:
            merger.write(f_out)
        merger.close()
        for p in (tmp_summary, tmp_body):
            try:
                os.remove(p)
            except OSError:
                pass
    else:
        os.replace(tmp_body, output_pdf)

# ------------------------
# UI Actions (as original)
# ------------------------
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
    # event.data may include paths in braces for spaces
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
        # show brief confirmation
        messagebox.showinfo("UPC→SKU mapping loaded",
                            f"Loaded {len(upc_to_sku_map)} UPC→SKU rows from:\n{human_basename(xlsx)}")
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
            # disable process button while working
            process_button.config(state=tk.DISABLED, text="Processing…")

            out_dir = downloads_dir()
            os.makedirs(out_dir, exist_ok=True)
            combined_path = os.path.join(out_dir, "CombinedPackingSlips.pdf")
            sorted_path = os.path.join(out_dir, "SortedPackingSlips.pdf")

            # 1) Merge
            merge_pdfs(packing_slips_paths, combined_path)
            # 2) Extract per-page records
            page_records = extract_skus_with_sources(combined_path, upc_to_sku_map)
            # 3) Sort + summary
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

# Top: two buttons side-by-side (new Excel button added here)
top_btn_frame = tk.Frame(root)
top_btn_frame.pack(fill=tk.X, padx=5, pady=5)

select_packing_slips_button = tk.Button(top_btn_frame, text="Select Packing Slip Files", command=select_packing_slips)
select_packing_slips_button.pack(side=tk.LEFT, padx=5, pady=5)

select_upc_button = tk.Button(top_btn_frame, text="Select UPC→SKU Excel File", command=on_select_upc_excel)
select_upc_button.pack(side=tk.LEFT, padx=5, pady=5)

# Treeview (same as original)
treeview_packing_slips = ttk.Treeview(root, columns=('Full Path',), displaycolumns=(), selectmode='extended')
treeview_packing_slips.pack(fill=tk.BOTH, expand=True)
treeview_packing_slips.heading('#0', text='Packing Slip Files - Page Count: 0', anchor='w')

# Drag & Drop (same as original)
treeview_packing_slips.drop_target_register(DND_FILES)
treeview_packing_slips.dnd_bind('<<Drop>>', on_drop_packing_slips_treeview)

# Bottom button row (same as original: red remove, green process)
button_frame = tk.Frame(root)
button_frame.pack(fill=tk.X, pady=0)

remove_button = tk.Button(button_frame, text="Remove Selected Files", command=remove_selected_items, fg="red")
remove_button.pack(side=tk.LEFT, padx=5, pady=5)

process_button = tk.Button(button_frame, text="Process Files", command=process_files, fg="green")
process_button.pack(side=tk.RIGHT, padx=5, pady=5)

# Window geometry (same as original)
width = 450
height = 600
root.geometry(f'{width}x{height}')

# Bottom banner strip (slightly darker, like original; updated year)
banner_label = tk.Label(root, text="2025, Created by Sean Mali", bg="gray", fg="white", height=1)
banner_label.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

root.mainloop()
