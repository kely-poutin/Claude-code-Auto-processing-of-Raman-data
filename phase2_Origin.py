#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Phase 2 - Import Excel data into Origin and generate Raman spectrum line chart.
================================================================================
Auto-detects the .xlsx file in the script's own directory, reads the "Data
Collection" sheet, writes data into an Origin workbook, generates a line chart,
and saves the Origin project as .opju.

Requirements: Origin 2021+ (with originpro package), openpyxl
Usage: Run after Phase 1 in the same folder:
       python phase2.py
"""

import os
import sys
import time
import glob
import openpyxl
import originpro as op

# ---- Fix Windows GBK stdout for scripts with CJK paths ----
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ---- Auto-detect script directory ----
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

FOLDER_NAME = os.path.basename(BASE_DIR)
PRIMARY_XLSX = os.path.join(BASE_DIR, FOLDER_NAME + ".xlsx")

# ---- Step 6: Locate .xlsx and read data ----
print(f"Working dir: {BASE_DIR}")

# Prefer xlsx matching the folder name, else pick the newest
if os.path.exists(PRIMARY_XLSX):
    xlsx_path = PRIMARY_XLSX
else:
    candidates = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
    if not candidates:
        print("[ERR] No .xlsx file found in this folder. Run Phase 1 first.")
        sys.exit(1)
    xlsx_path = max(candidates, key=os.path.getmtime)
    print(f"[WARN] No folder-named xlsx; using newest: {os.path.basename(xlsx_path)}")

xlsx_basename = os.path.splitext(os.path.basename(xlsx_path))[0]
print(f"Source file: {os.path.basename(xlsx_path)}")
print(f"Base name (for outputs): {xlsx_basename}")

# Read "Data Collection" sheet
wb = openpyxl.load_workbook(xlsx_path, data_only=True)
try:
    ws = wb["数据集合"]  # 数据集合
except KeyError:
    print('[ERR] Sheet "数据集合" not found in xlsx. Verify Phase 1 completed successfully.')
    wb.close()
    sys.exit(1)

n_cols = ws.max_column
n_rows = ws.max_row

if n_cols < 2:
    print("[ERR] Data collection needs at least 2 columns (Raman Shift + >=1 Raw data).")
    wb.close()
    sys.exit(1)

print(f"\nData size: {n_cols} columns x {n_rows} rows")
if n_rows == 0:
    print("[ERR] Data collection is empty.")
    wb.close()
    sys.exit(1)

# Read all data column by column
all_data = []
for ci in range(1, n_cols + 1):
    col_data = []
    for ri in range(1, n_rows + 1):
        v = ws.cell(row=ri, column=ci).value
        col_data.append(None if v is None else float(v))
    all_data.append(col_data)
    non_empty = sum(1 for v in col_data if v is not None)
    print(f"  Col {ci}: {len(col_data)} rows, {non_empty} non-empty")

wb.close()

# ---- Step 7: Launch Origin (twice, >=5s gap), adjust columns ----
print("\n=== Launching Origin ===")


def launch_origin():
    """Start Origin and return the active worksheet."""
    print("  Starting Origin ...")
    op.new()
    time.sleep(5)
    print("  Refreshing ...")
    op.new()
    time.sleep(3)
    wks = op.find_sheet("w")
    if wks is None:
        print("[ERR] Cannot get Origin worksheet. Check Origin installation and originpro config.")
        sys.exit(1)
    return wks


wks = launch_origin()
print(f"Default workbook: '{wks.name}', columns: {wks.cols}")

# Add columns if needed (default workbook has 2)
if wks.cols < n_cols:
    to_add = n_cols - wks.cols
    print(f"  Need {n_cols} cols, have {wks.cols}, adding {to_add} ...")
    for _ in range(to_add):
        op.lt_exec("wks.addcol()")
    wks = op.find_sheet("w")
    print(f"  After adding: {wks.cols} columns")

# ---- Step 8: Write data to Origin workbook ----
print("\n=== Writing data ===")
for ci in range(n_cols):
    col_data = all_data[ci]
    clean = [(v if v is not None else float("nan")) for v in col_data]
    wks.from_list(ci, clean)

    # Column letter label
    if ci < 26:
        letter = chr(65 + ci)
    else:
        letter = f"Col{ci + 1}"
    print(f"  {letter}: {len(clean)} values")

print("Data written successfully.")

# ---- Step 9: Generate line chart ----
print("\n=== Generating line chart ===")
wks.activate()

plot_cmd = f"plotxy iy:=(1,2:{n_cols}) plot:=200"
print(f"  LabTalk: {plot_cmd}")
op.lt_exec(plot_cmd)

time.sleep(3)

graph = op.find_graph()
if graph:
    print(f"  Graph created: {graph.name}")
    try:
        op.lt_exec('xb.text$ = "Raman Shift (cm\\-(1))"')
        op.lt_exec('yl.text$ = "Intensity (a.u.)"')
        print("  Axis labels set.")
    except Exception as e:
        print(f"  [WARN] Axis label error: {e}")

    # Check and set axis ranges
    try:
        op.lt_exec('layer.x.from = 400;')
        op.lt_exec('layer.x.to = 2000;')
        op.lt_exec('layer.y.from = 0;')
        op.lt_exec('layer.y.to = 30000;')
        print("  Axis ranges set: X[400, 2000], Y[0, 30000].")
    except Exception as e:
        print(f"  [WARN] Axis range error: {e}")

    # Check and set line width for all data plots
    try:
        # doc -e P iterates all data plots; -w uses 0.1 pt units (2000 = 200 pts)
        op.lt_exec('doc -e P {set %C -w 2000;};')
        print("  Line width set to 200 for all plots.")
    except Exception as e:
        print(f"  [WARN] Line width error: {e}")
else:
    print("  [WARN] No graph detected. Please check Origin manually.")

# ---- Step 10: Save Origin project ----
opju_path = os.path.join(BASE_DIR, xlsx_basename + ".opju")
print(f"\n=== Saving project ===")
print(f"  Path: {opju_path}")

try:
    op.save(opju_path)
except Exception as e:
    print(f"  [WARN] API save failed ({e}), trying LabTalk ...")
    escaped = opju_path.replace("\\", "\\\\")
    op.lt_exec(f'save -i "{escaped}";')

time.sleep(2)

if os.path.exists(opju_path):
    print(f"  [OK] Saved: {opju_path} ({os.path.getsize(opju_path):,} bytes)")
else:
    print(f"  [ERR] File not found: {opju_path}")

# ---- Step 11: Close Origin ----
print("\n=== Closing Origin ===")
try:
    op.exit()
except Exception:
    pass
print("Origin closed.")
print("\n=== Phase 2 Complete ===")
