---
name: excel-live-automation
description: >-
  Automate, inspect, edit, audit, format, and recalculate live Excel workbooks in-memory via Windows COM without closing Excel or triggering file lock conflicts. Use whenever interacting with open .xlsx files, editing formulas across thousands of cells, applying accounting formatting, or automating budget workbooks on the user's desktop.
---

# Live In-Memory Excel Automation Skill

A standard workflow and reference for interacting with, editing, auditing, and formatting open Excel workbooks on Windows in real-time, while Excel remains open on the user's desktop.

## Why In-Memory COM Automation?

1. **User Pair-Programming:** The user can keep their spreadsheet open and visible while the agent performs updates in the background.
2. **Zero Disk Conflicts:** Writing via `openpyxl`, `pandas.to_excel`, or `xlsxwriter` fails with `PermissionError: [Errno 13]` when a file is open in Excel. Furthermore, disk-rewriting libraries often strip custom shapes, macros, conditional formatting, and cached calculation chains.
3. **Instantaneous Execution:** Excel's native COM interface allows writing 10,000+ formulas in under 20 milliseconds when using 2D array assignments.

---

## Core Operational Rules

1. **Never Close the User's Excel Application:**
   * Always attach to the existing running Excel process.
   * Do not kill `excel.exe` or call `app.Quit()`.
2. **No Third-Party Disk Overwrites:**
   * Never use `openpyxl.load_workbook(..., save=True)` on workbooks that are currently open in Excel.
3. **Directory Cleanliness:**
   * Never output scratch scripts (`.py`, `.bat`), dumps, or logs into the user's document or budget folders.
   * Run one-off automation scripts from conversation scratch or designated project tools directories.
4. **Batch Operations Over Cell Loops:**
   * Never loop over cells with `ws.Range(f"A{r}").Value = val` for hundreds of cells.
   * Always construct a 2D Python list (`matrix`) and assign it in a single call: `ws.Range("A1:Z100").Formula = matrix`.
5. **Screen Updating & Calculation Control:**
   * Disable `app.ScreenUpdating = False` during batch updates to prevent flickering and speed up execution.
   * Force full calculation using `app.CalculateFull()` after major formula insertions.
   * Save live in-memory changes to disk using `wb.Save()`.

---

## Standard Python Module Setup

When installed via pip:
```python
import live_excel as xl
# or
from live_excel import run_in_live_excel, resolve_target, read_range, write_range, format_cells, save
```

Or directly from source / local repository:
```python
import sys
# Add repository root to sys.path
sys.path.insert(0, r"path/to/LiveExcelAutomation")
import live_excel as xl  # or: import excel_core
```

---

## Common Recipes & Workflows

### 1. Attaching to Live Excel
```python
def my_worker(app):
    # Resolve workbook and worksheet
    wb, ws = excel_core.resolve_target(app, workbook="MyWorkbook.xlsx", sheet="Sheet1")
    
    # Read status
    print(f"Active Sheet: {ws.Name}, UsedRange: {ws.UsedRange.Address}")
    return True

excel_core.run_in_live_excel(my_worker)
```

### 2. High-Speed 2D Matrix Formula Injection
To populate multiple columns and rows without lag:
```python
def populate_formulas(app):
    wb, ws = excel_core.resolve_target(app, workbook="MyWorkbook.xlsx", sheet="Sheet1")
    
    start_row = 10
    end_row = 100
    cols = ["AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ"]
    
    # Read base formulas from Column Z
    z_formulas = [ws.Range(f"Z{r}").Formula for r in range(start_row, end_row + 1)]
    
    import re
    matrix = []
    for z_f in z_formulas:
        row_f = []
        for c in cols:
            # Shift relative column references from Z to target column
            new_f = re.sub(r'\bZ(\d+)\b', f'{c}\\1', z_f) if z_f.startswith("=") else z_f
            row_f.append(new_f)
        matrix.append(row_f)
        
    prev = app.ScreenUpdating
    app.ScreenUpdating = False
    try:
        ws.Range(f"AA{start_row}:AJ{end_row}").Formula = matrix
    finally:
        app.ScreenUpdating = prev
        
    app.CalculateFull()
    wb.Save()
```

### 3. Professional Accounting Formatting
Apply standard formatting (borders, fills, bolding, number formatting):
```python
def format_table(app):
    wb, ws = excel_core.resolve_target(app, workbook="MyWorkbook.xlsx", sheet="Sheet1")
    
    # 1. Number Formatting for Data
    data_rng = ws.Range("AA10:AJ100")
    data_rng.NumberFormat = "#,##0"  # or "0"
    
    # 2. Thin Borders for Grid
    for edge in [excel_core.xlEdgeTop, excel_core.xlEdgeBottom, excel_core.xlEdgeLeft, excel_core.xlEdgeRight, excel_core.xlInsideHorizontal, excel_core.xlInsideVertical]:
        data_rng.Borders(edge).LineStyle = excel_core.xlContinuous
        data_rng.Borders(edge).Weight = excel_core.xlThin
        
    # 3. Total / Summary Row Styling (Double Bottom Underline)
    total_rng = ws.Range("AA101:AJ101")
    total_rng.Font.Bold = True
    total_rng.Interior.ColorIndex = 37 # Soft blue fill
    total_rng.Borders(excel_core.xlEdgeTop).LineStyle = excel_core.xlContinuous
    total_rng.Borders(excel_core.xlEdgeTop).Weight = excel_core.xlThin
    total_rng.Borders(excel_core.xlEdgeBottom).LineStyle = excel_core.xlDouble # Double underline
    
    wb.Save()
```

### 4. Color Removal or Manipulation
```python
# Remove fill (no fill / transparent)
ws.Range("AK10:AK50").Interior.ColorIndex = -4142  # xlNone

# Standard Header Fill (Dark Blue with White Text)
hdr_rng = ws.Range("A1:AO1")
hdr_rng.Interior.ColorIndex = 49
hdr_rng.Font.Color = 16777215 # Pure White
hdr_rng.Font.Bold = True
```

---

## References

* Detailed COM constants: [Excel Constants Reference](./references/excel_constants.md)
* Complete Python API: [API Reference](./references/api_reference.md)
