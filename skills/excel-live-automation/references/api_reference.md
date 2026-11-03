# `excel_core.py` Python API Reference

Module Location: `D:\Nishit\Softwares\My Projects\ExcelAutomation\excel_core.py`

---

## Functions

### `run_in_live_excel(worker_fn, *args, **kwargs)`
Runs a Python callable inside an isolated COM thread on the interactive Windows desktop session.
* Catches desktop busy states and retries gracefully.
* Ensures `pythoncom.CoInitialize()` and `pythoncom.CoUninitialize()` lifecycle.
* Passes the active `Application` object to `worker_fn(app)`.

### `resolve_target(app, workbook=None, sheet=None)`
Resolves and returns a tuple `(Workbook, Worksheet)`:
* `app`: The live Excel Application.
* `workbook`: Optional string name (e.g. `"1. BigSheet 2026-27 - 30.10.2026.xlsx"`). If `None`, uses `app.ActiveWorkbook`.
* `sheet`: Optional string name or index. If `None`, uses `wb.ActiveSheet`.

### `read_range(range_address, workbook=None, sheet=None)`
Reads a rectangular range from the active or specified sheet into a nested 2D Python list.
* Example: `data = excel_core.read_range("A1:M50")`

### `read_dataframe(range_address, header=True, workbook=None, sheet=None)`
Reads the specified range into a pandas `DataFrame`.

### `write_range(top_left_cell, matrix, workbook=None, sheet=None)`
Writes a 2D matrix of values starting at `top_left_cell`.

### `format_cells(range_address, number_format=None, font_name=None, font_size=None, font_bold=None, font_color=None, bg_color=None, halign=None, valign=None, workbook=None, sheet=None)`
Convenience wrapper to format a range without writing raw COM calls.

### `autofit(columns=None, rows=None, workbook=None, sheet=None)`
Autofits column widths or row heights.

### `save(workbook=None)`
Saves the live in-memory workbook directly to disk without closing Excel.
