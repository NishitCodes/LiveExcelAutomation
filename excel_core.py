r"""
excel_core.py - Universal Live Excel In-Memory COM Automation Engine
Location: D:\Nishit\Softwares\ExcelAutomation\excel_core.py

Provides instantaneous (<0.15s), non-invasive live Excel automation across 
all Excel workbooks on the system by bridging the Windows WindowStation/Desktop 
boundary directly to the user's interactive desktop.

Capabilities:
- Targets any active or named workbook and worksheet.
- Live in-memory reading & writing (zero disk rewrites, zero file-lock collisions).
- Bulk 2D array read/write and Pandas DataFrame integration.
- Full formatting toolkit (accounting formats, percentages, fonts, borders, alignments, autofit).
- Safe save directly from RAM without closing Excel.
"""

import sys
import os
import time
import queue
import threading
import ctypes
from ctypes import wintypes
import pythoncom
import win32com.client

user32 = ctypes.windll.user32
oleacc = ctypes.windll.oleacc

# Alignment constants
xlCenter = -4108
xlLeft = -4131
xlRight = -4152
xlTop = -4160
xlBottom = -4107
xlGeneral = 1

# Border & Line constants
xlContinuous = 1
xlDash = -4115
xlDashDot = 4
xlDashDotDot = 5
xlDot = -4118
xlDouble = -4119
xlSlantDashDot = 13
xlLineStyleNone = -4142

# Border weights
xlHairline = 1
xlThin = 2
xlMedium = -4138
xlThick = 4

# Border edge indices (for ranges)
xlEdgeLeft = 7
xlEdgeTop = 8
xlEdgeBottom = 9
xlEdgeRight = 10
xlInsideVertical = 11
xlInsideHorizontal = 12

# Calculation modes
xlCalculationAutomatic = -4105
xlCalculationManual = -4135

class GUID(ctypes.Structure):
    _fields_ = [
        ('Data1', wintypes.DWORD),
        ('Data2', wintypes.WORD),
        ('Data3', wintypes.WORD),
        ('Data4', ctypes.c_byte * 8)
    ]

# IID_IDispatch: {00020400-0000-0000-C000-000000000046}
IID_IDispatch = GUID(
    0x00020400, 0x0000, 0x0000,
    (ctypes.c_byte * 8)(0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)
)

def rgb(r: int, g: int, b: int) -> int:
    """Converts RGB (0-255) to Windows BGR COLORREF integer."""
    return int(r) + (int(g) * 256) + (int(b) * 65536)

def _find_live_excel_handles():
    """Locates XLMAIN and child EXCEL7 window on the interactive desktop."""
    h_desk = user32.OpenDesktopW('default', 0, False, 0x10000000)
    if h_desk:
        user32.SetThreadDesktop(h_desk)

    hwnd_main = user32.FindWindowW('XLMAIN', None)
    if not hwnd_main and h_desk:
        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def enum_cb(h, lp):
            nonlocal hwnd_main
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(h, buf, 256)
            if buf.value == 'XLMAIN':
                hwnd_main = h
                return False
            return True
        user32.EnumDesktopWindows(h_desk, WNDENUMPROC(enum_cb), 0)

    if not hwnd_main:
        return None, None

    hwnd_child = None
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum_ch(ch, lp):
        nonlocal hwnd_child
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(ch, buf, 256)
        if buf.value == 'EXCEL7':
            hwnd_child = ch
            return False
        return True
    user32.EnumChildWindows(hwnd_main, WNDENUMPROC(enum_ch), 0)

    return hwnd_main, hwnd_child

def _desktop_worker_runner(func, q, args, kwargs):
    """Executes func inside a fresh thread attached to the interactive desktop."""
    try:
        hwnd_main, hwnd_child = _find_live_excel_handles()
        app = None

        if hwnd_child:
            # Clear cell editing mode if blinking cursor
            user32.PostMessageW(hwnd_main, 0x0100, 0x1B, 0) # WM_KEYDOWN VK_ESCAPE
            user32.PostMessageW(hwnd_main, 0x0101, 0x1B, 0) # WM_KEYUP VK_ESCAPE

            pythoncom.CoInitialize()
            p_acc = ctypes.c_void_p()
            OBJID_NATIVEOM = -16
            res = oleacc.AccessibleObjectFromWindow(
                hwnd_child, OBJID_NATIVEOM, ctypes.byref(IID_IDispatch), ctypes.byref(p_acc)
            )
            if res == 0 and p_acc.value:
                try:
                    p_disp = pythoncom.ObjectFromAddress(p_acc.value, pythoncom.IID_IDispatch)
                    xl_win = win32com.client.Dispatch(p_disp)
                    if hasattr(xl_win, 'Application'):
                        app = xl_win.Application
                    elif hasattr(xl_win, 'Parent') and hasattr(xl_win.Parent, 'Application'):
                        app = xl_win.Parent.Application
                except Exception:
                    app = None

        # Fallback to ROT GetActiveObject
        if app is None:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                pass

        if app is None:
            raise RuntimeError("Could not connect to live Excel instance in RAM. Is Excel open?")

        result = func(app, *args, **kwargs)
        pythoncom.CoUninitialize()
        q.put(('ok', result))
    except Exception as e:
        q.put(('err', str(e)))

def run_in_live_excel(func, *args, **kwargs):
    """
    Dispatches a callable against the running Excel instance.
    The callable receives (app, *args, **kwargs).
    """
    op_timeout = kwargs.pop('_timeout', 30.0)
    q = queue.Queue()
    t = threading.Thread(target=_desktop_worker_runner, args=(func, q, args, kwargs))
    t.start()
    t.join(timeout=op_timeout)
    if t.is_alive():
        raise TimeoutError(f"Live Excel operation timed out after {op_timeout}s.")
    status, value = q.get()
    if status == 'err':
        raise RuntimeError(value)
    return value

# ==============================================================================
# Universal Target Resolution
# ==============================================================================

def resolve_target(app, workbook=None, sheet=None):
    """
    Resolves target Workbook and Worksheet dynamically.
    - If workbook is None -> app.ActiveWorkbook
    - If workbook is a str -> matches by filename, path, or partial name.
      If not open and points to an existing file -> opens it in Excel.
    - If sheet is None -> wb.ActiveSheet
    - If sheet is a str/int -> wb.Worksheets(sheet)
    """
    if app.Workbooks.Count == 0:
        if workbook and os.path.exists(workbook):
            target_wb = app.Workbooks.Open(os.path.abspath(workbook))
        else:
            raise RuntimeError("Excel is open, but no workbooks are loaded.")
    else:
        target_wb = None
        if workbook is None:
            target_wb = app.ActiveWorkbook
        else:
            # Search open workbooks
            wb_str = str(workbook).strip().lower()
            for wb in app.Workbooks:
                if wb_str == wb.Name.lower() or wb_str == wb.FullName.lower():
                    target_wb = wb
                    break
            if target_wb is None:
                # Try partial match
                for wb in app.Workbooks:
                    if wb_str in wb.Name.lower():
                        target_wb = wb
                        break
            # If still None and workbook is an existing file path, open it
            if target_wb is None and os.path.isfile(str(workbook)):
                target_wb = app.Workbooks.Open(os.path.abspath(str(workbook)))
                
        if target_wb is None:
            target_wb = app.ActiveWorkbook

    # Resolve sheet
    if sheet is None:
        target_ws = target_wb.ActiveSheet
    else:
        try:
            target_ws = target_wb.Worksheets(sheet)
        except Exception:
            # Try case-insensitive match
            sheet_str = str(sheet).strip().lower()
            target_ws = None
            for ws in target_wb.Worksheets:
                if sheet_str == ws.Name.lower():
                    target_ws = ws
                    break
            if target_ws is None:
                target_ws = target_wb.ActiveSheet

    return target_wb, target_ws

# ==============================================================================
# Universal High-Level Operations
# ==============================================================================

def get_status():
    """Returns metadata about the active Excel instance and open workbooks."""
    def _action(app):
        wbs = []
        for wb in app.Workbooks:
            wbs.append({
                "name": wb.Name,
                "path": wb.FullName,
                "saved": wb.Saved,
                "sheets_count": wb.Worksheets.Count,
                "active_sheet": wb.ActiveSheet.Name if wb.ActiveSheet else None
            })
        active_wb = app.ActiveWorkbook.Name if app.ActiveWorkbook else None
        active_ws = app.ActiveSheet.Name if app.ActiveSheet else None
        return {
            "excel_version": app.Version,
            "active_workbook": active_wb,
            "active_sheet": active_ws,
            "workbooks_open_count": app.Workbooks.Count,
            "workbooks": wbs
        }
    return run_in_live_excel(_action)

def list_workbooks():
    """Lists all currently open workbooks in Excel."""
    def _action(app):
        return [wb.Name for wb in app.Workbooks]
    return run_in_live_excel(_action)

def list_sheets(workbook=None):
    """Lists all sheets in the target (or active) workbook."""
    def _action(app):
        wb, _ = resolve_target(app, workbook=workbook)
        sheets = []
        for s in wb.Worksheets:
            sheets.append({
                "name": s.Name,
                "visible": s.Visible == -1, # -1 is xlSheetVisible
                "used_range": s.UsedRange.Address
            })
        return {"workbook": wb.Name, "sheets": sheets}
    return run_in_live_excel(_action)

def read_range(cell_range: str, sheet=None, workbook=None):
    """Reads values from a cell range in RAM in a single fast COM roundtrip."""
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        val = ws.Range(cell_range).Value
        return val
    return run_in_live_excel(_action)

def write_range(start_cell: str, values, sheet=None, workbook=None):
    """Writes 2D list/tuple to a range starting at start_cell in a single COM call."""
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        if not values or not isinstance(values, (list, tuple)):
            raise ValueError("values must be a non-empty 2D list or tuple.")
        
        num_rows = len(values)
        num_cols = len(values[0]) if isinstance(values[0], (list, tuple)) else 1
        
        top_left = ws.Range(start_cell)
        r_start = top_left.Row
        c_start = top_left.Column
        target_range = ws.Range(
            ws.Cells(r_start, c_start),
            ws.Cells(r_start + num_rows - 1, c_start + num_cols - 1)
        )
        
        prev_update = app.ScreenUpdating
        app.ScreenUpdating = False
        try:
            target_range.Value = values
        finally:
            app.ScreenUpdating = prev_update
            
        return {
            "success": True,
            "range": target_range.Address,
            "sheet": ws.Name,
            "workbook": wb.Name
        }
    return run_in_live_excel(_action)

def read_dataframe(cell_range: str, sheet=None, workbook=None, has_headers: bool = True):
    """Reads range directly into a Pandas DataFrame in <0.05s."""
    try:
        import pandas as pd
    except ImportError:
        raise ImportError("pandas is required for read_dataframe. Install via: pip install pandas")

    raw_data = read_range(cell_range, sheet=sheet, workbook=workbook)
    if not raw_data:
        return pd.DataFrame()
    
    if isinstance(raw_data, (int, float, str)):
        return pd.DataFrame([[raw_data]])

    data = list(raw_data)
    if not data or not isinstance(data[0], (list, tuple)):
        data = [[x] for x in data]

    if has_headers and len(data) >= 2:
        headers = [str(c) if c is not None else f"Col_{i}" for i, c in enumerate(data[0])]
        return pd.DataFrame(data[1:], columns=headers)
    elif has_headers and len(data) == 1:
        headers = [str(c) if c is not None else f"Col_{i}" for i, c in enumerate(data[0])]
        return pd.DataFrame(columns=headers)
    else:
        return pd.DataFrame(data)

def set_columns_hidden(columns: str, hidden: bool = True, sheet=None, workbook=None):
    """Hides or unhides specific columns (e.g. 'C:H')."""
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        prev_update = app.ScreenUpdating
        app.ScreenUpdating = False
        try:
            ws.Columns(columns).EntireColumn.Hidden = bool(hidden)
        finally:
            app.ScreenUpdating = prev_update
        return {"success": True, "columns": columns, "hidden": bool(hidden), "sheet": ws.Name, "workbook": wb.Name}
    return run_in_live_excel(_action)

def set_rows_hidden(rows: str, hidden: bool = True, sheet=None, workbook=None):
    """Hides or unhides specific rows (e.g. '10:15')."""
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        prev_update = app.ScreenUpdating
        app.ScreenUpdating = False
        try:
            ws.Rows(rows).EntireRow.Hidden = bool(hidden)
        finally:
            app.ScreenUpdating = prev_update
        return {"success": True, "rows": rows, "hidden": bool(hidden), "sheet": ws.Name, "workbook": wb.Name}
    return run_in_live_excel(_action)

def autofit(columns=None, rows=None, sheet=None, workbook=None):
    """Auto-fits columns and/or rows live in RAM (<0.05s)."""
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        prev_update = app.ScreenUpdating
        app.ScreenUpdating = False
        try:
            if columns:
                ws.Columns(columns).AutoFit()
            if rows:
                ws.Rows(rows).AutoFit()
            if not columns and not rows:
                ws.UsedRange.Columns.AutoFit()
                ws.UsedRange.Rows.AutoFit()
        finally:
            app.ScreenUpdating = prev_update
        return {"success": True, "sheet": ws.Name, "workbook": wb.Name}
    return run_in_live_excel(_action)

def set_column_width(columns: str, width: float, sheet=None, workbook=None):
    """Sets specific column width (e.g. columns='C:H', width=12.0)."""
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        prev_update = app.ScreenUpdating
        app.ScreenUpdating = False
        try:
            ws.Columns(columns).ColumnWidth = float(width)
        finally:
            app.ScreenUpdating = prev_update
        return {"success": True, "columns": columns, "width": width, "sheet": ws.Name, "workbook": wb.Name}
    return run_in_live_excel(_action)

def format_cells(cell_range: str, sheet=None, workbook=None, **kwargs):
    """
    Applies comprehensive formatting to a range in RAM.
    kwargs:
      number_format: str (e.g. '#,##0.00;(#,##0.00);"-"' or '0.0%')
      font_name: str (e.g. 'Arial')
      font_size: float (e.g. 10.0)
      bold: bool
      italic: bool
      font_color: int (RGB COLORREF)
      bg_color: int (RGB COLORREF)
      halign: int (xlCenter, xlLeft, xlRight)
      valign: int (xlCenter, xlTop, xlBottom)
      wrap_text: bool
      indent: int
      row_height: float
    """
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        prev_update = app.ScreenUpdating
        app.ScreenUpdating = False
        try:
            rng = ws.Range(cell_range)
            if "number_format" in kwargs:
                rng.NumberFormat = kwargs["number_format"]
            if "font_name" in kwargs:
                rng.Font.Name = kwargs["font_name"]
            if "font_size" in kwargs:
                rng.Font.Size = kwargs["font_size"]
            if "bold" in kwargs:
                rng.Font.Bold = bool(kwargs["bold"])
            if "italic" in kwargs:
                rng.Font.Italic = bool(kwargs["italic"])
            if "font_color" in kwargs:
                rng.Font.Color = kwargs["font_color"]
            if "bg_color" in kwargs:
                rng.Interior.Color = kwargs["bg_color"]
            if "halign" in kwargs:
                rng.HorizontalAlignment = kwargs["halign"]
            if "valign" in kwargs:
                rng.VerticalAlignment = kwargs["valign"]
            if "wrap_text" in kwargs:
                rng.WrapText = bool(kwargs["wrap_text"])
            if "indent" in kwargs:
                rng.IndentLevel = int(kwargs["indent"])
            if "row_height" in kwargs:
                rng.Rows.RowHeight = float(kwargs["row_height"])
        finally:
            app.ScreenUpdating = prev_update
        return {"success": True, "range": cell_range, "sheet": ws.Name, "workbook": wb.Name}
    return run_in_live_excel(_action)

def save(workbook=None):
    """Saves the target (or active) workbook in RAM directly to disk."""
    def _action(app):
        wb, _ = resolve_target(app, workbook=workbook)
        wb.Save()
        return {"success": True, "saved_workbook": wb.Name, "path": wb.FullName}
    return run_in_live_excel(_action)

def eval_live(code_str: str, sheet=None, workbook=None):
    """
    Executes an arbitrary Python code block against (app, wb, ws) in live Excel.
    Variables 'app', 'wb', 'ws' are pre-populated in the local scope.
    """
    def _action(app):
        wb, ws = resolve_target(app, workbook, sheet)
        loc = {"app": app, "wb": wb, "ws": ws, "result": None}
        exec(code_str, globals(), loc)
        return loc.get("result")
    return run_in_live_excel(_action)
