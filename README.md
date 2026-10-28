# Universal Excel Live In-Memory Automation System

**Location:** `D:\Nishit\Softwares\ExcelAutomation\`

This system provides instantaneous (<0.15s), safe, in-memory automation for any Excel workbook across the user's computer while Excel remains open on screen.

---

## 1. Directory Structure

```
D:\Nishit\Softwares\ExcelAutomation\
├── excel_core.py          # Universal live in-memory COM automation engine
├── README.md              # System architecture and agent usage documentation
└── context\               # Persistent knowledge bank for workbooks
    └── <WorkbookName>.xlsx.md   # Structural schema, rules, and history for specific workbooks
```

---

## 2. Core Operational Rules for AI Agents

1. **Keep User Project Folders Clean:**
   * NEVER leave `.py`, `.bat`, `.md`, or cache files inside user working folders (such as `D:\Nishit\Budget Section\BigSheet\`).
   * Working directories must contain **only the user's Excel files**.
2. **Context First:**
   * Before modifying any workbook, check `D:\Nishit\Softwares\ExcelAutomation\context\<WorkbookName>.xlsx.md`.
   * If a context file exists, follow all recorded rules (hidden columns, master date parameters, formulas, formats).
   * If working on a new workbook, create an initial context document in `context\` once understood.
3. **In-Memory Automation ONLY:**
   * Use `excel_core.py` to interact directly with Excel in RAM.
   * Do not rewrite workbooks to disk using third-party libraries (e.g. `openpyxl`) while they are open in Excel.
   * Never close the user's open Excel window.

---

## 3. Quick Python API Reference

```python
import sys
sys.path.insert(0, r"D:\Nishit\Softwares\ExcelAutomation")
import excel_core

# Inspect active Excel instance
info = excel_core.get_status()
sheets = excel_core.list_sheets()

# High-speed data operations
matrix = excel_core.read_range("A1:M50")
df = excel_core.read_dataframe("A1:M50")
excel_core.write_range("C5", [[10.5, 20.0], [30.0, 45.2]])

# Layout & Formatting
excel_core.set_columns_hidden("C:H", hidden=True)
excel_core.autofit(columns="A:Z")
excel_core.format_cells(
    "C5:M50",
    number_format='#,##0.00;(#,##0.00);"-"',
    halign=excel_core.xlRight
)

# Save live without closing
excel_core.save()
```
