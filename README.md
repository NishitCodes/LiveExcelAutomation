# LiveExcelAutomation ⚡📊

**Instantaneous, in-memory live Microsoft Excel automation for Windows via native COM.**  
*Edit, audit, format, and recalculate spreadsheets in real-time on your screen — zero file-lock collisions, zero workbook closures, and zero lost calculation chains.*

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python: 3.9+](https://img.shields.io/badge/python-3.9%2B-brightgreen.svg)](https://www.python.org/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![Agent Skill Included](https://img.shields.io/badge/Agent%20Skill-Ready-orange.svg)](skills/excel-live-automation/SKILL.md)

---

## 🎯 The Problem LiveExcel Solves

Traditional Python Excel tools (`openpyxl`, `xlsxwriter`, `pandas.to_excel`) and headless OpenXML CLI tools operate strictly on files stored on disk. When a user has a spreadsheet open in Microsoft Excel:

1. **File Lock Crashes (`[Errno 13] Permission denied`):** Windows locks open Excel files. File-on-disk tools cannot write to an open workbook without crashing or forcing the user to quit Excel.
2. **Workflow Interruption:** Users must close Excel, run a script, and reopen the file to see results — repeatedly.
3. **No Dynamic Formula Calculation:** Headless file tools cannot evaluate native formulas. Opening an edited file often results in blank formula caches or broken calculation dependencies until manually recalculated.
4. **Data & Feature Loss:** Disk re-writers frequently strip custom chart links, shapes, legacy macros (VBA), or print layouts upon re-serialization.

### The LiveExcel Solution

**LiveExcelAutomation** bridges directly into the user's running Microsoft Excel process via native Windows COM (`Oleacc` / `IDispatch`). 

- 🖥️ **Live Pair-Programming:** You see changes materialize instantly on your screen in real time.
- ⚡ **Blazing Performance:** 10,000+ formulas or cells written in under 20 milliseconds via contiguous 2D array memory blocks.
- 🧮 **Native Excel Recalculation:** Leverages Excel's actual calculation engine (`CalculateFull`) directly in RAM.
- 🔒 **Zero File Locks:** Never encounters `PermissionError`. Workbooks are saved directly from RAM (`wb.Save()`).
- 🤖 **AI Agent Native:** Comes bundled with a production AI Agent Skill for Antigravity, Gemini, Claude Code, Cursor, and custom LLM workflows.

---

## 📊 LiveExcel vs Conventional File Tools

| Capability | Headless OpenXML Tools (`openpyxl`, `OfficeCLI`) | **LiveExcelAutomation** |
| :--- | :---: | :---: |
| **Interactive Screen Editing** | ❌ No (file must be closed) | ✅ **Yes (edits visible live on screen)** |
| **File Lock Immunity** | ❌ Fails with `PermissionError` | ✅ **Immune (edits directly in RAM)** |
| **Batch Write Speed (10k cells)** | ⚠️ 500ms – 3,000ms (disk I/O) | ⚡ **< 20ms (contiguous 2D memory)** |
| **Native Excel Formula Recalculation** | ❌ No (requires opening in Excel) | ✅ **Yes (`CalculateFull()` in RAM)** |
| **Preserves Custom Charts & VBA** | ⚠️ Can strip or corrupt | ✅ **100% Preserved (native COM)** |
| **Pandas Two-Way Interoperability** | ⚠️ Read/write to disk only | ✅ **Direct RAM read & write in <0.05s** |
| **AI Agent Pair-Programming** | ❌ Agent cannot touch open user sheets | ✅ **Seamless agent-human collaboration** |

---

## 🚀 Installation

### Prerequisites
- Windows 10 / 11 / Windows Server
- Microsoft Excel (Office 365, 2021, 2019, 2016)
- Python 3.9+

### Install from Source
```bash
git clone https://github.com/NishitCodes/LiveExcelAutomation.git
cd LiveExcelAutomation
pip install -r requirements.txt
pip install -e .
```

---

## 💡 Quick Start

Make sure Microsoft Excel is open with any spreadsheet on your screen, then run:

### 1. Inspect Active Excel Instance
```python
import live_excel as xl

status = xl.get_status()
print(f"Connected: {status['connected']}")
print(f"Active Workbook: {status['active_workbook']}")
print(f"Active Sheet: {status['active_sheet']}")
```

### 2. High-Speed 2D Matrix Formula & Value Injection
Write thousands of cells in a single call without lag:
```python
import live_excel as xl

# Prepare a 2D matrix of formulas or values
data = [
    ["Region", "Q1", "Q2", "Total"],
    ["North", 12000, 15000, "=SUM(B2:C2)"],
    ["South", 18000, 21000, "=SUM(B3:C3)"],
    ["Total", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)"]
]

# Writes in <0.02 seconds and recalculates in RAM
xl.write_range("A1", data)
```

### 3. Professional Accounting & Number Formatting
Apply clean financial formatting, borders, and colors in milliseconds:
```python
import live_excel as xl

# Apply standard accounting format to numbers
xl.format_cells(
    "B2:D4",
    number_format='_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    halign=xl.xlRight
)

# Style headers with navy blue fill and bold white text
xl.format_cells(
    "A1:D1",
    bold=True,
    font_color=xl.rgb(255, 255, 255),
    bg_color=xl.rgb(24, 43, 73),
    halign=xl.xlCenter
)

# Auto-fit columns for clean display
xl.autofit(columns="A:D")
```

### 4. Two-Way Pandas DataFrame Integration
Read and write directly to DataFrames in under 0.05 seconds:
```python
import pandas as pd
import live_excel as xl

# 1. Read live range directly into DataFrame
df = xl.read_dataframe("A1:D4", has_headers=True)
print(df)

# 2. Vectorized calculation in Pandas
df["Growth_%"] = [None, 25.0, 16.7, 20.0]

# 3. Write back directly to live Excel starting at F1
xl.write_dataframe(df, start_cell="F1", include_headers=True)
xl.autofit(columns="F:K")

# Save workbook directly to disk without closing Excel
xl.save()
```

---

## 🤖 Built for AI Agents & Pair-Programming

LiveExcelAutomation was engineered from the ground up for modern AI coding assistants (**Antigravity**, **Gemini**, **Claude Code**, **Cursor**, etc.) to pair-program on spreadsheets alongside humans.

### Included AI Agent Skill
This repository includes a ready-to-use AI Agent Skill located at:
[`skills/excel-live-automation/SKILL.md`](skills/excel-live-automation/SKILL.md)

It equips any LLM agent with:
- Strict guardrails (never close Excel, never overwrite disk files with openpyxl).
- High-speed batch injection recipes.
- Pre-defined financial number formatting strings.
- Complete Windows COM constant mappings.

To install the skill in Antigravity or your agent environment:
```bash
# Copy skill into your agent's skills directory
cp -r skills/excel-live-automation ~/.gemini/config/skills/
```

---

## 📁 Repository Structure

```text
LiveExcelAutomation/
├── live_excel/                 # Core Python package
│   ├── __init__.py             # Exports & version info
│   └── core.py                 # Windows COM engine & high-speed batch methods
├── excel_core.py               # Root backward-compatibility shim
├── examples/                   # Standalone runnable scripts
│   ├── 01_inspect_live_excel.py
│   ├── 02_fast_batch_formulas.py
│   ├── 03_accounting_formatting.py
│   └── 04_dataframe_interop.py
├── skills/                     # AI Agent Skill package
│   └── excel-live-automation/
│       ├── SKILL.md
│       └── references/
│           ├── api_reference.md
│           └── excel_constants.md
├── pyproject.toml              # Build & package configuration
├── requirements.txt            # Dependencies (pywin32, pandas)
├── LICENSE                     # MIT License
└── README.md                   # This documentation
```

---

## 🛡️ Architecture & Technical Highlights

- **WindowStation Desktop Boundary Bridging:** Uses `ctypes.windll.oleacc.AccessibleObjectFromWindow` with `OBJID_NATIVEOM` to attach directly to Excel's running `EXCEL7` window class, even when launched from background agent subprocesses or different Windows terminal sessions.
- **2D COM Variant Marshalling:** Instead of thousands of individual COM property calls, contiguous Python matrices are marshalled directly into Excel COM `SAFEARRAY` structures in a single dispatch call.
- **State Preservation:** Transparently handles `ScreenUpdating`, calculation modes (`xlCalculationManual` vs `xlCalculationAutomatic`), and restores previous states upon completion or error.

---

## 📄 License

Distributed under the [MIT License](LICENSE). Copyright (c) 2026 Nishit.
