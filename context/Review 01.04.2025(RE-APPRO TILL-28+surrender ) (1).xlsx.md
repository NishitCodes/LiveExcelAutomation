# Review 01.04.2025 Context & Automation Guide

* **Target File:** `Review 01.04.2025(RE-APPRO TILL-28+surrender ) (1).xlsx`
* **File Location:** `D:\Nishit\Budget Section\Important Expenditure Updates\2024-25\`
* **Automation Engine:** `D:\Nishit\Softwares\ExcelAutomation\excel_core.py`

---

## 1. Overview & Purpose
* **Domain:** Expenditure Review & Re-appropriation statement up to 01.04.2025 for Grant No. 017 / 016 (Ministry of Corporate Affairs).
* **Sheets Count:** 18 functional sheets.
* **Key Sheets:**
  * `FRONT SHEET `: Executive summary expenditure statement.
  * `OH-wise`: Comprehensive Object Head-wise breakdown across all sections and heads (~6,195 rows).
  * `Object Heads`: Dedicated master reference sheet listing standard DFPR 2024 Object Heads (Codes 01 to 77), official nomenclature, classification, workbook presence, and local variants.
  * `Report`: Statement of Grant Transaction (Controller Wise).
  * `Scheme-wise sheet`: Breakdowns by schemes.
  * Section/Office Sheets: `MCA`, `ROC`, `OL`, `RD`, `NFRA`, `SPECIAL COURTS`, `IEPF`, `IBBI`, `ICLS`, `SFIO`, `NCLT`, `NCLAT`, `MCA(HQ)`.

---

## 2. Object Heads Reference (`Object Heads` sheet)
* **Location:** Directly follows `OH-wise` in the tab order.
* **Dimensions:** Rows 1 to 45, Columns A to F.
* **Columns:**
  * **Col A:** S.No. (1 to 40)
  * **Col B:** Object Code (`01` to `77`, stored as Text `@` to preserve leading zeroes)
  * **Col C:** Object Head Description (Official DFPR 2024 Nomenclature)
  * **Col D:** Expenditure Classification (`Revenue Expenditure`, `Capital Outlay`, `Transfers & Recoveries`)
  * **Col E:** Present in Workbook (dynamic count of sheets utilizing the head)
  * **Col F:** Common Variant(s) in Workbook (e.g. `(SALARIES)`, `ALLOWANCES`, `NIP-OFFICE EXPENSES`)

---

## 3. Automation Rules & Conventions
1. **Live In-Memory Execution:** Always interact via `D:\Nishit\Softwares\ExcelAutomation\excel_core.py`.
2. **Never Close Excel:** The user keeps the file open on screen.
3. **Keep Working Directory Clean:** Never save scripts, logs, or markdown notes in `D:\Nishit\Budget Section\Important Expenditure Updates\2024-25\`.
