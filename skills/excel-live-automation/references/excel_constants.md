# Excel COM Constants & ColorIndex Reference

This guide provides common constants used when automating Microsoft Excel via Python COM (`win32com` / `pythoncom`).

---

## 1. Border & LineStyle Constants

| Constant | Value | Description |
| :--- | :---: | :--- |
| `xlContinuous` | `1` | Solid continuous line |
| `xlDash` | `-4115` | Dashed line |
| `xlDot` | `-4118` | Dotted line |
| `xlDouble` | `-4119` | Double line (Accounting underline) |
| `xlLineStyleNone` | `-4142` | No border line |

### Border Edges
| Constant | Value | Description |
| :--- | :---: | :--- |
| `xlEdgeLeft` | `7` | Left edge border |
| `xlEdgeTop` | `8` | Top edge border |
| `xlEdgeBottom` | `9` | Bottom edge border |
| `xlEdgeRight` | `10` | Right edge border |
| `xlInsideVertical` | `11` | Interior vertical gridlines |
| `xlInsideHorizontal` | `12` | Interior horizontal gridlines |

### Border Weights
| Constant | Value | Description |
| :--- | :---: | :--- |
| `xlHairline` | `1` | Very thin hairline border |
| `xlThin` | `2` | Standard thin border |
| `xlMedium` | `-4138` | Medium bold border |
| `xlThick` | `4` | Heavy thick border |

---

## 2. Alignment Constants

| Constant | Value | Description |
| :--- | :---: | :--- |
| `xlGeneral` | `1` | Default alignment |
| `xlLeft` | `-4131` | Left horizontal alignment |
| `xlCenter` | `-4108` | Center horizontal or vertical alignment |
| `xlRight` | `-4152` | Right horizontal alignment |
| `xlTop` | `-4160` | Top vertical alignment |
| `xlBottom` | `-4107` | Bottom vertical alignment |

---

## 3. Fill Colors (`Interior.ColorIndex`)

Excel provides a standard 56-color palette accessible via `Interior.ColorIndex`:

| ColorIndex | Visual Color | Typical Budget Sheet Usage |
| :---: | :--- | :--- |
| `-4142` | **`xlNone`** | Transparent / No Fill |
| `49` | **Dark Blue** (`#003366`) | Primary Table Headers (paired with White Bold Text) |
| `37` | **Light Sky Blue** (`#99CCFF`) | Subtotal / Grand Total Rows |
| `24` | **Soft Lavender** (`#CCCCFF`) | Category Subheaders |
| `6` | **Bright Yellow** (`#FFFF00`) | Draft Notes / Review Highlights |
| `15` | **Light Gray** (`#C0C0C0`) | Inactive / Locked Columns |
