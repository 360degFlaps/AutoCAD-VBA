# AutoCAD Terrazzo Area Table Generator (VBA)

This VBA script for AutoCAD automates the process of calculating and exporting floor areas based on **closed polylines** and their corresponding **text labels** (such as plot numbers). It creates a table inside AutoCAD and also exports the data to **Microsoft Excel**.

Ideal for **quantity surveyors** and **CAD technicians** doing flooring area takeoffs or plot area summaries.

---

## Features

- Matches each **text label** to its enclosing **closed polyline**
- Calculates the **area** of each closed polyline
- Inserts an **AutoCAD table** listing Plot No. and Area
- Exports the same data to **Excel** for further use
- Highlights matched polylines and texts in **color**

---

## How It Works

1. **User selects** a group of `Text` and `LWPolyline` entities.
2. Script filters only:
   - Closed polylines (`AcadLWPolyline.Closed = True`)
   - `AcadText` entities
3. It checks which text lies **inside** which polyline.
4. Creates:
   - Table inside AutoCAD Model Space
   - Spreadsheet in Excel

---

## Example Output (in AutoCAD)

| Plot No. | Area (sq.units) |
|----------|------------------|
| A01      | 122.56           |
| A02      | 115.34           |
| A03      | 134.89           |

---

## Usage Instructions

1. Open AutoCAD.
2. Press `Alt + F11` to open the **VBA Editor**.
3. Insert a new **Module** and paste the code.
4. Run the macro: `GenerateTerrazzoAreaTable`
5. **Select** the `Text` and `Polyline` objects on screen when prompted.
6. Script will process and generate:
   - A table inside AutoCAD
   - A new Excel sheet with results

---

## Requirements

- **AutoCAD** with VBA support enabled.
- **Microsoft Excel** (installed and accessible via COM).
- Entities must include:
  - At least **one closed polyline**
  - At least **one text label** within that polyline

---

## File Overview

| File                        | Description                            |
|-----------------------------|----------------------------------------|
| `GenerateTerrazzoArea.bas` | VBA Module containing the macro        |
| `README.md`                 | Documentation                         |

---

## Notes

- The function `PointInPolyline` uses the **ray-casting algorithm** to determine if a text lies inside a p
