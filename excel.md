# 📊 Excel Practice Sheets

Personal Excel practice files — formulas, charts, pivot tables and more.



# 📊 Microsoft Excel — Practice & Revision

A hands-on Excel learning project covering everything from basic formulas to advanced functions, charts, and data analysis.

---

## 📁 Files in this Repo

| File | Description |
|---|---|
| `Excel_Practice_Exercise.xlsx` | Full practice dataset with 30 tasks |
| `Excel_Revision_Notes.pdf` | Complete revision notes (14 chapters) |

---

## ✅ What I Practiced

### 🗂️ Worksheet Basics
- Creating and navigating worksheets
- Entering data, using AutoFill and drag-to-fill
- Working with cell ranges (e.g. `A1:A5`)
- Understanding the Ribbon and Workbench

### 🎨 Formatting
- Bold text, font colours, fill colours
- Currency formatting (₹)
- Freezing panes and column widths

### 🔢 Formulas & Operators
- Math operators: `+`, `-`, `*`, `/`, `^`
- Comparison operators: `=`, `>`, `<`, `>=`, `<=`, `<>`
- Text joining with `&` → `=A2&" "&E2`
- Absolute vs relative cell references (`$A$1`)

### 🧮 Functions Practised

| Category | Functions |
|---|---|
| Math | `SUM`, `SUMIF`, `SUMIFS` |
| Counting | `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS` |
| Logical | `IF`, `Nested IF`, `IFS`, `AND`, `OR` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP` |
| Statistics | `AVERAGE`, `MEDIAN`, `MIN`, `MAX`, `MINIFS`, `MAXIFS`, `STDEV.P`, `STDEV.S`, `QUARTILE`, `MODE` |
| Text | `TEXTJOIN`, `TEXTSPLIT`, `RIGHT`, `LEFT`, `MID`, `FIND`, `SEARCH` |
| Date & Time | `DATE`, `DAY`, `MONTH`, `YEAR`, `TODAY`, `HOUR`, `MINUTE`, `TEXT` |
| Array | `UNIQUE`, `SORT` (dynamic arrays) |

### 📊 Charts Created
- **Line Chart** — trends over time
- **Pie Chart** — employee count per department
- **Column Chart** — Total Sales comparison across employees
- **Scatter Plot** — comparing two numeric values
- **Histogram** — frequency distribution
- **Box & Whisker** — data spread and outliers
- **Sparklines** — inline mini-charts per row
- **Map Chart** — geographic data visualisation

### 🔍 Data Tools
- AutoFilter and multi-column filtering
- Sorting (single and multi-level)
- `IFERROR` for handling errors gracefully

---

## 🗃️ Practice Dataset

The dataset contains **15 employees** with the following fields:

```
EmpID | Name | Department | Region | Gender | JoinDate
Sales_Jan | Sales_Feb | Sales_Mar | Product | UnitPrice | UnitsSold | Bonus%
```

### Tasks completed (30 total):
- ✅ Added calculated columns: `Total_Sales`, `Avg_Sales`, `Revenue`, `Bonus_Amount`
- ✅ Extracted data: Year from date, last 3 chars of EmpID
- ✅ Logical labels: Top Performer / Regular, Grades A–D
- ✅ Summary stats: MAX, MIN, MEDIAN, STDEV, QUARTILE
- ✅ Lookup: Found employee details using VLOOKUP & XLOOKUP
- ✅ Dynamic arrays: UNIQUE departments, SORT names
- ✅ Charts: Pie, Column, Box & Whisker, Sparklines

---

## 📝 Key Learnings

- `STDEV.S` is used for a **sample**, `STDEV.P` for the full **population**
- `XLOOKUP` is more powerful than `VLOOKUP` — searches any direction, no column index needed
- Dynamic array functions like `UNIQUE` and `SORT` **spill** results automatically
- When `YEAR()` shows a wrong date format, change the cell format to **General**
- Always start a formula with `=`, otherwise Excel treats it as plain text

---



*Practised as part of my data skills learning journey 🚀*

## 🛠️ Tools
- Microsoft Excel