# 📊 Excel Practice Sheets

Personal Excel practice files — formulas, charts, pivot tables and more.




# 📊 Microsoft Excel — Practice 2

An advanced hands-on Excel practice project covering SUBTOTAL, AGGREGATE, Conditional Formatting, and Pivot Tables using a real-world Store Sales dataset.

---

## 📁 Files in this Repo

| File | Description |
|---|---|
| `Excel_Practice_2.xlsx` | Full practice dataset with 30 tasks across 5 sheets |

---

## 🗃️ Dataset Overview

**Store Sales Data** — 20 orders from 5 salespersons across 4 regions and 4 months.

| Column | Description |
|---|---|
| OrderID | Unique order identifier |
| SalesPerson | Name of the salesperson |
| Category | Electronics / Clothing / Furniture |
| Region | North / South / East / West |
| Month | Jan / Feb / Mar / Apr |
| Units | Number of units sold |
| UnitPrice | Price per unit (₹) |
| Discount% | Discount applied |
| CustomerRating | Rating out of 5 |
| DeliveryDays | Days taken to deliver |
| Returned | Yes / No |

---

## ✅ What I Practised

### 🔢 Section 1 — Computed Columns
Added the following calculated columns to the dataset:

- **GrossRevenue** = Units × UnitPrice
- **DiscountAmt** = GrossRevenue × Discount% / 100
- **NetRevenue** = GrossRevenue − DiscountAmt
- **PerformanceTag** — High / Medium / Low using IFS based on NetRevenue
- **RatingGrade** — Excellent / Good / Poor using nested IF based on CustomerRating
- **Return or Not** — Returned / Kept using IF on the Returned column

---

### 📊 Section 2 — SUBTOTAL Function

`=SUBTOTAL(function_number, range)`

| Function# | What it does |
|---|---|
| 1 | AVERAGE |
| 2 | COUNT |
| 3 | COUNTA |
| 4 | MAX |
| 5 | MIN |
| 9 | SUM |

**Key learning:** SUBTOTAL automatically updates when a filter is applied — unlike SUM which always includes all rows.

Practised:
- `=SUBTOTAL(9, N2:N21)` → SUM of NetRevenue (updates with filter)
- `=SUBTOTAL(1, I2:I21)` → AVERAGE CustomerRating
- `=SUBTOTAL(4, N2:N21)` → MAX NetRevenue
- `=SUBTOTAL(5, J2:J21)` → MIN DeliveryDays
- `=SUBTOTAL(103, A2:A21)` → COUNT ignoring hidden rows

---

### ⚙️ Section 3 — AGGREGATE Function

`=AGGREGATE(function_number, option, range)`

| Option | What it ignores |
|---|---|
| 0 | Nothing |
| 1 | Hidden rows |
| 5 | Error cells |
| 7 | Hidden rows + Errors both |

**Key difference vs SUBTOTAL:** AGGREGATE can ignore ERROR cells — SUBTOTAL cannot!

Practised:
- `=AGGREGATE(9,5,N2:N21)` → SUM ignoring errors
- `=AGGREGATE(1,7,I2:I21)` → AVERAGE ignoring hidden rows AND errors
- `=AGGREGATE(4,5,N2:N21)` → MAX ignoring errors
- `=AGGREGATE(14,6,N2:N21,1)` → LARGE (1st biggest value)

---

### 🎨 Section 4 — Conditional Formatting

Applied 6 types of Conditional Formatting on the dataset:

| Type | Applied On | Effect |
|---|---|---|
| **Color Scale** | CustomerRating (I) | Green=high, Yellow=medium, Red=low ratings |
| **Data Bars** | NetRevenue (N) | Mini bar chart inside each cell |
| **Highlight Cell Rules** | CustomerRating (I) | Red fill if rating < 3.5 |
| **Icon Sets** | CustomerRating (I) | 🟢🟡🔴 traffic lights per rating |
| **Top/Bottom Rules** | NetRevenue (N) | Green highlight on top 3 values |
| **Formula Rule** | Entire row (A:P) | Yellow row if Returned = "Yes" |

Formula used for row highlight:
```
=$K2="Yes"    → applied to range A2:P21
```
> The $ before K locks the column but lets the row number adjust for each row.

---

### 🔄 Section 5 — Pivot Tables

**Created 4 Pivot Tables:**

| Pivot | Rows | Columns | Values |
|---|---|---|---|
| 1 | SalesPerson | — | Sum of NetRevenue |
| 2 | Category | — | Count of OrderID |
| 3 | Region | — | Average CustomerRating |
| 4 | Category | Month | Sum of NetRevenue (Cross Tab) |

**Additional Pivot features practised:**
- ✅ **Slicer** — Visual filter buttons for Category field
- ✅ **Value Field Settings** — Changed SUM to AVERAGE
- ✅ **Sort** — Sorted NetRevenue largest to smallest
- ✅ **Pivot Chart** — Bar chart linked to pivot table
- ✅ **Slicer + Chart** — Both update together when slicer is clicked

---

## 💡 Key Learnings

- `SUBTOTAL` recalculates automatically when rows are filtered — `SUM` does not
- `AGGREGATE` is safer than `SUBTOTAL` when data has error cells
- Conditional Formatting formula rules need `$` to lock the column: `=$K2="Yes"`
- Pivot Tables summarise large datasets instantly — no formulas needed
- Slicers make Pivot Tables interactive with one click
- Pivot Charts update automatically when the Pivot Table or Slicer changes



*Part of my Data Analyst Excel learning journey 🚀*


## 🛠️ Tools
- Microsoft Excel
