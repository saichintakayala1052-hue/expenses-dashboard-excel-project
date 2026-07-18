
# 📊 Excel VBA Expenses Dashboard (FY 2025)

Expense &amp; Budget Analysis Dashboard (Excel VBA &amp; Macros)  The Expense &amp; Budget Analysis Dashboard is a dynamic financial reporting solution developed in Microsoft Excel using VBA (Visual Basic for Applications) and advanced Excel features.A fully automated, self-contained VBA macro that turns a raw expenses list into a polished, interactive Excel dashboard — complete with PivotTables, KPI cards, charts, and slicers. No add-ins, no external dependencies, just one macro.

![Excel](https://img.shields.io/badge/Excel-2016%2B-217346?logo=microsoftexcel&logoColor=white)
![VBA](https://img.shields.io/badge/Language-VBA-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)

---

## ✨ Features

- **One-click build** — run a single macro (`BuildExpensesDashboard`) and the entire dashboard is generated from scratch
- **Live progress feedback** — status bar messages and screen animation show exactly what's happening at every step
- **7 PivotTables** — Department, Category, Monthly Trend, Approved Status, KPI Totals, Payment Method, Budget vs Actual
- **6 KPI cards** — Total Expenses, Budget Allocated, Budget Remaining, Transactions, Approval Rate, Avg Expense
- **5 charts** — column, doughnut, line, and bar charts, all styled with a consistent green theme
- **2 slicers** — Month and Department, connected to every PivotTable for synchronized filtering
- **Auto-formatting** — Excel Table, number formats, header styling, column widths, and print setup all handled automatically
- **Idempotent** — safe to re-run any time; old Dashboard/Pivot sheets are rebuilt cleanly

---

## 📋 Requirements

- Excel for Windows, **Office 2016 or later**
- Macros enabled (`.xlsm` workbook)

---

## 📁 Data Setup

Before running the macro, add a sheet named **`Data`** with these exact column headers in Row 1 (order doesn't matter, all must be present):

| Header |
|---|
| Expense ID |
| Date |
| Department |
| Category |
| Amount (INR) |
| Payment Method |
| Approved |
| Budget Allocated |
| Budget Remaining |

> A `Month` column is added automatically by the macro if it doesn't already exist.

---

## 🚀 Installation & Usage

1. Open your workbook and confirm the `Data` sheet is set up as described above.
2. Save the file as a macro-enabled workbook (**.xlsm**), if it isn't already.
3. Press **Alt + F11** to open the VBA editor.
4. **Insert → Module**, then paste in the full code from [`dashboard.bas`](./dashboard.bas).
5. Close the VBA editor.
6. Press **Alt + F8**, select **`BuildExpensesDashboard`**, and click **Run**.
7. Watch the status bar for live progress — a confirmation message appears when the build finishes.

To rebuild after updating your data, just run the macro again.

---

## 🗂️ What Gets Created

| Sheet | Purpose |
|---|---|
| **Data** | Your source data, converted into an Excel Table (`tblExpenses`) with formatting applied |
| **Pivot** | 7 PivotTables + hidden chart-helper ranges (not meant to be viewed directly) |
| **Dashboard** | The final presentation sheet — header banner, KPI cards, charts, and slicers |

---

## 🎨 Design Reference

| Property | Value |
|---|---|
| Font | Aptos, 10pt |
| Canvas background | `RGB(246, 250, 244)` |
| Header gradient | `RGB(6,66,25)` → `RGB(17,112,31)` |
| Primary chart green | `RGB(10,93,30)` |
| Zoom level | 70% |
| Slicer style | `SlicerStyleLight6` |
| Money format | `#,##0,"K"` |

---

## 🏗️ Architecture

```
BuildExpensesDashboard()          ← entry point
├── PrepareDataSheet()            ← builds tblExpenses, formats columns
├── BuildPivotSheet()             ← creates 7 PivotTables + chart helpers
│   ├── CreatePivot_ExpenseByDepartment
│   ├── CreatePivot_ExpenseByCategory
│   ├── CreatePivot_MonthlyTrend
│   ├── CreatePivot_ApprovedStatus
│   ├── CreatePivot_KPI
│   ├── CreatePivot_PaymentMethod
│   └── CreatePivot_BudgetVsActual
└── BuildDashboardSheet()         ← draws the dashboard
    ├── CreateHeader
    ├── CreateKPICards (×6)
    ├── CreateDashboardCharts (×5)
    ├── CreateMonthSlicer
    └── CreateDepartmentSlicer
```

---

## ⚠️ Notes

- `Application.ScreenUpdating` stays `True` throughout so you can watch the build happen live — this makes it slightly slower than a "silent" build, by design.
- Re-running the macro deletes and rebuilds the `Pivot` and `Dashboard` sheets from scratch; it does not touch your original `Data` sheet content beyond formatting and the added `Month` column.
- Column format is fixed to **INR** number formatting (`Amount (INR)`); relabel or adjust the format string in the code if you need a different currency.

---

## 📄 License

MIT — free to use, modify, and distribute.
