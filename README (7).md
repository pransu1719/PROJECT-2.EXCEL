<div align="center">

# -- ! Sales Data Analyzer ! --
### *Excel-Based Sales Performance, Statistics & What-If Analysis*

[![Excel](https://img.shields.io/badge/Excel-2007%2B-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Formulas](https://img.shields.io/badge/Formulas-SUMIFS%20%2F%20INDEX%20MATCH-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Pivot](https://img.shields.io/badge/Pivot-Dashboard-4CAF50?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Stats](https://img.shields.io/badge/Analysis-Descriptive%20%26%20Regression-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"Raw rows become insight the moment a formula asks the right question."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📊 Sheet 1 — Sales Data](#-sheet-1--sales-data)
- [👑 Sheet 2 — High Value Customers](#-sheet-2--high-value-customers)
- [📈 Sheet 3 — Descriptive Statistics](#-sheet-3--descriptive-statistics)
- [📉 Sheet 4 — Regression Analysis](#-sheet-4--regression-analysis)
- [🎛️ Sheet 5 — What-If Analysis](#️-sheet-5--what-if-analysis)
- [📋 Sheet 6 — Pivot Dashboard](#-sheet-6--pivot-dashboard)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)

---

## 📌 Overview

The **Sales Data Analyzer** is an Excel workbook (`PR_2_ANALYZER.xlsx`) built to turn a raw sales transaction log into a full analytical picture — profitability, top customers, descriptive statistics, a regression model, a what-if simulator, and a pivot-powered dashboard — all driven by native Excel formulas and PivotTables, no external tools required.

This project is designed to:
- Practice formula-driven analysis (`SUMIFS`, `INDEX`/`MATCH`, `AVERAGE`, `STDEV`, etc.) on real-shaped transactional data
- Build a repeatable structure for customer and regional sales analysis
- Apply descriptive and regression statistics to a business dataset
- Simulate "what-if" scenarios on pricing, discount, and quantity
- Summarize everything into a single interactive Pivot Dashboard

---

## 🎯 Problem Statement

> **Objective:** Analyze a multi-region sales transaction log to surface top customers, profitability drivers, statistical trends, and forecast scenarios — entirely within Excel.

Given 200 rows of transaction-level sales data (customer, region, product, sales, quantity, discount, profit, unit price), the workbook must classify and summarize the data across six purpose-built sheets, each answering a different business question.

| 📂 Sheet | 📄 Type | 🔍 Description |
|----------|---------|-----------------|
| Sales Data | Raw Data | Source transaction log — the single source of truth |
| High Value Customers | Analysis | Identifies top-spending / most profitable customers |
| Descriptive Statistics | Analysis | Summary statistics across sales, profit, and margin |
| Regression Analysis | Analysis | Models the relationship between sales drivers and profit |
| What-If Analysis | Simulation | Tests hypothetical pricing/discount/quantity scenarios |
| Pivot Dashboard | Reporting | Interactive PivotTable/PivotChart summary view |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🗃️ **200-Row Transaction Log** | Realistic multi-region, multi-product sales dataset |
| 👑 **High-Value Customer Detection** | Ranks customers by spend and profitability |
| 📊 **Descriptive Statistics** | Mean, median, standard deviation, and distribution summaries |
| 📉 **Regression Modeling** | Quantifies how sales/quantity/discount relate to profit |
| 🎛️ **What-If Simulation** | Lets you flex key inputs and see the projected outcome |
| 📋 **Pivot Dashboard** | One-view rollup by region, product, and customer |
| 🧮 **Derived Metrics** | `Margin Rate` and `Growth vs AVG` computed per transaction |
| 📐 **Formula-Driven** | No hardcoded results — every output recalculates from source data |

---

## 🏗️ Project Structure

```
📦 PR_2_ANALYZER/
│
├── 📄 PR_2_ANALYZER.xlsx          ← Main workbook (entry point)
│   ├── 📑 Sales Data              ← Raw transaction log (200 rows × 12 cols)
│   ├── 📑 High value customers    ← Top customer analysis
│   ├── 📑 Descriptive Statistics  ← Summary statistics
│   ├── 📑 Regression Analysis     ← Profit driver model
│   ├── 📑 WHAT IF Analysis        ← Scenario simulator
│   └── 📑 Pivot Dashboard         ← PivotTable/PivotChart summary
│
└── 📄 README.md                   ← Project documentation
```

---

## 🔄 Project Workflow

```
Raw Sales Data
      │
      ▼
┌─────────────────────────────┐
│   Sales Data (source sheet) │  ← Customer, Region, Product, Sales, Profit...
└────────────┬────────────────┘
             │
   ┌─────────┼───────────────┬────────────────┬──────────────┐
   ▼         ▼                ▼                ▼              ▼
┌────────┐ ┌───────────────┐ ┌─────────────┐ ┌────────────┐ ┌─────────────┐
│ High   │ │ Descriptive   │ │ Regression  │ │ What-If    │ │ Pivot       │
│ Value  │ │ Statistics    │ │ Analysis    │ │ Analysis   │ │ Dashboard   │
│Customer│ │               │ │             │ │            │ │             │
└────┬───┘ └───────┬───────┘ └──────┬──────┘ └─────┬──────┘ └──────┬──────┘
     │             │                │              │               │
     └─────────────┴────────────────┴──────────────┴───────────────┘
                                     │
                                     ▼
                        Consolidated Business Insights
```

---

## 📊 Sheet 1 — Sales Data

The foundation sheet. Each row is one transaction with 12 fields:

| Column | Field | Description |
|--------|-------|--------------|
| A | Customer ID | Unique customer identifier (e.g., `CUST019`) |
| B | Customer Name | Full customer name |
| C | Region | West / East / South / North / Central |
| D | Product | Books, Clothing, Electronics, Furniture, Office Supplies |
| E | Sales | Transaction sales value |
| F | Quantity | Units sold |
| G | Discount | Discount rate applied |
| H | OrderDate | Date of the transaction |
| I | Profit | Transaction profit |
| J | UnitPrice | Price per unit |
| K | Margin Rate | *Derived* — Profit ÷ Sales |
| L | Growth vs AVG | *Derived* — how the transaction compares to the average |

> `Margin Rate` and `Growth vs AVG` are calculated fields layered on top of the raw log, feeding the downstream analysis sheets.

---

## 👑 Sheet 2 — High Value Customers

Surfaces the customers driving the most revenue and profit, typically via `SUMIFS`/`INDEX`-`MATCH` aggregation and a ranking/sort of total spend per `Customer ID`.

---

## 📈 Sheet 3 — Descriptive Statistics

Summarizes the shape of the Sales, Profit, and Margin Rate columns — mean, median, standard deviation, min/max, and quartiles — to characterize typical transaction size and variability.

---

## 📉 Sheet 4 — Regression Analysis

Models the relationship between drivers (Sales, Quantity, Discount, UnitPrice) and Profit, to quantify which variables move profitability the most.

---

## 🎛️ Sheet 5 — What-If Analysis

A scenario sandbox — adjusting inputs like discount rate, unit price, or quantity to see the projected effect on sales and profit, using Excel's What-If tools (Data Tables / Goal Seek / Scenario Manager).

---

## 📋 Sheet 6 — Pivot Dashboard

A PivotTable/PivotChart rollup summarizing sales and profit by Region, Product, and Customer for at-a-glance reporting.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| 📊 **Microsoft Excel 2007+** | Core workbook platform |
| ➗ **SUMIFS / INDEX-MATCH** | Conditional aggregation and lookups |
| 📐 **Descriptive Stat Functions** | `AVERAGE`, `MEDIAN`, `STDEV`, `QUARTILE` |
| 📉 **Regression Tools** | `SLOPE`, `INTERCEPT`, `RSQ` / Data Analysis ToolPak |
| 🎛️ **What-If Tools** | Data Tables, Goal Seek, Scenario Manager |
| 📋 **PivotTables & PivotCharts** | Interactive dashboard summary |

---

## 📈 Results & Insights

After the workbook is fully built out, it produces:

- ✅ **Top Customer Rankings** — by total sales and total profit
- 📊 **Statistical Profile** — of sales, profit, and margin across all transactions
- 📉 **Profit Drivers** — quantified relationship between discount/quantity/price and profit
- 🎛️ **Scenario Outcomes** — projected profit under adjusted pricing/discount assumptions
- 📋 **Regional & Product Rollups** — via the Pivot Dashboard

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🖥️ **No External Tools** | Runs entirely inside Excel — no scripting required |
| 🔄 **Fully Recalculating** | Every metric updates automatically as source data changes |
| 📚 **Educational** | Covers lookups, statistics, regression, and pivoting in one file |
| 🧩 **Modular** | Each analysis lives on its own sheet, easy to extend |
| 📈 **Business-Relevant** | Mirrors a real sales-reporting workflow end to end |

---

## 📄 License

This project is licensed under the **MIT License** — free to use, modify, and distribute with attribution.

---

## 👤 Author

<div align="center">

**🎓 Role:** Data / Business Analytics Student \
**🛠️ Skills:** Excel · SUMIFS/INDEX-MATCH · Descriptive & Regression Statistics · PivotTables · What-If Analysis

</div>
