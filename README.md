# Retail Sales Analytics in Excel

**Power Pivot star schema | DAX | Revenue analysis by customer, category and region**

Cleaned a corrupted four-table retail database and rebuilt it as an interactive Excel model that ranks customers, categories and regions by revenue, and exposes a revenue leak caused by a missing customer record.

| | |
|---|---|
| **Tools** | Excel, Power Pivot, DAX, PivotTables, Timeline Slicers |
| **Data** | 4 related tables (1 fact, 3 dimensions) |
| **Deliverable** | Interactive dashboard with a single reusable Total Revenue measure |

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Dataset Overview](#2-dataset-overview)
3. [Data Quality Assessment](#3-data-quality-assessment)
4. [Data Cleaning and Error Correction](#4-data-cleaning-and-error-correction)
5. [Feature Engineering and DAX](#5-feature-engineering-and-dax)
6. [Data Modeling](#6-data-modeling)
7. [Analysis](#7-analysis)
8. [Key Findings](#8-key-findings)
9. [Recommendations](#9-recommendations)
10. [Conclusion](#10-conclusion)

---

## 1. Project Overview
Retail data usually arrives split across tables, with duplicates, inconsistent text and numbers stored as text. This project cleans that data and models it in Power Pivot so revenue can be explored by customer, category and region without lookup formulas.

## 2. Dataset Overview

| Table | Type | Contents |
|---|---|---|
| `Sales_Fact` | Fact | Order lines, quantities, store, sale date |
| `Customers_Dim` | Dimension | Customer ID, name, contact, registration date |
| `Product_Dim` | Dimension | Product ID, name, category, unit price |
| `Stores_Dim` | Dimension | Store ID, region, store manager |

## 3. Data Quality Assessment

| Issue | Impact |
|---|---|
| Blank product IDs and mixed date formats in `Sales_Fact` | Transactions at risk of dropping out of the model |
| Inconsistent casing ("Audio", "audio", "ELECTRONICS") | One category split into several groups |
| Duplicate customer IDs | Breaks the one-to-many relationship |
| Currency symbols stored inside price values | Prices cannot be summed or averaged |

## 4. Data Cleaning and Error Correction

**Casing and spacing:** one formula standardizes case and removes hidden spaces.
```excel
=PROPER(TRIM(A2))
```

**Blank product IDs:** replaced with an explicit placeholder so no transaction is lost in the join.
```excel
=IF(ISBLANK(B2), "UNKNOWN_PROD", B2)
```

**Duplicate customers:** removed with Excel's Remove Duplicates tool on the customer ID column.

**Prices stored as text:** currency symbols stripped and the column converted to numeric values with two decimal places.

## 5. Feature Engineering and DAX

One cross-table measure calculates revenue row by row, pulling each product's unit price across the relationship.

```powerquery
Total Revenue = SUMX(Sales_Fact, Sales_Fact[Qty_Order] * RELATED(Product_Dim[Unit_Price]))
```

- `SUMX` iterates every row of `Sales_Fact`.
- `RELATED` fetches the matching unit price from `Product_Dim`.
- Every visual reuses this one measure, so revenue is defined once.

## 6. Data Modeling

Star schema with `Sales_Fact` at the center and three one-to-many relationships:

| Dimension key | Fact key |
|---|---|
| `Customers_Dim[CustomerID]` | `Sales_Fact[CustomerID]` |
| `Product_Dim[ProductID]` | `Sales_Fact[ProductID]` |
| `Stores_Dim[StoreID]` | `Sales_Fact[StoreID]` |

<img width="959" height="503" alt="Power Pivot diagram view showing Sales_Fact connected to Customers_Dim, Product_Dim and Stores_Dim" src="https://github.com/user-attachments/assets/780dc13f-ff5d-48db-b7da-d40379b158a5" />

## 7. Analysis

| View | What it answers |
|---|---|
| Revenue by customer | Top 10 buyers and account value over time |
| Revenue by category | Units sold versus revenue, to compare efficiency |
| Revenue by region | Regional performance, filterable by month with a Timeline Slicer |

### Dashboard Preview

<img src="images/retail-sales-dashboard.png" alt="Retail Sales Performance Analysis Excel dashboard with revenue by customer, category and region" width="100%">

Revenue totals $27,189.91 on 359 units sold. Jane Smith is the top customer at $4,741, and Furniture is the top category at $8,000.

## 8. Key Findings

| Finding | Detail |
|---|---|
| South is the strongest region | 7,379.07 dollars, against 5,626.13 dollars in the East, a gap of about 24 percent |
| Audio is the most efficient category | 3,630.00 dollars from 46 units, while Accessories returned 368.86 dollars on 14 units |
| Untracked revenue | 1,133.41 dollars in sales tied to placeholder ID `CUST-9999`, which has no customer record |

## 9. Recommendations
1. **Close the South and East gap:** document what drives South performance and test it in the East.
2. **Shift procurement toward Audio:** it returns more revenue per unit than accessories.
3. **Validate customer IDs at checkout:** reject any sale without a matching customer record to stop the `CUST-9999` leak.

## 10. Conclusion
Cleaning the source tables and modeling them in Power Pivot produced a reusable revenue model, and surfaced a data capture gap worth fixing at the point of sale.

