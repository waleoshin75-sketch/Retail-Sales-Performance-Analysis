# Retail Sales Data Analysis Using Excel

## Revenue by Customer, Revenue by Category, and Revenue by Region Analysis

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Dataset Overview](#2-dataset-overview)
3. [Data Quality Assessment](#3-data-quality-assessment)
4. [Data Cleaning & Error Correction](#4-data-cleaning--error-correction)
5. [Feature Engineering](#5-feature-engineering)
6. [Revenue By Customer Analysis](#6-revenue-by-customer-analysis)
7. [Revenue By Category Analysis](#7-revenue-by-category-analysis)
8. [Revenue By Region Analysis](#8-revenue-by-region-analysis)
9. [Key Findings & Insights](#9-key-findings--insights)
10. [Recommendations](#10-recommendations)
11. [Conclusion](#11-conclusion)

---

## 1 Project Overview

Retail data rarely arrives ready to analyze. It usually shows up scattered across separate tables, full of broken entries, duplicate records, and numbers stored as text. This project takes exactly that kind of raw, multi-table retail database and turns it into a working analytical model, built entirely inside Microsoft Excel using Power Pivot.

Rather than relying on slow, sprawling lookup formulas, the cleaned tables were loaded into a proper star schema data model. The result is a dashboard that lets anyone explore revenue by customer, category, or region on demand, backed by data that has actually been checked and corrected first.

---

## 2 Dataset Overview

The model is built on four related tables, structured the way a retail data warehouse typically is.

| Table | What It Holds |
| --- | --- |
| `Sales_Fact` | Transaction line items, order IDs, store locations, and sale dates |
| `Customers_Dim` | Customer IDs, names, and contact details |
| `Product_Dim` | Product IDs, category, and unit price |
| `Stores_Dim` | Store IDs mapped to region and location |

`Sales_Fact` sits at the center as the transactional record; the other three describe who bought, what was bought, and where.

---

## 3 Data Quality Assessment

Before any modeling began, the raw tables were audited and several issues stood out clearly enough to distort any report built on top of them.

- **Structural gaps in the transaction log:** `Sales_Fact` contained blank product ID rows and a mix of inconsistent date formats.
- **Inconsistent text casing:** Category and name fields mixed cases freely, so "Audio", "audio", and "ELECTRONICS" were being treated as separate values instead of one.
- **Duplicate customer records:** The customer table contained repeated primary identifiers, breaking the one-to-many relationship a clean model depends on.
- **Prices stored as text:** Unit price values had currency symbols embedded in them, which made the column unusable for any calculation until it was cleaned.

---

## 4 Data Cleaning & Error Correction

Each issue above was resolved with a targeted formula or Excel tool, applied directly across the affected columns.

### A. Normalizing Case and Trimming Spaces

Erratic casing and hidden trailing spaces were both fixed in a single formula, restoring one consistent category value where there used to be several.

```excel
=PROPER(TRIM(A2))
```

### B. Patching Blank Product Mappings

Blank product ID cells were replaced with an explicit placeholder rather than left empty, so no transaction silently dropped out of the model during a lookup.

```excel
=IF(ISBLANK(B2), "UNKNOWN_PROD", B2)
```

### C. Removing Duplicate Records

Excel's built-in duplicate removal tool was run against the customer table's primary identifier column, stripping out repeated rows and restoring the unique keys the relational model depends on.

### D. Converting Prices to Real Numbers

Currency symbols embedded in the price column were stripped out using text-cleaning formulas, converting the column from text into proper two-decimal numeric values that Excel could actually sum and average.

---

## 5 Feature Engineering

With clean tables in hand, the next step was connecting them properly and building the one measure the entire analysis depends on, total revenue.

### A. Building the Relational Model

The four tables were connected in Power Pivot using one-to-many relationships, with the dimension tables feeding into the central `Sales_Fact` table.

- `Customers_Dim[CustomerID]` to `Sales_Fact[CustomerID]`
- `Product_Dim[ProductID]` to `Sales_Fact[ProductID]`
- `Stores_Dim[StoreID]` to `Sales_Fact[StoreID]`

![Power Pivot data model diagram showing Sales_Fact connected to Customers_Dim, Product_Dim, and Stores_Dim](https://github.com/user-attachments/assets/780dc13f-ff5d-48db-b7da-d40379b158a5)

### B. The Total Revenue Measure

With the relationships in place, a single DAX measure calculates revenue row by row across every transaction, pulling the matching unit price in from the product table for each item sold.

```dax
Total Revenue = SUMX(Sales_Fact, Sales_Fact[Qty_Order] * RELATED(Product_Dim[Unit_Price]))
```

`SUMX` iterates through every row in `Sales_Fact` individually, and `RELATED` reaches across the relationship to pull in the correct unit price for that row's product, so the measure stays accurate even as new transactions are added.

---

## 6 Revenue By Customer Analysis

Plotting the Total Revenue measure against the customer table produces an automatic buyer leaderboard, ranking customers from highest to lowest lifetime spend. Filtering by date range on top of that shows how an individual customer's value has actually changed over time, rather than just their all-time total.

---

## 7 Revenue By Category Analysis

Comparing units sold against revenue generated by category exposed a clear split in how efficiently different product lines perform.

- **High-efficiency categories:** Audio products generated strong revenue from a relatively small number of units sold, a sign of high per-unit value.
- **Low-efficiency categories:** Accessories required a much higher unit volume to generate a comparatively small amount of revenue.

---

## 8 Revenue By Region Analysis

Transactions were aggregated by store region to compare geographic performance directly. A timeline slicer was added on top, letting anyone filter the entire report by month and see how each region's revenue shifts over time, rather than looking at one static total.

---

## 9 Key Findings & Insights

- **The South region is the strongest market.** It generated 7,379.07 dollars in gross revenue, compared to 5,626.13 dollars in the East, the weakest region, a gap of roughly 24 percent between the two.
- **Audio is the standout category by efficiency.** It brought in 3,630 dollars in total revenue from just 46 units sold. Accessories, by contrast, needed 14 units to generate only 368.86 dollars, a far weaker return per unit.
- **A ghost customer profile is hiding untracked revenue.** 1,133.41 dollars in sales are tied to a placeholder customer ID, `CUST-9999`, that has no matching record in the customer table at all, meaning that revenue currently has no real buyer attached to it in the model.

---

## 10 Recommendations

1. **Study and replicate what's working in the South.** Whatever is driving the South region's stronger performance is worth documenting and testing in the East before assuming the gap is unfixable.
2. **Shift procurement toward higher-efficiency categories.** Audio's strong revenue-to-volume ratio makes it a better candidate for expanded inventory investment than low-margin accessories.
3. **Close the CUST-9999 gap at the point of sale.** Add validation at checkout so a transaction can't be recorded without a real, matching customer ID, preventing this kind of untracked revenue from recurring.

---

## 11 Conclusion

This project shows what a spreadsheet can do once it's treated like a real data model rather than a flat grid. Cleaning the raw tables, correcting the formatting issues that were fragmenting categories and breaking calculations, and connecting everything through a proper star schema turned a messy export into a dashboard that actually answers questions, and surfaced a genuine revenue leak in the process.
