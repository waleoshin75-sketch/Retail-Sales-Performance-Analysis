# Retail Sales Data Analysis Using Excel
## Revenue By Customer, Revenue By Category & Revenue By Region Analysis
### Consumer Electronics Retail – Synthetic Electronics Dataset (2026)

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

## 1. Project Overview
In modern retail environments, data is frequently captured across scattered systems, resulting in broken entries, duplicates, and calculation mismatches. The primary goal of this business intelligence project was to take a deeply corrupted, multi-table retail transactional database and transform it into a centralized, high-performance analytical application natively inside Microsoft Excel.

By bypassing system-slowing grid lookups and implementing a structured **Star Schema Data Model (Power Pivot)**, this project builds a self-service dashboard that allows executive leadership to explore revenue drivers dynamically while ensuring 100% data integrity.

---

## 2. Dataset Overview
The relational database utilized for this analysis mimics an enterprise electronics retail environment, consisting of 4 distinct tables:
*   **`Sales_Fact`**: Transactional logs tracking item quantities, order IDs, store locations, and sale timelines.
*   **`Customers_Dim`**: Master customer database capturing unique buyer IDs, names, and contact registration data.
*   **`Production_Dim`**: Product inventory definitions mapping item category tiers to standard unit catalog prices.
*   **`Stores_Dim`**: Operational facility directory mapping regional location codes to physical storefronts.

---

## 3. Data Quality Assessment
Before modeling, an extensive audit was performed on the raw tables, exposing critical data vulnerabilities that would heavily distort corporate financial reporting if left unaddressed:
*   **Structural Entry Failures:** The transactional fact table contained significant data gaps, including empty product ID rows and mixed text configurations inside the primary date parameters.
*   **Text Formatting Inconsistencies:** Text fields across the customer and product files suffered from erratic casing structures, leading to split groupings (e.g., separating "Audio", "audio", and "ELECTRONICS" into distinct entities).
*   **Primary Key Violations:** The master customer registry contained duplicate primary identifier records, compromising relational database logic.
*   **Financial Integrity Risks:** Retail price columns were contaminated with embedded text currency flags, rendering them mathematically uncalculable by Excel's standard numerical engine.

---

## 4. Data Cleaning & Error Correction
To restore complete structural database integrity, the following programmatic Excel cell formulas were applied directly across the grids:

### A. Normalizing Text Casing and Trimming Spaces
*   **Problem:** Erratic case entries and hidden text padding spaces generated fragmented categorical splits.
*   **Solution:** Unified text casing and stripped out invisible trailing space breaks simultaneously using nested string configurations:
   

### B. Patching Blank Product Mappings
*   **Problem:** Missing product identifier keys threatened to drop transactional entries out of the model entirely.
*   **Solution:** Built a conditional logic formula to isolate blank cells and assign them to a structured placeholder flag:
    

### C. Eliminating Duplicate Records
*   **Problem:** Repeated primary identifier records broke the unique constraints required for database lookups.
*   **Solution:** Deployed advanced filtering criteria rules via Excel's data extraction tool to isolate unique values, successfully stripping out redundant identifier rows.

### D. Formatting Financial Decimals
*   **Problem:** Text currency flags forced values to render as text fields, freezing all mathematical summary options.
*   **Solution:** Executed regular expression cleaning routines to strip out raw text symbols, converting the output columns into pure, clean numeric decimal fields formatted to two decimal points.

---

## 5. Feature Engineering & Relational Modeling
Rather than relying on heavy cell calculations like `VLOOKUP` or `XLOOKUP` which slow down workbook speeds, the cleaned tables were loaded directly into **Excel Power Pivot** to build a high-performance **Star Schema Relational Model**.

### A. The Relational Plumbing
A 1-to-Many (`1:*`) active data relationship map was established, drawing connection lines from the central transaction table up to the master descriptive lookups:
*   `Customers_Dim[CustomerID]` Connected to `Sales_Fact[CustomerID]`
*   `Production_Dim[ProductID]` Connected to `Sales_Fact[ProductID]`
*   `Stores_Dim[StoreID]` Connected to `Sales_Fact[StoreID]`

*   <img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/780dc13f-ff5d-48db-b7da-d40379b158a5" />


### B. Cross-Table Advanced DAX Calculation
To calculate our true monetary yields on a dynamic, row-by-row transactional matrix level, I authored a cross-table **DAX Measure** inside the model:
```dax
Total Revenue := SUMX(Sales_Fact, Sales_Fact[Qty_Order] * RELATED(Production_Dim[Unit price]))
```
*   *Why this works:* `SUMX` forces an iterative, row-by-row transaction calculation, while `RELATED` pulls the unit catalog pricing from the dimension table instantly for each matching item ID.

---

## 6. Revenue By Customer Analysis
By mapping our custom `Total Revenue` measure against our master buyer records, the system generated an automated executive buyer leaderboard. This analysis effectively isolates customer valuation concentrations, highlighting our top 10 highest-value purchasers while filtering historical timelines to analyze true account value over time.

---

## 7. Revenue By Category Analysis
Product catalog margins were audited by contrasting unit sales volumes directly against gross profitability returns. This evaluation exposed a significant variance in resource performance:
*   **High-Velocity Yields:** Premium inventory groups (such as **Audio**) generate immense financial yields while maintaining a tiny, highly efficient inventory footprint.
*   **Low-Velocity Drag:** Cheap accessory components demand significant inventory throughput but return negligible dollar values to the company ledger.

---

## 8. Revenue By Region Analysis
To evaluate sales channel strengths, transactional volumes were aggregated across geographic store parameters. Integrating interactive **Timeline Slicers** allows users to filter the entire model instantly, showing how regional cash generation changes between operational months.

---

## 9. Key Findings & Insights
*   **The Regional Engine:** The South Region serves as our primary market engine, generating a commanding **\$7,379.07** in gross performance. Conversely, the East Region represents our weakest market block at **\$5,626.13**, revealing an immediate 24% performance lag.
*   **Product Asset Discrepancies:** Exceptional asset efficiency was uncovered within the Audio catalog, securing a top-tier **\$3,630.00** in total revenue on just **46 units sold**. Meanwhile, low-margin Accessories required high transaction volume (**14 units**) yet returned a minimal **\$368.86**.
*   **The System Vulnerability Leak:** A massive data logging deficit was exposed inside our transactional registry, uncovering **\$1,133.41 in untracked sales** tied back to a ghost placeholder profile (`CUST-9999`) that completely lacks customer records. 

---

## 10. Recommendations
1.  **Synchronize Regional Strategies:** Extract the exact local marketing and sales playbooks currently driving performance in our high-yielding **South Region** and implement them across underperforming **East Region** operations.
2.  **Optimize Catalog Sourcing:** Reallocate seasonal procurement capital away from low-velocity accessory items and heavily scale investment into high-margin **Audio** components.
3.  **Enforce POS Database Hard Locks:** Update software validation protocols at all point-of-sale cash registers to automatically reject transaction processing whenever customer mapping fields fail to validate against active customer directories, permanently stopping the `CUST-9999` ghost leak.

---

## 11. Conclusion
This project demonstrates the power of shifting Excel from a standard spreadsheet tool into a true enterprise data modeling engine. By taking a corrupted, fragmented database, cleaning it using precise functional cell overrides, and architecting an optimized Power Pivot Star Schema, this build delivers a robust business intelligence tool. The application successfully surfaces data integrity vulnerabilities while giving corporate executives the dynamic visibility required to optimize sales velocity and protect revenue margins.
