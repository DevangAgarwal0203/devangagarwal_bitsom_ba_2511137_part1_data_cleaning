# Part 1 — Business Data Cleaning, Validation & Excel Reporting

**Student Name:** Devang Agarwal  
**Student ID:** bitsom_ba_2511137  
**Assignment:** Masai School — Final Capstone Project, Module 6  
**Part:** Part 1 of 4  
**Topic:** Data Cleaning, Validation, and Excel Pivot Reporting  

---

## Problem Summary

The retail company's order management system exports sales data from multiple internal systems. This leads to a single dataset that contains inconsistent text formatting, mixed date formats, duplicate records, missing values, invalid discount entries, and calculation mismatches between the reported sales figures and what can be verified from unit prices and quantities.

The goal of this part is to:
- Clean and standardise the raw dataset
- Validate it against defined business rules
- Document all data quality issues found
- Produce a data quality report and pivot-based summary reports for the business review team

---

## Dataset Description

**File:** `data/raw_orders.xlsx`  
**Sheet:** `raw_orders`  
**Total rows:** 932  
**Columns (21):** order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city, category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status

The dataset covers retail orders from 2024 and 2025 across four regions (North, South, East, West), three product categories (Furniture, Technology, Office Supplies), and five customer segments.

---

## Tools Used

- **Python 3** — main cleaning and processing language
- **pandas** — data loading, manipulation, and transformation
- **openpyxl** — writing formatted Excel output files
- **dateutil** — robust multi-format date parsing
- **Microsoft Excel / LibreOffice Calc** — for viewing and taking screenshots of the output files

---

## Cleaning Steps Performed

### Step 1 — Preserve Raw Data
The original file `raw_orders.xlsx` was not modified. All cleaning was done in a Python script (`clean_data.py`) that reads the raw file and produces separate output files.

### Step 2 — Text Field Cleaning
The following fields were cleaned across all rows:
- `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `product_name`, `ship_mode`, `payment_status`, `order_status`

Actions applied:
- Stripped leading and trailing whitespace
- Collapsed double (or more) internal spaces to single space
- Converted all values to Title Case (e.g., `NORTH` → `North`, `consumer` → `Consumer`)

### Step 3 — Date Cleaning and Standardisation
Both `order_date` and `ship_date` contained at least five different date formats mixed in the same column. A multi-format parser tried seven known patterns before falling back to `dateutil`. Parsed dates were used to:
- Compute `shipping_delay_days` = ship_date − order_date (in days)
- Extract `order_month` and `order_year`
- Identify and flag records where ship_date < order_date (22 records flagged as Invalid)

### Step 4 — Duplicate Handling
Two types of duplicates were identified:
- **Exact duplicates (19 order IDs, 38 rows):** All fields identical. Extra copies removed; first occurrence kept.
- **Conflicting duplicates (12 order IDs, 25 rows):** Same order_id, different sales amounts or statuses. All copies retained; flagged as `conflicting_duplicate_flagged` for manual review.

Net rows after removing exact duplicates: **913**

### Step 5 — Business Rule Application
| Rule | Action |
|---|---|
| Missing `region` (25 records) | Filled as `Unknown`; flagged |
| Missing `ship_mode` (21 records) | Filled as `Unknown`; flagged |
| Missing discount (18 records) | Set to `0.0` if other sales fields valid; else left NaN |
| Negative discount (15 records) | `cleaned_discount` set to NaN; flagged Invalid |
| Discount > 50% (7 records) | `cleaned_discount` set to NaN; flagged Invalid |
| Discount as text % (8 records) | Converted to decimal; if > 0.50, set to NaN; flagged Invalid |
| Cancelled / Returned / Failed orders | Excluded from completed sales pivot summaries |
| Ship date before order date (22 records) | Flagged Invalid |

### Step 6 — Calculated Columns Added
| Column | Formula |
|---|---|
| `cleaned_discount` | Standardised numeric discount; NaN if invalid |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)` |
| `calculated_profit` | `calculated_sales − cost` |
| `profit_margin` | `calculated_profit / calculated_sales` |
| `shipping_delay_days` | `ship_date − order_date` in days |
| `order_month` | Month number extracted from order_date |
| `order_year` | Year extracted from order_date |
| `data_quality_flag` | `Clean`, `Warning`, or `Invalid` per record |
| `duplicate_flag` | `exact_duplicate_removed` or `conflicting_duplicate_flagged` or blank |
| `discount_flag` | `ok`, `missing_discount`, `negative_discount`, `high_discount`, `converted_from_pct` |
| `sales_mismatch_flag` | `True` if `|calculated_sales − sales| > ₹1` |

---

## Business Rules Applied

1. Missing region → fill `Unknown`
2. Missing ship_mode → fill `Unknown`
3. Missing discount → treat as 0 only if quantity, unit_price, and sales are all valid
4. Negative discount → flag as invalid; set cleaned_discount to NaN
5. Discount > 50% → flag as unusually high; set cleaned_discount to NaN
6. Discount stored as percentage text (e.g., `70%`) → convert to decimal, then apply rule 4 or 5
7. Cancelled orders → exclude from completed sales summaries
8. Failed/Returned orders → exclude from completed sales summaries; Returned orders summarised separately
9. Ship date before order date → flag as Invalid
10. Exact duplicate rows → remove extra copies, keep first occurrence
11. Conflicting duplicate order IDs → retain all copies, flag for manual review

---

## Summary of Data Quality Issues Found

| Issue | Count |
|---|---|
| Total raw rows | 932 |
| Exact duplicate rows removed | 19 |
| Conflicting duplicate order IDs | 12 (25 rows) |
| Missing region (filled as Unknown) | 25 |
| Missing ship_mode (filled as Unknown) | 21 |
| Missing discount | 18 |
| Negative discounts | 15 |
| Discounts above 50% | 7 |
| Percentage-text discounts (e.g., 70%) | 8 |
| Ship date before order date | 22 |
| Sales calculation mismatches | 54 |
| **Records flagged Clean** | **818** |
| **Records flagged Warning** | **44** |
| **Records flagged Invalid** | **51** |
| **Net rows in cleaned dataset** | **913** |

---

## Summary of Pivot Reports

The file `outputs/pivot_summary.xlsx` contains six pivot sheets, all based on **Completed orders only** unless noted:

| Sheet | Content |
|---|---|
| `1_Sales_By_Region` | Total sales, total profit, order count, avg profit margin — sorted by sales (desc) |
| `2_Category_SubCat` | Sales and profit by category and sub-category — sorted by sales within category (desc) |
| `3_ShipMode_Orders` | Order count, total sales, average shipping delay — sorted by order count (desc) |
| `4_Segment_ProfitMargin` | Avg profit margin, total sales, total profit by customer segment — sorted by margin (desc) |
| `5_NonCompleted_ByRegion` | Cancelled and Returned orders by region — separate from completed sales summary |
| `6_Monthly_Sales_Trend` | Monthly sales and profit chronologically (Jan 2024 → latest) |

---

## Key Business Insights

1. **Sales quality is high but not perfect.** 818 of 913 records (89.6%) are clean. 51 records have critical issues (mainly invalid discounts and ship-date anomalies) that could distort revenue and margin analysis if included unchecked.

2. **Duplicate records point to a system integration problem.** 31 order IDs appeared more than once, with 12 showing conflicting data across copies. This strongly suggests the data was exported from multiple source systems without deduplication. The IT or data engineering team should standardise the export pipeline.

3. **Discount data is unreliable in a subset of records.** 30 records have clearly invalid discounts (negative or above 50%). An additional 8 records stored discounts as percentage text (e.g., `70%`). This suggests the discount input field in the order system lacks validation controls.

4. **22 orders have ship dates earlier than their order dates.** This is logistically impossible and indicates either a date entry error or a format mismatch during system export. These records should be corrected at source before use in any delivery performance reporting.

5. **The West region shows a notably high volume of both completed and cancelled orders,** suggesting possible fulfilment or customer satisfaction issues in that region (visible in pivot sheets 1 and 5).

6. **Technology as a category generates strong profit margins** relative to Furniture, which tends to have higher sales volume but thinner margins. This insight is visible in pivot sheet 2 and should inform product mix decisions.

7. **Standard Class and Second Class ship modes account for the majority of order volume.** Same Day deliveries are fewer but may carry higher customer expectations. Average shipping delay data (pivot sheet 3) can inform SLA reviews.

---

## Assumptions and Limitations

### Assumptions
- Discount valid range: 0% to 50% (0.00 to 0.50). This is typical for retail and consistent with the majority of values in the dataset.
- Date format ambiguity resolved using US-style MM/DD/YYYY where the day/month values were both ≤ 12 and the format was unclear. This may occasionally produce incorrect parsed dates.
- Missing discounts were set to 0.0 only where quantity, unit_price, and sales were all valid and non-zero.
- Conflicting duplicates could not be resolved without querying the source system; both copies were retained and flagged.

### Limitations
- Date parsing ambiguity means some edge-case dates may have been interpreted incorrectly.
- 12 conflicting duplicate order IDs remain unresolved and require manual review against source systems.
- 54 sales mismatches have been identified but their exact root cause cannot be determined from the dataset alone.
- No external product master or pricing catalogue was available to validate unit prices or costs.

---

## Screenshots

Screenshots are located in the `screenshots/` folder:

| File | Content |
|---|---|
| `raw_data_preview.png` | First rows of the raw dataset before any cleaning |
| `cleaned_data_preview.png` | Cleaned dataset with all calculated columns visible |
| `pivot_summary_1.png` | Sales and Profit by Region pivot (Sheet 1 of pivot_summary.xlsx) |
| `pivot_summary_2.png` | Monthly Sales Trend pivot (Sheet 6 of pivot_summary.xlsx) |

---

## Repository Structure

**GitHub Repository Name:** `devangagarwal_bitsom_ba_2511137_part1_data_cleaning`

```
devangagarwal_bitsom_ba_2511137_part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx          ← Original dataset (unchanged)
│   └── cleaned_orders.xlsx      ← Cleaned dataset with calculated columns
├── outputs/
│   ├── data_quality_report.xlsx ← 8-sheet data quality report
│   ├── pivot_summary.xlsx       ← 6-sheet pivot summary report
│   └── cleaning_log.md          ← Step-by-step cleaning log
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
└── README.md                    ← This file
```
