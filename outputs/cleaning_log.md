# Cleaning Log — Part 1: Business Data Cleaning, Validation & Excel Reporting

**Student:** Devang Agarwal (bitsom_ba_2511137)  
**Dataset:** raw_orders.xlsx  
**Cleaned Output:** cleaned_orders.xlsx  
**Total Raw Rows:** 932  
**Total Rows After Cleaning:** 913  

---

## 1. Issues Found

### 1.1 Text Formatting Issues

| Field | Issues Observed | Examples |
|---|---|---|
| `customer_name` | All-caps, trailing spaces, double internal spaces | `PRIYA MENON`, `Vikram  Iyer`, `Ananya Rao ` |
| `segment` | Leading/trailing spaces, lowercase, mixed case | `  Small Business `, `consumer`, `Corporate ` |
| `region` | All-caps, lowercase, leading/trailing spaces | `NORTH`, `west`, `  North `, `WEST` |
| `category` | All-caps, extra leading spaces, double spaces | `OFFICE SUPPLIES`, `  Furniture `, `office  supplies` |
| `sub_category` | All-caps, lowercase | `LABELS`, `storage` |
| `ship_mode` | All-caps, double internal spaces, trailing space | `STANDARD CLASS`, `Standard  Class`, `Second Class ` |
| `payment_status` | Trailing spaces | `Paid ` |
| `order_status` | Lowercase, leading/trailing spaces, trailing space | `completed`, `  Completed `, `Completed ` |

### 1.2 Missing Values

| Field | Count | Business Rule Applied |
|---|---|---|
| `region` | 25 | Filled as `Unknown`; flagged in quality report |
| `ship_mode` | 21 | Filled as `Unknown`; flagged in quality report |
| `discount` | 18 | Set to `0.0` if quantity, unit_price, and sales were all valid; else left as NaN |

### 1.3 Invalid and Unusual Discounts

| Issue | Count | Action |
|---|---|---|
| Negative discount (e.g., -0.19, -0.23, -0.14) | 15 | Flagged as `negative_discount`; `cleaned_discount` set to NaN; record marked `Invalid` |
| High discount > 50% (e.g., 0.55, 0.65) | 7 | Flagged as `high_discount`; `cleaned_discount` set to NaN; record marked `Invalid` |
| Discount stored as percentage text (e.g., `70%`, `85%`) | 8 | Converted to decimal (0.70, 0.85); flagged as `converted_from_pct`; since > 0.50, `cleaned_discount` set to NaN and marked `Invalid` |
| Missing discount | 18 | Set to 0.0 where valid, else NaN |

### 1.4 Date Format Inconsistencies

The `order_date` and `ship_date` columns contained at least five different date formats mixed together in the same column:

- `21 Jul 2024` (DD Mon YYYY)
- `07/27/2024` (MM/DD/YYYY)
- `2024-07-22` (ISO 8601)
- `28-11-2024` (DD-MM-YYYY)
- `05 Sep 2024` (DD Mon YYYY)

All dates were parsed using a multi-format parser (tried 7 known formats before falling back to `dateutil`). Unparseable values were left as `NaT` and flagged.

### 1.5 Ship Date Before Order Date

**Count:** 22 records  
Ship dates earlier than order dates were identified after parsing both columns to standardised datetime. These records are flagged with `ship_before_order` and assigned `data_quality_flag = Invalid`.

Sample affected records: ORD-2024-10025, ORD-2024-10073, ORD-2025-10128, ORD-2024-10143, ORD-2024-10155, ORD-2025-10161

### 1.6 Duplicate Records

**Exact Duplicate Rows (all fields identical):**  
19 order IDs had exact duplicate copies in the dataset (38 total rows involved). One copy was kept; the rest were removed. These appear to be import errors from multiple source systems.

Affected IDs: ORD-2024-10134, ORD-2024-10251, ORD-2024-10277, ORD-2024-10281, ORD-2024-10503, ORD-2024-10599, ORD-2024-10705, ORD-2024-10713, ORD-2024-10850, ORD-2025-10007, ORD-2025-10128, ORD-2025-10378, ORD-2025-10400, ORD-2025-10550, ORD-2025-10633, ORD-2025-10677, ORD-2025-10695, ORD-2025-10720, ORD-2025-10730

**Conflicting Duplicate Order IDs (same ID, different data):**  
12 order IDs had multiple entries with differing sales amounts, order statuses, or other fields. All copies were retained but flagged as `conflicting_duplicate_flagged`. These need manual review to determine the correct record.

Affected IDs: ORD-2024-10124, ORD-2024-10143, ORD-2024-10273, ORD-2024-10332, ORD-2024-10424, ORD-2025-10091, ORD-2025-10171, ORD-2025-10225, ORD-2025-10315, ORD-2025-10336, ORD-2025-10374, ORD-2025-10572

### 1.7 Sales / Profit Calculation Mismatches

**Count:** 54 records where `calculated_sales` (quantity × unit_price × (1 − cleaned_discount)) differed from the reported `sales` value by more than ₹1.00.

Primary causes identified:
- Discount was originally negative or unusually high, distorting the reported figure
- Discount was stored as percentage text and misapplied
- Possible manual override or data entry errors in the source system

---

## 2. Cleaning Actions Performed

| Step | Action |
|---|---|
| Text cleaning | Applied TRIM (strip whitespace), removed double internal spaces, converted all text fields to Title Case |
| Date standardisation | Parsed 7 date formats; stored standardised dates in `order_date_parsed` and `ship_date_parsed` internally; original raw date columns preserved as-is |
| Shipping delay | Computed `shipping_delay_days = ship_date_parsed − order_date_parsed` in days |
| Extracted date parts | Added `order_month` and `order_year` columns from parsed order date |
| Duplicate removal | Exact duplicates: removed all but the first occurrence. Conflicting duplicates: retained all, flagged. |
| Missing region | Filled 25 blanks with `Unknown` |
| Missing ship_mode | Filled 21 blanks with `Unknown` |
| Missing discount | Set to `0.0` for 18 records where quantity, unit_price, and sales were present and valid |
| Invalid discount | Set `cleaned_discount` to NaN for negative, >50%, or percentage-text discounts |
| Calculated columns | Added `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin` |
| Data quality flag | Assigned `Clean`, `Warning`, or `Invalid` per record based on all issues found |

---

## 3. Business Rules Applied

| Rule | Applied Action |
|---|---|
| Missing region | Filled as `Unknown`; flagged in data_quality_report |
| Missing ship_mode | Filled as `Unknown`; flagged in data_quality_report |
| Missing discount | Set to 0.0 only if all other sales fields (quantity, unit_price, sales) were valid |
| Negative discount | `cleaned_discount` set to NaN; record flagged `Invalid` |
| Discount > 50% | `cleaned_discount` set to NaN; record flagged `Invalid` |
| Cancelled orders | Excluded from completed sales pivot summaries |
| Failed payment orders | Excluded from completed sales pivot summaries |
| Returned / Refunded orders | Summarised separately in pivot Sheet 5 |
| Ship date before order date | Flagged as `Invalid` in `data_quality_flag` |
| Exact duplicate rows | Removed extra copies; kept first occurrence; documented in quality report |
| Conflicting duplicate order IDs | All copies retained; flagged `conflicting_duplicate_flagged`; listed in quality report |

---

## 4. Assumptions Made

1. **Date format priority:** When a date string like `01/07/2024` was ambiguous (could be Jan 7 or Jul 1), the parser was set to `dayfirst=False` (US-style MM/DD/YYYY). For unambiguous formats such as `07/27/2024`, the month was clear from the value itself.
2. **Discount valid range:** Discounts between 0% and 50% (0.00 to 0.50) were considered valid for this retail business. Any discount above 50% was flagged as unusually high.
3. **Missing discount = zero:** Where a discount was missing but all other sales fields (quantity, unit_price, sales) were valid and non-zero, the discount was treated as 0. This is consistent with a no-discount sale.
4. **Conflicting duplicates retained:** Since conflicting duplicate records have different sales amounts or order statuses, it is not possible to determine the correct record without consulting the source system. All copies were kept and flagged rather than silently deleting records.
5. **Calculated sales formula:** `calculated_sales = quantity × unit_price × (1 − cleaned_discount)`. Where `cleaned_discount` was NaN (due to invalid discount), the mismatch between calculated and reported sales was noted and flagged.
6. **Completed-only pivots:** All main pivot summaries (sales by region, category, segment) use only `order_status = Completed` records. Cancelled, Returned, and Failed-payment records were excluded from these summaries as per business rules.
7. **Text standardisation to Title Case:** All text fields were converted to Title Case (e.g., `NORTH` → `North`, `consumer` → `Consumer`). This is the most common standard for customer-facing fields and matches the majority of existing values in the dataset.

---

## 5. Records Removed

| Reason | Records Removed |
|---|---|
| Exact duplicate rows (keep first occurrence) | 19 |
| **Total removed** | **19** |

No records were silently deleted for any other reason. All flagged records (Invalid, Warning) remain in `cleaned_orders.xlsx` with clear flags.

---

## 6. Records Flagged

| Flag | Count | Meaning |
|---|---|---|
| `Clean` | 818 | No issues detected |
| `Warning` | 44 | Minor issues (e.g., conflicting duplicate, sales mismatch) — needs review |
| `Invalid` | 51 | Critical issues (invalid discount, ship-before-order, date parsing failure) |

---

## 7. Limitations

1. **Date ambiguity:** Mixed date formats in the same column mean some dates may have been parsed incorrectly if the format was ambiguous (e.g., `01/07/2024` interpreted as January 7 rather than July 1). Without a known single source format, some edge cases cannot be resolved automatically.
2. **Conflicting duplicates unresolved:** 12 order IDs have conflicting data across multiple rows. The correct record cannot be determined without querying the source operational system.
3. **Sales mismatch cause unknown:** 54 records have a gap between the reported `sales` figure and the recalculated `calculated_sales`. The root cause (whether it is a discount error, unit price error, or manual override in the source) is not determinable from the data alone.
4. **No external price master:** Unit prices and costs could not be validated against a product master or pricing catalogue, so price-level errors may not have been caught.
5. **Profit validation limited:** `calculated_profit = calculated_sales − cost` was used. If the `cost` column itself contains errors, the derived profit and margin figures will also be incorrect.
6. **Screenshots require manual capture:** Screenshots of raw and cleaned data previews and pivot summaries must be taken manually from Excel and placed in the `screenshots/` folder.
