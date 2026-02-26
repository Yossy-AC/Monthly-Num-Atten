---
name: aggregate-enrollment
description: Aggregate monthly enrollment data and generate summary reports. Use this skill whenever the user asks to process student enrollment Excel files, generate monthly statistics, create pivot tables, compile enrollment summaries, analyze enrollment trends, or create year-round attendance reports. Automates reading multiple Excel files from a lists directory, extracting enrollment data by month, and generating a consolidated Pivot format Excel file with monthly totals by grade, classroom, course, and instructor.
compatibility: Python 3.14+, pandas, openpyxl
---

# Enrollment Data Aggregator

Process and aggregate student enrollment data from Excel files into structured monthly reports.

## What This Does

Automatically processes all Excel files in the `lists/` directory and generates a **monthly enrollment pivot table** in Excel format. The skill:

- Reads all `*_YYMM.xlsx` files (where YYMM is fiscal year + month code)
- Extracts enrollment data for high school students (grades 1-3)
- Aggregates by month, filtering out incomplete or invalid records
- Generates a Pivot table: **Grade | Classroom | Course | M/C | Instructor** × **Monthly Columns**
- Outputs: `outputs/monthly_stats.xlsx` + console statistics

## When to Use

- User asks to "aggregate" or "compile" enrollment data
- User requests monthly enrollment reports, statistics, or trends
- User mentions processing multiple Excel files or "lists"
- User wants to create a pivot table or summary from enrollment data
- Any form of enrollment analysis or data compilation task

## Input Files

**Location**: `lists/` directory

**Format**: `*_YYMM.xlsx`
- YY = Fiscal year (last 2 digits, e.g., 25 = FY2025)
- MM = Month (01-12)
- Example: `〔定例報告〕2025AC受講者ﾘｽﾄ_2504.xlsx` → April 2025

**Data Structure** (Row 4 is header, 0-indexed columns):
- Column C: Enrollment add date
- Column G: Cancellation date
- Column J: Course name
- Column K: M/C (Master/Core classification)
- Column L: Classroom
- Column P: Grade code (31=High1, 32=High2, 33=High3)
- Column AA: Instructor

## Output

**File**: `outputs/monthly_stats.xlsx`

**Format**: Pivot table with fixed columns:
- **Static Columns** (8): Grade | Classroom | Course | M/C | Instructor | (removed: School, Department, Gender)
- **Month Columns** (4月～3月): April through March (fiscal year order)
- **Cell Values**: Enrollment count per month

**Console Output**: Annual summary showing total enrollees per month

## Execution

```bash
python scripts/aggregate.py
```

The script automatically:
1. Scans `lists/` for all `.xlsx` files
2. Extracts target month from filename (`_YYMM` pattern)
3. Filters to active students (High 1-3 only)
4. Applies active status cutoff (enrollment add date ≤ previous month-end)
5. Removes records where Instructor is "0", "-", or blank
6. Aggregates by the 5-dimensional key (grade × classroom × course × M/C × instructor)
7. Merges all monthly results into one Pivot Excel
8. Outputs `monthly_stats.xlsx` + prints annual totals

## Data Filtering Rules

**Included Students**:
- Grade codes: 31 (High 1), 32 (High 2), 33 (High 3)
- Active status: `add_date ≤ cutoff_date AND (cancel_date IS NULL OR cancel_date > cutoff_date)`
  - Cutoff date = last day of the month before the target month
  - Example: For May (05), cutoff is April 30

**Excluded Records**:
- Instructor field is "0", "-", or empty

**Aggregation Axes**:
- 学年 (Grade)
- 教室 (Classroom)
- 講座名 (Course name)
- M/C (Master/Core)
- 担当 (Instructor)

## Example Output

**Spreadsheet** (`monthly_stats.xlsx`):
```
| Grade | Classroom | Course | M/C | Instructor | Apr | May | Jun | ... | Feb |
|-------|-----------|--------|-----|------------|-----|-----|-----|-----|-----|
| 高1   | 505       | Math   |     | Tanaka     | 25  | 30  | 28  | ... | 32  |
| 高2   | 509       | English|【M】| Suzuki     | 15  | 18  | 17  | ... | 19  |
| ...   |           |        |     |            |     |     |     |     |     |
```

**Console Output**:
```
Processing 11 files from lists/
------------------------------------------------------------
Annual Summary:
  4月: 576
  5月: 752
  6月: 776
  ...
  2月: 769
  Total: 8,590
```

## Performance

- Processes 11 files (~3.5MB total) in ~5 seconds
- Vectorized pandas operations (no row-by-row loops)
- Outputs: 9.3 KB Excel file

## Implementation Details

**Core Functions** (`services/aggregator.py`):
- `parse_target_month()` — Extract month from filename (universal fiscal year logic)
- `aggregate()` — Month-level enrollment count with filtering
- `build_pivot()` — Merge all months into one DataFrame
- `to_excel_bytes()` — Generate Excel output

**Fiscal Year Logic**:
- Files named `_YYMM` where YY is fiscal year lower 2 digits
- If month ≥ 4: calendar year = fiscal year
- If month < 4: calendar year = fiscal year + 1
- Examples: `_2504` → 2025-04, `_2501` → 2026-01

## Notes

- Instructor ("担当") filtering: Records with "0", "-", or empty instructor names are excluded (represents unassigned/placeholder entries)
- School, Department, and Gender columns are removed from the pivot (simplified axis)
- Data is stored as monthly CSVs in `outputs/results/` before final Excel merge
- Re-running the script clears previous CSV cache and regenerates from source files
