# TMR Report Generator — How It Works

**Program:** `tmr_report_generator.py`
**Purpose:** Generates a unified monthly TMR (Trade Marketing Representative) attendance and KPI report from raw daily data files.

---

## 1. What You Provide (Inputs)

The program takes three types of input files:

| Input | Quantity | What it contains |
|---|---|---|
| TMR Daily Activity files | One per day (bulk upload) | Every individual agent visit made by every TMR on that day |
| TMR Datewise Summary files | One per day (bulk upload) | One aggregated row per TMR per day (total visits, working hours) |
| DH Wise Target file | Single file | The daily visit target for each Distribution House |

---

## 2. How to Use the Window

When you run the program, a window opens with four steps:

**Step 1 — Daily Activity Files:** Click "Browse & Add Files". A file picker opens where you can select multiple files at once (hold Ctrl to select a whole month's worth). You can click Browse again to add files from a different folder. A list below the button shows every file you have added. The "Clear" button empties the list so you can start over.

**Step 2 — Datewise Summary Files:** Same behaviour as Step 1, but for the summary files.

**Step 3 — DH Wise Target File:** Click "Browse" to select one file. The path appears next to the button.

**Step 4 — Output File:** Click "Browse" to choose where the finished report should be saved and what to name it.

Once all four steps are filled in, click **Generate Report**. The button grays out and a progress log at the bottom of the window shows each step as it runs. When done, a popup asks if you want to open the file in Excel immediately.

---

## 3. What the Program Does Internally — Step by Step

### Step A: Load the Daily Activity Files

All activity files are read and combined into one large table. For each row (each agent visit), the program:

- Parses `checkin_time` and `checkout_time` as proper date-times.
- Converts the `visit_time_h_m_s` column (stored as text like `"0:5:24"`) into total seconds by splitting on the colon character — so `"0:5:24"` becomes `0×3600 + 5×60 + 24 = 324 seconds`.
- Extracts a plain date from `created_at` for grouping (e.g. `2026-06-01`).
- Strips leading zeros from wallet numbers so `01763489307` and `1763489307` are treated as the same TMR.

### Step B: Load the Datewise Summary Files

All summary files are read and combined into one table. The `working_time_h_m_s` column is converted to seconds using the same method above. The date column is normalized to remove any time component.

### Step C: Load the DH Wise Target File

The program reads the DH target file and builds a simple lookup table:

```
DH Code  →  Daily Visit Target
DHDKN00814  →  35
DHCTG00548  →  30
DHGAZ00651  →  40
...
```

It does this by matching `DH Code` and `Daily visit target` column names in a case-insensitive, space-tolerant way (so `"DH Code"`, `"dh_code"`, and `"DH  Code"` all work). If a DH code from the activity data has no match in this file, the program falls back to a default target of 30.

### Step D: Evaluate Daily Attendance for Each TMR

This is the most important step. For each TMR on each date that appears in the activity data, the program:

1. **Filters to business hours only.** Only visits where the check-in time is on or after 10:00 AM and the check-out time is on or before 7:00 PM are kept. Visits outside this window are ignored entirely.

2. **Filters to valid duration only.** From the remaining visits, only those where the duration is between 3 minutes (180 seconds) and 10 minutes (600 seconds) are counted. Visits shorter than 3 minutes or longer than 10 minutes are ignored.

3. **Looks up the TMR's daily target.** The program finds the TMR's `dh_code` from the activity data and uses that to look up the target in the DH Wise Target file.

4. **Decides P or A.** If the number of valid visits is at least 75% of the daily target, the TMR is marked **P (Present)**. Otherwise they are marked **A (Absent)**. For example, if the daily target is 35, then the threshold is `35 × 0.75 = 26.25`, so 27 or more valid visits = P. Fewer than 27 = A.

   If a TMR has no activity data at all for a day but does appear in the summary file, the program falls back to using the `agent_visit_count` from the summary against the same 75% threshold. If neither file has data for a TMR on a given day, the cell is left blank. If the date is a Friday, the cell is always left blank because Friday is the weekly holiday.

### Step E: Calculate Summary Metrics Per TMR

For each TMR, the program aggregates across all the summary rows:

- **Average Work Time:** Adds up all `working_time_h_m_s` values across days where the TMR has any recorded working time, then divides by the number of those days. Stored as `H:MM:SS` text.
- **Average Strike:** Adds up all `agent_visit_count` values across days where the TMR has any visits, then divides by the number of those days. Rounded to two decimal places.

Then it computes the three KPI metrics:

- **Market Hour Achievement:** `Average Work Time ÷ 8 hours`. If the result exceeds 1.0 it is capped at 1.0 (you cannot score above 100%).
- **Strike Rate Achievement%:** `Average Strike ÷ Daily Visit Target`. Capped at 1.0 the same way.
- **Attendance%:** `Number of P days ÷ MTD Working Days` (working days = all non-Friday days in the period).

**MTD Work Day** is the count of non-Friday calendar days from the first date to the last date found across all uploaded files. **Friday Count** is the count of Fridays in the same range. **Total Month Days** is the total calendar days from first to last date.

**Weekly agent coverage target** is calculated as `Daily Visit Target × 6` (six working days in a week).

**Payable Day** equals the Present count. Approved Leave is set to 0 (the program does not receive a leave file, so leave is not tracked).

### Step F: Write the Excel Report

The program creates a new `.xlsx` file with one sheet called **TMR Daily Attendance**. The columns appear in this order:

```
TMR Wallet | TMR Name | Region | DH Code | Distributor House Name |
Daily visit target | Weekly agent coverage target |
Report Start Date | Report till Date |
MTD Work Day | Friday Count | Govt Holiday Count | Total Month Days |
Market Hour Target | Average Work Time | Average Strike |
[one column per date, e.g. 6/1/2026, 6/2/2026 ...] |
Present | Approved Leave | Payable Day |
Attendance% | Market Hour Achievement | Strike Rate Achievement%
```

Formatting applied:
- Header row: upay blue (#0054A5) background, white bold text.
- P cells: green background (#00B050), white bold text.
- A cells: dark red background (#C00000), white bold text.
- Friday / blank cells: gray (#D6D6D6).
- Present and Payable Day columns: upay yellow (#FFD504) background, bold text.
- Even and odd data rows alternate between white and a light blue tint for readability.
- Panes are frozen at column C, row 2 — so you can scroll right through the date columns without losing the TMR Wallet and Name, and scroll down without losing the header.

---

## 4. How to Prepare the DH Wise Target File

This file must be an Excel file (`.xlsx` or `.xls`) with the following four columns. The column names must match exactly (but capitalization and extra spaces are ignored):

| Column | What to put here |
|---|---|
| `DH Code` | The Distribution House code exactly as it appears in the activity and summary files, e.g. `DHDKN00814` |
| `Distributor House Name` | The full name of the DH, e.g. `GODHULY TRADERS` |
| `Market type` | Either `Metro` or `Non metro` |
| `Daily visit target` | A whole number: the number of unique agent visits a TMR under this DH is expected to complete each working day |

**One row per DH.** Each row covers all TMRs that belong to that DH. If two TMRs work under `DHDKN00814`, they both inherit a daily target of whatever is in that DH's row.

**The DH Code must match exactly** what appears in the `dh_code` column of your activity and summary files. If there is a mismatch (e.g. a typo or different formatting), that TMR's target will fall back to 30.

**Example structure:**

```
DH Code       | Distributor House Name        | Market type | Daily visit target
DHDKN00814    | GODHULY TRADERS               | Metro       | 35
DHDKN00936    | SEBA PAY                      | Metro       | 35
DHCTG00548    | MMFS ENTERPRISE               | Non metro   | 30
DHGAZ00651    | S I TRADING                   | Metro       | 40
DHBOG00859    | HUJAIFA ENTERPRISE            | Metro       | 40
DHMYM00905    | HASAN ENTERPRISE              | Non metro   | 30
```

You do not need to list TMR names or wallets in this file — the program links TMRs to their DH automatically using the `dh_code` column that already exists in the activity and summary files.

**No other columns are required.** Additional columns (e.g. notes, phone numbers) are simply ignored.

---

## 5. How Source Data Flows Into Each Output Column

| Output Column | Source | How |
|---|---|---|
| TMR Wallet | Summary file | Pulled from `tmr_wallet`, leading zero restored for display |
| TMR Name | Summary file | Pulled from `tmr_name` |
| Region | Summary file | Pulled from `region` |
| DH Code | Summary file | Pulled from `dh_code` |
| Distributor House Name | Summary file | Pulled from `distributor_house_name` |
| Daily visit target | DH Target file | Looked up using the TMR's `dh_code` |
| Weekly agent coverage target | Calculated | `Daily visit target × 6` |
| Report Start Date | All files | Earliest date found across all uploaded files |
| Report till Date | All files | Latest date found across all uploaded files |
| MTD Work Day | Calculated | Non-Friday days between start and end date |
| Friday Count | Calculated | Friday count between start and end date |
| Govt Holiday Count | Fixed | Always 0 (not tracked in this version) |
| Total Month Days | Calculated | Calendar days between start and end date |
| Market Hour Target | Fixed | `8:00:00` for every TMR |
| Average Work Time | Summary file | Sum of `working_time_h_m_s` ÷ days with recorded time |
| Average Strike | Summary file | Sum of `agent_visit_count` ÷ days with any visits |
| Date columns (e.g. 6/1/2026) | Activity file (primary) / Summary file (fallback) | P, A, or blank per the attendance logic above |
| Present | Calculated | Count of P cells in the date columns |
| Approved Leave | Fixed | Always 0 |
| Payable Day | Calculated | Equal to Present count |
| Attendance% | Calculated | `Present ÷ MTD Work Day` |
| Market Hour Achievement | Calculated | `Average Work Time ÷ 8h`, capped at 1.0 |
| Strike Rate Achievement% | Calculated | `Average Strike ÷ Daily visit target`, capped at 1.0 |