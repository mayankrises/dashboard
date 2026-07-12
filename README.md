# DCI Dashboard — BPCL Uran

## Setup (one time)
```bash
pip install -r requirements.txt
```

## Run
```bash
streamlit run dashboard.py
```
This opens in your browser. Use the sidebar to **upload your DCI workbook (.xlsx)** — any file with the same sheet/column layout as `BPCL URAN TEMPLATE.xlsx` will work. Nothing is uploaded anywhere else; it all stays on your computer.

## What it reads
- Only the `DCI` sheet is used (it's the source of truth — the other sheets are just old PivotTables built from it).
- It reads the "latest revision" columns (REV., UPLOADED DATE, CODE, FROM EIL DATE, etc.) for current status, and the Rev-0…Rev-10 history blocks for trend/turnaround charts.

## Key assumptions baked in (change in `dashboard.py` if wrong)
- **Approval codes:** `1` = Approved, `R` = Retained for Record (kept as its **own** bucket, not merged into Approved), `2` = comments/resubmit, `3` = rejected/resubmit, `V`/`v` = Void.
- **Allowed EIL review period:** defaults to 21 days — adjustable live in the sidebar.
- **Aging** starts from the *Schedule submission date* for documents not yet submitted, and from the *latest uploaded date* for documents awaiting an EIL review return. Closed (Approved/Retained) and Void documents are excluded from aging.
- **Root-cause notes** are the 12 items carried over from the old Excel dashboard's "Root Cause Analysis for Overshoot of Engineering Manhours" — shown as view-only text.

## Tabs
Executive Summary · Submission Progress · Review Progress · Approval Status · Discipline Analysis · Vendor Analysis · Delays & Aging · Manhours · Document Explorer (searchable, filterable, CSV export, overdue rows highlighted in red).
