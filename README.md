# Mistake Analysis

Automated processing and reporting of employee scanning errors in greenhouse production.

## Overview

This project collects daily error reports from greenhouses, normalizes and
anonymizes the data, and generates PDF reports with key metrics and
visualizations. It includes both a processing pipeline and an interactive
desktop application (GUI) for generating reports.

**Key features:**
- Automated processing of daily Excel reports from multiple locations
- Error normalization — 200+ raw input variants mapped to 14 standard categories
- Anonymization of employee names using deterministic MD5 hashing
- Multi-month reports (select 1 or more months, or full range)
- Per-stage reports (single greenhouse / Plant Protection / Old Warehouse)
- Configurable PDF report sections with title page and custom comments
- Charts: weekly totals & daily average, pie chart per location,
  error type distribution, monthly trend, correlation analysis
- Optional console log file and detailed error reporting

## Architecture

Three main scripts:

| Script | Role |
|--------|------|
| `process_errors.py` | Reads raw Excel files → produces cleaned, anonymized datasets |
| `report_engine.py` | PDF report generation engine (tables, charts, sections) |
| `gui.py` | Tkinter GUI to select period, sections and generate PDF reports |

## Project Structure
```
mistake-analysis/
├── s24-25/                          # season 2024/2025
│   ├── process_errors.py            # data processing + anonymization
│   ├── report.py                    # PDF report generation
│   ├── report_greenhouse.py         # PDF report per greenhouse
│   ├── sample_data/                 # sample input templates
│   ├── reports/                     # anonymized PDF reports
│   ├── errors_combined_anon.csv     # anonymized error dataset
│   ├── errors_combined_anon.xlsx
│   ├── harvest_normalized.csv
│   └── unrecognized_errors.xlsx
│
├── s25-26/                          # season 2025/2026
│   ├── process_errors.py            # step 1: process Excel files
│   ├── report_engine.py             # step 2: PDF generation engine
│   ├── gui.py                       # step 2: GUI (Tkinter)
│   ├── fonts/                       # optional custom fonts (Carlito)
│   ├── data/                        # input Excel files (gitignored)
│   ├── output/                      # processed datasets (gitignored)
│   ├── reports/2026/                # generated PDF reports
│   └── logs/                        # run logs (gitignored)
│
├── requirements.txt
├── .gitignore
└── README.md
```
## Input File Format

Before running the pipeline, place your daily error Excel files in the `data/` folder.  
Each file must follow this naming convention:
Błędy_YYYY-MM-DD.xlsx

## How to Run

**Requirements:**
```bash
pip install -r requirements.txt

Or manually:

pip install pandas matplotlib reportlab openpyxl numpy
```

**Step 1 — Prepare input data**
```
Place your daily error Excel files in s25-26/data/. Each file should:
have a date in the filename (DD.MM.YYYY or YYYY.MM.DD)
contain one sheet per location, with sheet names starting with
Etap, Stary Magazyn, or Ochrona roślin

have a column Imię i nazwisko (employee name) and Błąd (error description)
```
**Step 2 — Process raw data:**

```bash
cd s25-26
python process_errors.py
```
Reads all `.xlsx` files from `data/`, normalizes errors, anonymizes names,
and writes results to `output/`:

- `errors_combined.xlsx` — cleaned & anonymized error records
- `errors_combined.csv` — same, CSV
- `unrecognized.xlsx` — entries that could not be mapped
- `headcount.xlsx` — headcount and error % per location

### Step 3 — Generate PDF report (GUI)

```bash
cd s25-26
python gui.py
```

In the GUI:

- choose **report type** (Monthly / Per stage)
- for monthly: select **one or more months** (Ctrl+click for multi-select)
- for per stage: select a **stage** and optionally a **range of months**
- optionally expand **More options** to:
  - add a **title page** (author, position, email, title, subtitle)
  - toggle individual **report sections** (statistics, error types,
    breakdown, attendance, correlation, trend)
  - add a custom **comment**
- click **Generate PDF**

The PDF is saved to `s25-26/reports/<year>/`.

## Technologies

Python · pandas · matplotlib · ReportLab · openpyxl · tkinter

## Author
Built as a portfolio project combining personal data, SQL, Python, and web development.

Bartosz Grzybowski 

https://www.linkedin.com/in/bartosz-grzybowski-5444732bb/
