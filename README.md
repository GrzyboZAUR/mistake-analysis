# Mistake Analysis

Automated processing and reporting of employee scanning errors in greenhouse production.

## Overview

This project collects daily error reports from 6 greenhouses, normalizes and anonymizes the data, and generates a PDF report with key metrics and visualizations.

**Key features:**
- Automated data processing from multiple Excel files
- Error normalization — 200+ raw input variants mapped to 10 standard categories
- Anonymization of employee names using deterministic MD5 hashing
- Harvest data normalization to protect sensitive production figures
- PDF report generation with weekly and monthly breakdown

## Project Structure
```
mistake-analysis/
├── data/                      # Raw Excel files (gitignored)
├── sample_data/               # Sample data and template Excel 
├── reports/                   # Generated anonymized PDF reports
├── errors_combined_anon.csv   # Anonymized error dataset
├── harvest_normalized.csv     # Normalized harvest index data
├── process_errors.py          # Data processing + anonymization
├── report.py                  # PDF report generation
├── report_greenhouse.py       # PDF report generation per-greenhouse
└── README.md
```
## Input File Format

Before running the pipeline, place your daily error Excel files in the `data/` folder.  
Each file must follow this naming convention:
Błędy_YYYY-MM-DD.xlsx

## How to Run

**Requirements:**
```bash
pip install pandas matplotlib reportlab openpyxl numpy
```

**Step 1 — Place your daily Excel files in data/ using the format above**

**Step 2 — Process raw data:**

```bash
python process_errors.py
```
Reads all `.xlsx` files from `data/`, normalizes errors, anonymizes names and outputs `errors_combined_anon.csv`.

**Step 3a — Generate monthly report:**

Edit the variables at the top of `report.py`:
```python
MONTH = 4
WEEK  = "14-18"
YEAR  = 2026
```
Then run:
```bash
python report.py
```
Report is saved to `reports/report_errors_YYYY_MM.pdf`.

**Step 3b — Generate per-greenhouse report:**

Edit the variable at the top of `report_greenhouse.py`:
```python
TARGET_GREENHOUSE = None  # set to e.g. "Etap 1" for a single greenhouse
```
Then run:
```bash
python report_greenhouse.py
```
Generates one PDF per greenhouse in `reports/`.

## Technologies

Python · pandas · matplotlib · ReportLab · openpyxl

## Author
Built as a portfolio project combining personal data, SQL, Python, and web development.

Bartosz Grzybowski 

https://www.linkedin.com/in/bartosz-grzybowski-5444732bb/
