# Primebuild Keypay Location Automation

Streamlit app that processes raw Keypay timesheet exports and generates a categorised summary report.

## What it does

1. Reads a raw timesheet export (`.xlsx`) with an `All Timesheets` sheet
2. Classifies each row into one of six categories:
   - **Approved – Non-C Locations** (D, R, N project codes)
   - **Approved – Unallocated** (home/operations location)
   - **Unapproved – Unallocated** (Submitted + home location)
   - **Unapproved – Allocated** (Submitted + any project location)
   - **Others – Self Approved** (Reviewed By matches employee name)
   - **Others – Approved AL but C Costed** (Annual Leave on a C location)
3. Excludes: Processed rows, Approved C-location rows, Approved home-location rows
4. Generates a formatted `.xlsx` output with Summary, All Timesheets, and Report Details sheets

## Output filename

Auto-generated as `YYYYMMDD_<custom name>.xlsx` where the date is today's date.

## Deployment

- Deploy via Streamlit Cloud pointing to this repo
- Place `logo.jpg` in the repo root for branding

## Local development

```bash
pip install -r requirements.txt
streamlit run app.py
```
