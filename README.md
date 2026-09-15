# Excel Report Automation (Python + pandas + openpyxl) — sample project

Automated a messy monthly sales export (`messy_sales.xlsx`) into a
client-ready report (`clean_report.xlsx`):

- Cleaned 132 raw rows down to 120 (duplicates removed)
- Normalized mixed date formats, text casing, whitespace and numbers
- Handled bad quantities ("two", blanks) without crashing
- Generated a formatted workbook: styled Sales sheet, Summary sheet with
  per-region revenue, and a totals row ($110,190.71)

Turns a 30-minute manual cleanup into a one-click script.

## Run it

```bash
pip install pandas openpyxl
python clean_report.py
```

This is a demonstration sample showing my automation workflow.
No client, no fake data.
