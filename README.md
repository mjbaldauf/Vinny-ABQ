# Server Checkout Sheet Generator

This script automates the creation of a two-week server checkout Excel workbook for Vinaigrette. It uses a template .xlsx file and generates a copy with 14 dated sheets.

---

## What It Does

- Prompts the user to enter a start date in MM.DD format.
- Calculates a two-week date range based on the start date.
- Copies a template Excel file and renames it using the date range (e.g., 06.08-06.21.xlsx).
- Removes the "NO EXPO" sheet.
- Renames the "EXPO" sheet to the starting date (e.g., 06.08).
- Duplicates the sheet 13 times with incremented dates.
- Moves the "Summary of Employee Tips" sheet to the end.
- Saves the new workbook in the specified destination directory.

---

## Requirements

- Python 3.x
- openpyxl
- Standard library modules:
  - shutil
  - datetime

Install openpyxl if needed:
bash
pip install openpyxl
