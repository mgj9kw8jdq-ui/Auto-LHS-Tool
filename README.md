# Auto LHS Tool
A Python automation tool that generates Loop History Sheets (LHS) from a Master History Sheet.  
You enter a loop number, and the tool extracts the required rows and creates a formatted Excel file.

## Features
- Extracts loop-specific data from 3000+ records
- Automatically maps predefined fields into a structured LHS format
- GUI-based loop input using `tkinter` (optional).
- Reduces manual effort and errors

## Tech Used
Python, pandas, openpyxl, tkinter

## How to Run
pip install pandas openpyxl tqdm  
python auto_lhs.py

## Notes
- Only dummy Excel files should be used (no real project data).
- Adjust column names in code if needed.
