import pandas as pd
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except: pass

# Find the latest report
import glob
import os

reports = glob.glob("Report_*.xlsx")
reports.sort(key=os.path.getmtime)
latest = reports[-1]

print(f"Checking: {latest}\n")

# Read Executive Summary
df = pd.read_excel(latest, sheet_name='Executive Summary')

print("=== Executive Summary - First 5 Days ===")
print(df[['תאריך', 'כניסה', 'יציאה', 'סה"כ']].head())

print("\n=== Explanation ===")
print("כניסה = First entry of the day (from ALL projects)")
print("יציאה = Last exit of the day (from ALL projects)")
print("סה\"כ = Total hours = (Exit - Entry - Breaks)")
print("\nIf the fix is working, entry/exit should reflect ALL projects,")
print("not just active ones.")
