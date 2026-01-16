import pandas as pd
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except: pass

file = r"C:\Users\Shmuel\.gemini\antigravity\scratch\Timesheet_Automation\Report_2025-12_20260108_230115.xlsx"

print("=== Checking KM Column Issues ===\n")

# Read Executive Summary
df = pd.read_excel(file, sheet_name='Executive Summary')

print("Full Executive Summary:")
print(df[['תאריך', 'ק"מ']])

print("\n=== Rows with Missing/Zero KM ===")
missing_km = df[df['ק"מ'].isna() | (df['ק"מ'] == 0)]
print(missing_km[['תאריך', 'ק"מ']])

print(f"\nTotal rows: {len(df)}")
print(f"Rows with missing/zero KM: {len(missing_km)}")

# Also check Detailed Report
print("\n=== Detailed Report Sample ===")
df_detail = pd.read_excel(file, sheet_name='Detailed Report')
print(df_detail[['תאריך', 'ק"מ']].head(10))
