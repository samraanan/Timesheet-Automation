import pandas as pd
import os
from datetime import datetime
import shutil

BASE_DIR = r"c:\Users\Shmuel\Desktop\Timesheet_Automation"
CONFIG_PATH = os.path.join(BASE_DIR, 'Config.xlsx')

# Missing pairs from verification log (decoded from UTF-8 bytes)
missing_pairs = [
    ("בית (כפר אלדד)", "אור"),
    ("אור", "בית (כפר אלדד)"),
    ("בית (כפר אלדד)", "דרור"),
    ("דרור", "אור"),
    ("בית (כפר אלדד)", "כללי"),
    ("כללי", "דרור"),
    ("דרור", "מעיינות"),
    ("מעיינות", "דרור"),
    ("בית (כפר אלדד)", "חנות"),
    ("חנות", "בית (כפר אלדד)"),
    ("כללי", "שראל"),
    ("שראל", "מעיינות"),
    ("מעיינות", "נווה יעקב"),
    ("נווה יעקב", "בית (כפר אלדד)"),
    ("אור", "מעיינות"),
    ("חנות", "דרור"),
    ("דרור", "חנות"),
    ("אור", "חנות"),
    ("בית (כפר אלדד)", "מעיינות"),
    ("מחכה לפליקס", "שראל"),
    ("שראל", "חנות"),
    ("בית (כפר אלדד)", "נווה יעקב"),
    ("נווה יעקב", "בית (כפר אלדד)"),
]

def add_missing_distances():
    print("Loading Config.xlsx...")
    
    # Backup
    backup_path = CONFIG_PATH.replace('.xlsx', f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copyfile(CONFIG_PATH, backup_path)
    print(f"Backup created: {os.path.basename(backup_path)}")
    
    # Load all sheets
    xls = pd.ExcelFile(CONFIG_PATH)
    sheet_dict = {}
    for sheet_name in xls.sheet_names:
        sheet_dict[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Get Distances sheet
    df_dist = sheet_dict['Distances']
    print(f"Current distances: {len(df_dist)} rows")
    
    # Identify column names
    col_p1 = df_dist.columns[0]  # מיקום א'
    col_p2 = df_dist.columns[1]  # מיקום ב'
    col_dist = df_dist.columns[2]  # מרחק (ק"מ)
    
    print(f"Columns: {col_p1}, {col_p2}, {col_dist}")
    
    # Build existing pairs map
    existing = {}
    for _, row in df_dist.iterrows():
        p1 = str(row[col_p1]).strip()
        p2 = str(row[col_p2]).strip()
        dist = row[col_dist]
        existing[(p1, p2)] = dist
    
    print(f"Existing pairs in map: {len(existing)}")
    
    # Find missing and add them
    new_rows = []
    for p1, p2 in missing_pairs:
        if (p1, p2) in existing:
            continue  # Already exists
        
        # Check if reverse exists
        reverse_dist = existing.get((p2, p1))
        
        if reverse_dist is not None:
            # Use reverse distance
            new_rows.append({col_p1: p1, col_p2: p2, col_dist: reverse_dist})
            print(f"Adding: {p1} -> {p2} = {reverse_dist} (from reverse)")
        else:
            # No data, add with 0
            new_rows.append({col_p1: p1, col_p2: p2, col_dist: 0})
            print(f"Adding: {p1} -> {p2} = 0 (no data)")
    
    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_dist = pd.concat([df_dist, df_new], ignore_index=True)
        sheet_dict['Distances'] = df_dist
        print(f"\nAdded {len(new_rows)} new distance pairs.")
        print(f"Total distances now: {len(df_dist)} rows")
        
        # Write back
        with pd.ExcelWriter(CONFIG_PATH, engine='xlsxwriter') as writer:
            for name, df in sheet_dict.items():
                df.to_excel(writer, sheet_name=name, index=False)
        
        print(f"\nSUCCESS: Updated {CONFIG_PATH}")
    else:
        print("\nNo new pairs to add (all already exist).")

if __name__ == "__main__":
    add_missing_distances()
