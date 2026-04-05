
import pandas as pd
import os
import shutil
from datetime import datetime

# Paths
BASE_DIR = r"c:\Users\Shmuel\Desktop\Timesheet_Automation"
CONFIG_PATH = os.path.join(BASE_DIR, 'Config.xlsx')
DIST_PATH = os.path.join(BASE_DIR, 'טבלת מיקומים.xlsx')

def merge_files():
    print("Starting merge process...")
    
    if not os.path.exists(CONFIG_PATH):
        print("Config.xlsx not found!")
        return
    if not os.path.exists(DIST_PATH):
        print("Distance file not found!")
        return

    # Backup Config
    backup_path = CONFIG_PATH.replace('.xlsx', f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copyfile(CONFIG_PATH, backup_path)
    print(f"Backed up Config to: {os.path.basename(backup_path)}")

    try:
        # Load existing sheets
        # appending to existing excel properly often requires openpyxl mode='a' but pandas replace is easier:
        # read all, add new df, write all. Safe for small files.
        
        xls = pd.ExcelFile(CONFIG_PATH)
        sheet_dict = {}
        for sheet_name in xls.sheet_names:
            sheet_dict[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
            
        print(f"Loaded existing sheets: {list(sheet_dict.keys())}")
        
        # Load Distances
        dist_df = pd.read_excel(DIST_PATH)
        print(f"Loaded Distances: {dist_df.shape}")
        
        # Add to dict
        sheet_dict['Distances'] = dist_df
        
        # Write back
        with pd.ExcelWriter(CONFIG_PATH, engine='xlsxwriter') as writer:
            for name, df in sheet_dict.items():
                df.to_excel(writer, sheet_name=name, index=False)
                print(f"Wrote sheet: {name}")
                
        print("SUCCESS: Config.xlsx updated with Distances sheet.")
        
    except Exception as e:
        print(f"ERROR: {e}")
        # Restore backup if failed?
        shutil.copyfile(backup_path, CONFIG_PATH)
        print("Restored backup due to error.")

if __name__ == "__main__":
    merge_files()
