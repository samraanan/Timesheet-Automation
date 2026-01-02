import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime, timedelta
import sys
try:
    sys.stdout.reconfigure(encoding='utf-8')
except: pass
import warnings

# התעלמות מאזהרות עיצוב של קבצי מקור ישנים
warnings.simplefilter("ignore")

# --- GLOBAL CONFIG VARIABLES ---
CONFIG = {
    'MAPPING': {},
    'HOME_LOCATION': "בית (כפר אלדד)", 
    'MAPPING': {},
    'HOME_LOCATION': "בית (כפר אלדד)", 
    'IGNORE_KEYWORDS': ['סה"כ', 'Total', 'Grand Total', 'משך', 'משך יחסי'], # Expanded ignore list
    'ACTIVE_PROJECTS': []
}

ALLOWED_COLUMNS = [
    'תאריך', 'שעת התחלה', 'שעת סיום', 'פרויקט', 
    'תיאור', 'הפסקות', 'הערות'
]

# Constants
KM_COL_NAME = 'ק"מ'
TOTAL_COL_NAME = 'סה"כ'

# Helper for normalization (Global)
def normalize_str(s):
    if pd.isna(s): return ""
    s = str(s).strip()
    # Replace Gershayim with Quote
    s = s.replace('״', '"').replace("''", '"').replace('`', "'")
    # Clean invisible chars just in case
    s = s.replace('\u200b', '').replace('\ufeff', '')
    return s

def load_configuration():
    """Loads settings and mappings from Config.xlsx"""
    # חיפוש קובץ קונפיג בתיקייה הנוכחית
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_dir, 'Config.xlsx')
    
    if not os.path.exists(config_path):
        print("Config.xlsx not found automatically. Please select it.")
        root = tk.Tk()
        root.withdraw()
        config_path = filedialog.askopenfilename(title="Select Config.xlsx File")
        if not config_path:
            print("CRITICAL: No config file selected.")
            sys.exit(1)

    print(f"Loading configuration from: {config_path}")
    
    try:
        # 1. Load Projects Mapping
        df_map = pd.read_excel(config_path, sheet_name='Projects')
        print(f"[DEBUG] Loaded Projects Sheet. Shape: {df_map.shape}")
        print(f"[DEBUG] Columns: {df_map.columns.tolist()}")
        if not df_map.empty:
            print(f"[DEBUG] First Row: {df_map.iloc[0].tolist()}")
        
        mapping = {}
        active_list = []
        
        # זיהוי עמודות לפי שם (Dynamic Column matching)
        cols = [str(c).lower().strip() for c in df_map.columns]
        
        # defaults
        col_idx_proj = 0
        col_idx_map = -1
        col_idx_active = -1
        
        # Try to find headers
        cols = [str(c).lower().strip() for c in df_map.columns]
        for i, c in enumerate(cols):
            if any(k in c for k in ['project', 'פרויקט', 'שם']):
                col_idx_proj = i
            elif any(k in c for k in ['map', 'alias', 'מיפוי']):
                col_idx_map = i
            elif any(k in c for k in ['active', 'פעיל', 'לכלול']):
                col_idx_active = i

        # Fallback if specific headers not found but multiple columns exist
        if col_idx_active == -1 and len(cols) > 2: col_idx_active = 2 # A, B, C -> C is Active? or User said A is Active.
        # User screenshot: A=Active, B=Map, C=Project. 
        # If headers are missing/wrong, we might be in trouble. 
        # But if 'Active' is found (as per debug output `['Active']`), then `col_idx_active` should be 0.
        
        print(f"[DEBUG] Indices: Proj={col_idx_proj}, Active={col_idx_active}, Map={col_idx_map}")

        for _, row in df_map.iterrows():
            # Get Project Name
            if col_idx_proj < len(row):
                proj_name = row.iloc[col_idx_proj]
            else:
                # Fallback: if we only have 1 column and it's named 'Active' but contains 'Or' (project name)?
                # That means the header is wrong. 
                # Let's assume col 0 is ALWAYS project if explicit headers fail?
                # User screenshot had Project in Col C. 
                # If pandas sees only ['Active'], maybe it read the first row as header and it only had data in Col A?
                # But 'Or' is in the data.
                # If `row` has 'Or', and `col_idx_active`=0 (Header 'Active'), then `is_active`='Or'.
                # `is_active` check `in ['yes'...]`. 'Or' is not in list. Active=False.
                # Result: No projects loaded.
                # We need to correctly identify Project column.
                
                # Heuristic: If we have 1 column, it is likely the Project Name (legacy config).
                # But header says 'Active'.
                # Let's trust the content? No, strictly follow headers if present.
                proj_name = None

            # Get Active Status
            is_active = "yes" # Default to yes if column missing? No, default policy is INACTIVE. 
            
            if col_idx_active != -1 and col_idx_active < len(row):
                 is_active = str(row.iloc[col_idx_active]).lower()
                 
                 # SPECIAL HANDLING: If `col_idx_proj` == `col_idx_active` (e.g. only 1 col 'Active' which is actually Project Names)
                 # Then the value is the project name, NOT 'yes'/'no'.
                 # If we detect this overlap, we assume Active=Yes.
                 if col_idx_proj == col_idx_active:
                     is_active = 'yes'
            
            elif col_idx_active == -1:
                if len(cols) == 1:
                     # If only 1 column exists, assume it is list of ACTIVE projects.
                     is_active = 'yes' 

            # Get Map

            # Get Map
            map_name = None
            if col_idx_map != -1 and col_idx_map < len(row):
                map_name = row.iloc[col_idx_map]

            if pd.isna(proj_name): continue
            
            proj_name = normalize_str(proj_name)
            is_active = str(is_active).lower()
            
            # Special case: If header is 'Active' but values are 'Or', 'Dror'...
            # Then col 0 is Project.
            # If col_idx_proj points to same as active?
            # Let's refine:
            # If Project Not Found, but Active Found: Use Active Column as Project if values look like strings?
            # Dangerous.
            
            # Let's just fix the crash first.
            if is_active in ['yes', 'true', '1', 'כן', 'פעיל', '1.0']:
                active_list.append(proj_name)
            
            if map_name and pd.notna(map_name):
                 mapping[proj_name] = normalize_str(map_name)
            else:
                 mapping[proj_name] = proj_name
                
        CONFIG['MAPPING'] = mapping
        CONFIG['ACTIVE_PROJECTS'] = active_list
        
        # 2. Load General Settings
        try:
            df_settings = pd.read_excel(config_path, sheet_name='Settings')
            for _, row in df_settings.iterrows():
                key = str(row.iloc[0]).strip()
                val = str(row.iloc[1]).strip()
                
                if key == 'Home_Location':
                    CONFIG['HOME_LOCATION'] = val
                elif key == 'Ignore_Keywords':
                    keywords = [x.strip() for x in val.split(',')]
                    CONFIG['IGNORE_KEYWORDS'] = keywords
        except:
             print("Warning: Settings sheet not found or empty.")

        # 3. Load Distances (New)
        try:
            df_dist = pd.read_excel(config_path, sheet_name='Distances')
            CONFIG['DISTANCE_DF'] = df_dist
            print("Loaded Distances from Config.xlsx")
        except:
            CONFIG['DISTANCE_DF'] = None
            print("Distances sheet not found in Config.xlsx (will ask for file).")
            
    except Exception as e:
        print(f"CRITICAL ERROR loading Config: {e}")
        sys.exit(1)

def load_data_files():
    """Selects data files. Supports CLI args for automation."""
    
    dist_file = None
    timesheet_files = []

    # Check if Distances are already loaded from Config
    has_config_dist = CONFIG.get('DISTANCE_DF') is not None

    # Check for CLI args: script.py <timesheet_file1> ... (Dist optional if in config)
    if len(sys.argv) > 1:
        print("Using file paths from command line arguments...")
        args = sys.argv[1:]
        
        for path in args:
            # If we don't have config dist, we look for it in args
            if not has_config_dist and ("מיקומים" in path or "Distance" in path):
                dist_file = path
            else:
                # If we DO have config dist, but user passed it anyway, we can ignore or override.
                # Let's assume args are timesheets if config has dist.
                if "מיקומים" not in path and "Distance" not in path:
                    timesheet_files.append(path)
        
        if (has_config_dist or dist_file) and timesheet_files:
            return dist_file, timesheet_files

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    if has_config_dist:
        print("Please select Timesheet files...")
        title = "בחר קבצי שעות(המיקומים נטענו מהקונפיג)"
    else:
        print("Please select Data Files (Distance Matrix + Timesheets)...")
        title = "בחר את קובץ המיקומים ואת קבצי השעות"

    file_paths = filedialog.askopenfilenames(title=title)
    
    for path in file_paths:
        filename = os.path.basename(path)
        if "Config" in filename:
            continue
        
        # Identify distance file only if needed
        if not has_config_dist and ("מיקומים" in filename or "Distance" in filename):
            dist_file = path
        else:
            timesheet_files.append(path)
            
    return dist_file, timesheet_files

def create_distance_matrix(source):
    """Source can be file path or DataFrame"""
    if isinstance(source, pd.DataFrame):
        df = source
    elif isinstance(source, str) and source:
         try:
            df = pd.read_csv(source) if source.endswith('csv') else pd.read_excel(source)
         except: return {}
    else:
        return {}

    df.columns = df.columns.str.strip()
    dist_map = {}
    
    # Identify columns flexibly
    col_dist = next((c for c in df.columns if 'מרחק' in c), None)
    col_p1 = next((c for c in df.columns if "מיקום א" in c), None)
    col_p2 = next((c for c in df.columns if "מיקום ב" in c), None)
    
    if not (col_dist and col_p1 and col_p2):
        print("Warning: Distance columns not identified correctly.")
        return {}

    for _, row in df.iterrows():
        try:
            d_val = str(row[col_dist]).replace(KM_COL_NAME, '').strip()
            d = float(d_val)
        except: d = 0.0
        
        p1 = str(row[col_p1]).strip()
        p2 = str(row[col_p2]).strip()
        
        dist_map[(p1, p2)] = d
        dist_map[(p2, p1)] = d
        
    return dist_map

def calculate_daily_km(daily_projects, dist_map):
    home = CONFIG['HOME_LOCATION'].strip()
    valid_projects = []
    
    for p in daily_projects:
        if pd.isna(p): continue
        p_str = str(p).strip()
        
        # Apply mapping
        mapped_loc = CONFIG['MAPPING'].get(p_str, p_str)
        # Normalize
        mapped_loc = mapped_loc.strip()
        
        if mapped_loc: valid_projects.append(mapped_loc)
            
    if not valid_projects: return 0.0
    
    # מסלול: בית -> פרויקט 1 -> פרויקט 2 -> ... -> בית
    route = [home] + valid_projects + [home]
    
    total_km = 0.0
    for i in range(len(route) - 1):
        origin, dest = route[i], route[i+1]
        if origin == dest: continue
        
        dist = dist_map.get((origin, dest))
        
        # Check reverse direction if not found
        if dist is None:
            dist = dist_map.get((dest, origin))
        
        if dist is None:
            # print(f"[DEBUG] Missing distance: '{origin}' -> '{dest}'")
            try:
                # Log explicitly for the user - using repr to avoid encoding errors
                print(f"Warning: Missing distance definition between: {repr(origin)} <--> {repr(dest)}. Using 0.0km.")
            except Exception as e:
                print(f"Warning: Missing distance between points (CRITICAL display error: {e}). Using 0.0km.")
            dist = 0.0
        total_km += dist
        
    return total_km

def parse_duration(x):
    if pd.isna(x) or str(x).strip() == '': return timedelta(0)
    try:
        if isinstance(x, timedelta): return x
        if isinstance(x, datetime): return timedelta(hours=x.hour, minutes=x.minute)
        s = str(x).replace('null', '').strip()
        parts = s.split(':')
        if len(parts) == 3: return timedelta(hours=int(parts[0]), minutes=int(parts[1]), seconds=int(parts[2]))
        if len(parts) == 2: return timedelta(hours=int(parts[0]), minutes=int(parts[1]))
    except: pass
    return timedelta(0)

def main():
    # 1. LOAD CONFIG
    load_configuration()
    
    # 2. LOAD FILES
    dist_file, ts_files = load_data_files()
    
    # Check if we have distances from Config or File
    if CONFIG.get('DISTANCE_DF') is not None:
        print("Using Distances from Config.")
        dist_map = create_distance_matrix(CONFIG['DISTANCE_DF'])
    elif dist_file:
        dist_map = create_distance_matrix(dist_file)
    else:
        print("ERROR: Missing Distance Matrix file ('מיקומים') or Config sheet.")
        return

    if not ts_files:
        print("ERROR: No Timesheet files selected.")
        return
    all_data = []
    
    print("Processing files...")
    for ts in ts_files:
        try:
            df = pd.read_csv(ts) if ts.endswith('csv') else pd.read_excel(ts)
            df.columns = df.columns.str.strip()
            cols = [c for c in ALLOWED_COLUMNS if c in df.columns]
            all_data.append(df[cols])
        except Exception as e: print(f"Error reading {ts}: {e}")

    if not all_data: return
    
    full_df = pd.concat(all_data, ignore_index=True)
    
    # NORMALIZE PROJECT COLUMN IMMEDIATELY
    full_df['פרויקט'] = full_df['פרויקט'].apply(normalize_str)
    
    # NORMALIZE CONFIG MAPPING KEYS/VALUES + IGNORE LIST
    if CONFIG.get('MAPPING'):
        new_map = {}
        for k, v in CONFIG['MAPPING'].items():
            new_map[normalize_str(k)] = normalize_str(v)
        CONFIG['MAPPING'] = new_map
        
    if CONFIG.get('IGNORE_KEYWORDS'):
        CONFIG['IGNORE_KEYWORDS'] = [normalize_str(k) for k in CONFIG['IGNORE_KEYWORDS']]

    # --- 2. MAP ALIASES ("בית" -> "בית (כפר אלדד)") ---
    # Apply mapping immediately to the 'Project' column
    def apply_mapping(val):
        # val is already normalized above
        s = val 
        # Direct lookup
        if s in CONFIG['MAPPING']:
            return CONFIG['MAPPING'][s]
        # Check for "House" variations if requested
        if s in ['בית', 'הבית']:
            # If Home location is defined, map to it
            home = CONFIG.get('HOME_LOCATION')
            if home: return home
        return s
    
    full_df['פרויקט'] = full_df['פרויקט'].apply(apply_mapping)

    # --- 3. FILTER BY CONFIG ---
    if CONFIG['IGNORE_KEYWORDS']:
        # Ensure pattern also uses normalized form? 
        # We already normalized full_df and Ignore list.
        # Strict check
        full_df = full_df[~full_df['פרויקט'].isin(CONFIG['IGNORE_KEYWORDS'])]
        # Fuzzy check (careful not to filter "Or" if "Or Something" is ignored)
        # We'll rely on strict mostly, but if user wants fuzzy:
        # full_df = full_df[~full_df['פרויקט'].str.contains(pattern...)]
        # For now, stick to exact exclusion of "Total", "Grand Total", etc.
        pass
    # הסרת מילות מפתח
    if CONFIG['IGNORE_KEYWORDS']:
        pattern = '|'.join(CONFIG['IGNORE_KEYWORDS'])
        # 2026-01-02: Enhanced filtering to strictly drop rows where PROJECT column matches these
        full_df = full_df[~full_df['פרויקט'].astype(str).str.fullmatch(pattern, case=False, na=False)]
        # Also contains logic for robust safety
        full_df = full_df[~full_df['פרויקט'].astype(str).str.contains(pattern, case=False, na=False)]
    
    # --- 2. MAP ALIASES ("בית" -> "בית (כפר אלדד)") ---
    # Apply mapping immediately to the 'Project' column
    def apply_mapping(val):
        s = str(val).strip()
        # Direct lookup
        if s in CONFIG['MAPPING']:
            return CONFIG['MAPPING'][s]
        # Check for "House" variations if requested
        if s in ['בית', 'הבית']:
            # If Home location is defined, map to it
            home = CONFIG.get('HOME_LOCATION')
            if home: return home
        return s
    
    full_df['פרויקט'] = full_df['פרויקט'].apply(apply_mapping)

    full_df['פרויקט'] = full_df['פרויקט'].apply(apply_mapping)
    
    # --- 2.5 NUCLEAR FILTER FOR 'TOTAL' ROWS ---
    # The source file contains summary rows named 'סה"כ' or 'Total'.
    # We must drop them immediately to prevent them from becoming projects.
    # We use a hardcoded list in addition to Config to be safe.
    drop_names = ['סה"כ', 'סה״כ', 'Total', 'Grand Total', 'Start', 'End', 'Duration']
    # Normalize drop list
    drop_names = [normalize_str(x) for x in drop_names]
    
    # Filter
    full_df = full_df[~full_df['פרויקט'].isin(drop_names)]
    # Also fuzzy filter for robust safety
    for dn in drop_names:
        full_df = full_df[~full_df['פרויקט'].astype(str).str.contains(dn, case=False, na=False)]

    # --- 3. FILTER BY CONFIG ---
    if CONFIG['IGNORE_KEYWORDS']:
        pattern = '|'.join(CONFIG['IGNORE_KEYWORDS'])
        # STRICT FILTER: Drop rows where Project is exactly an ignored keyword
        full_df = full_df[~full_df['פרויקט'].isin(CONFIG['IGNORE_KEYWORDS'])]
        # Also drop rows where Project CONTAINS ignored keyword (fuzzy)
        full_df = full_df[~full_df['פרויקט'].astype(str).str.contains(pattern, case=False, na=False)]

    active_list = CONFIG['ACTIVE_PROJECTS']
    
    # Debug: Print found projects
    print(f"Projects in File: {full_df['פרויקט'].unique()}")
    print(f"Active List from Config: {active_list}")

    if active_list:
        # Filter: Keep if Project is in Active List
        full_df = full_df[full_df['פרויקט'].isin(active_list)]
    else:
        # If active list is empty, default to NOTHING (User requirement)
        print("WARNING: No active projects defined in Config. Result will be empty.")
        full_df = full_df[full_df['פרויקט'].isin([])] # Empty

    # המרות זמנים
    full_df['תאריך'] = pd.to_datetime(full_df['תאריך'], dayfirst=True, errors='coerce')
    full_df = full_df.dropna(subset=['תאריך'])
    
    def parse_time_str(s):
        """המרת טקסטים כמו '08:00' ל־datetime.time"""
        if pd.isna(s): return None
        
        # טיפול במספרים (שעות באקסל כחלק יחסי מיום)
        if isinstance(s, (float, int)):
            try:
                # המרה משבר של יום לשעה
                # 0.5 = 12:00
                total_seconds = int(float(s) * 24 * 3600)
                m, sec = divmod(total_seconds, 60)
                h, m = divmod(m, 60)
                return (datetime.min + timedelta(hours=h, minutes=m, seconds=sec)).time()
            except:
                pass

        s = str(s).strip()
        try:
            # אם יש תאריך מלא
            dt = pd.to_datetime(s, errors='coerce')
            if pd.notna(dt):
                return dt.time() if hasattr(dt, 'time') else dt
            # אם רק שעה ודקה
            parts = s.split(':')
            if len(parts) >= 2:
                h, m = int(parts[0]), int(parts[1])
                # Drop seconds for cleaner matching/deduplication
                return datetime.strptime(f"{h:02d}:{m:02d}", "%H:%M").time()
        except Exception as e:
            print(f"Error parsing time '{s}': {e}")
            return None
        return None

    # המרת שעת התחלה וסיום לפורמט זמן
    full_df['שעת התחלה'] = full_df['שעת התחלה'].apply(parse_time_str)
    full_df['שעת סיום'] = full_df['שעת סיום'].apply(parse_time_str)

    # הדפסת ערכים לבדיקה
    # print("Start Time Example:", full_df['שעת התחלה'].head())
    # print("End Time Example:", full_df['שעת סיום'].head())

    # הסרת שורות עם ערכים חסרים בשעות
    full_df = full_df.dropna(subset=['שעת התחלה', 'שעת סיום'])

    # הסרת שורות כפולות (אותו זמן בדיוק) - למניעת כפילות אור/דרור
    # עיגול לדקות ב-parse_time_str מאפשר זיהוי חפיפות
    # שינוי: בדיקת זמן התחלה בלבד כדי לתפוס חפיפות גם אם זמן סיום שונה במעט
    before_dedup = len(full_df)
    full_df = full_df.drop_duplicates(subset=['תאריך', 'שעת התחלה'], keep='last')
    after_dedup = len(full_df)
    if before_dedup > after_dedup:
        print(f"Removed {before_dedup - after_dedup} duplicate time overlaps from source data.")
    
    # חישוב משך
    def calc_duration(row):
        try:
            t1 = row['שעת התחלה']
            t2 = row['שעת סיום']
            
            if pd.isna(t1) or pd.isna(t2):
                return timedelta(0)

            # שימוש בתאריך שרירותי כדי לחשב הפרש
            dt1 = datetime.combine(datetime.today(), t1)
            dt2 = datetime.combine(datetime.today(), t2)
            
            duration = dt2 - dt1
            
            # אם יצא שלילי (עבודה מעבר לחצות), נוסיף יום
            if duration < timedelta(0):
                duration += timedelta(days=1)
                
            return duration
        except Exception as e:
            print(f"Error calculating duration row: {row.values} -> {e}")
            return timedelta(0)

    full_df['Duration'] = full_df.apply(calc_duration, axis=1)
    full_df['הפסקות_delta'] = full_df['הפסקות'].apply(parse_duration)

    # print("Duration Example:", full_df['Duration'].head())

    full_df = full_df.sort_values(by=['תאריך'])

    print(f"Active Projects Found: {len(full_df['פרויקט'].unique())}")

    # --- SPLIT BY MONTH ---
    # Create a Year-Month column for grouping
    full_df['Month'] = full_df['תאריך'].dt.to_period('M')
    
    monthly_groups = full_df.groupby('Month')
    
    print(f"Found data for {len(monthly_groups)} months. Generating reports...")

    for period, month_df in monthly_groups:
        period_str = str(period)
        print(f"\n--- Processing Month: {period_str} ---")
        
        # --- REPORT GENERATION (Per Month) ---
        
        # 1. Project Summary Data
        pivot_proj = month_df.pivot_table(
            index='תאריך', 
            columns='פרויקט', 
            values='Duration', 
            aggfunc='sum'
        ).fillna(timedelta(0))
        pivot_proj.loc['Grand Total'] = pivot_proj.sum()

        # 2. Daily Stats Calculation
        daily_stats = []
        grouped = month_df.groupby(month_df['תאריך'].dt.date)
        
        for date_obj, group in grouped:
            entry = group['שעת התחלה'].min()
            exit_time = group['שעת סיום'].max()

            # המרת השעות ל-datetime מלא (אותו יום)
            # base_date = datetime.combine(date_obj, datetime.min.time()) # Unused
            entry_dt = datetime.combine(date_obj, entry)
            exit_dt = datetime.combine(date_obj, exit_time)

            total_breaks = group['הפסקות_delta'].sum()
            gross_presence = exit_dt - entry_dt
            net_time = gross_presence - total_breaks

            # חישוב קילומטרים (לפי סדר כרונולוגי של התחלה)
            sorted_group = group.sort_values('שעת התחלה')
            projects_sequence = sorted_group['פרויקט'].tolist()
            km = calculate_daily_km(projects_sequence, dist_map)

            row_data = {
                'תאריך': date_obj,
                'כניסה': entry,
                'יציאה': exit_time,
                'הפסקות': total_breaks,
                'סה"כ נטו': net_time,
                KM_COL_NAME: km
            }

            # הוספת פירוט לכל פרויקט
            proj_sums = group.groupby('פרויקט')['Duration'].sum()
            for proj, duration in proj_sums.items():
                row_data[proj] = duration

            daily_stats.append(row_data)
            
        stats_df = pd.DataFrame(daily_stats)
        
        # --- EXPORT ---
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_filename = os.path.join(base_dir, f"Report_{period_str}_{timestamp}.xlsx")
        
        print(f"Generating Excel: {output_filename} ...")
        
        # --- HELPER: FORMAT ZEROS AS BLANK ---
        def fmt_zero(val):
            """Returns empty string if 0, else value"""
            if pd.isna(val) or val == "": return ""
            # Check for zero numeric
            if isinstance(val, (int, float)) and abs(val) < 0.000001: return ""
            if isinstance(val, timedelta) and val.total_seconds() == 0: return ""
            # Check strings
            s = str(val).strip()
            if s in ['0', '0.0', '00:00:00', '00:00', '0:00']: return ""
            return val

        # Helper to convert timedelta to decimal hours
        def to_hours(x):
            if isinstance(x, timedelta):
                return x.total_seconds() / 3600
            return x

        # Create export versions with decimal hours
        pivot_proj_export = pivot_proj.applymap(to_hours)
        
        # Calculate detailed report export
        detailed_df = stats_df.copy()
        
        # RENAME 'סה"כ נטו' -> TOTAL
        if 'סה"כ נטו' in detailed_df.columns:
            detailed_df.rename(columns={'סה"כ נטו': TOTAL_COL_NAME}, inplace=True)
        
        # Columns handling
        base_cols = ['תאריך', 'כניסה', 'יציאה']
        end_cols = ['הפסקות', TOTAL_COL_NAME, KM_COL_NAME]
        
        # Dynamically find project columns
        existing_cols = detailed_df.columns.tolist()
        
        # NORMALIZE COLUMNS: Replace Gershayim with Quote to handle potential encoding mismatches
        # This ensures 'ק״מ' becomes 'ק"מ'
        new_cols = []
        for c in existing_cols:
            c_norm = str(c).replace('״', '"').replace("''", '"')
            if c_norm != c:
               print(f"[DEBUG] Renaming column '{c}' to '{c_norm}'")
            new_cols.append(c_norm)
        detailed_df.columns = new_cols
        
        # CRITICAL: Remove duplicate columns immediately after normalization normalization
        detailed_df = detailed_df.loc[:, ~detailed_df.columns.duplicated()]
        
        existing_cols = detailed_df.columns.tolist() # Update list for filtering
        
        # FIX DUPLICATE KM: Aggressive filter
        # exclude anything that looks like KM or Total
        def is_prohibited(c_name):
            s = str(c_name).strip()
            if s in base_cols + end_cols: return True
            if s in [TOTAL_COL_NAME, KM_COL_NAME, 'Total', 'Grand Total', 'ק"מ.1']: return True
            # Check for KM variations
            if 'ק"מ' in s: return True 
            if 'KM' in s.upper(): return True
            return False

        proj_cols = [c for c in existing_cols if not is_prohibited(c)]
        
        # Deduplicate final columns list just in case
        final_cols_pre = base_cols + proj_cols + end_cols
        final_cols = list(dict.fromkeys(final_cols_pre))
        
        # Convert Timedeltas to Hours
        cols_to_convert = ['הפסקות'] + proj_cols
        for c in cols_to_convert:
             if c in detailed_df.columns:
                detailed_df[c] = detailed_df[c].apply(to_hours)
                
        # Fill 'Total' (סה"כ) with 0.0
        detailed_df[TOTAL_COL_NAME] = 0.0

        # Grand Total Row for Summary Tables
        sum_cols = proj_cols + ['הפסקות', TOTAL_COL_NAME, KM_COL_NAME]
        
        # Ensure only existing columns are summed
        valid_sum_cols = [c for c in sum_cols if c in detailed_df.columns]
        current_sum = detailed_df[valid_sum_cols].sum(numeric_only=True)
        detailed_df.loc['Total'] = pd.Series(current_sum)
        
        # CRITICAL: Remove duplicate columns from DF *BEFORE* reindex
        # This handles case where multiple columns have name 'ק"מ'
        detailed_df = detailed_df.loc[:, ~detailed_df.columns.duplicated()]
        
        # Also ensure 'KM' column exists if missing
        if KM_COL_NAME not in detailed_df.columns:
            detailed_df[KM_COL_NAME] = 0.0

        # Reorder columns manually to avoid duplicates
        # detailed_export = detailed_df.reindex(columns=final_cols) <-- Causes matching duplicates
        
        detailed_export = pd.DataFrame()
        for c in final_cols:
            if c in detailed_df.columns:
                # If column exists multiple times, it returns a DataFrame
                col_data = detailed_df[c]
                if isinstance(col_data, pd.DataFrame):
                    # Take the first one only
                    print(f"[WARN] Found duplicate column '{c}' in dataframe. Using first instance.")
                    detailed_export[c] = col_data.iloc[:, 0]
                else:
                    detailed_export[c] = col_data
            else:
                # Missing column (shouldn't happen for core ones, but maybe optional cols)
                pass

        # Ensure 'Total' and 'KM' are present (created above)
        
        # FINAL DEBUG: Check detailed_export columns
        print(f"[DEBUG] Final Export Columns: {detailed_export.columns.tolist()}")
        # Check for duplicates explicitly
        if detailed_export.columns.duplicated().any():
             print("[CRITICAL WARN] Duplicates found in detailed_export columns! Dropping...")
             detailed_export = detailed_export.loc[:, ~detailed_export.columns.duplicated()]
             print(f"[DEBUG] Corrected Columns: {detailed_export.columns.tolist()}")

        # Executive Summary
        # Rename there too
        exec_cols = ['תאריך', 'כניסה', 'יציאה', 'הפסקות', TOTAL_COL_NAME, KM_COL_NAME]
        exec_df = detailed_export[exec_cols].copy()

        # Clean Zeros Function for DataFrame (Values only)
        # We only apply this to non-formula, non-time columns that we write value-by-value
        # For columns where we write FORMULAS, we will handle zero-hiding via Number Format or Conditional Formatting if possible, 
        # but Excel formulas returning 0 show 0. We'll use custom number format to hide it: #,##0.00;-#,##0.00;;
        
        try:
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Formats
                date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                time_fmt = workbook.add_format({'num_format': 'hh:mm'})
                # Custom format: Positive; Negative; Zero (Hidden); Text
                decimal_fmt = workbook.add_format({'num_format': '#,##0.00;-#,##0.00;;'})   
                num_fmt = workbook.add_format({'num_format': '#,##0.0;-#,##0.0;;'})
                
                # --- SHEET 1: PROJECT SUMMARY ---
                # Clean Zeros for Display
                pivot_display = pivot_proj_export.applymap(lambda x: fmt_zero(x) if x == 0 else x)
                pivot_display.to_excel(writer, sheet_name='Project Summary')
                
                ws_summ = writer.sheets['Project Summary']
                ws_summ.right_to_left() # RTL
                ws_summ.set_column('A:A', 12, date_fmt)
                ws_summ.set_column('B:Z', 12, decimal_fmt)

                # --- SHEET 2: DETAILED REPORT ---
                # Write data (without formulas first)
                # Apply fmt_zero only to columns we are writing as VALUES (Projects, KM, Breaks)
                # We leave 'Total' alone as we overwrite it
                
                # --- REMOVE REDUNDANT 'GRAND TOTAL' ---
                # Fixed logic below using display_df
                
                # Let's fix the dataframe BEFORE writing
                if 'Total' in detailed_df.index:
                     detailed_df.at['Total', 'תאריך'] = 'Total'
                
                # Revert to using detailed_df for writing logic to ensure we control the 'Total' row
                # We need to re-apply the fmt_zero to a display copy
                display_df = detailed_df.copy()
                try:
                    display_df.drop(index='Grand Total', inplace=True, errors='ignore')
                except: pass
                
                # Format Zeros
                for c in display_df.columns:
                     if c in proj_cols + ['הפסקות', KM_COL_NAME]:
                          display_df[c] = display_df[c].apply(lambda x: fmt_zero(x) if x == 0 else x)

                # Write
                display_df.to_excel(writer, sheet_name='Detailed Report', index=False)
                
                ws_detail = writer.sheets['Detailed Report']
                ws_detail.right_to_left() # RTL
                ws_detail.set_column('A:A', 20, date_fmt) 
                ws_detail.set_column('B:C', 10, time_fmt)
                
                # Dynamic Column Indices
                headers = display_df.columns.tolist()
                
                # Find Notes column index for Text Wrap
                try:
                    col_notes = headers.index('הערות')
                    wrap_fmt = workbook.add_format({'text_wrap': True, 'align': 'right', 'valign': 'top'})
                    ws_detail.set_column(col_notes, col_notes, 40, wrap_fmt)
                except ValueError: pass
                
                # Find column indices (0-based)
                col_total_idx = -1
                try:
                    col_entry = headers.index('כניסה')
                    col_exit = headers.index('יציאה')
                    col_breaks = headers.index('הפסקות')
                    col_total_idx = headers.index(TOTAL_COL_NAME)
                except ValueError: pass
                
                # Write Formulas for 'Total' column
                # Rows range: 2 to len(df)+1.
                if col_total_idx != -1:
                    num_rows = len(display_df)
                    for r in range(num_rows):
                        xl_row = r + 1
                        
                        # Check if this is the Total row
                        val_date = str(display_df.iloc[r]['תאריך'])
                        
                        if val_date == 'Total':
                            # Write SUM formula (Bottom Row)
                            # =SUM(E2:E{last_data})
                            # This answers "Formula in summary row should be different"
                            col_letter = chr(ord('A') + col_total_idx)
                            last_data_row = xl_row # Row BEFORE this one (Excel row is 1-based, previous is row_idx)
                            formula = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
                            ws_detail.write_formula(xl_row, col_total_idx, formula, workbook.add_format({'bold': True, 'num_format': '0.00'}))
                            # Also write label bold
                            ws_detail.write(xl_row, 0, 'Total', workbook.add_format({'bold': True}))
                            continue

                        # Normal Row Formula
                        # (Exit-Entry)*24 - Breaks
                        let_entry = chr(ord('A') + col_entry)
                        let_exit = chr(ord('A') + col_exit)
                        let_break = chr(ord('A') + col_breaks)
                        formula = f"=(({let_exit}{xl_row+1}-{let_entry}{xl_row+1})*24)-{let_break}{xl_row+1}"
                        ws_detail.write_formula(xl_row, col_total_idx, formula, decimal_fmt)

                # Formats
                ws_detail.set_column(3, len(headers)-1, 12, decimal_fmt) 

                # --- SHEET 3: EXECUTIVE SUMMARY ---
                exec_export = exec_df.copy()
                
                # Remove existing 'Total' row from index if present (inherited from detailed_df)
                # to prevent double counting and duplicate rows (NaN date + Total date)
                try:
                    exec_export.drop(index='Total', inplace=True, errors='ignore')
                except: pass
                
                # Append TOTAL ROW to Exec Summary
                # Sum numeric cols
                sum_row = exec_export[['הפסקות', TOTAL_COL_NAME, KM_COL_NAME]].sum(numeric_only=True)
                exec_export.loc[len(exec_export)] = {
                    'תאריך': 'Total',
                    'כניסה': '',
                    'יציאה': '',
                    'הפסקות': sum_row['הפסקות'],
                    TOTAL_COL_NAME: sum_row[TOTAL_COL_NAME],
                    KM_COL_NAME: sum_row[KM_COL_NAME]
                }

                exec_export[TOTAL_COL_NAME] = exec_export[TOTAL_COL_NAME].apply(lambda x: fmt_zero(x) if x == 0 else x)
                exec_export[KM_COL_NAME] = exec_export[KM_COL_NAME].apply(lambda x: fmt_zero(x) if x == 0 else x)
                # Breaks is Decimal Hours now, don't format as time
                
                exec_export.to_excel(writer, sheet_name='Executive Summary', index=False)
                ws_exec = writer.sheets['Executive Summary']
                ws_exec.right_to_left() # RTL
                ws_exec.set_column('A:A', 20, date_fmt)
                ws_exec.set_column('B:C', 10, time_fmt)
                ws_exec.set_column('D:D', 10, decimal_fmt) # Breaks -> DECIMAL
                ws_exec.set_column('E:E', 12, decimal_fmt) # Total -> DECIMAL
                ws_exec.set_column('F:F', 10, num_fmt)     # KM

                # Add Formulas to Executive Summary
                # Loop all rows including Total
                for r in range(len(exec_export)):
                    xl_row = r + 1
                    
                    # Check if Total Row
                    if str(exec_export.iloc[r]['תאריך']) == 'Total':
                         # SUM Formula for Total Column (E)
                         # =SUM(E2:E{last_data})
                         col_idx = 4 # E
                         col_let = 'E'
                         last_data = xl_row 
                         formula = f"=SUM({col_let}2:{col_let}{last_data})"
                         ws_exec.write_formula(xl_row, col_idx, formula, workbook.add_format({'bold': True, 'num_format': '0.00'}))
                         # Bold Label
                         ws_exec.write(xl_row, 0, 'Total', workbook.add_format({'bold': True}))
                         continue

                    # Normal Row Formula
                    # (Out-In)*24 - Breaks
                    formula = f"=(C{xl_row+1}-B{xl_row+1})*24-D{xl_row+1}"
                    ws_exec.write_formula(xl_row, 4, formula, decimal_fmt)

                # --- SHEET 4+: PER PROJECT SHEETS ---
                monthly_projects = month_df['פרויקט'].unique()
                
                for proj in monthly_projects:
                    if pd.isna(proj): continue
                    proj_name = str(proj).strip()
                    if not proj_name: continue
                    
                    p_df = month_df[month_df['פרויקט'] == proj_name].copy()
                    if p_df.empty: continue
                    
                    p_df = p_df.sort_values(by=['תאריך', 'שעת התחלה'])
                    
                    proj_rows = []
                    
                    for idx, row in p_df.iterrows():
                        desc = str(row['תיאור']) if pd.notna(row['תיאור']) else ""
                        notes = str(row['הערות']) if pd.notna(row['הערות']) else ""
                        full_notes = f"{desc} {notes}".strip()
                        
                        proj_rows.append({
                            'תאריך': row['תאריך'],
                            'כניסה': row['שעת התחלה'],
                            'יציאה': row['שעת סיום'],
                            'הערות': full_notes,
                            'סה"כ': 0.0 # Placeholder
                        })
                            
                    if not proj_rows: continue
                    
                    proj_sheet_df = pd.DataFrame(proj_rows)
                    
                    # Safe sheet name
                    safe_name = "".join(c for c in proj_name if c not in '[]:*?/\')([')
                    safe_name = safe_name[:30]
                    
                    # Write Sheet
                    proj_sheet_df.to_excel(writer, sheet_name=safe_name, index=False)
                    ws_proj = writer.sheets[safe_name]
                    ws_proj.right_to_left() # RTL
                    
                    # Columns config
                    ws_proj.set_column('A:A', 12, date_fmt)
                    ws_proj.set_column('B:C', 10, time_fmt)
                    ws_proj.set_column('D:D', 40)
                    ws_proj.set_column('E:E', 12, decimal_fmt)
                    
                    # --- ADD FORMULAS & GRAND TOTAL ---
                    # Columns: A(Date), B(Entry), C(Exit), D(Notes), E(Total)
                    # Formula for E: (C-B)*24
                    
                    num_rows = len(proj_sheet_df)
                    col_total_idx = 4 # E is index 4
                    
                    for r in range(num_rows):
                        xl_r = r + 1 # Data starts row 2 (index 1)
                        # Formula: (C{r+1}-B{r+1}) * 24
                        # If values are missing, might result in error or 0.
                        # Handle blank visually via format.
                        
                        formula = f"=(C{xl_r+1}-B{xl_r+1})*24"
                        ws_proj.write_formula(xl_r, col_total_idx, formula, decimal_fmt)
                        
                    # Add Grand Total Row
                    # Row after last data row
                    total_row = num_rows + 1 # Excel row index (1-based, after header + data)
                    
                    ws_proj.write(total_row, 0, "Total", workbook.add_format({'bold': True}))
                    
                    # Sum Formula: =SUM(E2:E{last})
                    sum_formula = f"=SUM(E2:E{total_row})"
                    ws_proj.write_formula(total_row, col_total_idx, sum_formula, decimal_fmt)


    
            print(f"SUCCESS: Report saved successfully to:\n{os.path.abspath(output_filename)}")
    
        except Exception as e:
            print(f"ERROR Saving Excel: {e}")
            print("Tip: Close the Excel file if it's currently open!")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # Halt only if interactive
        if len(sys.argv) <= 1:
            input("\nPress Enter to close window...")