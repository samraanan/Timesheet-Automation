import pandas as pd

# Load Config
sheets = {s: pd.read_excel('Config.xlsx', sheet_name=s) for s in pd.ExcelFile('Config.xlsx').sheet_names}
df = sheets['Distances']

# Add missing pair
col_p1, col_p2, col_dist = df.columns[0], df.columns[1], df.columns[2]

new_rows = [
    {col_p1: 'בית (כפר אלדד)', col_p2: 'מחכה לפליקס', col_dist: 0},
    {col_p1: 'מחכה לפליקס', col_p2: 'בית (כפר אלדד)', col_dist: 0},
]

df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
sheets['Distances'] = df

# Write back
with pd.ExcelWriter('Config.xlsx', engine='xlsxwriter') as writer:
    for name, data in sheets.items():
        data.to_excel(writer, sheet_name=name, index=False)

print(f"Added 2 pairs. Total distances: {len(df)}")
