import openpyxl
import pandas as pd
import numpy as np

# Inspect the 3 new uploaded Excel files
f_master = 'Master_Financial_Data_F_2.xlsx'
f_cf = 'Cash Flow 12 Month.xlsx'
f_pl = 'P&L_Rent_Projects_F_2.xlsx'

wb_m = openpyxl.load_workbook(f_master, data_only=True)
print("Master Sheets:", wb_m.sheetnames)
ws_m = wb_m['Loans_&_Installments']
print("Master Named Tables:", list(ws_m.tables.keys()))

def parse_named_table(ws, table_name):
    tbl = ws.tables[table_name]
    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(tbl.ref)
    data = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
        data.append(list(row))
    headers = data[0]
    rows = data[1:]
    return pd.DataFrame(rows, columns=headers)

df_loans = parse_named_table(ws_m, 'القروض')
df_installments = parse_named_table(ws_m, 'الاقساط')
df_revenues = parse_named_table(ws_m, 'الايردات')
df_dev_projects = parse_named_table(ws_m, 'Units_Under_Construction')
df_banks = parse_named_table(ws_m, 'البنوك')
df_collections = parse_named_table(ws_m, 'تحصيلات_الايجار')

print("\n--- Loans Columns ---", df_loans.columns.tolist())
print("--- Dev Projects Columns ---", df_dev_projects.columns.tolist())

df_cf_raw = pd.read_excel(f_cf, sheet_name='Sheet1')
time_cols = [c for c in df_cf_raw.columns if c not in ['Unnamed: 0', 'Unnamed: 1']]
print("\nCF Time Columns Count:", len(time_cols), time_cols)

df_pl_raw = pd.read_excel(f_pl, sheet_name='Sheet1')
print("\nP&L Shape:", df_pl_raw.shape)
print("P&L Head:\n", df_pl_raw.head(2))
