
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.dimensions import RowDimension


date_str = '20250331'
current_date = pd.to_datetime(date_str, format='%Y%m%d')
formatted_date = f"{current_date.month}/{current_date.day}/{str(current_date.year)[-2:]}"
print(formatted_date)

file_path = f'S:/Audrey/Audrey/AAT.DCF/{date_str}/Status_Final_{date_str}.xlsx'
aat_pm_owner_path = f's:/Audrey/Audrey/AAT.DCF/AAT PM Owner.xlsx'
output_folder = f'S:/Audrey/Audrey/AAT.DCF/{date_str}'
output_filename = f'AAT vs ECF {date_str}.xlsx'
output_path = os.path.join(output_folder, output_filename)
os.makedirs(output_folder, exist_ok=True)

df = pd.read_excel(file_path)

df[f'{formatted_date} MV'] = df[f'{formatted_date} MV'].copy()

df.drop(columns=['Instrument','Abs IRR Change'], inplace=True)
df.drop_duplicates(subset='Deal Name', keep='first', inplace=True)

df.insert(df.columns.get_loc('IRR DCF Base') + 1, 'AAT&DCF IRR Diffs', df[f'{formatted_date} IRR'] - df['IRR DCF Base'])
df.insert(df.columns.get_loc('Duration DCF Base¹') + 1, 'Duration Diffs', df['Duration Base¹'] - df['Duration DCF Base¹'])
df.insert(df.columns.get_loc('Deal Owner') + 1, 'AAT PM Owner', df['Deal Owner'].map(pd.read_excel(aat_pm_owner_path, index_col='Deal Owner')['PM Owner']))

liq_cap_col_idx = df.columns.get_loc('Liq Cap')
market_value_col_idx = df.columns.get_loc(f'{formatted_date} MV')
df_liq_cap = df.pop('Liq Cap')
df_market_value = df.pop(f'{formatted_date} MV')
df.insert(df.columns.get_loc('Duration Diffs') + 1, 'Liq Cap', df_liq_cap)
df.insert(df.columns.get_loc('Liq Cap') + 1, f'{formatted_date} MV', df_market_value)


total_MV = df[f'{formatted_date} MV'].sum()
df['MV %'] = df[f'{formatted_date} MV'] / total_MV * 100
df['MV %'] = df['MV %'].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")
df.sort_values(by=f'{formatted_date} MV', ascending=False, inplace=True)
df['Cumulative MV %'] = df[f'{formatted_date} MV'].cumsum() / total_MV * 100
df['Cumulative MV %'] = df['Cumulative MV %'].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")


df.rename(columns={'IRR Change': 'MoM IRR Movements', 'Duration Base¹': 'Duration AAT Base', 'Duration DCF BaseDuration DCF Base¹': 'Duration DCF Base¹'}, inplace=True)
df.to_excel(output_path, index=False)


wb = load_workbook(output_path)
ws = wb.active

# Define the fills for highlighting
highlight_fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
highlight_fill_orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
highlight_fill_green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")


def highlight_and_collect(ws, column_name, threshold, fill, market_value_threshold=25000000):
    significant_rows = []
    market_value_col_idx = None

    # Find the column index for 'Market Value'
    for col in ws.iter_cols(1, ws.max_column):
        if col[0].value == f'{formatted_date} MV':
            market_value_col_idx = col[0].column

    if market_value_col_idx is None:
        raise KeyError(f"'{formatted_date} MV' column not found")

    for col in ws.iter_cols():
        if col[0].value == column_name:
            for cell in col[1:]:
                if cell.value is not None and isinstance(cell.value, (int, float)) and abs(cell.value) > threshold:
                    market_value = ws.cell(row=cell.row, column=market_value_col_idx).value
                    if market_value is not None and market_value >= market_value_threshold:
                        cell.fill = fill
                        row_values = [c.value for c in ws[cell.row]]
                        significant_rows.append(row_values)
    return significant_rows

significant_changes = highlight_and_collect(ws, 'MoM IRR Movements', 0.05, highlight_fill_yellow)
significant_diffs = highlight_and_collect(ws, 'AAT&DCF IRR Diffs', 0.05, highlight_fill_orange)
highlight_durations = highlight_and_collect(ws, 'Duration Diffs', 0.5, highlight_fill_green)


wb.save(output_path)

# Convert significant changes and diffs to DataFrames and save to new sheets
header = [cell.value for cell in ws[1]]
df_significant_changes = pd.DataFrame(significant_changes, columns=header)
df_significant_diffs = pd.DataFrame(significant_diffs, columns=header)
df_highlight_durations = pd.DataFrame(highlight_durations, columns=header)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_significant_changes.to_excel(writer, sheet_name='Significant AAT IRR Movers', index=False)
    df_significant_diffs.to_excel(writer, sheet_name='Significant AAT&DCF Diffs', index=False)
    df_highlight_durations.to_excel(writer, sheet_name='Highlight Duration Diffs', index=False)

# Load the sheets into DataFrames
df_significant_changes = pd.read_excel(output_path, sheet_name='Significant AAT IRR Movers')
df_significant_diffs = pd.read_excel(output_path, sheet_name='Significant AAT&DCF Diffs')


wb = load_workbook(output_path)

# Define the formats
header_font = Font(bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='00008B', end_color='00008B', fill_type='solid')
alignment_center = Alignment(horizontal='center', vertical='center')



ws = wb['Sheet1']
ws.title = 'Summary'


ws.insert_cols(ws.max_column + 1)
ws.cell(row=1, column=ws.max_column, value='Category')


irr_diff_col_idx = None
for col in ws. iter_cols(1, ws.max_column):
    if col[0].value == 'AAT&DCF IRR Diffs':
        irr_diff_col_idx = col[0].column
        break


if irr_diff_col_idx is not None:
    mv_col_idx = None
    for col in ws.iter_cols(1, ws.max_column):
        if col[0].value == f'{formatted_date} MV':
            mv_col_idx = col[0].column
            break

    if mv_col_idx is None:
        raise KeyError(f"'{formatted_date} MV' column not found")

    for row in range(2, ws.max_row + 1):
        irr_diff = ws.cell(row=row, column=irr_diff_col_idx).value
        mv_value = ws.cell(row=row, column=mv_col_idx).value
        if mv_value > 25000000:
            ws.cell(row=row, column=ws.max_column, value='Significant Discrepancy' if abs(irr_diff) > 0.05 else 'Alignment')
        else:
            ws.cell(row=row, column=ws.max_column, value='Significant discrepancy but ignore' if abs(irr_diff) > 0.05 else 'Alignment')

# Drop "Cumulative MV%" column from all sheets except "Summary"
for ws in wb.worksheets:
    if ws.title != 'Summary':
        for col in ws.iter_cols(1, ws.max_column):
            if col[0].value == 'Cumulative MV %':
                ws.delete_cols(col[0].column)
                break


# Reorder the sheets by putting "Common Significant" sheet in the 2nd position
sheet_names = wb.sheetnames
wb._sheets = [wb[name] for name in sheet_names]

def format_worksheet(ws):
    for col in ws.iter_cols():
        if 'IRR' in col[0].value:
            for cell in col[1:]:
                cell.number_format = '0.00%'
        if col[0].value in [f'{formatted_date} MV', 'Liq Cap']:
            for cell in col[1:]:
                cell.number_format = '#,##0_);(#,##0)'
        if 'Duration' in col[0].value:
            for cell in col[1:]:
                cell.number_format = '0.00'
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value is not None)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = alignment_center
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

for ws in wb.worksheets:
    format_worksheet(ws)


#formatting to all worksheets
wb.save(output_path)
wb.close()

print("Data processing and formatting completed successfully.")


wb = load_workbook(output_path)
ws = wb['Summary']

highlight_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# Find the column index for ‘Deal Name' and '{formatted_date} MV'
deal_name_col_idx = None
mv_col_idx = None
for col in ws.iter_cols(1, ws.max_column):
    if col[0].value == 'Deal Name':
        deal_name_col_idx = col[0].column
    if col[0].value == f'{formatted_date} MV':
        mv_col_idx = col[0].column
    if deal_name_col_idx and mv_col_idx:
        break

if deal_name_col_idx is None or mv_col_idx is None:
    raise KeyError(f"'Deal Name' or '{formatted_date} MV' column not found")

# Iterate over rows and apply fill conditionally based on the MV value
for row in range(2, ws.max_row + 1):
    mv_value = ws.cell(row=row, column=mv_col_idx).value
    deal_name_cell = ws.cell(row=row, column=deal_name_col_idx)

    # Apply highlight if MV > 25,000,000, otherwise group and hide the row
    if mv_value is not None and mv_value > 25000000:
        deal_name_cell.fill = highlight_fill
    else:
        ws.row_dimensions[row].outlineLevel = 1
        ws.row_dimensions[row].hidden = True

wb.save(output_path)
wb.close()