from gc import DEBUG_LEAK

import pandas as pd
import os

from matplotlib.patheffects import withSimplePatchShadow
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.dimensions import RowDimension
from sympy import andre
from sympy.utilities.exceptions import ignore_warnings
from utils import get_column_index, format_header_cell, format_all_sheets
from watchdog.events import PatternMatchingEventHandler

date_str = '20250731'
current_date = pd.to_datetime(date_str, format='%Y%m%d')
# Calculate the last date as one month before the current date
last_date = current_date - pd.offsets.MonthEnd(1)
formatted_date = f"{current_date.month}/{current_date.day}/{str(current_date.year)[-2:]}"
formatted_last_date = f"{last_date.month}/{last_date.day}/{str(last_date.year)[-2:]}"
print(formatted_date,formatted_last_date)

file_path = f'S:/Audrey/Audrey/AAT.DCF/{date_str}/Status_Final_{date_str}.xlsx'
# aat_data_path = f'S:/Audrey/Audrey/AAT.DCF/{date_str}/AAT_{date_str}.xlsx'
aat_data_path = f'//evprodfsg01/QR_Workspace/AssetAllocation/CFValidation/Prod/AATOutput/{date_str}/AATOutput.{date_str}.xlsx'
aat_pm_owner_path = f's:/Audrey/Audrey/AAT.DCF/AAT PM Owner.xlsx'
output_folder = f'S:/Audrey/Audrey/AAT.DCF/{date_str}'
output_filename = f'AAT vs ECF {date_str}.xlsx'
output_path = os.path.join(output_folder, output_filename)
os.makedirs(output_folder, exist_ok=True)

# === Global Cell Styles ===
HEADER_FONT = Font(bold=True, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='00008B', end_color='00008B', fill_type='solid')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')

def load_data():
    df_aat = pd.read_excel(aat_data_path)
    df_b = pd.read_excel(file_path)
    return pd.merge(df_aat, df_b, on = 'Deal Name', how = 'left')


def process_data(df):

    df.drop(columns=['Instrument','Abs IRR Change'], inplace=True)
    df.drop_duplicates(subset='Deal Name', keep='first', inplace=True)

    df.insert(df.columns.get_loc(f'{formatted_date} IRR') + 1, 'AAT&ECF IRR Diffs', df[f'{formatted_date} IRR'] - df[f'{formatted_date} AAT IRR'])
    df.insert(df.columns.get_loc('Duration DCF Base¹') + 1, 'Duration Diffs', df['Duration DCF Base¹'] - df['Duration AAT Base'])

    pm_map = pd.read_excel(aat_pm_owner_path).set_index('Sr. Portfolio Manager')['AAT PM Owner']
    df.insert(df.columns.get_loc('Sr. Portfolio Manager') + 1, 'AAT PM Owner', df['Sr. Portfolio Manager'].map(pm_map))

    liq_cap_col_idx = df.columns.get_loc('Liq Cap')
    market_value_col_idx = df.columns.get_loc(f'{formatted_date} MV')
    df_liq_cap = df.pop('Liq Cap')
    df_market_value = df.pop(f'{formatted_date} MV')
    df.insert(df.columns.get_loc('Duration Diffs') + 1, 'Liq Cap', df_liq_cap)
    df.insert(df.columns.get_loc('Liq Cap') + 1, f'{formatted_date} MV', df_market_value)

    df.rename(columns={f'{formatted_date} IRR': f'{formatted_date} ECF IRR',
                       f'{formatted_last_date} IRR': f'{formatted_last_date} ECF IRR',
                       'IRR Change': 'MoM ECF IRR Movements',
                       'Duration AAT Base': 'Duration AAT', 'Duration DCF Base¹': 'Duration ECF',
                       'Comments': 'AAT Comments'}, inplace=True)

    return df

def reorder_columns(df):

    columns_order = [
        'Deal Name', 'Sr. Portfolio Manager', 'AAT PM Owner', f'{formatted_date} AAT IRR',
        f'{formatted_date} ECF IRR', 'AAT&ECF IRR Diffs', f'{formatted_last_date} ECF IRR','MoM ECF IRR Movements',
        'Duration AAT', 'Duration ECF', 'Duration Diffs','Liq Cap',
        f'{formatted_date} MV', 'MV %', 'AAT Comments'
    ]
    return df[columns_order]

def calculate_metrics(df):
    total_MV = df[f'{formatted_date} MV'].sum()
    df['MV %'] = df[f'{formatted_date} MV'] / total_MV * 100
    df['MV %'] = df['MV %'].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")
    df.sort_values(by=f'{formatted_date} MV', ascending=False, inplace=True)
    df['Cumulative MV %'] = df[f'{formatted_date} MV'].cumsum() / total_MV * 100
    df['Cumulative MV %'] = df['Cumulative MV %'].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")
    return df

def save_to_excel(df, output_path):
    with pd.ExcelWriter(output_path, engine = 'openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Summary')


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

def significant_changes_and_diffs(ws):
    # Define the fills for highlighting
    highlight_fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    highlight_fill_orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    highlight_fill_green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    # Highlight significant changes and collect rows
    significant_changes = highlight_and_collect(ws, 'MoM ECF IRR Movements', 0.05, highlight_fill_yellow)
    significant_diffs = highlight_and_collect(ws, 'AAT&ECF IRR Diffs', 0.05, highlight_fill_orange)
    highlight_durations = highlight_and_collect(ws, 'Duration Diffs', 0.5, highlight_fill_green)

    return significant_changes, significant_diffs, highlight_durations


def format_worksheet(ws):
    for col in ws.iter_cols():
        header = col[0].value
        if header and 'IRR' in header:
            for cell in col[1:]:
                cell.number_format = '0.00%'
        elif header in [f'{formatted_date} MV', 'Liq Cap']:
            for cell in col[1:]:
                cell.number_format = '#,##0_);(#,##0)'
        elif header and 'Duration' in header:
            for cell in col[1:]:
                cell.number_format = '0.00'

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value is not None)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = ALIGN_CENTER

    for cell in ws[1]:
        format_header_cell(cell)


def create_highlighted_sheets(wb, significant_changes, significant_diffs, highlight_durations):
    header = [cell.value for cell in wb['Summary'][1]]

    def create_sheet(name, rows):
        ws = wb.create_sheet(title=name)
        ws.append(header)
        for row in rows:
            ws.append(row)
        return ws

    ws1 = create_sheet('Significant ECF IRR Movers', significant_changes)
    ws2 = create_sheet('Significant AAT&ECF Diffs', significant_diffs)
    ws3 = create_sheet('Highlight Duration Diffs', highlight_durations)

    format_all_sheets(ws1, ws2, ws3)


def summary_category(wb):
    ws = wb['Summary']

    mv_pct_col_idx = get_column_index(ws, 'MV %')
    category_col_idx = mv_pct_col_idx + 1
    ws.insert_cols(category_col_idx)

    cell = ws.cell(row=1, column=category_col_idx, value='Category')
    format_header_cell(cell)

    irr_diff_col_idx = get_column_index(ws, 'AAT&ECF IRR Diffs')
    mv_col_idx = get_column_index(ws, f'{formatted_date} MV')

    for row in range(2, ws.max_row + 1):
        irr_diff = ws.cell(row=row, column=irr_diff_col_idx).value
        mv_value = ws.cell(row=row, column=mv_col_idx).value

        if irr_diff is not None:
            if mv_value is not None and mv_value > 25_000_000:
                value = 'Significant Discrepancy' if abs(irr_diff) > 0.05 else 'Alignment'
            else:
                value = 'Significant discrepancy but ignore' if abs(irr_diff) > 0.05 else 'Alignment'
            ws.cell(row=row, column=category_col_idx, value=value)


def drop_cumulative_mv_column(wb):
    # Drop "Cumulative MV%" column from all sheets except "Summary"
    for ws in wb.worksheets:
        if ws.title != 'Summary':
            for col in ws.iter_cols(1, ws.max_column):
                if col[0].value == 'Cumulative MV %':
                    ws.delete_cols(col[0].column)
                    break



def highlight_summary(ws):
    # Define the fill for highlighting
    highlight_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
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
    return ws



def main():
    df = load_data()
    df = process_data(df)
    df = calculate_metrics(df)
    df = reorder_columns(df)
    save_to_excel(df, output_path)

    wb = load_workbook(output_path)
    ws = wb['Summary']

    format_worksheet(ws)

    significant_changes, significant_diffs, highlight_durations = significant_changes_and_diffs(ws)
    create_highlighted_sheets(wb, significant_changes, significant_diffs, highlight_durations)

    summary_category(wb)
    highlight_summary(ws)
    drop_cumulative_mv_column(wb)

    wb.save(output_path)
    wb.close()
    print("✅ Data processing and formatting completed successfully.")

if __name__ == "__main__":
    main()
    print("Data processing and formatting completed successfully.")

