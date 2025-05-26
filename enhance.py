import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Load the Excel file
file_path = 'Input.xlsx'
output54_df = pd.read_excel(file_path, sheet_name='output54', engine='openpyxl')
output97_df = pd.read_excel(file_path, sheet_name='output97', engine='openpyxl')

# Clean column headers
output54_df.columns = output54_df.columns.str.strip()
output97_df.columns = output97_df.columns.str.strip()

# Create a list to store rows for the comparison report
comparison_rows = []

# Track source and row number
for idx54, row54 in output54_df.iterrows():
    key = row54['KEY']
    matching_rows = output97_df[output97_df['KEY'] == key]
    if not matching_rows.empty:
        row97 = matching_rows.iloc[0]
        idx97 = matching_rows.index[0]

        row54_with_meta = row54.copy()
        row54_with_meta['Source'] = 'output54'
        row54_with_meta['RowNumber'] = idx54 + 2
        row97_with_meta = row97.copy()
        row97_with_meta['Source'] = 'output97'
        row97_with_meta['RowNumber'] = idx97 + 2

        comparison_rows.append(row54_with_meta)
        comparison_rows.append(row97_with_meta)

# Create DataFrame
comparison_df = pd.DataFrame(comparison_rows)
cols = ['Source', 'RowNumber'] + [col for col in comparison_df.columns if col not in ['Source', 'RowNumber']]
comparison_df = comparison_df[cols]

# Save to Excel
output_file = 'comparison_report_with_highlights.xlsx'
comparison_df.to_excel(output_file, index=False)

# Load workbook for formatting
wb = load_workbook(output_file)
ws = wb.active

# Apply highlights for differences
for i in range(2, ws.max_row, 2):
    for j in range(3, ws.max_column + 1):  # Skip Source and RowNumber
        cell1 = ws.cell(row=i, column=j)
        cell2 = ws.cell(row=i+1, column=j)
        if cell1.value != cell2.value:
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell1.fill = fill
            cell2.fill = fill

# Save the highlighted file
wb.save(output_file)
print(f"Comparison report with highlights saved as '{output_file}'.")
