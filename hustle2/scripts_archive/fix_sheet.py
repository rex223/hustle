import pandas as pd
import openpyxl
from openpyxl.styles import Font

file_path = 'RENTAL DATASHEET.xlsx'

# 1. Read existing data and drop corrupted summary columns
# We only want columns A through G (0 to 6)
df = pd.read_excel(file_path, usecols="A:G")

# Drop completely empty rows
df = df.dropna(how='all')

# 2. Sort the dataframe by the 'Group' column so tenants are stored together
# We ensure string conversion to avoid sorting errors if there are mixed types
df['Group_str'] = df['Group'].astype(str)
df = df.sort_values(by='Group_str').drop(columns=['Group_str']).reset_index(drop=True)

# 3. Save the clean sorted data back to a new file (we will overwrite original)
df.to_excel(file_path, index=False)

# 4. Open with openpyxl to apply formatting and safe formulas
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# Apply bold headers
bold_font = Font(bold=True)
for col_idx in range(1, 8):
    cell = ws.cell(row=1, column=col_idx)
    if cell.value:
        cell.font = bold_font

# 5. Build a safe summary table without dynamic arrays (to prevent corruption)
ws.cell(row=1, column=10, value="Group Summary").font = bold_font
ws.cell(row=1, column=11, value="Total Amount").font = bold_font

# Get unique groups from the dataframe
unique_groups = df['Group'].dropna().unique().tolist()
# Filter out any non-string or weird nan
unique_groups = [g for g in unique_groups if str(g).strip().upper() != 'NAN']
unique_groups.sort()

current_row = 2
# Add the existing unique groups
for group in unique_groups:
    ws.cell(row=current_row, column=10, value=group)
    ws.cell(row=current_row, column=11, value=f'=SUMIF(E:E, J{current_row}, C:C)')
    current_row += 1

# Add some extra blank rows for future groups, with safe formulas
for _ in range(10): # Room for 10 new groups
    ws.cell(row=current_row, column=11, value=f'=IF(J{current_row}<>"", SUMIF(E:E, J{current_row}, C:C), "")')
    current_row += 1

# 6. Adjust column widths for better readability
column_widths = {
    'A': 20,
    'B': 15,
    'C': 15,
    'D': 30,
    'E': 15,
    'F': 20,
    'G': 25,
    'J': 20,
    'K': 20
}
for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

wb.save(file_path)
print("File recovered, sorted by group, and safely summarized.")
