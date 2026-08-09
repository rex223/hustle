import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

file_path = 'RENTAL DATASHEET.xlsx'
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# 1. Rename columns based on context
headers = [
    "Tenant Name", 
    "Property / Location", 
    "Rent Amount", 
    "Rented Unit / Place", 
    "Group", 
    "Duration (Months)", 
    "Subscription End Date"
]

for col_idx, header in enumerate(headers, start=1):
    ws.cell(row=1, column=col_idx, value=header)

# Clear the old "GRP1 TOTAL GRP2 TOTAL AND SO ON" header
ws.cell(row=1, column=8, value="")

# 2. Set up the dynamic group summary
ws.cell(row=1, column=10, value="Group Summary")
ws.cell(row=1, column=11, value="Total Amount")

# Use Office 365 dynamic arrays to automatically pull unique groups and sum them
# If a new group is added, it automatically expands!
ws.cell(row=2, column=10, value='=UNIQUE(FILTER(E2:E1000, E2:E1000<>""))')
ws.cell(row=2, column=11, value='=SUMIF(E:E, J2#, C:C)')

# 3. Apply formatting
bold_font = Font(bold=True)
for col_idx in range(1, 12):
    cell = ws.cell(row=1, column=col_idx)
    if cell.value:
        cell.font = bold_font

# Adjust column widths for better readability
column_widths = {
    'A': 20,
    'B': 15,
    'C': 15,
    'D': 30,
    'E': 10,
    'F': 20,
    'G': 25,
    'J': 20,
    'K': 20
}
for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

# Save the workbook back
wb.save(file_path)
print("Excel file successfully updated.")
