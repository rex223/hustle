import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill

file_path = 'RENTAL DATASHEET.xlsx'

def format_excel():
    # 1. Read existing data
    df = pd.read_excel(file_path, usecols="A:G")
    df = df.dropna(how='all')
    
    # 2. Remove any existing "TOTAL" rows from previous runs so we don't double-count
    # We'll identify total rows by checking if 'Tenant Name' or 'Group' starts with 'TOTAL'
    df['Tenant Name'] = df['Tenant Name'].fillna('')
    df = df[~df['Tenant Name'].astype(str).str.startswith('TOTAL ')]
    
    # 3. Sort by Group
    df['Group_str'] = df['Group'].astype(str)
    df = df.sort_values(by='Group_str').drop(columns=['Group_str'])
    
    # 4. Calculate Subtotals and insert them
    final_rows = []
    current_group = None
    group_sum = 0
    
    for _, row in df.iterrows():
        grp = str(row['Group'])
        
        # If we switch to a new group (and it's not the first row)
        if current_group is not None and grp != current_group:
            # Append a total row for the previous group
            total_row = {
                'Tenant Name': f'TOTAL {current_group}',
                'Property / Location': '',
                'Rent Amount': group_sum,
                'Rented Unit / Place': '',
                'Group': '',
                'Duration (Months)': '',
                'Subscription End Date': ''
            }
            final_rows.append(total_row)
            # Reset sum
            group_sum = 0
            
        final_rows.append(row.to_dict())
        current_group = grp
        
        # Add to sum
        try:
            amt = float(row['Rent Amount'])
            if not pd.isna(amt):
                group_sum += amt
        except:
            pass

    # Don't forget the total for the very last group
    if current_group is not None:
        total_row = {
            'Tenant Name': f'TOTAL {current_group}',
            'Property / Location': '',
            'Rent Amount': group_sum,
            'Rented Unit / Place': '',
            'Group': '',
            'Duration (Months)': '',
            'Subscription End Date': ''
        }
        final_rows.append(total_row)
        
    final_df = pd.DataFrame(final_rows)
    
    # 5. Save back to Excel
    final_df.to_excel(file_path, index=False)
    
    # 6. Apply Formatting (Bold and color for totals)
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    bold_font = Font(bold=True)
    highlight_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    
    # Format Headers
    for col_idx in range(1, 8):
        cell = ws.cell(row=1, column=col_idx)
        if cell.value:
            cell.font = bold_font
            
    # Format Total Rows
    for row_idx in range(2, ws.max_row + 1):
        tenant_val = str(ws.cell(row=row_idx, column=1).value or '')
        if tenant_val.startswith('TOTAL '):
            for col_idx in range(1, 8):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = bold_font
                cell.fill = highlight_fill

    # Adjust column widths
    column_widths = {'A': 25, 'B': 15, 'C': 15, 'D': 30, 'E': 15, 'F': 20, 'G': 25}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    wb.save(file_path)
    print("Excel file successfully auto-formatted with subtotals!")

if __name__ == '__main__':
    format_excel()
