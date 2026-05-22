import openpyxl
wb = openpyxl.load_workbook('Механики.xlsx', data_only=True)
ws = wb.active
for i, row in enumerate(ws.iter_rows(max_row=20)):
    vals = [(cell.column_letter, repr(cell.value)[:50]) for cell in row if cell.value is not None]
    if vals:
        print(f'Row {i+1}:', vals)
