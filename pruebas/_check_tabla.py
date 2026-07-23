import openpyxl
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r'C:\Users\Near\Documents\Chamba Panama\excels\Iphone (1).xlsx', read_only=True, data_only=True)
ws = wb['Tabla']

print('=== Tabla ===')
count = 0
for row in ws.iter_rows(min_row=1, max_row=50, values_only=True):
    count += 1
    vals = []
    for i, v in enumerate(row):
        if v is not None:
            vals.append(f'C{i+1}={v}')
    if vals:
        print(f'R{count}: {vals}')
wb.close()
