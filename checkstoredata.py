from openpyxl import load_workbook


wb = load_workbook('./processed_store_data/10_10_2023_DB.xlsx')
print('loaded')

ws = wb.active

last_store = ''
for row in ws:
    if last_store != row[0]:
        last_store = row[0]
        print(f'[{last_store}]')
    if row[0].value == '015':
        for cell in row:
            print(cell.value, end=' ')
        print('\n')
