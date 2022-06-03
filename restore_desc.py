
import pickle
from openpyxl import load_workbook, Workbook


fname_store_data = f'01-05-2022_DB.xlsx'
fname_targ_result = f'./restore/통합작업본_01122022_작업완료원본.xlsx'

f = open('./bin_store_data/' + fname_store_data + '.pkl', 'rb')
dict_vendor = pickle.load(f)
f.close()

wb = load_workbook(fname_targ_result)
ws = wb['Sheet1']
idx_row = 0
for row in ws:
    if idx_row > 0:
        upc = row[4].value
        str_cd = row[10].value
        vend_cd = row[11].value
        item_cd = row[13].value

        print()
    idx_row += 1
