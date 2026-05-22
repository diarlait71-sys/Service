import pandas as pd
import os

files = ['Факт Услуги.xlsx', 'Факт Запасные части.xlsx']
for fname in files:
    fp = os.path.join(os.getcwd(), fname)
    print('FILE', fname)
    x = pd.read_excel(fp, sheet_name='Sheet1', header=None)
    hdr = None
    for i, row in x.iterrows():
        if 'Название ДЦ' in str(list(row.values)):
            hdr = i
            print('HEADER row', i)
            print(list(row.values))
            break
    print('shape', x.shape)
    print(x.head(12).to_string())
    print('---')
