import pandas as pd
import openpyxl as xl

df = pd.read_excel('src\OddEvenCells.xlsx')

df['cells'] = df['cells'].str.replace(r'[^0-9,]+','')

for i, data in df.iterrows():
    list = data.cells.split(",")
    for num in list:
        # print(num)
        # print("Type:",type(num), "num:", num)
        if num.strip() != "":
            if int(num) % 2 == 0:
                data.even = int(data.even) + 1
                df.loc[i, 'even'] = data.even
            else:
                data.odd = int(data.odd) + 1
                df.loc[i, 'odd'] = data.odd
                


df.to_excel('src\cellnums.xlsx', index=False)
