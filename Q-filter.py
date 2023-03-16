import openpyxl

book = openpyxl.load_workbook('scimagojr 2021.xlsx')
sheet = book['scimagojr 2021']
title_colomn = 'B'
Q_colomn = 'F'
title_Q = {}
for i in range(1, sheet.max_row+1):
    title_Q[sheet[title_colomn + str(i)].value] = sheet[Q_colomn + str(i)].value

for key, val in title_Q.items():
    print(key, val)
