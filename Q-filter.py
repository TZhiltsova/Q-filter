import openpyxl

book = openpyxl.load_workbook('scimagojr 2021.xlsx')
sheet = book['scimagojr 2021']
title_column = 'B'
Q_column = 'F'
Q_factor = 'Q1'
title_Q = {}
for i in range(1, sheet.max_row+1):
    title_Q[sheet[title_column + str(i)].value] = sheet[Q_column + str(i)].value

Q_required = []
for key, val in title_Q.items():
    if val == Q_factor:
        Q_required.append(key)

with open('Q_factor_list.txt', 'w') as q_list:
    for elem in Q_required:
        if Q_required.index(elem) == len(Q_required) - 1:
            q_list.write('SRCID (' + elem + ')')
        else:
            q_list.write('SRCID (' + elem + ')' + ' OR ')
