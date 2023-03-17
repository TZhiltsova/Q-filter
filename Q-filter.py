import openpyxl


def q_filter(path, sheet_name, title_column, q_column, q_factor, path_output):
    '''
        :param path: takes the pass to input exel file
        :param sheet_name: takes the name of sheet for analysis
        :param title_column: letter of  column with title of journals
        :param q_column: letter of column with Q factor
        :param q_factor: required Q factor
        :param path_output: pass for saving output file
        :return: file with list of journals with corresponding Q factor, in format for SQL searching
    '''
    book = openpyxl.load_workbook(path)
    sheet = book[sheet_name]
    title_q = {}
    for i in range(1, sheet.max_row+1):
        title_q[sheet[title_column + str(i)].value] = sheet[q_column + str(i)].value

    q_required = []
    for key, val in title_q.items():
        if val == q_factor:
            q_required.append(key)

    with open(path_output + 'Q_factor_list.txt', 'w') as q_list:
        for elem in q_required:
            if q_required.index(elem) == len(q_required) - 1:
                q_list.write('SRCID (' + elem + ')')
            else:
                q_list.write('SRCID (' + elem + ')' + ' OR ')
    return


path_to_file = input('Print path to file: ')
sheet_name_in_file = input('Print sheet name: ')
title_column_in_file = input('Print column with journal titles (A, B, C itc.): ')
Q_column_in_file = input('Print column with Q factor (A, B, C itc.): ')
Q_factor_in_file = input('Print required Q factor (Q1, Q2 itc.): ')
path_output_to_file = input('Print path to directory for output file: ')
q_filter(path_to_file, sheet_name_in_file, title_column_in_file, Q_column_in_file, Q_factor_in_file,
         path_output_to_file)
