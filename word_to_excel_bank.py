import docx
from docx import Document
import pandas as pd
import os
from datetime import datetime, timedelta, timezone

today = datetime.now()
this_month = datetime.strftime(today, '%m')
previous_month = datetime.strftime((today - timedelta(days=70)), '%m')
this_year = datetime.strftime(today, '%Y')
previous_year = datetime.strftime((today-timedelta(days=70)), '%Y')
previous_month1 = datetime.strftime((today - timedelta(days=today.day)), '%m')
previous_year1 = datetime.strftime((today-timedelta(days=today.day)), '%Y')

main_folder = '//172.16.30.100/AsterShare/Aster/All/Выписки банка/2023'

docx_files = []

for root, dirs, files in os.walk(main_folder):
    for file in files:
        if (file.endswith('.docx') and file.find(f'.{this_month}.') != -1 and file.find('ААБ') != -1) or (file.endswith('.docx') and file.find(f'.{previous_month}.') != -1 and file.find('ААБ') != -1) or (file.endswith('.docx') and file.find(f'.{previous_month1}.') != -1 and file.find('ААБ') != -1):
            docx_files.append(root+"/"+file)

docx_files = [file.replace("\\", "/") for file in docx_files]

data_word = []
for path in docx_files:
    file = docx.Document(path)
    table = file.tables[1]
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            row_data.append(cell.text)
        if row_data.count('\xa0') > 1:
            row_data.remove('\xa0')
        if row_data.count('Оборот Дебет') > 1:
            row_data.remove('Оборот Дебет')
        if row_data.count('Итоговый оборот за период:') > 1:
            row_data.remove('Итоговый оборот за период:')
        data_word.append(row_data)

data_word = pd.DataFrame(data_word, columns=data_word[0])

data_word = data_word[(data_word['Счет/ИНН']!='Счет/ИНН')&(data_word['Счет/ИНН']!='Итоговый оборот за период:')].reset_index()

data_word.to_excel(f'C://Users/s.zakarin/Desktop/Telegram chat/Банковские выписки/Выписка_банка_DOC_{previous_month}.{previous_year}_{this_month}.{this_year}.xlsx')