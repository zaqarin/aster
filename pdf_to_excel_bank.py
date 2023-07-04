import pdfplumber
import pandas as pd
import os
from datetime import datetime, timedelta, timezone

today = datetime.now()
this_month = datetime.strftime(today, '%m')
previous_month1 = datetime.strftime((today - timedelta(days=61)), '%m')
previous_month = datetime.strftime((today - timedelta(days=today.day)), '%m')
this_year = datetime.strftime(today, '%Y')
previous_year1 = datetime.strftime((today-timedelta(days=61)), '%Y')
previous_year = datetime.strftime((today-timedelta(today.day)), '%Y')


main_folder = 'path with pdf files'

pdf_files = []

for root, dirs, files in os.walk(main_folder):
    for file in files:
        if (file.endswith('.pdf') and file.find(f'.{this_month}.') != -1 and file.lower().find('кпк')==-1) or \
                (file.endswith('.pdf') and file.find(f'.{previous_month}.') != -1 and file.lower().find('кпк')==-1) or \
                (file.endswith('.pdf') and file.find(f'.{previous_month1}.') != -1 and file.lower().find('кпк')==-1):
            pdf_files.append(root+"/"+file)

pdf_files = [file.replace("\\", "/") for file in pdf_files]

tables = []

for file in pdf_files:
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()

            tables.extend(extracted_tables)

        data = pd.DataFrame()
        for table in tables:
            data = data._append(pd.DataFrame(table), ignore_index=True)

cols = {0:'NN', 1:'Дата', 2:'Док', 3:'МФО',
        4:'Счет корреспондента', 5:'Наименование корреспондента', 6:'Дебет', 7:'Кредит', 8:'Детали платежа'}

data = data.rename(columns=cols)
data['Счет корреспондента'] = data['Счет корреспондента'].apply(lambda x: x if type(x) == str else '')
data = data[((data['Счет корреспондента']!='')&(data['Счет корреспондента']!='Счет корреспондента'))].reset_index()

data.to_excel(f'path to save {previous_month1}.{previous_year1}-{this_month}.{this_year}.xlsx')
