import os
import pandas as pd
import numpy as np
import pdfplumber
from datetime import datetime, timedelta, timezone
import docx
from transliterate import translit

path = 'C:/Users/s.zakarin/Desktop/Telegram chat/Чаты'

chats = []

for file in os.listdir(path):
    chats.append(f'{path}/{file}')

CHAT = pd.DataFrame()

for chat in chats:
    data = pd.read_excel(chat)
    CHAT = pd.concat([CHAT, data])

KONTRAGENT = []
SUMMA = []
for i in CHAT['Заявка']:
    if type(i) == str:
        if i.find('1.') != -1 and not i[i.find('1.')+2].isdigit():
             zayavka = i[i.find('1.'):]
             zayavka = zayavka.split('\n')
             if len(zayavka) == 1:
                 zayavka1 = ' '.join(zayavka)
                 zayavka1 = zayavka1.split(';')
                 zayavka = zayavka1
                 if len(zayavka) > 2:
                     KONTRAGENT.append(translit(zayavka[0].replace('ООО', '').replace('MCHJ', '').replace('"', '').replace('«', '').replace('”',
                                                                                                                     '').replace(
                             '»', '').replace('»', '').replace('Контрагент-', '').replace('Контрагент -', '').replace('ИП',
                                                                                                                      '').replace(
                             'АО', '').replace(';', '').replace('»', '').replace('OOO', '').replace('OK', '').replace('ОК',
                                                                                                                      '').replace(
                             'ХК', '').replace('“', '').replace('XK', '').replace('POLIPLAST TRADE', 'POLIPLAST TRAID').replace('00О', '').replace('ооо', '').replace('НОУ', '').replace('ЯТТ', '').replace('Qmj', '').replace('ооо', '').replace('Контрагент', '').replace('ЧП', '').replace('AJ', '').replace('1.', ' ').replace('     ', ' ').replace(
                             '    ', ' ').replace('   ', ' ').replace('  ', ' ').strip().upper(), 'ru', reversed=True).upper())
                     SUMMA.append(zayavka[2])
             elif len(zayavka) > 2:
                     KONTRAGENT.append(translit(zayavka[0].replace('ООО', '').replace('MCHJ', '').replace('"', '').replace('«', '').replace('”',
                                                                                                                     '').replace(
                             '»', '').replace('»', '').replace('Контрагент-', '').replace('Контрагент -', '').replace('ИП',
                                                                                                                      '').replace(
                             'АО', '').replace(';', '').replace('»', '').replace('OOO', '').replace('OK', '').replace('ОК',
                                                                                                                      '').replace(
                             'ХК', '').replace('“', '').replace('XK', '').replace('POLIPLAST TRADE', 'POLIPLAST TRAID').replace('00О', '').replace('ооо', '').replace('НОУ', '').replace('ЯТТ', '').replace('Qmj', '').replace('ооо', '').replace('Контрагент', '').replace('ЧП', '').replace('AJ', '').replace('1.', ' ').replace('     ', ' ').replace(
                             '    ', ' ').replace('   ', ' ').replace('  ', ' ').strip().upper(), 'ru', reversed=True).upper())
                     SUMMA.append(zayavka[2])
             else:
                 KONTRAGENT.append('')
                 SUMMA.append('')
        elif i.find('1.') != -1 and i[i.find('1.')+2].isdigit():
             new_i = i[i.find('1.') + 2:]
             zayavka = new_i[new_i.find('1.'):]
             zayavka = zayavka.split('\n')
             if len(zayavka) == 1:
                 zayavka1 = ' '.join(zayavka)
                 zayavka1 = zayavka1.split(';')
                 zayavka = zayavka1
                 if len(zayavka) > 2:
                     KONTRAGENT.append(translit(zayavka[0].replace('ООО', '').replace('MCHJ', '').replace('"', '').replace('«', '').replace('”',
                                                                                                                     '').replace(
                             '»', '').replace('»', '').replace('Контрагент-', '').replace('Контрагент -', '').replace('ИП',
                                                                                                                      '').replace(
                             'АО', '').replace(';', '').replace('»', '').replace('OOO', '').replace('OK', '').replace('ОК',
                                                                                                                      '').replace(
                             'ХК', '').replace('“', '').replace('XK', '').replace('POLIPLAST TRADE', 'POLIPLAST TRAID').replace('00О', '').replace('ооо', '').replace('НОУ', '').replace('ЯТТ', '').replace('Qmj', '').replace('ооо', '').replace('Контрагент', '').replace('ЧП', '').replace('AJ', '').replace('1.', ' ').replace('     ', ' ').replace(
                             '    ', ' ').replace('   ', ' ').replace('  ', ' ').strip().upper(), 'ru', reversed=True).upper()
                                       )
                     SUMMA.append(zayavka[2])
                 else:
                    KONTRAGENT.append('')
                    SUMMA.append('')
             elif len(zayavka) > 2:
                 KONTRAGENT.append(translit(zayavka[0].replace('ООО', '').replace('MCHJ', '').replace('"', '').replace('«', '').replace('”',
                                                                                                                 '').replace(
                         '»', '').replace('»', '').replace('Контрагент-', '').replace('Контрагент -', '').replace('ИП',
                                                                                                                  '').replace(
                         'АО', '').replace(';', '').replace('»', '').replace('OOO', '').replace('OK', '').replace('ОК',
                                                                                                                  '').replace(
                         'ХК', '').replace('“', '').replace('XK', '').replace('POLIPLAST TRADE', 'POLIPLAST TRAID').replace('00О', '').replace('ооо', '').replace('НОУ', '').replace('ЯТТ', '').replace('Qmj', '').replace('ооо', '').replace('Контрагент', '').replace('ЧП', '').replace('AJ', '').replace('1.', ' ').replace('     ', ' ').replace(
                         '    ', ' ').replace('   ', ' ').replace('  ', ' ').strip().upper(), 'ru', reversed=True).upper()
                                   )
                 SUMMA.append(zayavka[2])
             else:
                 KONTRAGENT.append('')
                 SUMMA.append('')
        else:
            KONTRAGENT.append('')
            SUMMA.append('')
    else:
        KONTRAGENT.append('')
        SUMMA.append('')

for i, n in enumerate(KONTRAGENT):
    if n.upper().find('JANGI') != -1:
        KONTRAGENT[i] = 'YANGI TEXNOLOGIYALAR ILMIY-AXBOROT MARKAZI'
    elif n.upper().find('IZBA') != -1:
        KONTRAGENT[i] = 'IZBA ADVERTISING COMPANY'
    elif n.upper().find('KANTsLER') != -1:
        KONTRAGENT[i] = 'KANSLER'
    elif n.upper().find('_') != -1:
        index = n.upper().find('_')
        KONTRAGENT[i] = n[:index] + ' ' +  n[index+1:]
CHAT['Контрагент'] = KONTRAGENT
CHAT['Сумма'] = SUMMA
CHAT.drop_duplicates(inplace=True, ignore_index=True)
CHAT = CHAT[CHAT['Контрагент']!=''].reset_index()
CHAT.drop(columns=['Unnamed: 0', 'index'], inplace=True)


path1 = '//172.16.30.100/AsterShare/Aster UZ/Closing_documents'

documents = []

for root, dirs, files in os.walk(path1):
    if root.find('Корп') == -1:
        for file in files:
            if file.endswith('.pdf') or file.endswith('.jpeg'):
                documents.append(file)

error_documents = pd.DataFrame()
error_list = []

closing_documents = pd.DataFrame()
cl_data = []
cl_contr = []
cl_sum = []

for doc in documents:
    new_doc = doc.replace('.pdf','').replace('.jpeg','').split('_')
    if len(new_doc)>3:
        error_list.append(doc)
    elif len(new_doc)==3:
        cl_data.append(new_doc[0])
        cl_contr.append(new_doc[1])
        cl_sum.append(new_doc[2])

    else:
        error_list.append(doc)

closing_documents['Дата'] = cl_data
closing_documents['Контрагент'] = cl_contr
closing_documents['Сумма'] = cl_sum

error_documents['Документ'] = error_list

for i, n in enumerate(closing_documents['Контрагент']):
    closing_documents.loc[i, 'Контрагент'] = closing_documents.loc[i, 'Контрагент'].upper().replace('MCHJ', '').replace('ооо', '').replace('OK', '').replace('ЯТТ', '').replace('Qmj', '').replace('ОК', '').replace('MCHJ', '').replace('OOO', '').replace('ооо', '').replace('ООО', '').replace('OOO', '').replace('XK', '').replace('ХК', '').replace('ЧП', '').replace("'", "").replace('"', '').replace("QK", "").replace("ГУП", "").replace("QMJ", "").replace("AJ", "").replace("ИП", "").replace("     ", " ").replace("    ", " ").replace("   ", " ").replace("  ", " ").strip()

for i, n in enumerate(closing_documents['Дата']):
    closing_documents.loc[i, 'Дата'] = closing_documents.loc[i, 'Дата'].replace(',', '.').replace('29052023', '29.05.2023').replace(' ', '.')[:10]

for i, n in enumerate(closing_documents['Сумма']):
    closing_documents.loc[i, 'Сумма'] = closing_documents.loc[i, 'Сумма'].replace('.', ',').replace('     ', '').replace('    ', '').replace('   ', '').replace('  ', '').replace(' ', '')
    closing_documents.loc[i, 'Сумма'] = pd.to_numeric(closing_documents.loc[i, 'Сумма'], errors='coerce')

closing_documents = closing_documents[closing_documents['Сумма'].notnull()]
closing_documents.reset_index(inplace=True)
closing_documents.drop(columns=['index'], inplace=True)

today = datetime.now()
this_month = datetime.strftime(today, '%m')
previous_month = datetime.strftime((today - timedelta(days=70)), '%m')
this_year = datetime.strftime(today, '%Y')
previous_year = datetime.strftime((today-timedelta(days=70)), '%Y')
previous_month1 = datetime.strftime((today - timedelta(days=today.day)), '%m')
previous_year1 = datetime.strftime((today-timedelta(days=today.day)), '%Y')

main_folder = '//172.16.30.100/AsterShare/Aster/All/Выписки банка/2023'

pdf_files = []

for root, dirs, files in os.walk(main_folder):
    for file in files:
        if (file.endswith('.pdf') and file.find(f'.{this_month}.') != -1) or (file.endswith('.pdf') and file.find(f'.{previous_month}.') != -1) or (file.endswith('.pdf') and file.find(f'.{previous_month1}.') != -1):
            pdf_files.append(root+"/"+file)

pdf_files = [file.replace("\\", "/") for file in pdf_files]
kontragent = 'Номенклатура'
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
for i,n in enumerate(data['Наименование корреспондента']):
    data.loc[i, 'Наименование корреспондента'] = data.loc[i, 'Наименование корреспондента'].replace('Mchj', '').replace('ГУП Научно-информационный', '').replace('центр', '').replace('Частное Предприятие', '').replace('СП', '').replace('Фk', '').replace('filiali', '').replace('АЖ', '').replace('ЧП', '').replace('S CLASS', 'S-CLASS').replace('MChJ', '').replace('счет', '').replace('ссудный', '').replace('Основной', '').replace('JAMIYAT', '').replace("CHEKLANGAN", '').replace("MAS'ULIYATI", '').replace('Jamiyat', '').replace('Cheklangan', '').replace('Mas`Uliyati', '').replace('OOO', '').replace('ООО', '').replace('"', '').replace('МЧЖ', '').replace('ИП', '').replace('MCHJ', '').replace('OK', '').replace('ОК', '').replace('AK Samarkand filiali', '').replace('Aksiyadorlik', '').replace('Jamiyati', '').replace('Филиал', '').replace('ТШТТ', '').replace('QMJ', '').replace('АТФ', '').replace('cheklangan jamiyati', '').replace('mas`uliyati', '').replace('LIMITED LIABILITY CO', '').replace('филиал', '').replace('жамият', '').replace('маъсулияти чекланган', '').replace('ХУСУСИЙ КОРХ', '').replace('>', '').replace('<', '').replace('XK', '').replace('»', '').replace('«', '').replace('МЧЖ', '').replace('OK', '').replace('YATT', '').replace('YATT', '').replace('     ', ' ').replace('    ', ' ').replace('   ', ' ').replace('  ', ' ').strip()
for i, n in enumerate(data['Наименование корреспондента']):
    data.loc[i, 'Наименование корреспондента'] = translit(data.loc[i, 'Наименование корреспондента'], 'ru', reversed=True).upper()
data.drop(columns=['index', 'NN'], inplace=True)

data.to_excel(f'C:/Users/s.zakarin/Desktop/Telegram chat/Банковские выписки/Выписка_банка {previous_month}.{previous_year}-{this_month}.{this_year}.xlsx')

main_folder = '//172.16.30.100/AsterShare/Aster/All/Выписки банка/2023'

# docx_files = []
#
# for root, dirs, files in os.walk(main_folder):
#     for file in files:
#         if (file.endswith('.docx') and file.find(f'.{this_month}.') != -1) or (file.endswith('.docx') and file.find(f'.{previous_month}.') != -1) or (file.endswith('.docx') and file.find(f'.{previous_month1}.') != -1):
#             docx_files.append(root+"/"+file)
#
# docx_files = [file.replace("\\", "/") for file in docx_files]
#
# data_word = []
# for path in docx_files:
#     file = docx.Document(path)
#     table = file.tables[1]
#     for row in table.rows:
#         row_data = []
#         for cell in row.cells:
#             row_data.append(cell.text)
#         data_word.append(row_data)
#
# data_word = pd.DataFrame(data_word, columns=data_word[0])
#
# data_word = data_word[(data_word['Счет/ИНН']!='Счет/ИНН')&(data_word['Счет/ИНН']!='Итоговый оборот за период:')].reset_index()
# data_word.drop(columns=['index', '№ док-та', 'Оп'], inplace=True)

with pd.ExcelWriter('C://Users/s.zakarin/Desktop/Telegram chat/Finance UZ.xlsx', mode='w') as writer:
    CHAT.to_excel(writer, sheet_name='Заявки')
    error_documents.to_excel(writer, sheet_name='Docs_with_mistakes')
    closing_documents.to_excel(writer, sheet_name='Closing_docs')
    data.to_excel(writer, sheet_name=f'Выписка_банка {previous_month}.{previous_year}_{this_month}.{this_year}')
    # data_word.to_excel(writer, sheet_name= f'Выписка_банка_DOC {this_month}_{previous_month}')