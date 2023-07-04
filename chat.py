import pandas as pd
import os

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
NAZNACHENIE = []
LOCATION = []
USLOVIE = []
POSTAVKA = []
for i in CHAT['Заявка']:
    if type(i) == str:
        if i.find('1.') != -1 and not i[i.find('1.') + 2].isdigit():
            zayavka = i[i.find('1.'):]
            zayavka = zayavka.split(';')
            if len(zayavka) == 1:
                zayavka1 = ' '.join(zayavka)
                zayavka1 = zayavka1.split('\n')
                zayavka = zayavka1
                if len(zayavka) > 5:
                    KONTRAGENT.append(zayavka[0])
                    SUMMA.append(zayavka[2])
                    NAZNACHENIE.append(zayavka[1])
                    LOCATION.append(zayavka[3])
                    USLOVIE.append(zayavka[4])
                    POSTAVKA.append(zayavka[5])
            elif len(zayavka) > 5:
                KONTRAGENT.append(zayavka[0])
                SUMMA.append(zayavka[2])
                NAZNACHENIE.append(zayavka[1])
                LOCATION.append(zayavka[3])
                USLOVIE.append(zayavka[4])
                POSTAVKA.append(zayavka[5])
            else:
                KONTRAGENT.append('')
                SUMMA.append('')
                NAZNACHENIE.append('')
                LOCATION.append('')
                USLOVIE.append('')
                POSTAVKA.append('')
        elif i.find('1.') != -1 and i[i.find('1.') + 2].isdigit():
            new_i = i[i.find('1.') + 2:]
            zayavka = new_i[new_i.find('1.'):]
            zayavka = zayavka.split(';')
            if len(zayavka) == 1:
                zayavka1 = ' '.join(zayavka)
                zayavka1 = zayavka1.split('\n')
                zayavka = zayavka1
                if len(zayavka) > 5:
                    KONTRAGENT.append(zayavka[0])
                    SUMMA.append(zayavka[2])
                    NAZNACHENIE.append(zayavka[1])
                    LOCATION.append(zayavka[3])
                    USLOVIE.append(zayavka[4])
                    POSTAVKA.append(zayavka[5])
                else:
                    KONTRAGENT.append('')
                    SUMMA.append('')
                    NAZNACHENIE.append('')
                    LOCATION.append('')
                    USLOVIE.append('')
                    POSTAVKA.append('')
            elif len(zayavka) > 5:
                KONTRAGENT.append(zayavka[0])
                SUMMA.append(zayavka[2])
                NAZNACHENIE.append(zayavka[1])
                LOCATION.append(zayavka[3])
                USLOVIE.append(zayavka[4])
                POSTAVKA.append(zayavka[5])
            else:
                KONTRAGENT.append('')
                SUMMA.append('')
                NAZNACHENIE.append('')
                LOCATION.append('')
                USLOVIE.append('')
                POSTAVKA.append('')
        else:
            KONTRAGENT.append('')
            SUMMA.append('')
            NAZNACHENIE.append('')
            LOCATION.append('')
            USLOVIE.append('')
            POSTAVKA.append('')
    else:
        KONTRAGENT.append('')
        SUMMA.append('')
        NAZNACHENIE.append('')
        LOCATION.append('')
        USLOVIE.append('')
        POSTAVKA.append('')

for i, n in enumerate(KONTRAGENT):
    KONTRAGENT[i] = n.replace('1.', '').replace(';', '').strip()

for i, n in enumerate(SUMMA):
    SUMMA[i] = n.replace('3.', '').replace(';', '').strip()

CHAT['Контрагент'] = KONTRAGENT
CHAT['Сумма'] = SUMMA
CHAT['Локация'] = LOCATION
CHAT['Назначение'] = NAZNACHENIE
CHAT['Условие оплаты'] = USLOVIE
CHAT['Срок поставки/оказание'] = POSTAVKA
CHAT = CHAT[CHAT['Сумма'].str.find('2.') == -1]
CHAT.drop_duplicates(inplace=True, ignore_index=True)
CHAT = CHAT[CHAT['Контрагент'] != ''].reset_index()
CHAT.drop(columns=['Unnamed: 0', 'index'], inplace=True)
CHAT.to_excel('C://Users/s.zakarin/Desktop/FINANCE_FINAL/Выгрузки/chat.xlsx')