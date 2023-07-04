import pandas as pd
from telethon.sync import TelegramClient
from telethon.tl.types import InputPeerChat
from datetime import datetime, timedelta, timezone

# Create the client and connect
api_id = 29360599
api_hash = 'ec2bf9ca90bb465c9959378b62f426a7'
phone = ''

client = TelegramClient('session_name', api_id, api_hash)
client.start()

# Ensure you're authorized
if not client.is_user_authorized():
    client.send_code_request(phone)
    client.sign_in(phone, input('Enter the code: '))

# Get the entity of the group
dialogs = client.get_dialogs()

start_data = input('Пожалуйста введите начальную дату: ')
stop_data = input('Пожалуйста введите конечную дату: ')

start_data += ' 00:00'
stop_data = datetime.strftime(datetime.strptime(stop_data, '%d.%m.%Y') + timedelta(days=1), '%d.%m.%Y')
stop_data += ' 00:00'

start_data = datetime.strptime(start_data, '%d.%m.%Y %H:%M').replace(tzinfo=timezone.utc)
stop_data = datetime.strptime(stop_data, '%d.%m.%Y %H:%M').replace(tzinfo=timezone.utc)


for dialog in dialogs:
    if dialog.name == 'Финансы и расходы Узбекистан':
        group_entity = dialog
        break

# # Extract messages from the group
group_input_entity = InputPeerChat(group_entity.message.peer_id.channel_id)

messages = client.get_messages(group_input_entity.chat_id, limit=2000)


# yesterday = datetime.today() - timedelta(days=1)
# yesterday = datetime.strftime(yesterday, '%d_%m_%Y')
#
# yesterday_date = datetime.strptime(yesterday, '%d_%m_%Y')

name = []
msg = []
reply = []
vremya = []
replier = []
data = pd.DataFrame(columns=['Дата','Утвердитель', 'Сообщение','Заявка', 'Инициатор'])

for i in messages:
    if not i.message is None:
        if not i.reply_to is None:
            for j in messages:
                if j.id == i.reply_to.reply_to_msg_id:
                    if j.date >=start_data and j.date <=stop_data:
                        vremya.append(j.date.strftime('%Y-%m-%d %H:%M'))
                        name.append(f'{j.sender.first_name}')
                        msg.append(j.message)
                        replier.append(i.sender.first_name)

                        if i.media:
                            reply.append('')
                        else:
                            reply.append(i.message)


data['Дата'] = vremya
data['Инициатор'] = name
data['Заявка'] = msg
data['Сообщение'] = reply
data['Утвердитель'] = replier

start_data_save = datetime.strftime(start_data, '%d-%m-%Y')
stop_data_save = datetime.strftime(stop_data, '%d-%m-%Y')

data.to_excel(f'C:/Users/s.zakarin/Desktop/Telegram chat/Чаты/chat_{start_data.date()}_{stop_data.date()}.xlsx')

# 149.154.167.40:443
# -----BEGIN RSA PUBLIC KEY-----
# MIIBCgKCAQEAyMEdY1aR+sCR3ZSJrtztKTKqigvO/vBfqACJLZtS7QMgCGXJ6XIR
# yy7mx66W0/sOFa7/1mAZtEoIokDP3ShoqF4fVNb6XeqgQfaUHd8wJpDWHcR2OFwv
# plUUI1PLTktZ9uW2WE23b+ixNwJjJGwBDJPQEQFBE+vfmH0JP503wr5INS1poWg/
# j25sIWeYPHYeOrFp/eXaqhISP6G+q2IeTaWTXpwZj4LzXq5YOpk4bYEQ6mvRq7D1
# aHWfYmlEGepfaYR8Q0YqvvhYtMte3ITnuSJs171+GDqpdKcSwHnd6FudwGO4pcCO
# j4WcDuXc2CTHgH8gFTNhp/Y8/SpDOhvn9QIDAQAB
# -----END RSA PUBLIC KEY-----


# 149.154.167.50:443
# -----BEGIN RSA PUBLIC KEY-----
# MIIBCgKCAQEA6LszBcC1LGzyr992NzE0ieY+BSaOW622Aa9Bd4ZHLl+TuFQ4lo4g
# 5nKaMBwK/BIb9xUfg0Q29/2mgIR6Zr9krM7HjuIcCzFvDtr+L0GQjae9H0pRB2OO
# 62cECs5HKhT5DZ98K33vmWiLowc621dQuwKWSQKjWf50XYFw42h21P2KXUGyp2y/
# +aEyZ+uVgLLQbRA1dEjSDZ2iGRy12Mk5gpYc397aYp438fsJoHIgJ2lgMv5h7WY9
# t6N/byY9Nw9p21Og3AoXSL2q/2IJ1WRUhebgAdGVMlV1fkuOQoEzR7EdpqtQD9Cs
# 5+bfo3Nhmcyvk5ftB0WkJ9z6bNZ7yxrP8wIDAQAB
# -----END RSA PUBLIC KEY-----