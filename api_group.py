import pandas as pd
from telethon.sync import TelegramClient
from telethon.tl.types import InputPeerChat
from datetime import datetime, timedelta, timezone

# Create the client and connect
api_id = 'your api id'
api_hash = 'your api hash'
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


yesterday = datetime.today() - timedelta(days=1)
yesterday = datetime.strftime(yesterday, '%d_%m_%Y')

yesterday_date = datetime.strptime(yesterday, '%d_%m_%Y')

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

data.to_excel(f'Path_to_file/chat_{start_data.date()}_{stop_data.date()}.xlsx')