from telethon import TelegramClient, events, sync, errors
from telethon.errors import SessionPasswordNeededError
from telethon.tl.types import Message
from telethon.tl.types import MessageMediaWebPage
from telethon.tl.types import Chat, Channel
from telethon.errors.rpcerrorlist import UserBannedInChannelError

from loguru import logger

from datetime import datetime, timedelta
from time import sleep

from openpyxl.styles import NamedStyle
from openpyxl import load_workbook

import pandas as pd
import pytz
import os

# Consts
DATA_FILENAME = f'data_{datetime.now().strftime(r'%d-%m-%Y')}.xlsx'
FILE_EXISTS = True if os.path.exists(DATA_FILENAME) else False
COLUMNS = ['USERNAME ГРУППЫ', 'ВСЕГО СООБЩЕНИЙ ЗА ЧАС', 'ВСЕГО ПОЛЬЗОВАТЕЛЕЙ ЗА ЧАС', 'ОТПРАВКА СООБЩЕНИЙ']
API_ID = 20419714
API_HASH = "feee2b161028b2ce71cc92ede611e4ed"

with open('groups.txt', 'r') as file:
    LINES = [line for line in file.read().split('\n') if line]

class CollectionStats:

    def __init__(self) -> None:
        self.total_users = 0
        self.total_messages = 0
        self.write_to_chat = None

        self.client = TelegramClient(
            session='my_account', 
            api_id=API_ID, 
            api_hash=API_HASH)
    
    def telegram_client_connect(self) -> None:
        self.client.connect()

        if not self.client.is_user_authorized():
            logger.error("Сессия не найдена")
            print("Авторизация:")
            print("Введите номер телефона")
            
            self.client.sign_in(input('>>> '))

            try:
                print("Введите отправленный код")
                self.client.sign_in(code=input('>>> '))
            except SessionPasswordNeededError:
                print("Введите пароль")
                self.client.sign_in(password=input('>>> '))

            logger.success("Авторизация прошла успешно")

            os.system('cls')
            print('Developer - https://t.me/daniilprg\n')

    def get_saved_message(self) -> Message:
        for message in self.client.get_messages(
            self.client.get_me().id, 
            limit=2):
            if isinstance(message, Message):
                return message

    def check_write_to_chat(self, *, group, message) -> bool:
        try:
            self.client.send_message(group, message)
            logger.info('Проверочное сообщение успешно отправлено')
            return True
        except errors.ChatWriteForbiddenError:
            logger.error('Нет прав на отправку сообщений')
            return False
        except errors.FloodWaitError as e:
            logger.error(f'Сработал лимит отправки сообщений. Ожидаю {e.seconds} секунд')
            sleep(e.seconds + 1)
            self.check_write_to_chat(group=group, message=message)
        except errors.RPCError as e:
            logger.error('Нет прав на отправку сообщений')
            return False
        except Exception as e:
            logger.error(f'Возникла непредвиденная ошибка: {e}')
            return False

    def check_total_users_for_hour(self, *, group) -> int:
        timezone = pytz.timezone('Europe/Moscow')
        one_hour_ago = datetime.now(timezone) - timedelta(hours=1)
        participants = set()

        try:
            for message in self.client.iter_messages(group):
                if isinstance(message, Message) and message.date >= one_hour_ago: 
                    participants.add(message.sender_id)
                else: return len(participants)
        except Exception as e:
            logger.error(f'Возникла непредвиденная ошибка: {e}')
            return 0

    def check_total_messages_for_hour(self, *, group) -> int:
        timezone = pytz.timezone('Europe/Moscow')
        one_hour_ago = datetime.now(timezone) - timedelta(hours=1)
        number_of_messages = 0

        try:
            for message in self.client.iter_messages(group):
                if isinstance(message, Message) and message.date >= one_hour_ago:
                    number_of_messages += 1
                else: return number_of_messages
        except Exception as e:
            logger.error(f'Возникла непредвиденная ошибка: {e}')
            return 0
    
    def run(self, FILE_EXISTS) -> None:
        for line in LINES:
            logger.info(f'Проверка статистики группы @{line}')

            total_messages = self.check_total_messages_for_hour(group=line)
            total_users = self.check_total_users_for_hour(group=line)
            write_to_chat = self.check_write_to_chat(group=line, message=self.get_saved_message())
            
            data = {
                'USERNAME ГРУППЫ': [line], 
                'ВСЕГО СООБЩЕНИЙ ЗА ЧАС': [total_messages], 
                'ВСЕГО ПОЛЬЗОВАТЕЛЕЙ ЗА ЧАС': [total_users], 
                'ОТПРАВКА СООБЩЕНИЙ': ['Разрешена' if write_to_chat else 'Запрещена']
            }

            if not FILE_EXISTS:
                df = pd.DataFrame(columns=COLUMNS)
                df.to_excel(DATA_FILENAME, index=False)
                FILE_EXISTS = True

            df = pd.read_excel(DATA_FILENAME)
            df2 = pd.DataFrame(data)
            combined_df = pd.concat([df, df2], ignore_index=True)

            with pd.ExcelWriter(DATA_FILENAME, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']

                worksheet.set_column('A:A', 25)
                worksheet.set_column('B:B', 30)
                worksheet.set_column('C:C', 30)
                worksheet.set_column('D:D', 30)
                worksheet.set_column('E:E', 25)

            logger.success(f'Статистика группы @{line} успешно собрана')

if __name__ == "__main__":
    print('Developer - https://t.me/daniilprg\n')
    obj = CollectionStats()
    obj.telegram_client_connect()
    obj.run(FILE_EXISTS)


