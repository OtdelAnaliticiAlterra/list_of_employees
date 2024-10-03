import time
import os
from datetime import datetime
from education import education
from employees import employees
from work_exp import work_exp
from competence import competence
from photo import photos
from profile import profile
from rec import rec

from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

from constants import CHATS_IDS
from telegram_bot_logger import TgLogger

logger = TgLogger(
            name='ListOfEmployees',
            filename='Errors.log',
            token=os.environ.get('LOGGER_BOT_TOKEN'),
            chats_ids_filename=CHATS_IDS,
        )

# while True:
try:
    # if datetime.now().hour == 7:
        # 7
        # print('Старт:', datetime.now())

        # logger.info('Старт обработки списка сотрудников:', datetime.now())
        logger.info(f'Sтарт обработки списка сотрудников: {datetime.now()}')


        employees()
        education()
        work_exp()
        photos()
        profile()
        rec()
        competence()
        # print('Завершено:', datetime.now())

        # logger.info('Обработка списка сотрудников завершена:', datetime.now())
        logger.info(f'Обработка списка сотрудников завершена: {datetime.now()}')

        # time.sleep(60 * 60 * 12)
    # else:
    #     time.sleep(60 * 30)
except PermissionError as e:
    logger.error("Список сотрудников. Ошибка доступа к файлу. Проверьте права доступа. Ошибка: %s", e)
    # print("Ошибка доступа к файлу. Проверьте права доступа. Ошибка:", e)
except FileNotFoundError as e:
    logger.error("Список сотрудников. Файл не найден. Ошибка: %s", e)
    # print("Файл не найден. Ошибка:", e)
except Exception as e:
    logger.error("Список сотрудников. Произошла ошибка: %s", e)
    # print("Произошла ошибка:", e)