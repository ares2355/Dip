import os
from datetime import datetime, timedelta

import pymongo
import telebot
import telebot.types as tg
from dotenv import load_dotenv
from openpyxl import Workbook
from pymongo.collection import Collection
import string
import os

load_dotenv()

BOT_API = os.getenv('BOT_API')
MONGO_URL = os.getenv('MONGO_URL')

bot = telebot.TeleBot(BOT_API)
data_users = {}


def get_user_data(user_id):
    if user_id not in data_users.keys():
        data_users[user_id] = {}
    return data_users[user_id]


@bot.message_handler(content_types=['text'], commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, parse_mode='HTML', text=f'''
    Приветствую Вас:{message.chat.username}
<code>Данный бот создан для введения отчетности водителям!</code>
<code>\tВходные данные:</code>
<code>\t-Номер машины, Пример: 946</code>
<code>\t-Пробег машины перед выездом на работу!</code>
<code>\t-Наименования Организации</code>
<code>\t-Зарпалата за смену</code>
<code>\t-Заправленное топливо за смену в Литрах! Пример как указать: 40 или же 0 если нет заправки за смену</code>
''')
    bot.send_message(message.chat.id, '<b>Выбери:</b>', reply_markup=keyboard, parse_mode='HTML')
    bot.register_next_step_handler(message, select)


def select(message: tg.Message):
    if message.text == 'Начал смену!':
        bot.send_message(message.chat.id, '<em>Введите Номер машины:</em>', parse_mode='HTML')
        bot.register_next_step_handler(message, get_number_car)
    elif message.text == 'Закончил смену!':
        bot.send_message(message.chat.id, '<em>Введите пробег машины в конце смены:</em>', parse_mode='HTML')
        bot.register_next_step_handler(message, get_end_shift)
    elif message.text == 'Получить статистику':
        bot.send_message(message.chat.id, '<em>Выбирите период статистики:</em>',
                         reply_markup=keyboard_2, parse_mode='HTML')
        bot.register_next_step_handler(message, get_static)
    else:
        bot.send_message(message.chat.id, '<b>Слудуйте конпкам!</b>', reply_markup=keyboard,
                         parse_mode='HTML')
        bot.register_next_step_handler(message, select)


def get_number_car(message: tg.Message):
    data_user = get_user_data(message.from_user.id)
    if message.text.isdigit() and len(message.text) == 3:
        data_user['time'] = datetime.now()
        data_user['user'] = message.chat.id
        data_user['number_car'] = int(message.text)
        bot.send_message(message.chat.id, '<em>Введите Пробег машины перед выездом на работу:</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_mileage_car)
    else:
        bot.send_message(message.chat.id, '<em>Вы некоректно ввели данные! Попробуем еще раз</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_number_car)


def get_mileage_car(message: tg.Message):
    data_user = get_user_data(message.from_user.id)
    if message.text.isdigit():
        data_user['mileage_car'] = int(message.text)
        bot.send_message(message.chat.id, '<em>Введите Наименования Организации текущей смены:</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_organization)
    else:
        bot.send_message(message.chat.id, '<em>Вы некоректно ввели данные! Попробуем еще раз</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_mileage_car)


def get_organization(message: tg.Message):
    data_user = get_user_data(message.from_user.id)
    if message.text.isalpha():
        data_user['organization'] = message.text.strip()
        bot.send_message(message.chat.id, '<em>Введите Зарпалата за смену:</em>', parse_mode='HTML')
        bot.register_next_step_handler(message, get_salary)
    else:
        bot.send_message(message.chat.id, '<em>Вы некоректно ввели данные! Попробуем еще раз</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_organization)


def get_salary(message: tg.Message):
    data_user = get_user_data(message.from_user.id)
    if message.text.isdigit():
        data_user['salary'] = int(message.text)
        bot.send_message(message.chat.id, '<em>Введите Заправленное топливо за смену в Литрах:</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_fuel)
    else:
        bot.send_message(message.chat.id, '<em>Вы некоректно ввели данные! Попробуем еще раз</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_salary)


def get_fuel(message):
    data_user = get_user_data(message.from_user.id)
    if message.text.isdigit():
        data_user['fuel'] = int(message.text)
        print(data_user)
        bot.send_message(message.chat.id, '<em>Спасибо! Данные приняты!</em>\n'
                                          '\t<em>Нажми Конец смены по завршению работы!</em>', reply_markup=keyboard,
                         parse_mode='HTML')
        coll.insert_one(data_user)
        bot.register_next_step_handler(message, select)
    else:
        bot.send_message(message.chat.id, '<em>Вы некоректно ввели данные! Попробуем еще раз</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_fuel)


def get_end_shift(message: tg.Message):
    data_user = get_user_data(message.from_user.id)
    if message.text.isdigit():
        msg = int(message.text)
        current = data_user['mileage_car']
        coll.update_one({'mileage_car': current}, {'$set': {'end_shift': msg}})
        data_user.clear()
        bot.send_message(message.chat.id, '<em>Спасибо за успешную работу и до скорых встреч!</em>',
                         reply_markup=keyboard, parse_mode='HTML')
        bot.register_next_step_handler(message, select)
    else:
        bot.send_message(message.chat.id, '<em>Вы некоректно ввели данные! Попробуем еще раз</em>',
                         parse_mode='HTML')
        bot.register_next_step_handler(message, get_end_shift)


def data_processing(message, last_day, now_day, period_text):
    header = ['Вермя', 'Пользователь', 'Номер машины', 'Пробег', 'Организация', 'Зарплата', 'Топливо',
              'Конечный пробег']
    lists_statistic = []
    ws1 = wb.create_sheet(f'{period_text}', 0)
    ws1.append(header)
    for post in coll.find({"$and": [{"time": {"$gt": last_day, "$lte": now_day}},
                                    {"user": {"$eq": message.chat.id}}]}):
        lists_statistic.append(post)
    for elm in lists_statistic:
        lists_value = list(elm.values())
        alphabet = list(string.ascii_lowercase)
        for i in alphabet:
            ws1.column_dimensions[i.upper()].width = 18
        ws1.append(lists_value[1:])
        wb.save(f'{message.from_user.id}.xlsx')
        wb.close()


def get_static(message: tg.Message):
    now_day = datetime.now()
    if message.text == 'За неделю':
        period_time = now_day - timedelta(days=7)
    elif message.text == 'За месяц':
        period_time = now_day - timedelta(days=30)
    else:
        bot.send_message(message.chat.id, 'Неверно ввел команду! Потвори!', reply_markup=keyboard_2)
        bot.register_next_step_handler(message, get_static)
        return
    data_processing(message, period_time, now_day, message.text)
    bot.send_document(message.chat.id, open(f'{message.from_user.id}.xlsx', 'rb'),
                      caption=f'Ваша статистика {message.text}!', reply_markup=keyboard)
    bot.register_next_step_handler(message, select)
    # TODO Присылать отчет каждую неделю!
    # TODO добавить xlxs  в отдельную папку или использовать временные файлы в  python!


if __name__ == '__main__':
    keyboard = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    keyboard_2 = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)

    # buttons
    started = tg.KeyboardButton('Начал смену!')
    finished = tg.KeyboardButton('Закончил смену!')
    static = tg.KeyboardButton('Получить статистику')
    week = tg.KeyboardButton('За неделю')
    mount = tg.KeyboardButton('За месяц')

    keyboard.add(started, finished, static)
    keyboard_2.add(week, mount)

    client = pymongo.MongoClient(MONGO_URL)
    db = client.Data_Driver
    coll: Collection = db.Users

    wb = Workbook()
    ws = wb.active

    # ws.title = 'Primer'
    # wb.save('Test_1.xlsx')
    # wb.close()
    bot.infinity_polling()
