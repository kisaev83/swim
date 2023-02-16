# pip install pytelegrambotapi
# pip install pandas
# pip install openpyxl
# pip install XlsxWriter
# pip install telebot-calendar
# pip install cryptocode
import os
import re
# import logging
import telebot
import datetime as dt
import cryptocode
from telebot_calendar import Calendar, CallbackData, RUSSIAN_LANGUAGE
from telebot.types import ReplyKeyboardRemove, CallbackQuery, ForceReply
from transliterate import translit

import instruction
from data import create_app, list_swimmers_from_excel, create_start_protocol, create_track_protocol_for_organization, \
    results_to_db, results_from_excel, create_points, tournament_year_categories, \
    statistic_years, create_list_2, statistic_by_cat, years_in_categories, \
    create_pre_results, create_final_protocol, create_medals, create_start_protocol_for_organization
from data_group import write_to_group_open_excel, group_write_to_excel_and_db
from read_config import read_db_config
from db import connect_to_db, insert_db, close_connect_db, update_db, select_db

pdf_create = False

token = read_db_config(section='bot')["token"]
bot = telebot.TeleBot(token)

# Creates a unique calendar
calendar = Calendar(language=RUSSIAN_LANGUAGE)
calendar_1_callback = CallbackData("calendar_1", "action", "year", "month", "day")


@bot.message_handler(commands=['start'])
def command_start(message, editting=False):
    # print(message)
    name_user = message.chat.first_name
    second_name_user = message.chat.last_name
    # print(second_name_user)
    bot.clear_step_handler(message)

    def now_hour():
        now = dt.datetime.now()
        if 11 > now.hour >= 4:
            return 'Доброе утро'
        elif 17 >= now.hour >= 11:
            return 'Добрый день'
        elif 22 >= now.hour > 17:
            return 'Добрый вечер'
        else:
            return 'Доброй ночи'

    text = now_hour() + ', ' + name_user + '!'
    conn = connect_to_db()
    tables = select_db(conn, None, False, 'users', 'table_name', id_telegram=message.chat.id)
    keyboard = telebot.types.InlineKeyboardMarkup()
    try:
        if tables['table_name']:
            keyboard.add(telebot.types.InlineKeyboardButton(text='Турниры в работе', callback_data='Турниры в работе'))
            keyboard.add(telebot.types.InlineKeyboardButton(text='Архивные турниры', callback_data='Архивные турниры'))
    except:
        pass
    if not tables or tables['table_name'] is None:
        text += '\nВ данный момент времени бот-помощник @BeSwimBot проходит тестирование у определенного ' \
                'круга организаторов. Мы оповестим вас, когда бот начнёт полноценную работу. Спасибо.'

    else:
        keyboard.add(telebot.types.InlineKeyboardButton(text='Создать новый турнир', callback_data='Создать турнир'))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Присоединиться к турниру', callback_data='присоединиться'))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Справка по боту', callback_data='instruction'))

    if editting:
        bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)
    else:
        bot.send_message(message.chat.id, text, reply_markup=keyboard)

    query = f'DELETE FROM temp_tournament WHERE id_telegram={message.chat.id}'

    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    query = f"INSERT INTO users (id_telegram, first_name, second_name, count_enter, last_date) " \
            f"VALUES ({message.chat.id}, '{name_user}', '{second_name_user}', 1, '{dt.datetime.today()}') " \
            f"ON DUPLICATE KEY UPDATE count_enter = count_enter + 1, last_date = '{dt.datetime.today()}'"
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    close_connect_db(conn)


def send_pdf(message, file_name, caption, keyboard=None):
    command = "libreoffice --headless --convert-to pdf " + file_name
    os.system(command)
    file_name_pdf = file_name[:-4] + 'pdf'
    with open(file_name_pdf, 'rb') as file:
        bot.send_document(message.chat.id, file, caption=caption, parse_mode='markdown', reply_markup=keyboard)
    try:
        os.remove(file_name_pdf)
    except:
        pass


def create_new_tournament(message):
    text = 'Напишите название нового турнира. Соблюдайте регистр! Название будет публиковаться в протоколах, ' \
           'проверьте опечатки, орфографические и пунктуационные ошибки перед отправкой.\n' \
           'Не указывайте в названии дату турнира. Выбор даты будет на следующем шаге.'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.edit_message_text(chat_id=message.chat.id, message_id=message.message_id,
                          text=text, reply_markup=keyboard)
    bot.register_next_step_handler(message, save_tour_name)


def save_tour_name(message):
    msg_text = message.text  # сохранение текста от клиента
    # print(msg_text)
    conn = connect_to_db()
    insert_db(conn, table='temp_tournament', id_telegram=message.chat.id, name=msg_text)
    close_connect_db(conn)
    calendar_message(message)


def calendar_message(message):
    now = dt.datetime.now()  # Get the current date
    bot.send_message(
        message.chat.id,
        "Выберите дату проведения турнира",
        reply_markup=calendar.create_calendar(
            name=calendar_1_callback.prefix,
            year=now.year,
            month=now.month,  # Specify the NAME of your calendar
        ),
    )


def place_of_tour(message):
    text = 'Напишите место проведения турнира. Город и название бассейна. (Например: г. Воскресенск, п/б "Дельфин"). ' \
           'Не указывайте длину бассейна и количество дорожек, эти вопросы будет далее.'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.send_message(message.chat.id, text, reply_markup=keyboard)
    bot.register_next_step_handler(message, save_place_name)


def save_place_name(message):
    msg_text = message.text  # сохранение текста от клиента
    conn = connect_to_db()
    update_db(conn, None, table="temp_tournament", id_telegram=message.chat.id, place=msg_text)
    close_connect_db(conn)
    swimming_pool_lenght(message)


def swimming_pool_lenght(message):
    text = 'Выберите длину бассейна'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='25м', callback_data='25м'),
                 telebot.types.InlineKeyboardButton(text='50м', callback_data='50м'))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.send_message(message.chat.id, text, reply_markup=keyboard)


def number_tracks(message):
    text = 'Выберите количество дорожек'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='4', callback_data='tracks4'),
                 telebot.types.InlineKeyboardButton(text='5', callback_data='tracks5'),
                 telebot.types.InlineKeyboardButton(text='6', callback_data='tracks6'),
                 telebot.types.InlineKeyboardButton(text='7', callback_data='tracks7'),
                 telebot.types.InlineKeyboardButton(text='8', callback_data='tracks8'), row_width=5)
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)


def sessions_func(message):
    text = 'Сколько сессий планируется?'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='1 сессия', callback_data='session_1'),
                 telebot.types.InlineKeyboardButton(text='2 сессии', callback_data='session_2'))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='markdown')


def years_swimmers(message, session, old_mess=None, text_=''):
    old_message = message
    text = text_ + 'Напишите возрастные группы(категории) участников'
    if session == 1:
        text = text + '\n*2-й сессии*'
        mess = old_mess.message_id
    elif session == 2:
        text = text + '\n*1-й сессии*'
        mess = message.message_id
    else:
        mess = message.message_id
    text = text + ' через запятую, соблюдая пример ниже.\n*Пример: 2007+, 2008, 2009-2010, 2011-2013, 2014-*\n' \
                  '_Что означает: 2007 г.р. и старше, 2008 г.р., объединенные года 2009 и 2010, ' \
                  'объединенные года с 2011 по 2013, с 2014 г.р. и младше)_'
    # if old_mess:
    #     print(old_mess.message_id)
    # print(text)
    if session == 1:
        conn = connect_to_db()
        years = select_db(conn, None, False, 'temp_tournament', 'years', id_telegram=message.chat.id)['years']
        close_connect_db(conn)
        text = text + '\n' + '*1 сессия: ' + years + '*'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    if text_ == '':
        bot.edit_message_text(text, message.chat.id, mess, reply_markup=keyboard, parse_mode='markdown')
    else:
        old_message = bot.send_message(message.chat.id, text, 'markdown', reply_markup=keyboard)
    bot.register_next_step_handler(message, save_years, session, old_message)


def save_years(message, session, old_message):
    msg_text = message.text.replace(' ', '')  # сохранение текста от клиента
    regex = "^[1234567890\-,+ ]+$"
    pattern = re.compile(regex)
    if pattern.search(msg_text) is not None:
        conn = connect_to_db()
        if session == 2 or session == 0:
            update_db(conn, None, table="temp_tournament", id_telegram=message.chat.id, years=msg_text)
        elif session == 1:
            query = f"UPDATE temp_tournament SET years = CONCAT(years, '|', '{msg_text}') WHERE id_telegram={message.chat.id}"

            cursor = conn.cursor()
            cursor.execute(query)
            conn.commit()
        session -= 1
        close_connect_db(conn)
        if session == -1 or session == 0:
            tour_is_ready(message)
        elif session == 1:
            years_swimmers(message, session, old_message)
    else:
        years_swimmers(message, session, old_message, "*Вы ввели неправильный формат данных. В сообщении могут быть"
                                                      " только цифры от 0 до 9, знак ','(запятая) между категориями, "
                                                      "и при необходимости знаки '-' и '+'*\n")


def choose_distances(message, choosen):
    # print(choosen)

    freestyle = ['50м в/с', '100м в/с', '200м в/с']
    freestyle_2 = ['400м в/с', '800м в/с', '1500м в/с']
    backstyle = ['50м на спине', '100м на спине', '200м на спине']
    brass = ['50м брасс', '100м брасс', '200м брасс']
    batt = ['50м батт.', '100м батт.', '200м батт.']
    kompleks = ['100м комплекс', '200м комплекс', '400м комплекс']
    classic_last = ['25м в/с кл.ласты', '50м в/с кл.ласты', '100м в/с кл.ласты', '200м в/с кл.ласты']
    small_distance = ['25м в/с', '25м на спине', '25м брасс', '25м батт.']
    greblya_1 = ['250м', '500м', '750м', '1000м']
    greblya_2 = ['1250м', '1500м', '1750м', '2000м']
    greblya_3 = ['4000м', '6000м', '10000м']

    conn = connect_to_db()
    swimming_pool = select_db(conn, None, False, 'temp_tournament', 'swimming_pool', id_telegram=message.chat.id)
    if swimming_pool['swimming_pool'] == '50м':
        kompleks.remove('100м комплекс')
    distance = ''
    for i in range(1, 11):
        distance = distance + 'distance_' + str(i) + ', '
    distance = distance[:-2]
    query = f"SELECT {distance} FROM temp_tournament WHERE id_telegram={message.chat.id}"
    data = select_db(conn, query)
    close_connect_db(conn)
    choosen_distances = []
    for values in data.values():
        if values is None:
            break
        choosen_distances.append(values)
        try:
            freestyle.remove(values)
        except:
            pass
        try:
            freestyle_2.remove(values)
        except:
            pass
        try:
            backstyle.remove(values)
        except:
            pass
        try:
            brass.remove(values)
        except:
            pass
        try:
            batt.remove(values)
        except:
            pass
        try:
            kompleks.remove(values)
        except:
            pass
        try:
            classic_last.remove(values)
        except:
            pass
        try:
            small_distance.remove(values)
        except:
            pass
        try:
            greblya_1.remove(values)
        except:
            pass
        try:
            greblya_2.remove(values)
        except:
            pass
        try:
            greblya_3.remove(values)
        except:
            pass
    # print('choosen_distances', choosen_distances)
    text = 'Выберите дистанции для турнира в порядке их проведения. После выбора всех дистанций нажмите "Готово".\n'
    for dist_ in choosen_distances:
        text = text + str(choosen_distances.index(dist_) + 1) + '. ' + dist_ + '\n'
    keyboard = telebot.types.InlineKeyboardMarkup()
    button = []
    for dist in freestyle:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])
    button = []
    for dist in freestyle_2:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])
    button = []
    for dist in backstyle:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])

    button = []
    for dist in brass:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])

    button = []
    for dist in batt:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])

    button = []
    for dist in kompleks:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])

    button = []
    for dist in classic_last:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 4:
        for i in range(1, 5 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2], button[3])
    button = []
    for dist in small_distance:
        button.append(telebot.types.InlineKeyboardButton(text=dist, callback_data='choose ' + dist))
    if len(button) < 3:
        for i in range(1, 4 - len(button)):
            button.append(telebot.types.InlineKeyboardButton(text='', callback_data=' '))
    keyboard.add(button[0], button[1], button[2])
    keyboard.add(telebot.types.InlineKeyboardButton(text='Готово', callback_data='Готово дистанции'))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)


def tour_is_ready(message):
    conn = connect_to_db()
    query = f"SELECT * FROM temp_tournament WHERE id_telegram={message.chat.id}"
    data = select_db(conn, query)
    count_distance = 0
    for key, value in data.items():
        if key.startswith('distance'):
            if value is not None:
                count_distance += 1
    table_name = 'a' + data['date'].strftime('%d_%m_%Y_') + str(data['id_telegram'])

    insert_db(conn, 'tournaments', id_telegram=data['id_telegram'], table_name=table_name, name=data['name'],
              date=data['date'], place=data['place'], swimming_pool=data['swimming_pool'], tracks=data['tracks'],
              years=data['years'], distance_1=data['distance_1'], distance_2=data['distance_2'],
              distance_3=data['distance_3'], distance_4=data['distance_4'], distance_5=data['distance_5'],
              distance_6=data['distance_6'], distance_7=data['distance_7'], distance_8=data['distance_8'],
              distance_9=data['distance_9'], distance_10=data['distance_10'], added_date=dt.datetime.now(),
              quantity_dist=count_distance)
    tables = select_db(conn, None, False, 'users', 'table_name', 'rights', id_telegram=message.chat.id)
    if tables['table_name']:
        table_names = tables['table_name'] + '|' + table_name
        rights = tables['rights'] + '|' + 'all'
    else:
        table_names = table_name
        rights = 'all'
    update_db(conn, table='users', id_telegram=message.chat.id, table_name=table_names, rights=rights)
    query = f'DELETE FROM temp_tournament WHERE id_telegram={message.chat.id}'
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()

    distances = ''
    for i in range(1, count_distance + 1):
        distances = distances + 'distance_' + str(i) + ' varchar(20), distance_' + str(i) + \
                    '_time float, distance_' + str(i) + \
                    '_result float, distance_' + str(i) + '_fina int(4), distance_' + str(i) + \
                    '_dsq varchar(12), distance_' + str(i) + '_dsq_id int, FOREIGN KEY (distance_' \
                    + str(i) + '_dsq_id) REFERENCES dsq (id),'
    query = f'CREATE TABLE {table_name} (id int NOT NULL AUTO_INCREMENT, second_name varchar(30), ' \
            f'first_name  varchar(30), gender varchar(1), year int(4), command varchar(50), ' \
            f'{distances} gold int(2), silver int(2), bronze int(2), PRIMARY KEY (id))'
    # print(query)
    try:
        cursor = conn.cursor()
        cursor.execute(query)
        conn.commit()
        query = f'ALTER TABLE `swimming`.`{table_name}` ADD UNIQUE (`second_name`, `first_name`, `command`)'
        cursor = conn.cursor()
        cursor.execute(query)
        conn.commit()
        text = 'Турнир создан. Он доступен в разделе "Турниры в работе"'
    except:
        text = 'Турнир с этой датой уже существует.'
    close_connect_db(conn)
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Турниры в работе', callback_data='Турниры в работе'))
    bot.send_message(message.chat.id, text, reply_markup=keyboard)


def tournaments_in_work(message, archive=False):
    conn = connect_to_db()
    tables = select_db(conn, None, False, 'users', 'table_name', 'rights', id_telegram=message.chat.id)
    table_names = tables['table_name'].split('|')
    rights = tables['rights'].split('|')
    tab_name = ''
    for table in table_names:
        tab_name = tab_name + "table_name = '" + table + "' OR "
    tab_name = tab_name[:-3]
    if not archive:
        query = f"SELECT * FROM tournaments WHERE ({tab_name}) AND date >= curdate() ORDER BY added_date"
        text = '*Список предстоящих турниров.*\n\n'
    else:
        query = f"SELECT * FROM tournaments WHERE ({tab_name}) AND date < curdate() ORDER BY date DESC"
        text = '*Список прошедших турниров.*\n\n'
    data = select_db(conn, query, True)
    close_connect_db(conn)

    count_tour = 0
    keyboard = telebot.types.InlineKeyboardMarkup()
    for j, tournament in enumerate(data):
        count_tour += 1
        distances = ''
        for i in range(8, len(tournament) - 2):
            if tournament[i] is not None:
                distances = distances + tournament[i] + ', '
        distances = distances[:-2]
        categories = tournament[7].replace(' ', '')
        categories = categories.replace(',', ', ')
        if '|' in categories:
            categories = '1 сессия: ' + categories
            categories = categories.replace('|', '; 2 сессия: ')
        text += str(count_tour) + '. ' + tournament[3].strftime('%d.%m.%Y, ') + tournament[2] + '.\n'
        if not archive:
            text += tournament[4] + ', ' + tournament[5] + ', ' + str(tournament[6]) + ' дор.' + '\n' + \
                    'Возрастные категории: ' + categories + ' г.р.\n' + 'Дистанции: ' + distances + '\n\n'
        text_call = str(count_tour) + '. ' + tournament[3].strftime('%d.%m.%Y, ') + tournament[2]
        rights_str = ''
        for n, table_name in enumerate(table_names):
            if tournament[1] == table_name:
                rights_str = rights[n]
                break
        keyboard.add(telebot.types.InlineKeyboardButton(text=text_call, callback_data='_' + tournament[1]
                                                                                      + '-' + rights_str))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='Markdown')


def tournament_menu(message, table_name, rights):
    bot.clear_step_handler(message)
    conn = connect_to_db()
    data = select_db(conn, None, False, 'tournaments', '*', table_name="'" + table_name + "'")
    query = f"SELECT count(*) FROM {table_name}"
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    count = cursor.fetchone()[0]
    cursor.close()
    query = f"SELECT count(*) FROM {table_name} WHERE gender = 'м'"
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    count_boys = cursor.fetchone()[0]
    cursor.close()
    query = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='swimming' AND TABLE_NAME='{table_name}' AND COLUMN_NAME LIKE '%result'"
    distances = select_db(conn, query, True)
    query = ''
    for distance in distances:
        query = query + distance[0] + ' > 0 OR '
    query = f"SELECT count(*) FROM {table_name} WHERE " + query[:-3]
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    count_results = cursor.fetchone()[0]
    cursor.close()
    close_connect_db(conn)
    count_girls = count - count_boys
    distances = ''
    for i in range(1, 11):
        if data['distance_' + str(i)] is not None:
            distances = distances + data['distance_' + str(i)] + ', '
    distances = distances[:-2]
    text = '*' + data['date'].strftime('%d.%m.%Y, ') + data['name'] + \
           '.*\n' + data['place'] + ', ' + data['swimming_pool'] + ', ' + str(data['tracks']) + ' дор.' + '\n' \
           + 'Дистанции: ' + distances + \
           '\nКоличество заявившихся участников: ' + str(count) + ' (' + str(count_boys) + ' юн. и ' + \
           str(count_girls) + ' дев.)'

    keyboard = telebot.types.InlineKeyboardMarkup()
    text_button = ['Текущие результаты', 'Загрузить результаты',
                   'Заявочный протокол', 'Стартовый протокол',
                   'Итоговый протокол', 'Медалисты',
                   'Многоборье по очкам FINA2022', 'Загрузить техническую заявку от команды',
                   'Сформировать техническую заявку', 'Списки участников',
                   'Начать турнир',
                   'Изменить/удалить турнир', 'Добавить помощников в турнир',
                   'загрузить левый excel']
    callback_data = ['pre_' + table_name + '-' + rights, 'uploadresults_' + table_name + '-' + rights,
                     'createlist_' + table_name + '-' + rights, 'startprotocol_' + table_name + '-' + rights,
                     'createprotocol_' + table_name + '-' + rights, 'medals_' + table_name + '-' + rights,
                     'points_' + table_name + '-' + rights, 'uploadapp_' + table_name + '-' + rights,
                     'createapp_' + table_name + '-' + rights, 'statswim_' + table_name + '-' + rights,
                     'begin_' + table_name + '-' + rights,
                     'changetour_' + table_name + '-' + rights, 'helper(' + table_name,
                     'excel_' + table_name + '-' + rights]

    # 0 - 'Текущие результаты'
    # 1 - 'Загрузить результаты'
    # 2 - 'Заявочный протокол'
    # 3 - 'Стартовый протокол',
    # 4 - 'Итоговый протокол'
    # 5 - 'Медалисты'
    # 6 - 'Многоборье по очкам FINA2022'
    # 7 - 'Загрузить техническую заявку от команды'
    # 8 - 'Сформировать техническую заявку',
    # 9 - 'Списки участников'
    # 10 - 'Начать турнир'
    # 11 - 'Изменить/удалить турнир'
    # 12 - 'Добавить помощников в турнир'
    # 13 - 'загрузить левый excel'

    rights_number = []
    count_swimmers_is_0 = [8, 7, 11, 12]
    count_results_is_0 = [1, 2, 3, 7, 9, 11, 12]  # [1, 2, 3, 7, 8, 9, 11, 12]
    count_results_is_0_and_date_no_today = [2, 3, 7, 9, 11, 12]
    count_results_is_not_0 = [0, 1, 4, 5, 6, 9, 11, 12]
    date_old = [0, 1, 4, 5, 6, 9, 11]
    # date_today = [1, 2, 3, 7, 9, 11, 12]  # 10
    all_rights = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12]  # , 13 левый эксель
    temp_rights = []
    if rights == 'all':
        temp_rights = all_rights
    else:
        for i in rights:
            temp_rights.append(int(i))
    if count == 0:
        rights_number = list(set(temp_rights) & set(count_swimmers_is_0))
    # elif data['date'] == dt.date.today() and count_results == 0:
    #     print('2')
    #     rights_number = list(set(temp_rights) & set(date_today))
    elif data['date'] != dt.date.today() and count_results == 0:
        rights_number = list(set(temp_rights) & set(count_results_is_0_and_date_no_today))
    elif count_results == 0:
        rights_number = list(set(temp_rights) & set(count_results_is_0))
    elif data['date'] < dt.date.today():
        rights_number = list(set(temp_rights) & set(date_old))
    elif count_results > 0:
        rights_number = list(set(temp_rights) & set(count_results_is_not_0))
    else:
        rights_number = temp_rights
    # print(rights_number)

    for i in rights_number:
        keyboard.add(telebot.types.InlineKeyboardButton(text=text_button[i], callback_data=callback_data[i]))

    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню',
                                                    callback_data='Назад в меню'))
    try:
        bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='markdown')
    except:
        bot.send_message(message.chat.id, text, reply_markup=keyboard, parse_mode='markdown')


def statistic_menu(message, table_name, rights):
    conn = connect_to_db()
    data = select_db(conn, None, False, 'tournaments', '*', table_name="'" + table_name + "'")
    close_connect_db(conn)
    text = '*Списки участников*\n' + data['date'].strftime('%d.%m.%Y, ') + data['name'] + \
           '.\n'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Список участников по категориям',
                                                    callback_data='by cat_' +
                                                                  table_name + '-' + rights))
    # keyboard.add(telebot.types.InlineKeyboardButton(text='Список участников по командам',
    #                                                 callback_data='by com_' +
    #                                                               table_name + '-' + rights))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Общее количество участников',
                                                    callback_data='by all_' +
                                                                  table_name + '-' + rights))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                    callback_data='меню тур_' +
                                                                  table_name + '-' + rights))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='markdown')


def upload_app_from_command(message, table_name, rights):
    bot.send_message(message.chat.id,
                     'Заявки в бот необходимо загружать строго по одному файлу.\nОтправьте один файл xlsx с технической заявкой от команды.')
    bot.register_next_step_handler(message, handle_docs_photo, table_name, rights)


def handle_docs_photo(message, table_name, rights):  # функция сохранения файла для печати из сообщения
    file_info = bot.get_file(message.document.file_id)  # получение файла из сообщения
    downloaded_file = bot.download_file(file_info.file_path)
    file_name = message.document.file_name

    with open(file_name, 'wb') as new_file:
        new_file.write(downloaded_file)
    text, command = list_swimmers_from_excel(file_name, table_name)
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Загрузить еще одну заявку',
                                                    callback_data='uploadapp_' +
                                                                  table_name + '-' + rights))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                    callback_data='меню тур_' +
                                                                  table_name + '-' + rights))
    bot.send_message(message.chat.id, f'Файл загружен.\n{text}', 'markdown')
    bot.send_message(message.chat.id, f'Вернуться в меню', reply_markup=keyboard)
    try:
        os.remove(file_name)
    except:
        pass
    bot.send_message(360693297, 'Загружена заявка от команды в таблицу ' + table_name + '\n' + text)


def createprotocol_(message, table_name, rights):
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Сгруппировать участников по годам',
                                                    callback_data='По годам_' + table_name + '-' + rights))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Все участники по времени',
                                                    callback_data='По времени_' + table_name + '-' + rights))
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                    callback_data='меню тур_' +
                                                                  table_name + '-' + rights))
    text = 'Определите порядок формирования стартового протокола'
    bot.edit_message_text(message.text, message.chat.id, message.message_id)
    bot.send_message(message.chat.id, text, reply_markup=keyboard)


def createprotocol_years_2(message, table_name, rights, all=None, session=1):
    def text_for_sessions(text2, session_):
        try:
            if message.text.startswith("Объедините"):
                text2 = text2 + '*'
                for number_, groups in enumerate(session_choose_years[session_], start=1):
                    # print('number-group', number_, groups)
                    text2 = text2 + str(number_) + ' группа. ' + groups + '\n'
                text2 = text2 + '*'
        except:
            return text2[:-1]
        return text2

    conn = connect_to_db()
    session_categories, sessions_ = tournament_year_categories(conn, table_name)
    query = f"INSERT IGNORE INTO `temp_protocol` SET `table_name` = '{table_name}'"
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    conn.commit()
    groups_choosen = select_db(conn, None, False, 'temp_protocol', 'groups', table_name="'" + table_name + "'")
    query = f'SELECT DISTINCT year FROM {table_name} ORDER BY year'
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    years_in_ = cursor.fetchall()
    years_in_tournament = []
    cursor.close()
    # print(session_categories)
    if all:
        years = ''
        for session_ in session_categories.values():
            # print(session_)
            for year_ in session_:
                if '-' in year_ and len(year_) > 7:
                    year1 = int(year_[:4])
                    year2 = int(year_[5:])
                    for y in range(year1, year2 + 1):
                        years = years + str(y) + ','
                else:
                    years = years + year_ + ','
            years = years[:-1]
            years += '/'
            # print('years', years)
        years = years[:-1]
        # print('years2', years)
        query = f"UPDATE `temp_protocol` SET `groups`='{years}' WHERE " \
                f"`table_name`='{table_name}'"
        update_db(conn, query)
        close_connect_db(conn)
        return
    for years_ in years_in_:
        years_in_tournament.append(years_[0])
    # print('groups_choosen', groups_choosen)
    # print('session_categories', session_categories)
    # print('years_in_tournament', years_in_tournament)
    text = 'Объедините года в общие группы заплывов в порядке их прохождения дистанций. Выберите первый год, затем второй и т.д. ' \
           'для группы. После формирования группы нажмите "Группа готова", и продолжайте пока все года не будут использованы.\n'
    keyboard = telebot.types.InlineKeyboardMarkup()

    session_choose_years = {}
    for key, value in session_categories.items():
        for n, val_ in enumerate(value):
            if '-' in val_ and len(val_) > 7:
                session_categories[key][n] = val_[:4]
                session_categories[key].append(val_[5:])
        session_categories[key].sort()
    if groups_choosen['groups'] is not None:
        session_list_of_choose_years = groups_choosen['groups'].split('/')
        if '' in session_list_of_choose_years:
            session_list_of_choose_years.remove('')
        for number, sg in enumerate(session_list_of_choose_years, start=1):
            session_choose_years[number] = sg.split('|')
        if groups_choosen['groups'][-1] == '/':
            groups_choose_years = groups_choosen['groups'][:-1]
        else:
            groups_choose_years = groups_choosen['groups'].replace('/', ',')
        groups_choose_years = groups_choose_years.replace('|', ',')
        list_of_choose_years = groups_choose_years.split(',')
        for key, value in session_categories.items():
            for year_remove in list_of_choose_years:
                try:
                    session_categories[key].remove(year_remove)
                except:
                    pass
    if sessions_ == 2:
        if session == 1:
            text = text + '*1 сессия*\n'
            for year in session_categories[1]:
                keyboard.add(telebot.types.InlineKeyboardButton(text=year,
                                                                callback_data='year1_' + str(year) + '_' +
                                                                              table_name + '-' + rights))
            if not keyboard.keyboard:
                keyboard.add(telebot.types.InlineKeyboardButton(text='1 сессия готова',
                                                                callback_data='ГруппыОК2_' + table_name + '-' + rights))
            else:
                keyboard.add(
                    telebot.types.InlineKeyboardButton(text='Группа готова',
                                                       callback_data='Группа готова1_' + table_name + '-' + rights))
            text = text_for_sessions(text, 1)

        elif session == 2:
            text = text + '*1 сессия*\n'
            text = text_for_sessions(text, 1)
            text = '\n' + text + '*2 сессия*\n'
            for year in session_categories[2]:
                keyboard.add(telebot.types.InlineKeyboardButton(text=year,
                                                                callback_data='year2_' + str(year) + '_' +
                                                                              table_name + '-' + rights))
            if not keyboard.keyboard:
                callback_data = 'ГруппыОК_' + table_name + '-' + rights
                # print('callback_data', callback_data)
                keyboard.add(telebot.types.InlineKeyboardButton(text='Готово',
                                                                callback_data=callback_data))
            else:
                callback_data = 'Группа готова2_' + table_name + '-' + rights
                # print('callback_data', callback_data)
                keyboard.add(
                    telebot.types.InlineKeyboardButton(text='Группа готова',
                                                       callback_data=callback_data))
            text = text_for_sessions(text, 2)
    elif sessions_ == 1:
        for year in session_categories[1]:
            keyboard.add(telebot.types.InlineKeyboardButton(text=year,
                                                            callback_data='year_' + str(year) + '_' +
                                                                          table_name + '-' + rights))
        text = text_for_sessions(text, 1)
        if not keyboard.keyboard:
            keyboard.add(telebot.types.InlineKeyboardButton(text='Готово',
                                                            callback_data='ГруппыОК_' + table_name + '-' + rights))
        else:
            keyboard.add(
                telebot.types.InlineKeyboardButton(text='Группа готова',
                                                   callback_data='Группа готова_' + table_name + '-' + rights))
    if not keyboard.keyboard:
        keyboard.add(
            telebot.types.InlineKeyboardButton(text='Готово', callback_data='ГруппыОК_' + table_name + '-' + rights))

    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='markdown')
    close_connect_db(conn)


def createprotocol_years(message, table_name, rights, all=None):
    conn = connect_to_db()
    query = f'SELECT DISTINCT year FROM {table_name} ORDER BY year'
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    query = f"INSERT IGNORE INTO `temp_protocol` SET `table_name` = '{table_name}'"
    cursor = conn.cursor(buffered=True)
    cursor.execute(query)
    conn.commit()
    if all:
        years = ''
        for year in rows:
            years = years + str(year[0]) + ','
        years = years[:-1]
        # print('years', years)
        query = f"UPDATE `temp_protocol` SET `groups`='{years}' WHERE `id_telegram`= {message.chat.id} " \
                f"AND `table_name`='{table_name}'"
        update_db(conn, query)
        close_connect_db(conn)
        return
    data = select_db(conn, None, False, 'temp_protocol', 'groups', table_name="'" + table_name + "'")
    close_connect_db(conn)
    years = []
    groups_list_of_choose_years = []
    for year in rows:
        years.append(str(year[0]))
    if data['groups'] is not None:
        groups_list_of_choose_years = data['groups'].split('|')
        list_of_choose_years = []
        for list_of_years in groups_list_of_choose_years:
            list_of_choose = list_of_years.split(',')
            for i in list_of_choose:
                list_of_choose_years.append(i)
        # print('list_of_choose_years', list_of_choose_years)
        # print('groups_list_of_choose_years', groups_list_of_choose_years)
        for year in list_of_choose_years:
            try:
                years.remove(year)
            except:
                pass

    text = 'Объедините года в общие группы заплывов в порядке их прохождения дистанций. Выберите первый год, затем второй и т.д. ' \
           'для группы. После формирования группы нажмите "Группа готова", и продолжайте пока все года не будут использованы.\n\n'
    if message.text.startswith("Объедините"):
        text = text + '*'
        for number, groups in enumerate(groups_list_of_choose_years, start=1):
            # print('number-group', number, groups)
            text = text + str(number) + ' группа. ' + groups + '\n'
        text = text + '*'
    # print(text)
    keyboard = telebot.types.InlineKeyboardMarkup()
    for year in years:
        keyboard.add(telebot.types.InlineKeyboardButton(text=year,
                                                        callback_data='year_' + str(
                                                            year) + '_' + table_name + '-' + rights))

    if not keyboard.keyboard:
        keyboard.add(
            telebot.types.InlineKeyboardButton(text='Готово', callback_data='ГруппыОК_' + table_name + '-' + rights))
    else:
        keyboard.add(
            telebot.types.InlineKeyboardButton(text='Группа готова',
                                               callback_data='Группа готова_' + table_name + '-' + rights))

    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='markdown')


def women_or_men(message, table_name, rights):
    text = 'Кто первые начинают турнир?'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(
        telebot.types.InlineKeyboardButton(text='Девушки', callback_data='gender|ж_' + table_name + '-' + rights),
        telebot.types.InlineKeyboardButton(text='Юноши', callback_data='gender|м_' + table_name + '-' + rights))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)


def how_many_swimmmers_in_last_swim(message, table_name, rights):
    text = 'Какое минимальное количество участников должно быть в последнем заплыве группы?'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='1', callback_data='lastswim|1_' + table_name + '-' + rights),
                 telebot.types.InlineKeyboardButton(text='2', callback_data='lastswim|2_' + table_name + '-' + rights),
                 telebot.types.InlineKeyboardButton(text='3', callback_data='lastswim|3_' + table_name + '-' + rights))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)


def remove_swimmer(message, table_name, rights):
    name_swimmer = message.text.split()
    first_name = ''
    second_name = ''
    try:
        first_name = name_swimmer[1]
        second_name = name_swimmer[0]
    except:
        text = 'Неправильно ввели данные. Напишите фамилию и имя.'
        bot.send_message(message.chat.id, text)
        bot.register_next_step_handler(message, remove_swimmer, table_name, rights)

    conn = connect_to_db()
    query = f"SELECT `id`, `second_name`, `first_name`, `year`, `command` FROM {table_name} " \
            f"WHERE second_name = '{second_name}' AND first_name = '{first_name}'"
    data = select_db(conn, query, True)
    keyboard = telebot.types.InlineKeyboardMarkup()
    if not data:
        text = 'Участник не найден'
    else:
        if len(data) > 1:
            text = 'Найдено несколько участников по фамилии и имени. Кого из них надо убрать?'
            for swimmers in data:
                keyboard.add(telebot.types.InlineKeyboardButton(text=swimmers[1] + ' ' + swimmers[2] + ', ' +
                                                                     str(swimmers[3]) + ', ' + swimmers[4],
                                                                callback_data='swimremove|' + str(
                                                                    swimmers[0]) + '_' + table_name + '-' + rights))
        else:
            text = 'Найден участник: ' + data[0][1] + ' ' + data[0][2] + ', ' + str(data[0][3]) + ', ' + data[0][4] + \
                   '\nУбрать участника из турнира?'
            keyboard.add(telebot.types.InlineKeyboardButton(text='Убрать ' + data[0][1] + ' ' + data[0][2],
                                                            callback_data='swimremove|' + str(
                                                                data[0][0]) + '_' + table_name + '-' + rights))
            keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                            callback_data='меню тур_' +
                                                                          table_name + '-' + rights))
    bot.send_message(message.chat.id, text, reply_markup=keyboard)
    # print(text)


################## Загрузка результатов

def upload_results(message, table_name, rights, excel=False):
    file_info = bot.get_file(message.document.file_id)  # получение файла из сообщения
    downloaded_file = bot.download_file(file_info.file_path)
    file_name = message.document.file_name

    with open(file_name, 'wb') as new_file:
        new_file.write(downloaded_file)
    if not excel:
        results_to_db(file_name, table_name)
    else:
        results_from_excel(file_name, table_name)

    try:
        os.remove(file_name)
    except:
        pass
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                    callback_data='меню тур_' +
                                                                  table_name + '-' + rights))
    bot.send_message(message.chat.id, "Файл с результатами успешно обработан.", reply_markup=keyboard)
    bot.send_message(360693297, "Загружены результаты в таблицу " + table_name)



############### Добавление помощников

def add_helpers(message, table_name, menu_=''):
    str_encoded = cryptocode.encrypt(table_name, "swim")
    text = 'Выберите пункты меню, которые будут доступны для помощника. Помощникам нельзя добавить следующие пункты меню:\n' \
           '"Сформировать техническую заявку",\n"Изменить/удалить турнир",\n"Добавить помощников в турнир".\nЭти пункты ' \
           'доступны только Вам.' \
           '\n*После выбора нужных пунктов нажмите "Готово"*\n'
    # Отправьте ему код \n`' + str_encoded + '`'
    text_menu = ['1. Текущие результаты\n', '2. Загрузить результаты\n', '3. Заявочный протокол\n',
                 '4. Стартовый протокол\n',
                 '5. Итоговый протокол\n', '6. Медалисты\n', '7. Многоборье по очкам FINA2022\n',
                 '8. Загрузить техническую заявку от команды\n', '', '9. Списки участников\n']
    menu = []
    menu_now = [0, 1, 2, 3, 4, 5, 6, 7, 9]
    if menu_:
        for i in menu_:
            menu.append(int(i))
        for i in menu:
            text = text + text_menu[i]
            if i in menu_now:
                menu_now.remove(i)

    keyboard = telebot.types.InlineKeyboardMarkup()

    for i in menu_now:
        keyboard.add(telebot.types.InlineKeyboardButton(text=text_menu[i],
                                                        callback_data='helper|' + menu_ + str(i) + '(' + table_name))

    keyboard.add(telebot.types.InlineKeyboardButton(text='Готово',
                                                    callback_data='helper|ok-' + menu_ + '(' + table_name))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard, parse_mode='markdown')


def send_message_to_helper(message, table_name, rights):
    text = 'Перешлите своему помощнику или помощникам сообщение ниже. У всех, кому Вы перешлете сообщение, ' \
           'по этому приглашению будут доступны пункты меню, ' \
           'которые Вы выбрали. Если Вы хотите назначить другим помощникам другие привилегии - начните выбор заново.\n' \
           'Сами не присоединяйтесь к турниру по этому коду, иначе Ваши права администратора турнира аннулируются.'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                    callback_data='меню тур_' +
                                                                  table_name + '-' + 'all'))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)
    conn = connect_to_db()
    tournament = select_db(conn, None, False, 'tournaments', 'name', 'date', table_name="'" + table_name + "'")

    close_connect_db(conn)
    str_encoded = cryptocode.encrypt(table_name + '|' + rights, "s@im")
    text = f"Здравствуйте! Вы приглашены для помощи в турнире:\n*{tournament['name']}, " \
           f"{tournament['date'].strftime('%d.%m.%Y')}*\n" \
           f"Перейдите по ссылке в бота: @BeSwimBot , выберите пункт меню 'Присоединиться к турниру', и " \
           f"вставьте код ниже.\n_Он, кстати, скопируется в память устройства, при нажатии на него. " \
           f"Это можно сделать сразу)_\U0001F609\n\n" \
           f"`{str_encoded}`"

    bot.send_message(message.chat.id, text, 'markdown')


def join_to_tournament(message):
    str_decoded = cryptocode.decrypt(message.text, "s@im")
    # print(str_decoded)
    try:
        table_name = str_decoded[:str_decoded.index('|')]
        rights_get = str_decoded[str_decoded.index('|') + 1:]
        conn = connect_to_db()
        tables = select_db(conn, None, False, 'users', 'table_name', 'rights', id_telegram=message.chat.id)
        if tables['table_name']:
            tables_list = tables['table_name'].split('|')
            rights_list = tables['rights'].split('|')
            if table_name in tables_list:
                for n, table_ in enumerate(tables_list):
                    if table_name == table_:
                        rights_list[n] = rights_get
                        break
                table_names = '|'.join(tables_list)
                rights = '|'.join(rights_list)
            else:
                table_names = tables['table_name'] + '|' + table_name
                rights = tables['rights'] + '|' + rights_get
        else:
            table_names = table_name
            rights = rights_get
        update_db(conn, table='users', id_telegram=message.chat.id, table_name=table_names, rights=rights)
        close_connect_db(conn)
        text = 'Вы добавились в турнир, как помощник. Все ваши турниры доступны по кнопке ниже.'
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Турниры в работе', callback_data='Турниры в работе'))
        bot.send_message(message.chat.id, text, reply_markup=keyboard)
    except:
        pass


def delete_tournament(message, table_name, rights):
    conn = connect_to_db()
    data_user = select_db(conn, None, False, 'users', 'table_name', 'rights', id_telegram=message.chat.id)
    # print(data_user)
    tables_names = data_user['table_name'].split('|')
    rights_names = data_user['rights'].split('|')
    # print(tables_names, rights_names)
    for number, name in enumerate(tables_names):
        if name == table_name:
            tables_names.pop(number)
            rights_names.pop(number)
            break
    # print(tables_names, rights_names)
    tables_names_str = '|'.join(tables_names)
    rights_names_str = '|'.join(rights_names)
    # print(tables_names_str, rights_names_str)
    update_db(conn, None, 'users', id_telegram=message.chat.id, table_name=tables_names_str, rights=rights_names_str)
    query = f"DELETE FROM `tournaments` WHERE table_name = '{table_name}' AND id_telegram={message.chat.id}"
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    query = f"DELETE FROM `temp_protocol` WHERE table_name = '{table_name}'"
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    query = f"DROP TABLE IF EXISTS {table_name}"
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    close_connect_db(conn)
    text = 'Турнир удален. Вернитесь в главное меню.'
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
    bot.edit_message_text(text, message.chat.id, message.message_id, reply_markup=keyboard)


def referi_main(message, table_name, rights):
    msg_text = message.text  # сохранение текста от клиента
    conn = connect_to_db()
    insert_db(conn, 'temp_protocol', id_telegram=message.chat.id, table_name=table_name, groups=msg_text)
    close_connect_db(conn)
    bot.send_message(message.chat.id, 'Напишите фамилию и инициалы Рефери турнира')
    bot.register_next_step_handler(message, referi_second, table_name, rights)


def referi_second(message, table_name, rights):
    msg_text = message.text
    conn = connect_to_db()
    query = f"UPDATE temp_protocol SET groups = CONCAT(groups, '|', '{msg_text}') WHERE table_name='{table_name}'"
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    close_connect_db(conn)
    bot.send_message(message.chat.id, 'Напишите фамилию и инициалы Главного секретаря турнира')
    bot.register_next_step_handler(message, referi_secretar, table_name, rights)


def referi_secretar(message, table_name, rights):
    msg_text = message.text
    conn = connect_to_db()
    query = f"UPDATE temp_protocol SET groups = CONCAT(groups, '|', '{msg_text}') WHERE table_name='{table_name}'"
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()
    close_connect_db(conn)
    mess = bot.send_message(message.chat.id, 'Ожидайте, идёт формирование протокола...')
    file_name, caption = create_final_protocol(table_name, True)
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                    callback_data='меню тур_' +
                                                                  table_name + '-' + rights))
    if pdf_create:
        send_pdf(message, file_name, caption, keyboard)

    else:
        with open(file_name, 'rb') as file:
            bot.send_document(message.chat.id, file, caption=caption, parse_mode='markdown')
    bot.delete_message(mess.chat.id, mess.message_id)
    try:
        os.remove(file_name)
    except:
        pass
    bot.send_message(360693297, 'Сформирован итоговый с подписями ' + table_name)

@bot.callback_query_handler(
    func=lambda call: call.data.startswith(calendar_1_callback.prefix)
)
def callback_inline(call: CallbackQuery):
    name, action, year, month, day = call.data.split(calendar_1_callback.sep)
    # Processing the calendar. Get either the date or None if the buttons are of a different type
    date = calendar.calendar_query_handler(
        bot=bot, call=call, name=name, action=action, year=year, month=month, day=day
    )
    # There are additional steps. Let's say if the date DAY is selected, you can execute your code. I sent a message.
    if action == "DAY":
        bot.send_message(
            chat_id=call.from_user.id,
            text=f"Дата турнира: {date.strftime('%d.%m.%Y')}",
            reply_markup=ReplyKeyboardRemove(),
        )
        conn = connect_to_db()
        update_db(conn, None, table="temp_tournament", id_telegram=call.from_user.id, date=date)
        close_connect_db(conn)
        place_of_tour(call.message)
    elif action == "CANCEL":
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
        bot.send_message(
            chat_id=call.from_user.id,
            text="Отмена",
            reply_markup=keyboard,
        )


def write_to_group(table_name, group_id):
    message_text = write_to_group_open_excel(table_name)
    bot.send_message(group_id, '*' + message_text[0] + '\n' + message_text[1] + '*', 'markdown')
    for text in message_text[2:]:
        bot.send_message(group_id, text)


@bot.message_handler(func=lambda message: message.chat.id == -1001759839349)  # call.message.chat.id == -1001759839349
def group_handler(message):
    if message.reply_to_message is not None:
        if 'дорожка' not in message.reply_to_message.text:
            bot.send_message(message.chat.id, 'Ответьте на сообщение, содержащее слово "дорожка"',
                             reply_to_message_id=message.message_id)
            return
        result = message.text
        regex = r'\d{1,2}.\d{2},\d\d'
        pattern = re.compile(regex)
        if pattern.search(result):
            text = message.reply_to_message.text.split('\n')
            text_for_reply, swim_bool = group_write_to_excel_and_db(message.chat.id, text[1], result)
            bot.send_message(message.chat.id, text_for_reply,
                             reply_to_message_id=message.message_id)
            if swim_bool:
                write_to_group('a28_11_2022_360693297', -1001759839349)
        else:
            bot.send_message(message.chat.id, 'Напишите результат по шаблону: мин.сек,сотые',
                             reply_to_message_id=message.message_id)
    else:
        bot.send_message(message.chat.id, 'Пришлите ответ на сообщение с информацией об участнике',
                         reply_to_message_id=message.message_id)
    # bot.delete_message(call.message.chat.id, call.message.message_id)
    # bot.send_message(call.message.chat.id, call.message.text, reply_markup=ForceReply())
    # bot.edit_message_text(call.message.text, call.message.chat.id, call.message.message_id, reply_markup=ForceReply())
    # print(message)


@bot.callback_query_handler(func=lambda call: True)
# func=lambda call: call.message.chat.id == int(groups['design_group_id'])
def query_handler(call):
    # print('call.data', call.data)
    if call.data == 'Турниры в работе':
        tournaments_in_work(call.message)
    elif call.data == 'Архивные турниры':
        tournaments_in_work(call.message, True)
    elif call.data == 'Создать турнир':
        create_new_tournament(call.message)
    elif call.data == 'присоединиться':
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
        bot.edit_message_text("Отправьте код из сообщения с приглашением.", call.message.chat.id,
                              call.message.message_id, reply_markup=keyboard)
        bot.register_next_step_handler(call.message, join_to_tournament)
    elif call.data == 'Назад в меню':
        bot.clear_step_handler(call.message)
        command_start(call.message, True)
    elif call.data == '25м' or call.data == '50м':
        conn = connect_to_db()
        update_db(conn, None, table="temp_tournament", id_telegram=call.message.chat.id, swimming_pool=call.data)
        close_connect_db(conn)
        number_tracks(call.message)
    elif call.data.startswith("tracks"):
        conn = connect_to_db()
        update_db(conn, None, table="temp_tournament", id_telegram=call.message.chat.id, tracks=call.data[-1])
        close_connect_db(conn)
        choose_distances(call.message, '')
    elif call.data.startswith("choose"):
        choosen = call.data[7:]
        conn = connect_to_db()
        distance = []
        for i in range(1, 11):
            query = "SELECT distance_" + str(i) + f" FROM `temp_tournament` WHERE id_telegram={call.message.chat.id}"
            distance = select_db(conn, query)
            # print(distance)
            if distance['distance_' + str(i)] is None:
                break
        for key in distance.keys():
            key = key
        query = f"UPDATE temp_tournament SET {key} = '{choosen}' WHERE id_telegram={call.message.chat.id}"
        # print(query)
        update_db(conn, query)
        close_connect_db(conn)
        choose_distances(call.message, choosen)

    elif call.data == 'Готово дистанции':
        sessions_func(call.message)
    elif call.data.startswith("session"):
        session = int(call.data[-1])
        if session == 1:
            years_swimmers(call.message, 0)
        else:
            years_swimmers(call.message, 2)
    elif call.data.startswith("_a") or call.data.startswith("меню тур"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        tournament_menu(call.message, table_name, rights)
    elif call.data.startswith("pres"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-', 15)]
        rights = call.data[call.data.index('-', 15) + 1:]
        distance = translit(call.data[call.data.index('|') + 1:call.data.index('_')], "ru")
        gender = translit(call.data[call.data.index('pres') + 4:call.data.index('|') - 1], "ru")
        session = call.data[call.data.index('|') - 1:call.data.index('|')]
        file_name, caption = create_pre_results(table_name, distance, gender, session)

        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))

        if pdf_create:
            send_pdf(call.message, file_name, caption, keyboard)
        else:
            with open(file_name, 'rb') as file:
                bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown',
                                  reply_markup=keyboard)
        bot.delete_message(call.message.chat.id, call.message.message_id)
        try:
            os.remove(file_name)
        except:
            pass

    elif call.data.startswith("pre"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        keyboard = telebot.types.InlineKeyboardMarkup()
        conn = connect_to_db()
        categories_, sessions = tournament_year_categories(conn, table_name)

        try:
            session = int(call.data[call.data.index('pre') + 3:call.data.index('_')])
            sessions = 1
        except:
            session = 1

        text = 'Получить текущие результаты можно только по дистанциям и категориям, где добавлены результы.\n'

        if sessions == 1:
            quantity_distances = select_db(conn, None, False, 'tournaments', 'quantity_dist',
                                           table_name="'" + table_name + "'")['quantity_dist']
            for n in range(1, quantity_distances + 1):
                for gender in ['ж', 'м']:
                    years = []
                    cat_bool = False
                    for category in categories_[session]:
                        # cat_bool = False
                        count_time = 0
                        year_1, year_2, name_cat = years_in_categories(category)
                        years.append(int(year_1))
                        years.append(int(year_2))
                        # print(year_1, year_2, name_cat)
                        query = f"SELECT count(*) AS count, distance_{n} FROM {table_name} WHERE distance_{n}_time > 0 " \
                                f"and gender = '{gender}' and (year BETWEEN {year_1} AND {year_2}) GROUP BY distance_{n}"
                        data = select_db(conn, query, False)
                        # print(query)
                        if data:
                            count_time = data['count']
                            distance = data['distance_' + str(n)]
                        query = f"SELECT count(*) AS count FROM {table_name} WHERE distance_{n}_result > 0 " \
                                f"and gender = '{gender}' and (year BETWEEN {year_1} AND {year_2})"
                        count_result = select_db(conn, query, False)['count']
                        # print(query)
                        # print('count', count_time, count_result)
                        if count_time == count_result and count_time != 0 and count_result != 0:
                            cat_bool = True

                        if count_time != 0 and count_result != 0 and count_time > count_result:
                            query = f"INSERT INTO temp_protocol (table_name, groups) " \
                                    f"VALUES ('{table_name}', '{category}_{distance}_{gender}') " \
                                    f"ON DUPLICATE KEY UPDATE groups = " \
                                    f"CONCAT(groups, '|' '{category}_{distance}_{gender}')"
                            cursor = conn.cursor()
                            cursor.execute(query)
                            conn.commit()
                            # years.append(int(year_1))
                            # years.append(int(year_2))
                            cat_bool = True
                            text += '\n*Результаты загружены не полностью в категории:*\n' + category + ', ' + \
                                    ('девушки' if gender == 'ж' else 'юноши') + ', ' + distance + '\n'

                    if cat_bool:
                        if min(years) == 1950:
                            age = str(max(years)) + ' и старше'
                        elif max(years) == 2025:
                            age = str(min(years)) + ' и младше'
                        else:
                            if min(years) != max(years):
                                age = str(min(years)) + '-' + str(max(years))
                            else:
                                age = str(min(years))
                        dist_lat = translit(distance, "ru", reversed=True)
                        gender_lat = translit(gender, "ru", reversed=True)
                        keyboard.add(
                            telebot.types.InlineKeyboardButton(
                                text=distance + ', ' + ('девушки' if gender == 'ж' else 'юноши') + ', ' + age,
                                callback_data='pres' + gender_lat + str(session) + '|' + dist_lat +
                                              '_' + table_name + '-' + rights))
        else:
            text = 'Результаты какой сессии вас интересуют?'
            keyboard.add(telebot.types.InlineKeyboardButton(text='1 сесссия',
                                                            callback_data='pre1_' + table_name + '-' + rights),
                         telebot.types.InlineKeyboardButton(text='2 сесссия',
                                                            callback_data='pre2_' + table_name + '-' + rights))

        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        close_connect_db(conn)
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id, reply_markup=keyboard,
                              parse_mode='markdown')


    elif call.data.startswith("uploadapp_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        upload_app_from_command(call.message, table_name, rights)

    elif call.data.startswith("uploadresults_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        bot.edit_message_text(call.message.text, call.message.chat.id, call.message.message_id, parse_mode='markdown')
        bot.send_message(call.message.chat.id, 'Загрузите файл с результатами. Поддерживается только файл, '
                                               'созданный при формировании стартового протокола '
                                               '"ОРГАНИЗАТОР......xlsx', reply_markup=keyboard)
        bot.register_next_step_handler(call.message, upload_results, table_name, rights)
    elif call.data.startswith("excel_"):
        # print('excel')
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        bot.send_message(call.message.chat.id, 'Загрузите файл с результатами. Поддерживается только файл с названием: '
                                               '"ОРГАНИЗАТОР_Стартовый_(дата).xlsx')
        bot.register_next_step_handler(call.message, upload_results, table_name, rights, True)
    elif call.data.startswith("createapp_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        file_name, caption = create_app(table_name)
        with open(file_name, 'rb') as file:
            bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown')
        try:
            os.remove(file_name)
        except:
            pass
        bot.send_message(360693297, 'Сформирована техническая заявка ' + table_name)
    elif call.data.startswith("createlist_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        file_name, caption = create_list_2(table_name)

        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' + table_name + '-' + rights))

        # with open(file_name, 'rb') as file:
        #     bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown')
        if pdf_create:
            send_pdf(call.message, file_name, caption, keyboard)
        try:
            os.remove(file_name)
        except:
            pass
        bot.send_message(360693297, 'Сформирован заявочный ' + table_name)
    elif call.data.startswith("startprotocol_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        createprotocol_(call.message, table_name, rights)
    elif call.data.startswith("createprotocol_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        text = 'Нужно ли добавлять в протоколе внизу страницы строки с фамилиями главного судьи, ' \
               'рефери и главного секретаря, как на примере на картинке?'

        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Да',
                                                        callback_data='referi_' +
                                                                      table_name + '-' + rights),
                     telebot.types.InlineKeyboardButton(text='Нет',
                                                        callback_data='noreferi_' +
                                                                      table_name + '-' + rights)
                     )
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        bot.edit_message_text(call.message.text, call.message.chat.id, call.message.message_id, reply_markup=None)
        image = open('footer.png', 'rb')
        bot.send_photo(call.message.chat.id, image, text, reply_markup=keyboard)
    elif call.data.startswith("referi_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]

        bot.delete_message(call.message.chat.id, call.message.message_id)
        mess = bot.send_message(call.message.chat.id, 'Напишите фамилию и инициалы Главного судьи турнира')
        bot.register_next_step_handler(call.message, referi_main, table_name, rights)

    elif call.data.startswith("noreferi_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]

        bot.delete_message(call.message.chat.id, call.message.message_id)
        mess = bot.send_message(call.message.chat.id, 'Ожидайте, идёт формирование протокола...')
        file_name, caption = create_final_protocol(table_name, False)
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        if pdf_create:
            send_pdf(call.message, file_name, caption, keyboard)

        else:
            with open(file_name, 'rb') as file:
                bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown')
        bot.delete_message(mess.chat.id, mess.message_id)
        try:
            os.remove(file_name)
        except:
            pass
        bot.send_message(360693297, 'Сформирован итоговый ' + table_name)
    elif call.data.startswith("medals_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        file_name, caption = create_medals(table_name, 'медалисты.xlsx')
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))

        if pdf_create:
            send_pdf(call.message, file_name, caption, keyboard)
        else:
            with open(file_name, 'rb') as file:
                bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown')
        bot.delete_message(call.message.chat.id, call.message.message_id)
        try:
            os.remove(file_name)
        except:
            pass
        bot.send_message(360693297, 'Медалисты ' + table_name)
    elif call.data.startswith("points_"):
        # print('call.data', call.data)
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        conn = connect_to_db()
        data = select_db(conn, None, False, 'tournaments', '*', table_name="'" + table_name + "'")
        close_connect_db(conn)
        keyboard = telebot.types.InlineKeyboardMarkup()
        for key, value in data.items():
            if key.startswith('distance') and value is not None:
                callback = 'dp|' + key[key.index('_') + 1:] + '/' + value + '_' + table_name + '-' + rights
                keyboard.add(
                    telebot.types.InlineKeyboardButton(text=value, callback_data=callback))
        text = 'Выберите дистанции для подсчета очков в многоборье.\n'
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id, reply_markup=keyboard)

    elif call.data.startswith("dp"):
        # print('call.data', call.data)
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        distance = call.data[call.data.index('/') + 1:call.data.index('_')] + '\n'
        distance_number = call.data[call.data.index('|') + 1:call.data.index('/')]

        if distance_number == 'ok':
            dist = ''
            for index, i in enumerate(call.message.text):
                if i == '\n':
                    dist = call.message.text[index + 1:]
                    break
            distances = dist.split('\n')
            # tournament_menu(call.message, table_name, rights)
            file_name, caption = create_points(table_name, 'многоборье.xlsx', distances)
            keyboard = telebot.types.InlineKeyboardMarkup()
            keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                            callback_data='меню тур_' +
                                                                          table_name + '-' + rights))

            if pdf_create:
                send_pdf(call.message, file_name, caption, keyboard)
            else:
                with open(file_name, 'rb') as file:
                    bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown')
            bot.delete_message(call.message.chat.id, call.message.message_id)
            try:
                os.remove(file_name)
            except:
                pass
            bot.send_message(360693297, 'Многоборье ' + table_name)
            return
        conn = connect_to_db()
        data = select_db(conn, None, False, 'tournaments', '*', table_name="'" + table_name + "'")
        close_connect_db(conn)
        keyboard = telebot.types.InlineKeyboardMarkup()
        text = call.message.text + '\n' + distance
        for key, value in data.items():
            if key.startswith('distance') and value is not None and value not in text:
                callback = 'dp|' + key[key.index('_') + 1:] + '/' + value + '_' + table_name + '-' + rights
                keyboard.add(
                    telebot.types.InlineKeyboardButton(text=value, callback_data=callback))
        keyboard.add(
            telebot.types.InlineKeyboardButton(text='Готово', callback_data='dp|ok/_' + table_name + '-' + rights))
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id, reply_markup=keyboard)
    elif call.data.startswith("by cat"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        file_name, caption = statistic_by_cat(table_name)
        if pdf_create:
            send_pdf(call.message, file_name, caption, keyboard)
        else:
            with open(file_name, 'rb') as file:
                bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown',
                                  reply_markup=keyboard)

        try:
            os.remove(file_name)
        except:
            pass
        bot.send_message(360693297, 'Список участников по категориям ' + table_name)
        return
    elif call.data.startswith("by com"):
        print('swim by com')
    elif call.data.startswith("by all"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        file_name, caption = statistic_years(table_name)

        if pdf_create:
            send_pdf(call.message, file_name, caption, keyboard)
        else:
            with open(file_name, 'rb') as file:
                bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown',
                                  reply_markup=keyboard)

        try:
            os.remove(file_name)
        except:
            pass
        bot.send_message(360693297, 'Общий список участников ' + table_name)
        return
    elif call.data.startswith("statswim_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        statistic_menu(call.message, table_name, rights)
    elif call.data.startswith("removeswimmer_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        bot.send_message(call.message.chat.id, 'Напишите фамилию и имя участника (Сидоров Иван).',
                         reply_markup=keyboard)
        bot.register_next_step_handler(call.message, remove_swimmer, table_name, rights)
    elif call.data.startswith("swimremove"):
        # print(call.data)
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        id_swimmer = call.data[call.data.index('|') + 1:call.data.index('_')]
        conn = connect_to_db()
        query = f"DELETE FROM {table_name} WHERE id = {id_swimmer}"
        cursor = conn.cursor()
        cursor.execute(query)
        conn.commit()
        cursor.close()
        close_connect_db(conn)
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        bot.send_message(call.message.chat.id, 'Участник удален.', reply_markup=keyboard)
    elif call.data.startswith("changetour_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Изменить название турнира',
                                                        callback_data='changename_' +
                                                                      table_name + '-' + rights))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Изменить возрастные категории',
                                                        callback_data='changecategory_' +
                                                                      table_name + '-' + rights))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Удалить турнир',
                                                        callback_data='deltour_' +
                                                                      table_name + '-' + rights))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        bot.edit_message_text('Выберите действие.\nДоступен только пункт "Удалить турнир", '
                              'остальные находятся в разработке.',
                              call.message.chat.id, call.message.message_id, reply_markup=keyboard)
    elif call.data.startswith("deltour_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        text = 'Для удаления турнира нажмите "Удалить турнир".\n*Внимание! Это действие необратимо. ' \
               'Из базы данных полностью удалятся все данные по этому турниру.*'
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Удалить турнир',
                                                        callback_data='deldeltour_' +
                                                                      table_name + '-' + rights))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data='меню тур_' +
                                                                      table_name + '-' + rights))
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id,
                              reply_markup=keyboard, parse_mode='markdown')
    elif call.data.startswith("deldeltour_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        delete_tournament(call.message, table_name, rights)
    elif call.data.startswith("helper"):
        # print(call.data)
        table_name = call.data[call.data.index('(') + 1:]

        try:
            menu_ = call.data[call.data.index('|') + 1:call.data.index('(')]
        except:
            menu_ = ''
        # print(menu_)
        if menu_.startswith('a'):
            add_helpers(call.message, table_name)
        elif menu_.startswith('ok'):
            rights = menu_[menu_.index('-') + 1:]
            rights_list = sorted(rights)
            rights = ''.join(rights_list)
            # print(rights)
            # print(table_name)
            send_message_to_helper(call.message, table_name, rights)
        else:
            add_helpers(call.message, table_name, menu_)

    elif call.data.startswith("По годам_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        createprotocol_years_2(call.message, table_name, rights)
    elif call.data.startswith("По времени_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        createprotocol_years_2(call.message, table_name, rights, True)
        women_or_men(call.message, table_name, rights)
    elif call.data.startswith("year"):
        # print(call.data)
        choose_year = call.data[call.data.index('_') + 1:call.data.index('_', 6)]
        table_name = call.data[call.data.index('_', 6) + 1:call.data.index('-', 12)]
        session = call.data[call.data.index('year') + 4:call.data.index('_')]
        try:
            session = int(session)
        except:
            pass
        # print(choose_year, session)

        rights = call.data[call.data.index('-', 20) + 1:]
        conn = connect_to_db()
        data = select_db(conn, None, False, 'temp_protocol', 'groups', table_name="'" + table_name + "'")
        # print(data)
        if data['groups'] is None:
            groups = choose_year
        elif session == 1:
            if data['groups'][-1] == '|':
                groups = data['groups'] + choose_year
            else:
                groups = data['groups'] + ',' + choose_year
        elif session == 2:
            if '/' not in data['groups']:
                groups = data['groups'] + '/'
            if data['groups'][-1] == '|' or data['groups'][-1] == '/':
                groups = data['groups'] + choose_year
            else:
                groups = data['groups'] + ',' + choose_year
        else:
            if data['groups'][-1] == '|':
                groups = data['groups'] + choose_year
            else:
                groups = data['groups'] + ',' + choose_year
        update_db(conn, None, 'temp_protocol', table_name="'" + table_name + "'", groups=groups)
        close_connect_db(conn)
        createprotocol_years_2(call.message, table_name, rights, session=session)
    elif call.data.startswith("Группа готова"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        session = call.data[call.data.index('готова') + 6:call.data.index('_')]
        try:
            session = int(session)
        except:
            pass
        conn = connect_to_db()
        data = select_db(conn, None, False, 'temp_protocol', 'groups', table_name="'" + table_name + "'")
        groups = None
        try:
            groups = data['groups'] + '|'
        except:
            pass
        update_db(conn, None, 'temp_protocol', table_name="'" + table_name + "'", groups=groups)
        close_connect_db(conn)
        createprotocol_years_2(call.message, table_name, rights, session=session)
    elif call.data.startswith("ГруппыОК"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        session = call.data[call.data.index('К') + 1:call.data.index('_')]
        try:
            session = int(session)
        except:
            pass
        if session:
            conn = connect_to_db()
            data = select_db(conn, None, False, 'temp_protocol', 'groups', table_name="'" + table_name + "'")
            groups = None
            try:
                groups = data['groups'] + '/'
            except:
                pass
            update_db(conn, None, 'temp_protocol', table_name="'" + table_name + "'", groups=groups)
            close_connect_db(conn)
            createprotocol_years_2(call.message, table_name, rights, session=session)
        else:
            women_or_men(call.message, table_name, rights)
    elif call.data.startswith("gender"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        gender = call.data[call.data.index('|') + 1:call.data.index('|') + 2]
        conn = connect_to_db()
        update_db(conn, None, 'temp_protocol', table_name="'" + table_name + "'", gender=gender)
        close_connect_db(conn)
        how_many_swimmmers_in_last_swim(call.message, table_name, rights)
    elif call.data.startswith("lastswim"):
        bot.edit_message_text('Ожидайте, идёт формирование протокола...', call.message.chat.id, call.message.message_id)
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        last_swim = call.data[call.data.index('|') + 1:call.data.index('|') + 2]
        file_name, caption, file_name_track_session_1, file_name_track_session_2 = \
            create_start_protocol(last_swim, table_name)

        if pdf_create:
            send_pdf(call.message, file_name, caption)
            bot.delete_message(call.message.chat.id, call.message.message_id)
            send_pdf(call.message, file_name_track_session_1, caption="Дорожки для судей на секундомерах")

            if file_name_track_session_2:
                send_pdf(call.message, file_name_track_session_2, caption="Дорожки 2 сессия для судей на секундомерах")

        else:
            with open(file_name, 'rb') as file:
                bot.send_document(call.message.chat.id, file, caption=caption, parse_mode='markdown')

        conn = connect_to_db()
        categories, sessions_ = tournament_year_categories(conn, table_name)
        query = f"DELETE FROM `temp_protocol` WHERE table_name = '{table_name}'"
        cursor = conn.cursor()
        cursor.execute(query)
        conn.commit()
        # update_db(conn, None, 'temp_protocol', table_name="'" + table_name + "'", groups=None, gender=None)
        close_connect_db(conn)
        keyboard = telebot.types.InlineKeyboardMarkup()
        callback_data = 'меню тур_' + table_name + '-' + rights
        call_instruction = 'instr start_' + table_name + '-' + rights
        # print(callback_data)
        keyboard.add(telebot.types.InlineKeyboardButton(text='Показать справку полностью',
                                                        callback_data=call_instruction))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data=callback_data))
        if sessions_ == 2:
            file_name_organization, caption = create_track_protocol_for_organization(file_name_track_session_1)
            with open(file_name_organization, 'rb') as f:
                bot.send_document(call.message.chat.id, f, caption=caption, parse_mode='markdown')
            try:
                os.remove(file_name_organization)
            except:
                pass
            file_name_organization, caption = create_track_protocol_for_organization(file_name_track_session_2)

            with open(file_name_organization, 'rb') as f:
                bot.send_document(call.message.chat.id, f, caption=caption, parse_mode='markdown')
            try:
                os.remove(file_name_organization)
            except:
                pass

        else:
            file_name_organization, caption = create_track_protocol_for_organization(file_name_track_session_1)
            with open(file_name_organization, 'rb') as f:
                bot.send_document(call.message.chat.id, f, caption=caption, parse_mode='markdown')
            try:
                os.remove(file_name_organization)
            except:
                pass
        file_name_organization_start, caption = create_track_protocol_for_organization(file_name)
        with open(file_name_organization_start, 'rb') as f:
            bot.send_document(call.message.chat.id, f, caption=caption, parse_mode='markdown')
        text = f'*Справка. Описание сформированных файлов:*\n\n' \
               f'Файл *Стартовый_....pdf* - файл для отправки участникам и для вывода на печать.\n' \
               f'_(читать далее...)_'
        bot.send_message(call.message.chat.id, text, 'markdown', reply_markup=keyboard)
        try:
            os.remove(file_name_organization_start)
        except:
            pass
        try:
            os.rename(file_name, table_name + '.xlsx')
        except:
            pass
        try:
            os.remove(file_name_track_session_1)
        except:
            pass
        try:
            os.remove(file_name_track_session_2)
        except:
            pass
        bot.send_message(360693297, "Сформирован стартовый " + table_name)
    elif call.data == "instruction":
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Далее', callback_data='instr_menu_tur_1'))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
        text = instruction.instruction_menu('Главное меню')
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id,
                              parse_mode='markdown', reply_markup=keyboard)
    elif call.data == "instr_menu_tur_1":
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад', callback_data='instruction'),
                     telebot.types.InlineKeyboardButton(text='Далее', callback_data='instr_menu_tur_2'))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
        text = instruction.instruction_menu('Меню турнира 1')
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id,
                              parse_mode='markdown', reply_markup=keyboard)
    elif call.data == "instr_menu_tur_2":
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад', callback_data='instr_menu_tur_1'),
                     telebot.types.InlineKeyboardButton(text='Далее', callback_data='instr_menu_tur_3'))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
        text = instruction.instruction_menu('Меню турнира 2')
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id,
                              parse_mode='markdown', reply_markup=keyboard)
    elif call.data == "instr_menu_tur_3":
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад', callback_data='instr_menu_tur_2'))
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню', callback_data='Назад в меню'))
        text = instruction.instruction_menu('Меню турнира 3')
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id,
                              parse_mode='markdown', reply_markup=keyboard)
    elif call.data.startswith("instr start_"):
        table_name = call.data[call.data.index('_') + 1:call.data.index('-')]
        rights = call.data[call.data.index('-') + 1:]
        callback_data = 'меню тур_' + table_name + '-' + rights
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.add(telebot.types.InlineKeyboardButton(text='Назад в меню турнира',
                                                        callback_data=callback_data))
        text = '*Справка. Описание сформированных файлов:*\n\n' \
               'Файл *Стартовый_....pdf* - файл для отправки участникам и для вывода на печать.\n' \
               'Файл *Дорожки_....pdf* для вывода на печать. Используется при необходимости. ' \
               'Для записи результатов с секундомеров судьями на дорожках и для дальнейшего ' \
               'переноса результатов с бумаги в файл excel.\n' \
               '*Файлы, начинающиеся с "ОРГАНИЗАТОР"* - файлы в формате excel. Основные файлы для загрузки ' \
               'результатов в бот. Файлы "ОРГАНИЗАТОР Дорожки" и "ОРГАНИЗАТОР Стартовый" - обладают ' \
               'одинаковыми свойствами, ' \
               'и в работе можно использовать как один, так и другой. Различие лишь в виде исполнения файлов.\n' \
               '"ОРГАНИЗАТОР Дорожки" для удобства сформирован наподобие файла "Для судей на секундомерах" и ' \
               'используется в основном при ручном хронометраже и дальнейшего переноса данных с листов бумаги ' \
               'от судей на дорожках. В файле есть листы по количеству дорожек.\n' \
               '"ОРГАНИЗАТОР Стартовый" выполнен в виде стартового протокола. Удобнее вносить данные при наличии ' \
               'электронного табло.\n' \
               'Сохраните любой из файлов "ОРГАНИЗАТОР" в удобном месте на устройстве и заполняйте его результатами. ' \
               'Учтите, все результаты должны быть заполнены по окончании турнира, пропуски недопустимы. ' \
               'Если у участника нет результата (дисквалификация, неявка или не финишировал), ' \
               'выберите в соответствующей ячейке значение из выпадающего списка.\n' \
               'Если участник плывет лично(вне конкурса), не забудьте указать EXH и его время.\n' \
               'Файл можно загружать в бот сколько угодно раз и в любое время турнира.\n' \
               'Не забывайте сохранить файл перед отправкой в бот!\n' \
               'После первых загруженных результатов в меню турнира появится кнопка "Текущие результаты", где вы ' \
               'сможете сформировать файл с результатами по категориям.\nХорошего вам проведения турнира!'
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id,
                              parse_mode='markdown', reply_markup=keyboard)


# write_to_group('a28_11_2022_360693297', -1001759839349)

bot.infinity_polling()
