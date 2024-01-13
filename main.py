import telebot
import my_token
from openpyxl import load_workbook
from telebot import types

book = load_workbook(filename="Book.xlsx")
genre_sheet = book["Genres"]
user_sheet = book["User-Genre"]

bot = telebot.TeleBot(my_token.telebot_token)
bot.remove_webhook()


@bot.message_handler(commands=['start'])
def start(message):

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)

    support = types.KeyboardButton("Поддержка")
    genres = types.KeyboardButton("Жанры")
    predict = types.KeyboardButton("Получить плейлист!")

    markup.add(support, genres, predict)

    bot.send_message(message.chat.id, f'Добрый день, {message.from_user.first_name}!\n'
                                      f'\nЭтот бот создан для того, чтобы помочь Вам осознать собственное настроение и '
                                      f'на его основе составить музыкальный плейлист, который Вам точно понравится.\n'
                                      f'\nВы можете начать работу с ним прямо сейчас или сначала настроить свои любимые '
                                      f'жанры, чтобы наша подборка понравилась Вам еще сильнее!', reply_markup=markup)
    bot.send_sticker(message.chat.id, 'CAACAgIAAxkBAAELKLZloTEXxO9gsgWaPcTka2f-xskchQAC-BYAAstG0UpJuy3W8rP2XDQE')


def add_user(xlsx_user_id, user_id):
    for i in range(15):
        if i == 0:
            user_sheet.cell(row=xlsx_user_id, column=i+1).value = user_id
        else:
            user_sheet.cell(row=xlsx_user_id, column=i+1).value = 0
    book.save('Book.xlsx')


def find_or_add_user(user_id):
    j = 0
    for row in user_sheet.iter_rows(max_col=1):
        j += 1
        for cell in row:

            if cell.value is None:
                print(f'current user = {j}, id = {user_id}')
                add_user(j, user_id)
                return j

            elif cell.value == user_id:
                print(f'current user = {j}, id = {user_id}')
                return j


def find_user(user_id):
    j = 0
    for row in user_sheet.iter_rows(max_col=1):
        j += 1
        for cell in row:
            if cell.value == user_id:
                print(f'current user = {j}, id = {user_id}')
                return j


def genres_cmd(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    support = types.KeyboardButton("Добавить")
    genres = types.KeyboardButton("Убрать")
    menu = types.KeyboardButton("Вернуться в меню")
    markup.add(support, genres, menu)

    user_id = message.from_user.id
    user = find_or_add_user(user_id)
    print(user)

    user_genres = get_user_genres(user)

    bot.send_message(message.chat.id, f'Вот список ваших любимых жанров: \n'
                                      f'{user_genres}\n'
                                      f'Что бы вы хотели с ними сделать?', reply_markup=markup)


def predict_cmd(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    menu = types.KeyboardButton("Вернуться на главную страницу")
    markup.add(menu)

    bot.send_message(message.chat.id, f'Если Вы пришлете свою фотографию в этот диалог, наша нейронная сеть '
                                      f'сможет определить Ваше настроение и составить плейлист по настроению!',
                     reply_markup=markup)


def return_cmd(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)

    support = types.KeyboardButton("Поддержка")
    genres = types.KeyboardButton("Жанры")
    predict = types.KeyboardButton("Получить плейлист!")

    markup.add(support, genres, predict)

    bot.send_message(message.chat.id, f'Еще раз здравствуйте, {message.from_user.first_name}! 🩵\n'
                                      f'Хотите начать работу с ботом или настроить свои любимые жанры?',
                     reply_markup=markup)

    bot.send_sticker(message.chat.id, 'CAACAgIAAxkBAAELKNFloT5zfTrwuqGw95BKsyz3_ytmxQACHxcAAkuY0Eo26-gZoGwtPDQE')


def support_cmd(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    menu = types.KeyboardButton("Вернуться на главную страницу")
    markup.add(menu)

    bot.send_message(message.chat.id, f'Бот был создан совместно студентами третьего курса Университета ИТМО '
                                      f'Пряничниковым К. С. и Серебренниковой В. В.!🕊 \n'
                                      f'По любому вопросу можете обратиться сюда: @needlessbeating\n'
                                      f'Страница бота в гитхабе: https://github.com/larevies/Music-Bot',
                     reply_markup=markup)
    bot.send_sticker(message.chat.id, "CAACAgIAAxkBAAELK5VlopZR9Q_e89V8bPfoA0jZV-tnbQACShUAApYz0UrkbnOpOGIBNzQE")


def add_cmd(message):
    all_genres = ""
    for i in range(13):
        c = genre_sheet.cell(row=i+1, column=1)
        v = c.value
        all_genres += v
        if i < 12:
            all_genres += ", "

    bot.send_message(message.chat.id, f'Следующие жанры учтены в боте: {all_genres}.\n'
                                      f'\nЧтобы добавить жанр, отправьте его название в следующем сообщении:')

    bot.register_next_step_handler(message, pick_genre)


def find_genre(message):

    i = 1
    for column in genre_sheet.iter_cols(1):
        for cell in column:
            i += 1
            if cell.value == message.text:
                return i
    return 0


def find_first_vacant():
    j = 1
    for _ in user_sheet["A"]:
        j += 1
    print(j)
    return j


def pick_genre(message):

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    genres = types.KeyboardButton("Жанры")
    menu = types.KeyboardButton("Вернуться на главную страницу")

    markup.add(genres, menu)

    user_id = message.from_user.id
    genre = find_genre(message)
    user = find_user(user_id)

    if genre == 0:
        bot.send_message(message.chat.id, "Неверно введен жанр!")
        bot.send_sticker(message.chat.id, "CAACAgIAAxkBAAELK41lopW_7BoLbchvDWDqe9AyCyAungACCxMAAmUaQEv-syxD_8aWvzQE")
        bot.send_message(message.chat.id, "Чтобы начать заново, ответьте \"Жанры\"", reply_markup=markup)

    print(f'user = {user}, genre = {genre}')
    user_sheet.cell(row=user, column=genre).value = 1
    book.save('Book.xlsx')

    user_genres = get_user_genres(user)

    bot.send_message(message.chat.id, f'Успех! \n\n'
                                      f'Жанр \"{message.text}\"'
                                      f'добавлен в любимые жанры. \n\nВаш обновленный список любимых жанров:\n'
                                      f'{user_genres}')
    bot.send_sticker(message.chat.id, "CAACAgIAAxkBAAELK5FlopYFSAABVW40d_c2odXFWWLRJX8AAmAWAALsAAHoSgUP7GaqRHiVNAQ")
    bot.send_message(message.chat.id, "Чтобы начать заново, ответьте \"Жанры\"", reply_markup=markup)


@bot.message_handler(content_types='text')
def message_reply(message):
    if message.text == "Жанры":
        genres_cmd(message)

    elif message.text == "Получить плейлист!":
        predict_cmd(message)

    elif message.text == "Вернуться на главную страницу":
        return_cmd(message)

    elif message.text == "Поддержка":
        support_cmd(message)

    elif message.text == "Добавить":
        add_cmd(message)

    elif message.text == "Убрать":
        remove_cmd(message)


def get_user_genres(user_id):
    user_genres = []
    user_genres_names = ""
    for cell in user_sheet[user_id]:
        user_genres.append(cell.value)

    for i in range(len(user_genres)):
        if i == 0 or user_genres[i] == 0:
            continue
        else:
            user_genres_names += user_sheet.cell(2, i+1).value
            user_genres_names += ", "
    if len(user_genres_names) > 2:
        user_genres_names = user_genres_names[:-2]

    print("user genres")
    print(user_genres)
    print("user genres names")
    print(user_genres_names)

    return user_genres_names


def get_all_genres(sheet):
    col = sheet.col_values(0)
    return col


def remove_cmd(message):

    bot.send_message(message.chat.id, f'Чтобы убрать жанр из избранного, отправьте его название в следующем сообщении:')
    bot.register_next_step_handler(message, remove_genre)


def remove_genre(message):

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    genres = types.KeyboardButton("Жанры")
    menu = types.KeyboardButton("Вернуться на главную страницу")

    markup.add(genres, menu)

    user_id = message.from_user.id
    genre = find_genre(message)
    user = find_user(user_id)

    if genre == 0:
        bot.send_message(message.chat.id, "Неверно введен жанр!")
        bot.send_sticker(message.chat.id, "CAACAgIAAxkBAAELK41lopW_7BoLbchvDWDqe9AyCyAungACCxMAAmUaQEv-syxD_8aWvzQE")
        bot.send_message(message.chat.id, "Чтобы начать заново, ответьте \"Жанры\"", reply_markup=markup)

    print(f'user = {user}, genre = {genre}')
    user_sheet.cell(row=user, column=genre).value = 0
    book.save('Book.xlsx')

    user_genres = get_user_genres(user)

    bot.send_message(message.chat.id, f'Успех! \n\n'
                                      f'Жанр \"{message.text}\" убран из списка избранных жанров.\n'
                                      f'\nВаш обновленный список любимых жанров:\n{user_genres}')
    bot.send_sticker(message.chat.id, "CAACAgIAAxkBAAELK5NlopY3RYNsXY1SqmTy5AqLjcVtTgACpRcAAslE0Uvb7l5uWawiazQE")
    bot.send_message(message.chat.id, "Чтобы начать заново, ответьте \"Жанры\"", reply_markup=markup)




if __name__ == '__main__':
    try:
        bot.polling(none_stop=True)
    except Exception as e:
        print(e)
