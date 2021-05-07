import telebot
import openpyxl


def get_celebrations(message):
    a = f'Мероприятия на {message.text}:' + '\n\n'
    bot.send_message(message.from_user.id, a)
    for row in range(1, fd.max_row + 1):
        if message.text[:3].lower() in fd[f"B{row}"].value:
            a = '✧𓆏 ' + fd[f"A{row}"].value + ' 𓆏✧' + '\n\n' + 'Организаван: ' + fd[f"C{row}"].value
            bot.send_message(message.from_user.id, a)
    a = f"Это все мероприятия на {message.text}"
    bot.send_message(message.from_user.id, a)


api_key = '1731472073:AAGToqbwBTnsAgJjX0xIZ-y5gDJ46FP9I2k'
d = openpyxl.load_workbook('pk.xlsx')
bot = telebot.TeleBot(api_key)
fsh = d.get_sheet_names()[0]
fd = d.get_sheet_by_name(fsh)

keyboard = telebot.types.ReplyKeyboardMarkup()
keyboard.row('Январь', 'Февраль', 'Март')
keyboard.row('Апрель', 'Май', 'Июнь')
keyboard.row('Июль', 'Август', 'Сентябрь')
keyboard.row('Октябрь', 'Ноябрь', 'Декабрь')


@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, 'Добрый) Выберите месяц, а я пришлю Вам мероприятия.', reply_markup=keyboard)


@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == 'Январь':
        get_celebrations(message)

    elif message.text == 'Февраль':
        get_celebrations(message)

    elif message.text == 'Март':
        get_celebrations(message)

    elif message.text == 'Апрель':
        get_celebrations(message)

    elif message.text == 'Май':
        get_celebrations(message)

    elif message.text == 'Июнь':
        get_celebrations(message)

    elif message.text == 'Июль':
        get_celebrations(message)

    elif message.text == 'Август':
        get_celebrations(message)

    elif message.text == 'Сентябрь':
        get_celebrations(message)

    elif message.text == 'Октябрь':
        get_celebrations(message)

    elif message.text == 'Ноябрь':
        get_celebrations(message)

    elif message.text == 'Декабрь':
        get_celebrations(message)

    else:
        a = "Введите коректно"
        bot.send_message(message.from_user.id, a)


bot.polling(none_stop=True, interval=0)
