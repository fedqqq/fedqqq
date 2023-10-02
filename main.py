import openpyxl
import telebot
from telebot import types
import xlsxwriter as xw

i = 0
num_of_reg = 2

API = "6575862399:AAELeNy_CIVdvaSCplEto7rFSI8WnNyOvYQ"
bot = telebot.TeleBot(API)

table = xw.Workbook('Photo.xlsx')
table_sheet = table.add_worksheet()

dates_table = openpyxl.load_workbook("dates.xlsx")
dates_sheet = dates_table.active
open_dates = [dates_sheet['A1'].value, dates_sheet['A2'].value, dates_sheet['A3'].value]

@bot.message_handler(commands=['start'])
def hello(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item1 = types.KeyboardButton("записаться на фотосессию")
    markup.add(item1)
    bot.send_message(message.chat.id,
                     'Привет! Я профессиональный фотограф из Санкт-Петербурга, ниже ты можешь увидеть примеры моих работ)))',
                     reply_markup=markup)
    photo = open('profi-fotograf.jpg', 'rb')
    bot.send_photo(message.chat.id, photo)
    bot.register_next_step_handler(message, user_name)

@bot.message_handler(content_types=['text'])
def user_name(message):
    global i
    text = message.text.lower()
    if text == "записаться на фотосессию":
        if i == num_of_reg:
            bot.send_message(message.chat.id, "Извините, пока мест нет")
            table.close()
        else:
            i += 1
            bot.send_message(message.chat.id, "Расскажи о себе, как тебя зовут, (имя и фамилия)?")
            bot.register_next_step_handler(message, register_user_name)
    else:
        bot.send_message(message.chat.id, "Не понимаю тебя")

def register_user_name(message):
    global i
    name = message.text.lower()
    if i != num_of_reg:
        table_sheet.write(f"A{i}", f"{name}")
    bot.send_message(message.chat.id, f"Отлично, {name}, теперь выбери дату)")
    bot.send_message(message.chat.id,
                     f"Вот возможные варианты:\n 1){open_dates[0]}\n 2){open_dates[1]}\n В ответном сообщении укажите номер подходящей для Вас даты)")
    bot.register_next_step_handler(message, register_date)

def register_date(message):
    global i
    date = int(message.text.lower()) - 1
    if date == 0 or date == 1 or date == 2:
        table_sheet.write(f"B{i}", str(open_dates[date]))
        open_dates.remove(open_dates[date])
        bot.send_message(message.chat.id, f"Cпасибо, вы успешно зарегистрировались!")
    else:
        bot.send_message(message.chat.id, "Вы ошиблись")

bot.polling()