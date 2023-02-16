import telebot
import openpyxl

SKU_dict = {}
wb = openpyxl.load_workbook('Price_MB_070622.xlsx', data_only=True)
sheet = wb.active

#Создаем словарь
for i in range(2, sheet.max_row + 1):
    sku = str(sheet.cell (row = i, column = 1).value)
    name = str(sheet.cell (row = i, column = 4).value)
    drt = str(sheet.cell (row = i, column = 5).value)
    price = str(sheet.cell (row = i, column = 9).value)
    in_stock = str(sheet.cell (row = i, column = 10).value)
    SKU_dict[sku] = [name, drt, in_stock, price]
print ('Словарь создан')

#Запускаем бот
testbot = '5672516358:AAF97zqYoNcF9P9rAkCnLB1yhmoSDzJP4SE'
publicbot = '5818851427:AAGrq9ticMpvXi_tM9cN96mXpJ46-XNmmmo'

bot = telebot.TeleBot(testbot)
print('Бот запущен')

@bot.message_handler(commands=['start'])

def start(message):
    mess = f'<b>Добрый день </b>, <b>{message.from_user.first_name} {message.from_user.last_name}</b>, <b> Теперь Вы можете присылать мне номера деталей по катологу Mercedes, к сожалению пока только по одной штуке, используя латинский алфавит и без пробелов.</b>'
    bot.send_message(message.chat.id, mess, parse_mode='html')

@bot.message_handler()
def parts(message):
    prt = str(message.text)
    if prt in SKU_dict:
        mess = f'<b>Номер запчасти: </b> {prt} \n<b>Наименование: </b> {SKU_dict.get(prt)[0]} \n<b>Признак DRT: </b> {SKU_dict.get(prt)[1]} \n<b>Наличие на складе: </b>{SKU_dict.get(prt)[2]} <b> шт.</b> \n<b>Цена: </b>{SKU_dict.get(prt)[3]} <b>руб. </b>'
    else:
        mess = 'Не корректный номер'
    bot.send_message(message.chat.id, mess, parse_mode='html')



bot.polling(non_stop=True)