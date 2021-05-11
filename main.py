import telebot
import config # config.py
import urllib.request 
from function import gift_generation, result_format # function.py
from os import remove



bot = telebot.TeleBot(config.token)
message_global = None


@bot.message_handler(commands=['help'])
def welcome_help(message):
    text = 'Hi, I only work with excel files, so in order for me to answer, send me an excel file'
    bot.send_message(message.chat.id, text)



@bot.message_handler(content_types=["document"])
def handle_docs(message):

    global message_global
    message_global = message
    document_id = message.document.file_id
    file_info = bot.get_file(document_id)

    #Inline keyboard 
    markup = telebot.types.InlineKeyboardMarkup()
    markup.add(telebot.types.InlineKeyboardButton(text='Import in GIFT', callback_data=1))
    markup.add(telebot.types.InlineKeyboardButton(text='Formatting results', callback_data=2))

    if '.xlsx' in file_info.file_path:
        bot.send_message(message.chat.id, text="What operation do you want to perform?", reply_markup=markup)
    else:
        bot.send_message(message.chat.id, text="Unfortunately, I do not work with this format")


@bot.callback_query_handler(func=lambda call: True)
def query_handler(call):
    
    document_id = message_global.document.file_id
    file_info = bot.get_file(document_id)
    bot.answer_callback_query(callback_query_id=call.id, text='Thanks')

    if '.xlsx' in  file_info.file_path:
        urllib.request.urlretrieve(f'http://api.telegram.org/file/bot{config.token}/{file_info.file_path}', file_info.file_path)
            
        if call.data == '1':
            gift_generation(file_info.file_path)
            file = open('documents/import.txt', 'rb')
            bot.send_message(call.message.chat.id, 'The file in GIFT format')
            bot.send_document(message_global.chat.id, file)


        elif call.data == '2':
            
            try:
                result_format(file_info.file_path)
                file = open('documents/test_result.xlsx', 'rb')
                bot.send_message(call.message.chat.id, 'Your test results')
                bot.send_document(message_global.chat.id, file)
            except AttributeError:
                bot.send_message(call.message.chat.id, 'There is no split character in the question \":\"')
            except IndexError:
                bot.send_message(call.message.chat.id, 'The number of columns in the submitted file does not match')


        bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id)
        remove(file_info.file_path)


if __name__ == '__main__':
    bot.polling(none_stop=True, interval=0)
