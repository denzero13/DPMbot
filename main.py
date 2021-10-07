from telebot import types, TeleBot
import config
import urllib.request
from os import remove
from Classes import MoodleTestFormation, Gift
from function import purge


bot = TeleBot(config.token)
message_file = None
message_text = None


# Inline keyboard
markup = types.InlineKeyboardMarkup()
markup.add(types.InlineKeyboardButton(text="Import in GIFT. I am work with xlsx", callback_data=1))
markup.add(types.InlineKeyboardButton(text="Formatting results.I am work with csv ", callback_data=2))
markup.add(types.InlineKeyboardButton(text="Send to Cloud. I am work with csv", callback_data=3))
markup.add(types.InlineKeyboardButton(text="Send to Users. I am work with csv", callback_data=4))


@bot.message_handler(commands=["status", "help"])
def welcome_help(message):
    text = "Hi, I work with excel or csv files, so in order for me to answer, send me file." \
           "If you send just text I use it as the collection name in the cloud"
    bot.send_message(message.chat.id, text)


@bot.message_handler(content_types=["document"])
def handle_docs(message):
    global message_file
    message_file = message

    if ".csv" or ".xlsx" in file_info.file_path:
        bot.send_message(message.chat.id, text="What operation do you want to perform?", reply_markup=markup)
    else:
        bot.send_message(message.chat.id, text="Unfortunately, I do not work with this format")


@bot.message_handler(func=lambda message: True, content_types=["text"])
def command_default(message):
    global message_text
    message_text = message.text
    bot.send_message(message.chat.id, f"I named {message.text} your collection"
                                      f"\nMaybe try the help page at /help")


@bot.callback_query_handler(func=lambda call: True)
def query_handler(call):
    try:
        file_info = bot.get_file(message_file.document.file_id)
        bot.answer_callback_query(callback_query_id=call.id, text="Thanks")
        file_format = ""

        if call.data == "1" and ".xlsx" in file_info.file_path:
            urllib.request.urlretrieve(f"http://api.telegram.org/file/bot"
                                       f"{config.token}/{file_info.file_path}", file_info.file_path)
            file = Gift(file_info.file_path)
            file.data_formation()
            file = open("documents/import.txt", "rb")
            bot.send_message(call.message.chat.id, "The file in GIFT format")
            bot.send_document(message_file.chat.id, file)
            file_format = ".xlsx"

        elif ".csv" in file_info.file_path:
            urllib.request.urlretrieve(f"http://api.telegram.org/file/bot"
                                       f"{config.token}/{file_info.file_path}", file_info.file_path)
            file_format = ".csv"

            if call.data == "1":
                bot.send_message(call.message.chat.id, "I do not work with csv")

            elif call.data == "2":
                data = MoodleTestFormation(file_info.file_path)
                data.to_exel()
                file = open("documents/test_result.xlsx", "rb")
                bot.send_message(call.message.chat.id, "Your test results")
                bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id)
                bot.send_document(message_file.chat.id, file)

            elif call.data == "3":
                data = MoodleTestFormation(file_info.file_path)
                bot.send_message(call.message.chat.id, "In process")
                global message_text
                data.to_mongo(str(message_text))
                bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id)
                bot.send_message(call.message.chat.id, "Date loaded in Cloud")

            elif call.data == "4":
                data = MoodleTestFormation(file_info.file_path)
                bot.send_message(call.message.chat.id, "In process")
                data.to_mail()
                bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id)
                bot.send_message(call.message.chat.id, "I sent all message to users")
                file_format = ".html"

        purge("documents", file_format)

    except AttributeError:
        bot.send_message(call.message.chat.id, text="I found error .Please try again")
        bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id)


if __name__ == "__main__":
    bot.polling(none_stop=True, interval=2)