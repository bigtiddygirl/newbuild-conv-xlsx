import telebot
import os
import requests
import time

import newconv
import config

bot = telebot.TeleBot(config.token)

@bot.message_handler(commands=['start'])
def start_handler(message):
    try:
        bot.send_message(message.chat.id, "Бот готов к работе")
    except Exception as ex:
        bot.send_message(message.chat.id, "[!] ошибка - {}".format(str(ex)))
        telegram_polling()
    
    
    
@bot.message_handler(content_types=["text"]) 
def first(message):
    bot.send_message(message.chat.id, "Отправь мне excel-таблицу, чтобы создать фид")

@bot.message_handler(content_types=["document"])
def handle_docs_audio(message):
    try:
        save_dir = os.getcwd()
        file_name = message.document.file_name
        file_id = message.document.file_name
        file_id_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_id_info.file_path)
        src = file_name
        with open(save_dir + "/" + src, 'wb') as new_file:
            new_file.write(downloaded)
        bot.send_message(message.chat.id, "[*] Файл добавлен\nСоздаю фид")
        os.rename(file_name, "name.xlsx")
        newconv.converter()

        file_to_send = open("filename.xml")
        bot.send_document(message.chat.id, file_to_send)
        file_to_send.close()
        bot.send_message(message.chat.id, "Когда понадобится создать новый фид нажми --> /start и дождитесь ответа от бота")
        
        os.remove("filename.xml")
        os.remove("name.xlsx")

    except Exception as ex:
        bot.send_message(message.chat.id, "[!] ошибка - {}".format(str(ex)))

def telegram_polling():
    try:
        bot.polling(none_stop=True) 
    except Exception as ex:
        with open("Error.txt", "a") as myfile:
            myfile.write("\r\n<<ERROR polling>>\r\n - {}".format(str(ex)))
        #bot.stop_polling()
        time.sleep(10)
        telegram_polling()

if __name__ == '__main__':
    telegram_polling()


