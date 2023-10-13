from datetime import datetime
import logging
logging.basicConfig(level=logging.INFO, filename=r'.\log.txt', format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')
import os
import sys


from botcity.core import Backend
# Import for the Desktop Bot
from botcity.core import DesktopBot
from botcity.plugins.email import BotEmailPlugin
from botcity.plugins.telegram import BotTelegramPlugin
from telebot.apihelper import ApiTelegramException
from botcity.plugins.excel import BotExcelPlugin

# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

from credencial import segredos

def main():


    desktop_bot = DesktopBot()
    bot_excel = BotExcelPlugin()

    #VARIAVEIS DE COMUNICACAO
    token_telegram = segredos.get('segredo_token')
    grupo_telegram = segredos.get('segredo_group')
    username_telegram = segredos.get('segredo_username')
    email_user = segredos.get('segredo_useremail')
    pwd_email = segredos.get('segredo_senhaemail')
    lista_email = segredos.get('segredo_listaemail')
    user_register = segredos.get('segredo_userregister')
    senha_register = segredos.get('segredo_senharegister')
    token = segredos.get('segredo_tokenA3')
    telegram = BotTelegramPlugin(token=token_telegram)

    # COLETANDO DATA PARA PASTA

    data = datetime.today().strftime("%d-%m-%Y")
    data2 = str(data)
    data3 = data2[:10]
    logging.info(f'Coletando a Data: {data3}')
    texto_assunto_protocolo = f'ORDENS CNIB PROTOCOLADAS {data3} - {datetime.now()}'
    texto_email_protocolo = f'Segue em anexo os arquivos que foram protocolados - {data3} - {datetime.now()}.'
    texto_assunto_vazio = f'NENHUMA ORDEM PROTOCOLADA - {data3} - {datetime.now()}'
    texto_email_vazio = f'Nenhuma ordem para a serventia foi encontrada'



    app_path = r"C:\Windows\explorer.exe"

    desktop_bot.execute(app_path)
    desktop_bot.wait(2000)
    desktop_bot.maximize_window()
    desktop_bot.wait(1000)

    desktop_bot.type_keys(['ctrl', 'l'])
    desktop_bot.wait(1000)
    #desktop_bot.delete()
    #desktop_bot.wait(1000)
    desktop_bot.paste(rf'\\safira\REGISTRO\CNIB')
    desktop_bot.wait(2000)
    desktop_bot.enter()
    desktop_bot.type_keys(['ctrl', 'l'])
    desktop_bot.tab(presses=3)
    desktop_bot.wait(500)
    desktop_bot.control_a()
    desktop_bot.wait(2000)
    desktop_bot.type_keys(['ctrl','x'])
    desktop_bot.wait(1000)
    desktop_bot.type_keys(['ctrl', 'l'])
    desktop_bot.wait(1000)
    desktop_bot.kb_type(text=r'\\safira\REGISTRO\BACKUP_CNIB')
    desktop_bot.enter()
    desktop_bot.wait(500)
    desktop_bot.control_v()
    desktop_bot.wait(1000)
    desktop_bot.alt_f4()















def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
