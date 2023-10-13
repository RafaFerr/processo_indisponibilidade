
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


    # Instantiate the plugin
    bot_excel = BotExcelPlugin()
    bot_excel.read('protocolos.xlsx')
    dados = bot_excel.as_list()

    for index, dados in enumerate(dados):
        colunaA = dados[0]
        colunaB= dados[1]
        colunaC = dados[2]

        marco = 'X'
        marco = str(marco)
        bot_excel.set_cell('d',index,marco)
        print(f'{index} - {colunaA} - {colunaB} - {colunaC}')
        #break
    bot_excel.write('protocolos.xlsx')










def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()

