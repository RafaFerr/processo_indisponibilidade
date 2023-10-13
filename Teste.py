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

# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

from credencial import segredos
from botcity.plugins.excel import BotExcelPlugin

def main():

    '''
    bot.wait(3000)
    # To get the mouse coordinate.
    x = bot.get_last_x()
    y = bot.get_last_y()
    print(f'The last saved mouse position is: {x}, {y}')
    '''
    bot_excel = BotExcelPlugin()

    # VARIAVEIS DE COMUNICACAO
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
    #bot_excel.read('protocolos.xlsx')

    # COLETANDO DATA PARA PASTA

    data = datetime.today().strftime("%d-%m-%Y")
    data2 = str(data)
    data3 = data2[:10]
    logging.info(f'Coletando a Data: {data3}')
    texto_assunto_protocolo = f'ORDENS CNIB PROTOCOLADAS {data3} - {datetime.now()}'
    texto_email_protocolo = f'Segue em anexo os arquivos que foram protocolados - {data3} - {datetime.now()}.'
    texto_assunto_vazio = f'NENHUMA ORDEM PROTOCOLADA - {data3} - {datetime.now()}'
    texto_email_vazio = f'Nenhuma ordem para a serventia foi encontrada'

    desktop_bot = DesktopBot()
    webbot = WebBot()


    if not desktop_bot.find("barra_status", matching=0.97, waiting_time=10000):
        not_found("barra_status")
    desktop_bot.click()
    desktop_bot.wait(1500)
    # SUBIR A BARRA DE ROLAGEM PARA O INICIO DO FILTRO
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    desktop_bot.page_up()
    # FIM DA SUBIDA DA BARRA
    # INICIO DO PROCESSO PARA IMPRIMIR A PRIMEIRA ORDEM
    print('IMPRIMINDO A PRIMEIRA ORDEM')
    desktop_bot.click_at(x=1000, y=422)  # NOTEBOOK
    desktop_bot.wait(1000)
    desktop_bot.click_at(x=1000, y=422)  # NOTEBOOK
    desktop_bot.wait(1000)
    desktop_bot.control_c()
    num_comunicado = desktop_bot.get_clipboard()
    desktop_bot.wait(1000)
    desktop_bot.right_click_at(x=1000, y=422)  # NOTEBOOK
    desktop_bot.wait(500)
    desktop_bot.type_key('i')
    print('IMPRIMIR')
    desktop_bot.wait(1000)
    desktop_bot.tab()
    desktop_bot.type_key('m')
    desktop_bot.tab()
    desktop_bot.wait(500)
    desktop_bot.type_keys(['alt', 'i'])
    # SALVANDO O ARQUIVO PDF DA FOLHA DE ROSTO
    desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{num_comunicado}')
    desktop_bot.wait(2000)
    desktop_bot.enter()
    print('PDF SALVO')
    desktop_bot.wait(4000)
    # FIM DO PROCESSO DE IMPRESSAO DA PRIMEIRA ORDEM
    if not desktop_bot.find("aba_campos", matching=0.97, waiting_time=10000):
        not_found("aba_campos")
    desktop_bot.click()
    print('ABA CAMPOS PARA COLETAR DADOS')
    desktop_bot.wait(1500)

    # ESTRUTURA DE REPETIÇÃO PARA COLETAR OS DADOS DO QUE FOI PROTOCOLADO E IMPRIMIR AS ORDENS
    while not desktop_bot.find('protocolo_vazio'):
        bot_excel.read('protocolos.xlsx')
        print('ACHOU NUMERO DO PROTOCOLO LIVRO 1')
        if not desktop_bot.find("campo_protocolo", matching=0.97, waiting_time=10000):
            not_found("campo_protocolo")
        desktop_bot.click_relative(27, 35)
        desktop_bot.wait(500)
        desktop_bot.type_keys(['ctrl', 'home'])
        desktop_bot.wait(500)
        desktop_bot.type_keys(['shift', 'end'])
        desktop_bot.wait(500)
        desktop_bot.control_c()
        num_comunicado = desktop_bot.get_clipboard()
        desktop_bot.enter(wait=500, presses=8)
        desktop_bot.control_c()
        num_protocolo = desktop_bot.get_clipboard()
        bot_excel.add_row([data3, num_comunicado, num_protocolo])
        bot_excel.write('protocolos.xlsx')
        desktop_bot.wait(1500)

        if not desktop_bot.find("seta_avançar", matching=0.97, waiting_time=10000):
            not_found("seta_avançar")
        desktop_bot.click()
        print('AVANÇADO')
        desktop_bot.wait(1500)
        if not desktop_bot.find("aba_tabela", matching=0.97, waiting_time=10000):
            not_found("aba_tabela")
        desktop_bot.click()
        print('ABA TABELA')
        desktop_bot.wait(2000)
        if not desktop_bot.find("seta_selecao", matching=0.95, waiting_time=10000):
            not_found("seta_selecao")
        desktop_bot.right_click()
        desktop_bot.wait(2000)
        print('IMPRIMINDO A ORDEM')
        desktop_bot.type_down()
        desktop_bot.type_down()
        desktop_bot.enter()
        desktop_bot.tab()
        desktop_bot.type_key('m')
        desktop_bot.tab()
        desktop_bot.wait(1000)
        print(f'GERANDO PDF DA ORDEM {num_comunicado}')
        desktop_bot.type_keys(['alt', 'i'])
        desktop_bot.wait(1500)
        # SALVANDO O ARQUIVO PDF DA FOLHA DE ROSTO
        desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{num_comunicado}')
        desktop_bot.wait(2000)
        desktop_bot.enter()
        desktop_bot.wait(3000)
        if not desktop_bot.find("aba_campos", matching=0.97, waiting_time=10000):
            not_found("aba_campos")
        desktop_bot.click()
        print('CLIQUE ABA CAMPOS PARA RETORNAR')
        desktop_bot.wait(2500)



    else:
        desktop_bot.wait(2000)
        print('SEM MAIS ORDENS')
        bot_excel.write('protocolos.xlsx')
        print('PLANILHA SALVA')

    





def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()

