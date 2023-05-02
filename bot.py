import datetime

from botcity.core import Backend
# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    #maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    #execution = maestro.get_execution()

    #print(f"Task ID is: {execution.task_id}")
    #print(f"Task Parameters are: {execution.parameters}")

    desktop_bot = DesktopBot()

    # Conectando com a instância do aplicativo aberto


    # Execute operations with the DesktopBot as desired
    # desktop_bot.control_a()
    # desktop_bot.control_c()
    # value = desktop_bot.get_clipboard()

    webbot = WebBot()


    # Configure whether or not to run on headless mode
    webbot.headless = False

    # Uncomment to change the default Browser to Firefox
    webbot.browser = Browser.FIREFOX

    webbot.download_folder_path = 'C:\RPA\Indisponibilidade\ordens'

    # Uncomment to set the WebDriver path
    webbot.driver_path = "C:\RPA\Indisponibilidade\geckodriver.exe"

    # Setting the extension path
    extension_path = "C:\RPA\Indisponibilidade\webpki-firefox-ext.xpi"

    #Instalação da Extensão
    webbot.install_firefox_extension(extension=extension_path)

    # Opens the BotCity website.
    webbot.browse("https://www.indisponibilidade.org.br/autenticacao/")
    webbot.maximize_window()
    webbot.wait(5000)

    bt_certificado = webbot.find_element('cnib-autenticador',By.CLASS_NAME)
    bt_certificado.click()


    #if not webbot.find( "autenticacao", matching=0.97, waiting_time=10000):
    #    not_found("autenticacao")
    #webbot.click()

    webbot.wait(2000)

    if not desktop_bot.find( "autenticar", matching=0.97, waiting_time=10000):
        not_found("autenticar")
    desktop_bot.click_relative(415, 209)

    if not desktop_bot.find( "confirma_login", matching=0.97, waiting_time=10000):
        not_found("confirma_login")
    desktop_bot.click_relative(235, 216)
    desktop_bot.wait(4000)
    desktop_bot.type_keys('563929')
    desktop_bot.enter()


    #Botão de Baixar XML
    botaoxml = webbot.find_element('button.combo-bot:nth-child(2)',By.CSS_SELECTOR)
    botaoxml.click()
    desktop_bot.wait(1000)
    app_path = r"C:\EscribaTeste\Register3\sqlreg.exe"
    desktop_bot.execute(app_path)
    desktop_bot.connect_to_app(backend=Backend.WIN_32, path=app_path)

    desktop_bot.wait(26000)
    webbot.stop_browser()

    if not desktop_bot.find( "campo_nome", matching=0.97, waiting_time=10000):
        not_found("campo_nome")
    desktop_bot.double_click_relative(100, 7)
    desktop_bot.delete()
    desktop_bot.paste('rafael')
    desktop_bot.enter()
    desktop_bot.paste('0274')
    desktop_bot.enter()
    desktop_bot.wait(10000)

    desktop_bot.type_keys(['ctrl', 'f6'])
    
    if not desktop_bot.find( "bt_TJ", matching=0.97, waiting_time=10000):
        not_found("bt_TJ")
    desktop_bot.click()
    desktop_bot.wait(8000)
    desktop_bot.type_keys(['alt', 'c'])
    desktop_bot.wait(1000)
    desktop_bot.type_keys(['alt', 'n'])
    desktop_bot.paste(r'C:\RPA\Indisponibilidade\ordens')
    desktop_bot.type_keys(['alt', 'o'])
    desktop_bot.wait(1000)
    desktop_bot.shift_tab()
    desktop_bot.type_key('o')
    desktop_bot.type_keys(['alt', 'o'])
    desktop_bot.wait(500)

    if desktop_bot.find("atencao", matching=0.97, waiting_time=50000):
        not_found("atencao")
    desktop_bot.enter()
    desktop_bot.wait(5000)
    desktop_bot.type_keys(['alt', 'p', 'r'])
    desktop_bot.wait(500)
    desktop_bot.type_keys(['alt', 's'])

    while desktop_bot.find('aguardar_processamento'):
        desktop_bot.wait(5000)
        print(datetime.time)
    else:
        desktop_bot.type_keys(['alt', 'c'])
        desktop_bot.type_keys(['alt', 'n'])
        desktop_bot.shift_tab()
        desktop_bot.type_key('o')
        desktop_bot.delete()
        desktop_bot.key_esc()
        desktop_bot.type_keys(['alt', 's'])
        desktop_bot.wait(2000)
        desktop_bot.key_esc()
        desktop_bot.alt_f4()



def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()

