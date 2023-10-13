
import logging
logging.basicConfig(level=logging.INFO, filename=r'.\log.txt', format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')



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
    webbot = WebBot()
    desktop_bot = DesktopBot()

    # Configure whether or not to run on headless mode
    webbot.headless = False

    # Uncomment to change the default Browser to Firefox
    webbot.browser = Browser.FIREFOX

    webbot.download_folder_path = '.\ordens'

    # Uncomment to set the WebDriver path
    webbot.driver_path = ".\geckodriver.exe"

    # Setting the extension path
    extension_path = ".\webpki-firefox-ext.xpi"

    # Instalação da Extensão
    webbot.install_firefox_extension(extension=extension_path)

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

    # Opens the BotCity website.
    webbot.browse("https://www.indisponibilidade.org.br/autenticacao/")
    webbot.maximize_window()

    webbot.wait(5000)

    

    if webbot.find_element('/html/body/section/article', By.XPATH, waiting_time=10000):
        fechar = webbot.find_element('/html/body/section/article/a', By.XPATH)
        fechar.click()
        webbot.wait(2000)
    else:
        webbot.wait(2000)
        pass

    bt_certificado = webbot.find_element('cnib-autenticador', By.CLASS_NAME)
    bt_certificado.click()
    webbot.wait(1500)

    selecionar_certificado = webbot.find_element('#cnib-auth-dropdown > optgroup:nth-child(1) > option:nth-child(2)',
                                                 By.CSS_SELECTOR)
    selecionar_certificado.click()
    webbot.wait(1500)

    bt_autenticar = webbot.find_element(
        '.cnib-auth-visible > div:nth-child(1) > div:nth-child(2) > button:nth-child(4)', By.CSS_SELECTOR)
    bt_autenticar.click()
    webbot.wait(2500)

    desktop_bot.click_at(x=949, y=539)
    # desktop_bot.click()
    desktop_bot.tab()
    desktop_bot.tab()
    desktop_bot.enter()

    desktop_bot.wait(5000)
    desktop_bot.type_keys(token)
    desktop_bot.enter()
    webbot.wait(5000)


    visualizado_sim = webbot.find_element('/html/body/div[1]/div[2]/form/span[1]/select/option[1]',By.XPATH)
    visualizado_sim.click()

    webbot.tab()
    planilha = BotExcelPlugin().read(r"C:\RPA\Indisponibilidade\protocolos.xlsx").set_nan_as(value='')
    desktop_bot = DesktopBot()
    desktop_bot.wait(1000)
    dados = planilha.as_list()
    for index, dados in enumerate(dados):
        importado = dados[2]
        importado = str(importado)

        prot_cnib = dados[1]
        prot_cnib = str(prot_cnib)
        if importado =='x':
            print('ACHOU O X')
            print(prot_cnib)
            continue
        else:
            print('SEM X')
            print(prot_cnib)
            ##planilha.add_row([, , num_protocolo])

            planilha.write('protocolos.xlsx')


            webbot.paste(prot_cnib)
            webbot.wait(1000)

            bt_pesquisa = webbot.find_element('/html/body/div[1]/div[2]/form/span[5]/button',By.XPATH)
            bt_pesquisa.click()
            webbot.wait(3000)
            bt_imprimir = webbot.find_element('/html/body/div[1]/div[2]/center/button[1]',By.XPATH)
            bt_imprimir.click()
            webbot.wait(5000)
            desktop_bot.type_key('m')
            webbot.wait(2000)
            desktop_bot.enter()
            desktop_bot.wait(3000)
            desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{prot_cnib}_SITE')
            desktop_bot.wait(2000)
            desktop_bot.enter()
            print('ARQUIVO SALVO')
            #planilha.remove_row(row=1,sheet='Sheet1')
            desktop_bot.wait(3000)
            #planilha = BotExcelPlugin().write(r"C:\RPA\Indisponibilidade\protocolos.xlsx").set_nan_as(value='')
            print('AGUARDANDO PARA FECHAR PAGINA')
            desktop_bot.alt_f4()
            print('ALT F4')
            desktop_bot.wait(5000)
            campo_protocolo = webbot.find_element('//*[@id="ordem-protocolo-numero"]',By.XPATH)
            campo_protocolo.click()
            desktop_bot.wait(3000)
            desktop_bot.control_home()
            desktop_bot.wait(2000)
            #desktop_bot.type_keys(['ctrl', 'home'])
            desktop_bot.type_keys(['shift', 'end'])
            desktop_bot.wait(2000)
            desktop_bot.delete()



def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()




