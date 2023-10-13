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
    def setup_images(self):
        # Adiciona imagens se pyinstaller
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            res_path = os.path.join(sys._MEIPASS, "Indisponibilidade", "resources")
            self.add_image("acessar_register", os.path.join(res_path, "acessar_register.png"))
            self.add_image("aguardar_import_arquivo", os.path.join(res_path, "aguardar_import_arquivo.png"))
            self.add_image("aguardar_processamento", os.path.join(res_path, "aguardar_processamento.png"))
            self.add_image("atencao", os.path.join(res_path, "atencao.png"))
            self.add_image("autenticacao", os.path.join(res_path, "autenticacao.png"))
            self.add_image("autenticar", os.path.join(res_path, "autenticar.png"))
            self.add_image("bt_TJ", os.path.join(res_path, "bt_TJ.png"))
            self.add_image("buscar_isencao", os.path.join(res_path, "buscar_isencao.png"))
            self.add_image("cadastro", os.path.join(res_path, "cadastro.png"))
            self.add_image("campo_nome", os.path.join(res_path, "campo_nome.png"))
            self.add_image("carregar", os.path.join(res_path, "carregar.png"))
            self.add_image("confirma_login", os.path.join(res_path, "confirma_login.png"))
            self.add_image("confirmar_pasta", os.path.join(res_path, "confirmar_pasta.png"))
            self.add_image("imprimir", os.path.join(res_path, "imprimir.png"))
            self.add_image("num_comunicado", os.path.join(res_path, "num_comunicado.png"))
            self.add_image("protocolar", os.path.join(res_path, "protocolar.png"))
            self.add_image("selecionar_certificado", os.path.join(res_path, "selecionar_certificado.png"))
            self.add_image("teste", os.path.join(res_path, "teste.png"))

    def action(self, execution=None):
        # Fetch the Activity ID from the task:
        # task = self.maestro.get_task(execution.task_id)
        # activity_id = task.activity_id
        self.setup_images()


    '''
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    # Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()
    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")
    '''




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

    webbot = WebBot()

    # Configure whether or not to run on headless mode
    webbot.headless = False

    # Uncomment to change the default Browser to Firefox
    webbot.browser = Browser.FIREFOX

    webbot.download_folder_path = 'C:\RPA\Indisponibilidade\ordens'

    # Uncomment to set the WebDriver path
    webbot.driver_path = ".\geckodriver.exe"

    # Setting the extension path
    extension_path = ".\webpki-firefox-ext.xpi"

    #Instalação da Extensão
    webbot.install_firefox_extension(extension=extension_path)

    def acessar_site():
        # Opens the BotCity website.
        webbot.driver_path = ".\geckodriver.exe"
        extension_path = ".\webpki-firefox-ext.xpi"
        webbot.install_firefox_extension(extension=extension_path)
        webbot.browse("https://www.indisponibilidade.org.br/autenticacao/")
        webbot.maximize_window()
        webbot.wait(5000)

        if webbot.find_element('/html/body/section/article', By.XPATH,waiting_time=10000):
            fechar = webbot.find_element('/html/body/section/article/a',By.XPATH)
            fechar.click()
            webbot.wait(2000)
        else:
            webbot.wait(2000)
            pass

        bt_certificado = webbot.find_element('cnib-autenticador',By.CLASS_NAME)
        bt_certificado.click()
        webbot.wait(1500)

        selecionar_certificado = webbot.find_element('#cnib-auth-dropdown > optgroup:nth-child(1) > option:nth-child(2)',By.CSS_SELECTOR)
        selecionar_certificado.click()
        webbot.wait(1500)

        bt_autenticar = webbot.find_element('.cnib-auth-visible > div:nth-child(1) > div:nth-child(2) > button:nth-child(4)', By.CSS_SELECTOR)
        bt_autenticar.click()
        webbot.wait(2500)

        desktop_bot.click_at(x=949, y=539)
        #desktop_bot.click()
        desktop_bot.tab()
        desktop_bot.tab()
        desktop_bot.enter()

        desktop_bot.wait(5000)
        desktop_bot.type_keys(token)
        desktop_bot.enter()
        webbot.wait(5000)

    def backup_xml():
        app_path = r"C:\Windows\explorer.exe"

        desktop_bot.execute(app_path)
        desktop_bot.maximize_window()
        desktop_bot.wait(1000)

        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.delete()
        desktop_bot.wait(1000)
        desktop_bot.paste(r'C:\RPA\Indisponibilidade\ordens')
        desktop_bot.wait(1000)
        desktop_bot.enter()
        desktop_bot.tab(presses=3)
        desktop_bot.wait(500)
        desktop_bot.control_a()
        desktop_bot.wait(2000)
        desktop_bot.type_keys(['ctrl','x'])
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.kb_type(text=r'C:\RPA\Indisponibilidade\backup_ordens')
        desktop_bot.enter()
        #desktop_bot.tab(presses=3)
        desktop_bot.wait(500)
        desktop_bot.control_v()
        desktop_bot.wait(1000)
        desktop_bot.alt_f4()

    def backup_pastas():
        app_path = r"C:\Windows\explorer.exe"
        desktop_bot.execute(app_path)
        desktop_bot.wait(2000)
        desktop_bot.maximize_window()
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.paste(rf'\\safira\REGISTRO\CNIB')
        desktop_bot.wait(2000)
        desktop_bot.enter()
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.tab(presses=3)
        desktop_bot.wait(500)
        desktop_bot.control_a()
        desktop_bot.wait(2000)
        desktop_bot.type_keys(['ctrl', 'x'])
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.kb_type(text=r'\\safira\REGISTRO\BACKUP_CNIB')
        desktop_bot.enter()
        desktop_bot.wait(500)
        desktop_bot.control_v()
        desktop_bot.wait(1000)
        desktop_bot.alt_f4()

    def arquivo_excel():
        app_path = r"C:\Windows\explorer.exe"
        desktop_bot.execute(app_path)
        desktop_bot.wait(2000)
        desktop_bot.maximize_window()
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.paste(rf'C:\RPA\Indisponibilidade\modelo')
        desktop_bot.wait(2000)
        desktop_bot.enter()
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.tab(presses=3)
        desktop_bot.wait(500)
        desktop_bot.control_a()
        desktop_bot.wait(2000)
        desktop_bot.type_keys(['ctrl', 'c'])
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.kb_type(text=rf'C:\RPA\Indisponibilidade')
        desktop_bot.enter()
        desktop_bot.wait(500)
        desktop_bot.control_v()
        desktop_bot.wait(1000)
        desktop_bot.enter()
        desktop_bot.wait(2000)
        desktop_bot.alt_f4()


    logging.info('ACESSANDO SITE INDISPONIBILIDADE')
    #acessar_site()

    #ESTA PARTE SO PARA TESTE
    '''
    visualizado_sim = webbot.find_element('/html/body/div[1]/div[2]/form/span[1]/select/option[1]', By.XPATH)
    visualizado_sim.click()
    webbot.tab()
    webbot.tab()
    webbot.paste(text='26/09/2023')
    webbot.tab()
    webbot.paste(text='26/09/2023')
    webbot.wait(500)
    bt_pesquisa = webbot.find_element('/html/body/div[1]/div[2]/form/span[5]/button', By.XPATH)
    bt_pesquisa.click()
    webbot.wait(5000)
    '''
    #FIM DO TESTE




    '''
    #Botão de Baixar XML
    botaoxml = webbot.find_element('button.combo-bot:nth-child(2)',By.CSS_SELECTOR)
    botaoxml.click()
    filepath = webbot.wait_for_new_file(path=webbot.download_folder_path, file_extension=".xml", timeout=60000)
    logging.info('BAIXANDO XML')
    '''
    webbot.stop_browser()
    '''
    #ACESSANDO O SISTEMA REGISTER
    app_path = r"C:\EscribaTeste\Register3\sqlreg.exe"
    #app_path = r"C:\EscribaTeste\Register3\sqlreg.exe"
    desktop_bot.execute(app_path)
    logging.info('ACESSANDO REGISTER')


    desktop_bot.connect_to_app(backend=Backend.UIA, path=app_path)

    #desktop_bot.wait(26000)

    #AGUARDANDO TELA DE LOGIN
    desktop_bot.find('acessar_register',matching=0.97, waiting_time=35000)

    #webbot.stop_browser()

    if not desktop_bot.find( "campo_nome", matching=0.97, waiting_time=10000):
        not_found("campo_nome")
    desktop_bot.double_click_relative(100, 7)
    desktop_bot.delete()
    desktop_bot.paste(user_register)
    desktop_bot.enter()
    desktop_bot.paste(senha_register)
    desktop_bot.enter()
    '''
    if desktop_bot.find('cadastro', waiting_time=20000):
        #Abrir tela de cadastro de indisponibilidade
        desktop_bot.type_keys(['ctrl', 'f6'])
        #Acessar botao do TJ
        if not desktop_bot.find("bt_TJ", matching=0.97, waiting_time=10000):
            not_found("bt_TJ")
        desktop_bot.click()
        if desktop_bot.find('carregar',matching=0.97, waiting_time=30000):
            #CRIANDO PASTA
            logging.info(rf'CRIANDO PASTA_{data3}')
            print('CRIANDO PASTA')
            desktop_bot.type_keys(['alt', 'c'])
            desktop_bot.type_keys(['alt', 'n'])
            desktop_bot.paste(r'\\safira\registro\cnib')
            desktop_bot.wait(500)
            desktop_bot.type_keys(['alt', 'o'])

            desktop_bot.right_click_at(x=956, y=552)
            #desktop_bot.right_click_at(x=663, y=430)
            desktop_bot.type_down()
            desktop_bot.type_down()
            desktop_bot.type_down()
            desktop_bot.wait(1000)
            desktop_bot.enter()
            desktop_bot.wait(1000)
            desktop_bot.enter()
            desktop_bot.wait(1000)
            desktop_bot.kb_type(text=data3)
            desktop_bot.wait(1000)
            desktop_bot.enter()
            if desktop_bot.find('confirmar_pasta', waiting_time=10000):
                desktop_bot.type_keys(['alt', 's'])
                print('PASTA JA EXISTE')
            else:
                pass
            #desktop_bot.key_esc()
            desktop_bot.wait(1500)
            #CARREGAR ARQUIVO XML
            desktop_bot.type_keys(['alt', 'n'])
            logging.info('CARREGANDO ARQUIVO XML')
            print('CARREGANDO XML')
            #desktop_bot.right_click_at(x=956, y=552)

            desktop_bot.paste(r'C:\RPA\Indisponibilidade\ordens')
            #Comando para abrir a pasta
            desktop_bot.type_keys(['alt', 'o'])
            desktop_bot.wait(500)
            desktop_bot.type_keys(['shift', 'tab'])
            #desktop_bot.shift_tab()
            desktop_bot.type_keys(['shift', 'end'])
            desktop_bot.type_keys(['alt', 'o'])
            logging.info('AGUARDANDO CARREGAMENTO')
            desktop_bot.wait(2000)
            while desktop_bot.find('aguardar_import_arquivo', waiting_time=10000):
                desktop_bot.wait(10000)
                print(datetime.now())
            else:
                logging.info('FINALIZADO CARREGAMENTO')
                print('FINALIZANDO CARREGAMENTO')
                desktop_bot.wait(5000)

            #if desktop_bot.find("atencao", matching=0.97, waiting_time=50000):
            #    not_found("atencao")
            desktop_bot.enter()
            #desktop_bot.wait(5000)
            #Processar o arquivo
            desktop_bot.type_keys(['alt', 'p', 'r'])
            print('PROCESSANDO ARQUIVO')
            desktop_bot.wait(500)
            desktop_bot.type_keys(['alt', 's'])
            desktop_bot.wait(2500)
            #Aguardar o processamento do arquivo
            logging.info('PROCESSANDO O XML')
            while desktop_bot.find('aguardar_processamento',waiting_time=10000):
                desktop_bot.wait(10000)
                print(datetime.now())

            else:
                print('FINALIZANDO PROCESSAMENTO')
                logging.info('FINALIZADO O PROCESSAMENTO')
                desktop_bot.wait(5000)
        num_comunicado = ''

        if desktop_bot.find('protocolar', waiting_time=10000):
            # ESTRUTURA DE REPETIÇÃO PARA PROTOCOLAR
            print('ORDENS ENCONTRADAS PARA PROTOCOLAR')

            i = 0
            while desktop_bot.find('protocolar', waiting_time=10000):
                i=i+1
                logging.info(rf'{i} PROTOCOLOS')
                print('INICIANDO PROTOCOLO')

                desktop_bot.right_click_at(x=1000, y=425)  # NOTEBOOK
                desktop_bot.wait(2000)
                desktop_bot.type_key('p')
                desktop_bot.click_at(x=590, y=327)  # NOTEBOOK
                desktop_bot.wait(1000)
                desktop_bot.page_down()
                desktop_bot.page_down()
                desktop_bot.page_down()
                desktop_bot.page_down()
                desktop_bot.wait(1000)
                desktop_bot.type_up()
                desktop_bot.wait(1000)
                print('SELECIONANDO ISENÇÃO')
                desktop_bot.type_keys(['alt', 'l'])
                desktop_bot.wait(7000)

            else:
                print('FIM DO PROTOCOLO ')

                print('FIM DO BACKUP')
                logging.info('----------------')
                logging.info(rf'{i} PROTOCOLOS ENCONTRADOS')
                print('FILTRANDO PROTOCOLOS DO DIA')
                desktop_bot.type_keys(['alt', 'l'])
                desktop_bot.wait(200)
                desktop_bot.type_key('i')
                desktop_bot.enter()
                desktop_bot.kb_type(text=data3)
                desktop_bot.wait(200)
                desktop_bot.enter()
                desktop_bot.kb_type(text=data3)
                desktop_bot.wait(200)
                desktop_bot.enter()

                desktop_bot.wait(4000)
                print('ORDENS A SEREM IMPRESSAS')
                if not desktop_bot.find("barra_status", matching=0.97, waiting_time=10000):
                    not_found("barra_status")
                desktop_bot.click()
                desktop_bot.wait(1500)

                # SUBIR A BARRA DE ROLAGEM PARA O INICIO DO FILTRO
                print('SUBIR ROLAGEM')
                print(datetime.now())
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
                desktop_bot.page_up()
                desktop_bot.page_up()
                desktop_bot.page_up()
                desktop_bot.page_up()
                print(datetime.now())
                # FIM DA SUBIDA DA BARRA
                print('TOPO DA ROLAGEM')

                # INICIO DO PROCESSO PARA IMPRIMIR ORDEM
                print('LIBERANDO CAMPOS')
                logging.info(rf'LIBERANDO CAMPOS')
                desktop_bot.wait(1500)
                desktop_bot.right_click_at(x=1000, y=422)  # NOTEBOOK
                desktop_bot.wait(500)
                desktop_bot.type_key('i')
                desktop_bot.wait(1000)
                desktop_bot.key_esc()
                desktop_bot.key_esc()
                print('ENCERRA LIBERACAO DOS CAMPOS')
                desktop_bot.wait(1000)

                if not desktop_bot.find("aba_campos", matching=0.97, waiting_time=10000):
                    not_found("aba_campos")
                desktop_bot.click()
                print('ABA CAMPOS PARA COLETAR DADOS')
                desktop_bot.wait(1500)

                # ESTRUTURA DE REPETIÇÃO PARA COLETAR OS DADOS DO QUE FOI PROTOCOLADO E IMPRIMIR AS ORDENS
                while not desktop_bot.find('protocolo_vazio'):
                    bot_excel.read('protocolos.xlsx')
                    logging.info(rf'LENDO PLANILHA')

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
                    print('COPIADO NUMERO DO PROT CNIB')
                    desktop_bot.enter(wait=500, presses=8)
                    desktop_bot.control_c()
                    num_protocolo = desktop_bot.get_clipboard()
                    num_protocolo = str(num_protocolo)


                    print('COPIADO NUMERO DO PROT')

                    bot_excel.add_row([data3, num_comunicado, num_protocolo])

                    bot_excel.write('protocolos.xlsx')
                    print('PLANILHA SALVA')

                    logging.info(rf'GRAVADO {num_comunicado} e {num_protocolo}')

                    desktop_bot.wait(1500)
                    if not desktop_bot.find("aba_tabela", matching=0.97, waiting_time=10000):
                        not_found("aba_tabela")
                    desktop_bot.click()
                    print('ABA TABELA')
                    desktop_bot.wait(1500)

                    if not desktop_bot.find("seta_selecao", matching=0.95, waiting_time=10000):
                        not_found("seta_selecao")
                    desktop_bot.right_click()
                    print('SETA SELECAO')
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
                    logging.info(rf'GERADO PDF {num_comunicado} e {num_protocolo}')
                    if not desktop_bot.find("aba_campos", matching=0.97, waiting_time=10000):
                        not_found("aba_campos")
                    desktop_bot.click()
                    print('CLIQUE ABA CAMPOS PARA RETORNAR')
                    desktop_bot.wait(2500)

                    if not desktop_bot.find("seta_avançar", matching=0.97, waiting_time=10000):
                        not_found("seta_avançar")
                    desktop_bot.click()
                    print('AVANÇADO')
                    desktop_bot.wait(1500)
                else:
                    desktop_bot.wait(2000)
                    print('SEM MAIS ORDENS')
                    logging.info('FINALIZADA AS IMPRESSOES')
                    bot_excel.write('protocolos.xlsx')
                    print('PLANILHA SALVA')

            #INICIO DO PROCESSO DE IMPRESSAO DOS PROTOCOLOS
            print('SAINDO DA INDISPONIBILIDADE')
            desktop_bot.type_keys(['alt','s'])
            desktop_bot.wait(5000)
            desktop_bot.key_esc()
            desktop_bot.wait(2000)
            desktop_bot.type_keys(['ctrl', 'r'])
            logging.info('ACESSANDO A RECEPCAO DE REGISTRO')

            desktop_bot.wait(5000)


            #INSTANCIA COM A PLANILHA DAS ORDENS PROTOCOLADAS
            planilha = BotExcelPlugin().read(r"C:\RPA\Indisponibilidade\protocolos.xlsx").set_nan_as(value='')


            dados = planilha.as_list()[1:]
            #ESTRUTURA DE REPETIÇÃO PARA IMPRESSÃO DOS PROTOCOLOS DAS ORDENS
            logging.info('INICIO DAS IMPRESSOES DE PROTOCOLO')
            for index, dados in enumerate(dados):
                prot = dados[2]
                prot_cnib = dados[1]
                prot = str(prot)
                prot_cnib = str(prot_cnib)

                desktop_bot.type_keys(['alt', 'l'])
                desktop_bot.type_key('p')
                desktop_bot.wait(1000)
                desktop_bot.enter()
                print('LANÇANDO PROTOCOLO PARA IMPRIMIR')
                desktop_bot.kb_type(text=prot)

                desktop_bot.wait(500)
                desktop_bot.enter()

                if not desktop_bot.find("imprimir", matching=0.97, waiting_time=10000):
                    not_found("imprimir")
                desktop_bot.click()
                desktop_bot.wait(500)
                desktop_bot.type_key('e')
                desktop_bot.tab()
                desktop_bot.tab()
                desktop_bot.tab()
                desktop_bot.tab()
                desktop_bot.type_key('cn')
                desktop_bot.wait(500)
                desktop_bot.tab()
                desktop_bot.type_key('m')
                desktop_bot.wait(1000)
                desktop_bot.tab()
                desktop_bot.type_keys(['alt', 'i'])
                desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{prot_cnib}_{prot}')
                desktop_bot.wait(1000)
                desktop_bot.enter()
                logging.info(rf'IMPRESSAO DO {prot} COM SUCESSO')
                desktop_bot.wait(2000)

            #ACESSO AO SITE PARA BAIXAR A DOCUMENTAÇÃO DAS ORDENS PROTOCOLADAS
            acessar_site()
            visualizado_sim = webbot.find_element('/html/body/div[1]/div[2]/form/span[1]/select/option[1]', By.XPATH)
            visualizado_sim.click()

            webbot.tab()
            planilha = BotExcelPlugin().read(r"C:\RPA\Indisponibilidade\protocolos.xlsx").set_nan_as(value='')

            desktop_bot.wait(1000)
            dados = planilha.as_list()
            logging.info('INICIO DA IMPRESSAO DA ORDEM NO SITE')

            #ESTUTURA DE REPETIÇÃO PARA BAIXAR AS ORDENS DO SITE DA CNIB
            for index, dados in enumerate(dados):

                dados = planilha.as_list()

                prot_cnib = dados[1]

                prot_cnib = str(prot_cnib)
                webbot.paste(prot_cnib)
                webbot.wait(1000)
                bt_pesquisa = webbot.find_element('/html/body/div[1]/div[2]/form/span[5]/button', By.XPATH)
                bt_pesquisa.click()
                webbot.wait(3000)
                bt_imprimir = webbot.find_element('/html/body/div[1]/div[2]/center/button[1]', By.XPATH)
                bt_imprimir.click()
                webbot.wait(5000)
                desktop_bot.type_key('m')
                webbot.wait(2000)
                desktop_bot.enter()
                desktop_bot.wait(3000)
                desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{prot_cnib}_SITE')
                desktop_bot.wait(2000)
                desktop_bot.enter()
                print('ARQUIVO SALVO')
                logging.info(rf'ARQUIVO SALVO {prot_cnib}')
                #planilha.remove_row(row=1,sheet='Sheet1')
                desktop_bot.wait(2000)

                desktop_bot.wait(3000)
                print('AGUARDANDO PARA FECHAR PAGINA')
                desktop_bot.alt_f4()
                print('ALT F4')
                desktop_bot.wait(5000)
                campo_protocolo = webbot.find_element('//*[@id="ordem-protocolo-numero"]', By.XPATH)
                campo_protocolo.click()
                desktop_bot.wait(3000)
                desktop_bot.control_home()
                desktop_bot.wait(2000)
                # desktop_bot.type_keys(['ctrl', 'home'])
                desktop_bot.type_keys(['shift', 'end'])
                desktop_bot.wait(2000)
                desktop_bot.delete()

            #GRAVANDO A PLANILHA
            planilha.write(r"C:\RPA\Indisponibilidade\protocolos.xlsx")
            # ENVIAR EMAIL DE ORDENS PROTOCOLADAS
            webbot.browse('https://webmail.3ribh.com/')
            webbot.maximize_window()
            webbot.wait(2500)
            logging.info('ACESSO AO WEBMAIL')
            usuario = webbot.find_element(selector='#rcmloginuser')
            usuario.click()
            webbot.paste(email_user)
            senha = webbot.find_element(selector='#rcmloginpwd')
            senha.click()
            webbot.paste(pwd_email)
            webbot.enter()
            webbot.wait(3000)
            if webbot.find_element(selector='#compose-plus'):
                # Botao Criar Email
                criar_email = webbot.find_element(selector='#compose-plus')
                criar_email.click()
                webbot.wait(25000)
                webbot.paste(lista_email)
                webbot.wait(1500)
                webbot.tab()
                webbot.tab()
                webbot.tab()
                webbot.paste(texto_assunto_protocolo)
                webbot.tab()
                webbot.paste(texto_email_protocolo)
                # Botão Anexar
                anexar = webbot.find_element(selector='#compose-attachments > div > button')
                anexar.click()
                webbot.wait(2500)
                desktop_bot.paste(fr'\\192.168.0.5\d\REGISTRO\CNIB\{data3}')
                desktop_bot.wait(500)
                desktop_bot.enter()
                desktop_bot.wait(1000)
                desktop_bot.shift_tab()
                desktop_bot.wait(500)
                desktop_bot.control_a()
                desktop_bot.wait(500)
                desktop_bot.type_keys(['alt', 'a'])
                logging.info('ANEXADO ARQUIVOS')
                webbot.wait(5000)
                print('AGUARDANDO CARREGAMENTO DOS ANEXOS')
                # Botão Enviar
                enviar = webbot.find_element(selector='#rcmbtn115')
                enviar.click()
                webbot.wait(2500)
                # webbot.stop_browser()

                # Enviar Msg Via Telegram
                try:
                    response = telegram.send_message(
                        text=f'ORDENS PROTOCOLADAS',
                        group=grupo_telegram,
                        username=[username_telegram]
                    )

                except ApiTelegramException as e:
                    if e.error_code == 400 and "chat not found" in e.description:
                        # Trate o erro de chat não encontrado aqui
                        print("Erro: Chat não encontrado!")
                        print(f"Erro: {e}")

                    elif e.error_code == 400 and "Bad Request" in e.description:
                        print("Erro: Falta de Requisição")
                        print(f"Erro: {e}")

                    else:
                        pass

        #ESSA CONDIÇÃO CASO NAO TENHA NADA PROTOCOLADO, SO ENVIAR O EMAIL
        else:

            # ENVIAR EMAIL
            webbot.browse('https://webmail.3ribh.com/')
            webbot.maximize_window()
            webbot.wait(2500)
            usuario = webbot.find_element(selector='#rcmloginuser')
            usuario.click()
            webbot.paste(email_user)
            senha = webbot.find_element(selector='#rcmloginpwd')
            senha.click()
            webbot.paste(pwd_email)
            webbot.enter()
            webbot.wait(3000)
            if webbot.find_element(selector='#compose-plus'):
                # Botao Criar Email
                criar_email = webbot.find_element(selector='#compose-plus')
                criar_email.click()
                webbot.wait(25000)
                webbot.paste(lista_email)
                webbot.wait(1500)
                webbot.tab()
                webbot.tab()
                webbot.tab()
                webbot.paste(texto_assunto_vazio)
                webbot.tab()
                webbot.paste((texto_email_vazio))
            webbot.wait(1500)

            # Botão Enviar
            enviar = webbot.find_element(selector='#rcmbtn115')
            enviar.click()
            webbot.wait(2500)
            webbot.stop_browser()



            # Enviar Msg Via Telegram
            try:
                response = telegram.send_message(
                    text=f'NENHUMA ORDEM RECEBIDA',
                    group=grupo_telegram,
                    username=[username_telegram]
                )

            except ApiTelegramException as e:
                if e.error_code == 400 and "chat not found" in e.description:
                    # Trate o erro de chat não encontrado aqui
                    print("Erro: Chat não encontrado!")
                    print(f"Erro: {e}")

                elif e.error_code == 400 and "Bad Request" in e.description:
                    print("Erro: Falta de Requisição")
                    print(f"Erro: {e}")

                else:
                    pass
    print('BACKUP XML')
    backup_xml()
    desktop_bot.wait(5000)
    print('BACKUP PASTAS')
    backup_pastas()
    desktop_bot.wait(5000)
    print('BACKUP PASTAS')
    arquivo_excel()

    # Enviar Msg Via Telegram
    try:
        response = telegram.send_message(
            text=f'BACKUPS REALIZADOS',
            group=grupo_telegram,
            username=[username_telegram]
        )

    except ApiTelegramException as e:
        if e.error_code == 400 and "chat not found" in e.description:
            # Trate o erro de chat não encontrado aqui
            print("Erro: Chat não encontrado!")
            print(f"Erro: {e}")

        elif e.error_code == 400 and "Bad Request" in e.description:
            print("Erro: Falta de Requisição")
            print(f"Erro: {e}")

        else:
            pass




def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()

