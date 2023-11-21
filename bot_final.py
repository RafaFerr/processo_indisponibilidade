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
    webbot = WebBot()
    desktop_bot.click_at(x=949, y=539)

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


    bot_excel = BotExcelPlugin()
    import pandas as pd
    planilha = pd.read_excel(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx", 'Sheet1', keep_default_na=False)

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
    datahora = datetime.now()
    logging.info(f'Coletando a Data: {data3}')
    texto_assunto_protocolo = f'ORDENS CNIB PROTOCOLADAS {data3} - {datetime.now()}'
    texto_email_protocolo = f'Segue em anexo os arquivos que foram protocolados - {data3} - {datetime.now()}.'
    texto_assunto_vazio = f'NENHUMA ORDEM PROTOCOLADA - {data3} - {datetime.now()}'
    texto_email_vazio = f'Nenhuma ordem para a serventia foi encontrada'

    # Configure whether or not to run on headless mode
    webbot.headless = False

    # Uncomment to change the default Browser to Firefox
    webbot.browser = Browser.FIREFOX

    webbot.download_folder_path = r'D:\RPA\cnib\indisponibilidade\ordens'

    # Uncomment to set the WebDriver path
    webbot.driver_path = r"D:\RPA\cnib\indisponibilidade\geckodriver.exe"

    # Setting the extension path
    extension_path = r"D:\RPA\cnib\indisponibilidade\webpki-firefox-ext.xpi"

    # Instalação da Extensão
    webbot.install_firefox_extension(extension=extension_path)
    # Opens the BotCity website.


    webbot.browse("https://www.indisponibilidade.org.br/autenticacao/")
    print('ABRIR SITE')
    logging.info('ACESSO SITE INDISPONIBILIDADE')
    webbot.maximize_window()
    print('MAXIMIXADO')
    desktop_bot.click_at(x=1763, y=808)
    webbot.wait(5000)

    if webbot.find_element('/html/body/section/article', By.XPATH,waiting_time=2000):
        fechar = webbot.find_element('/html/body/section/article/a',By.XPATH)
        fechar.click()
        logging.info('BANNER FECHADO')
    else:
        print('SEM BANNER')
        logging.info('SEM BANNER')
        #webbot.wait(2000)
        pass

    bt_certificado = webbot.find_element('cnib-autenticador', By.CLASS_NAME)
    bt_certificado.click()
    webbot.wait(1500)

    selecionar_certificado = webbot.find_element('#cnib-auth-dropdown > optgroup:nth-child(1) > option:nth-child(2)', By.CSS_SELECTOR)

    selecionar_certificado.click()
    logging.info(f'SELECIONADO {selecionar_certificado}')
    webbot.wait(1500)

    bt_autenticar = webbot.find_element('.cnib-auth-visible > div:nth-child(1) > div:nth-child(2) > button:nth-child(4)', By.CSS_SELECTOR)
    bt_autenticar.click()
    webbot.wait(2500)

    desktop_bot.click_at(x=949, y=539)
    desktop_bot.tab()
    desktop_bot.tab()
    desktop_bot.enter()

    desktop_bot.wait(5000)
    desktop_bot.type_keys(token)
    logging.info('TOKEN')
    desktop_bot.enter()
    webbot.wait(5000)

    def backup_pastas():
        logging.info('INICIADO BACKUP DAS PASTAS E DO ARQUIVO XML')
        app_path = r"C:\Windows\explorer.exe"
        desktop_bot.execute(app_path)
        desktop_bot.wait(2000)
        desktop_bot.maximize_window()
        desktop_bot.wait(1000)
        desktop_bot.click_at(x=949, y=539)
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
        if desktop_bot.find("substituir_arquivos", matching=0.97, waiting_time=10000):
            print('SUBSTITUINDO')
            desktop_bot.enter()
        else:
            pass
        desktop_bot.wait(5000)
        # ENCERRA BACKUP PASTA
        # COMECA BACKUP ORDEM
        desktop_bot.click_at(x=949, y=539)
        desktop_bot.type_keys(['ctrl', 'l'])
        print('BACKUP DE ORDENS')
        desktop_bot.wait(1000)
        desktop_bot.paste(r'D:\RPA\cnib\indisponibilidade\ordens', wait=1000)
        desktop_bot.enter()
        desktop_bot.click_at(x=949, y=539)
        desktop_bot.wait(500)
        desktop_bot.control_a()
        desktop_bot.wait(2000)
        desktop_bot.type_keys(['ctrl', 'x'])
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.kb_type(text=r'D:\RPA\cnib\indisponibilidade\backup_ordens')
        desktop_bot.enter()
        desktop_bot.wait(500)
        desktop_bot.control_v()
        if desktop_bot.find("substituir_arquivos", matching=0.97, waiting_time=10000):
            print('SUBSTITUINDO')
            desktop_bot.enter()
        else:
            pass
        logging.info('FIM DO BACKUP')
        desktop_bot.wait(5000)
        desktop_bot.alt_f4()

    def arquivo_excel():
        app_path = r"C:\Windows\explorer.exe"
        desktop_bot.execute(app_path)
        desktop_bot.wait(2000)
        desktop_bot.maximize_window()
        desktop_bot.wait(1000)
        desktop_bot.type_keys(['ctrl', 'l'])
        desktop_bot.wait(1000)
        desktop_bot.paste(rf'D:\RPA\cnib\indisponibilidade\modelo')
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
        desktop_bot.kb_type(text=rf'D:\RPA\cnib\indisponibilidade')
        desktop_bot.enter()
        desktop_bot.wait(500)
        desktop_bot.control_v()
        desktop_bot.wait(1000)
        desktop_bot.enter()
        desktop_bot.wait(2000)
        desktop_bot.alt_f4()


    print('1 ACESSANDO SITE INDISPONIBILIDADE')




    #ESTA PARTE SO PARA TESTE
    print('3 VISUALIZADO SIM')
    visualizado_sim = webbot.find_element('/html/body/div[1]/div[2]/form/span[1]/select/option[1]', By.XPATH)
    visualizado_sim.click()
    webbot.tab()
    webbot.tab()
    webbot.paste(text='14/11/2023')
    webbot.tab()
    webbot.paste(text='14/11/2023')
    webbot.wait(500)
    bt_pesquisa = webbot.find_element('/html/body/div[1]/div[2]/form/span[5]/button', By.XPATH)
    bt_pesquisa.click()
    webbot.wait(5000)

    #FIM DO TESTE





    #Botão de Baixar XML
    botaoxml = webbot.find_element('button.combo-bot:nth-child(2)',By.CSS_SELECTOR)
    botaoxml.click()
    webbot.wait_for_new_file(path=webbot.download_folder_path, file_extension=".xml", timeout=60000)

    logging.info('BAIXANDO XML')

    webbot.stop_browser()

    #ACESSANDO O SISTEMA REGISTER
    app_path = r"C:\EscribaTeste\Register3\sqlreg.exe"
    #app_path = r"C:\EscribaTeste\Register3\sqlreg.exe"
    desktop_bot.execute(app_path)
    logging.info('ACESSANDO REGISTER')


    desktop_bot.connect_to_app(backend=Backend.UIA, path=app_path)

    #desktop_bot.wait(26000)

    #AGUARDANDO TELA DE LOGIN
    desktop_bot.find( 'acessar_register',matching=0.97, waiting_time=35000)

    #webbot.stop_browser()

    if not desktop_bot.find( "campo_nome", matching=0.97, waiting_time=10000):
        not_found("campo_nome")
    desktop_bot.double_click_relative(100, 7)
    desktop_bot.delete()
    desktop_bot.paste(user_register)
    desktop_bot.enter()
    desktop_bot.paste(senha_register)
    desktop_bot.enter()

    if desktop_bot.find( 'cadastro', waiting_time=20000):
        #Abrir tela de cadastro de indisponibilidade
        desktop_bot.type_keys(['ctrl', 'f6'])
        desktop_bot.wait(1200)
        #Acessar botao do TJ
        if not desktop_bot.find( "bt_TJ", matching=0.97, waiting_time=10000):
            not_found("bt_TJ")
        desktop_bot.click()
        logging.info('ACESSADO BOTAO TJ')
        if desktop_bot.find( 'carregar',matching=0.97, waiting_time=30000):
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
            if desktop_bot.find( 'confirmar_pasta', waiting_time=10000):
                desktop_bot.type_keys(['alt', 's'])
                print('PASTA JA EXISTE')
            else:
                pass
            #desktop_bot.key_esc()
            desktop_bot.wait(1500)
            #CARREGAR ARQUIVO XML
            desktop_bot.type_keys(['alt', 'n'])


            #desktop_bot.right_click_at(x=956, y=552)

            desktop_bot.paste(r'D:\RPA\cnib\indisponibilidade\ordens')
            #Comando para abrir a pasta
            desktop_bot.type_keys(['alt', 'o'])
            desktop_bot.wait(500)
            desktop_bot.type_keys(['shift', 'tab'])
            #desktop_bot.shift_tab()
            desktop_bot.type_keys(['shift', 'end'])
            desktop_bot.type_keys(['alt', 'o'])
            logging.info('AGUARDANDO CARREGAMENTO')
            print('CARREGANDO XML')
            desktop_bot.wait(2000)

            while desktop_bot.find( 'aguardar_import_arquivo', waiting_time=10000):
                desktop_bot.wait(2000)
                print(datetime.now())
            else:
                logging.info('FINALIZADO CARREGAMENTO')
                print('FINALIZANDO CARREGAMENTO')
                desktop_bot.wait(5000)

            #if desktop_bot.find( "atencao", matching=0.97, waiting_time=50000):
            #    not_found("atencao")
            desktop_bot.enter()
            desktop_bot.wait(1000)
            #Processar o arquivo
            desktop_bot.type_keys(['alt', 'p', 'r'])

            desktop_bot.wait(500)
            desktop_bot.type_keys(['alt', 's'])
            print('PROCESSANDO ARQUIVO')
            desktop_bot.wait(1000)
            #Aguardar o processamento do arquivo
            logging.info('PROCESSANDO O XML')
            while desktop_bot.find( 'aguardar_processamento',waiting_time=10000):
                desktop_bot.wait(2000)
                print(datetime.now())

            else:
                print('FINALIZANDO PROCESSAMENTO')
                logging.info('FINALIZADO O PROCESSAMENTO')
                desktop_bot.wait(5000)
        num_comunicado = ''

        if desktop_bot.find( 'protocolar', waiting_time=10000):
            # ESTRUTURA DE REPETIÇÃO PARA PROTOCOLAR
            print('ORDENS ENCONTRADAS PARA PROTOCOLAR')

            i = 0
            while desktop_bot.find( 'protocolar', waiting_time=10000):
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
                if not desktop_bot.find( "barra_status", matching=0.97, waiting_time=10000):
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

                if not desktop_bot.find( "aba_campos", matching=0.97, waiting_time=10000):
                    not_found("aba_campos")
                desktop_bot.click()
                print('ABA CAMPOS PARA COLETAR DADOS')
                desktop_bot.wait(1500)

                avancar = 0

                # ESTRUTURA DE REPETIÇÃO PARA COLETAR OS DADOS DO QUE FOI PROTOCOLADO E IMPRIMIR AS ORDENS

                while not desktop_bot.find('protocolo_vazio2', waiting_time=3000):
                    print('1 entrou no while')
                    emailvazio = 'protocolo_vazio2'

                    bot_excel.read(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx")
                    lista = len(planilha['COMUNICADO'])

                    # print('ACHOU NUMERO DO PROTOCOLO LIVRO 1')
                    if not desktop_bot.find( "campo_protocolo", matching=0.97, waiting_time=3000):
                        not_found("campo_protocolo")
                    # desktop_bot.click_relative(27, 35)
                    desktop_bot.click_relative(65, 89)
                    desktop_bot.wait(500)
                    desktop_bot.type_keys(['ctrl', 'home'])
                    desktop_bot.wait(500)
                    desktop_bot.type_keys(['shift', 'end'])
                    desktop_bot.wait(500)
                    desktop_bot.control_c()
                    num_comunicado = desktop_bot.get_clipboard()
                    print('3 COPIADO NUMERO DO COMUNICADO')

                    if lista == 0:
                        print('4 PLANILHA VAZIA')

                        desktop_bot.enter(wait=100, presses=8)
                        desktop_bot.control_c()
                        num_protocolo = desktop_bot.get_clipboard()
                        num_protocolo = str(num_protocolo)
                        print('5 COPIADO NUMERO DO PROTOCOLO')
                        bot_excel.add_row([data3, num_comunicado, num_protocolo, datahora])
                        bot_excel.write(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx")
                        print('6')

                        if not desktop_bot.find( "aba_tabela_cnib", matching=0.97, waiting_time=10000):
                            not_found("aba_tabela_cnib")
                        desktop_bot.click()

                        desktop_bot.wait(1500)

                        if not desktop_bot.find( "seta_selecao", matching=0.95, waiting_time=10000):
                            not_found("seta_selecao")
                        desktop_bot.right_click()
                        print('7 SETA SELECAO')
                        desktop_bot.wait(2000)
                        print('8 IMPRIMINDO A ORDEM')
                        desktop_bot.type_down()
                        desktop_bot.type_down()
                        desktop_bot.enter()
                        desktop_bot.tab()
                        desktop_bot.type_key('m')
                        desktop_bot.tab()
                        desktop_bot.wait(1000)
                        print(f'9 GERANDO PDF DA ORDEM {num_comunicado}')
                        desktop_bot.type_keys(['alt', 'i'])

                        desktop_bot.wait(1500)
                        # SALVANDO O ARQUIVO PDF DA FOLHA DE ROSTO
                        desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{num_comunicado}')
                        desktop_bot.wait(2000)
                        desktop_bot.enter()
                        desktop_bot.wait(3000)

                        if not desktop_bot.find( "aba_campos", matching=0.97, waiting_time=3000):
                            not_found("aba_campos")
                        desktop_bot.click()
                        desktop_bot.wait(1500)
                        if not desktop_bot.find( "seta_avançar", matching=0.97, waiting_time=3000):
                            not_found("seta_avançar")
                        desktop_bot.click()
                    else:
                        print('PLANILHA COM DADOS')
                        for i in range(lista):
                            print(f'FOR PARA VALIDAR SE COMUNICADO JA EXISTE {i}')
                            ordem = str(planilha['COMUNICADO'][i])

                            print(f'14. VARIAVEL AVANCAR Nº {avancar}')

                            if ordem == num_comunicado:
                                print(f'15. JA EXISTE {ordem} / {num_comunicado} IGUAIS')

                                avancar = 1
                                if not desktop_bot.find( "seta_avançar", matching=0.97, waiting_time=3000):
                                    not_found("seta_avançar")
                                desktop_bot.click()
                                print(f'16. BREAK NO FOR DA PLANILHA DE DADOS {avancar}')
                                break

                            else:
                                print('17. PROXIMO')
                                continue

                        if avancar == 1:
                            if not desktop_bot.find('protocolo_vazio2'):
                                print('AVANÇAR IGUAL A 1')
                                if not desktop_bot.find( "campo_protocolo", matching=0.97, waiting_time=3000):
                                    not_found("campo_protocolo")
                                # desktop_bot.click_relative(27, 35)
                                desktop_bot.click_relative(65, 89)
                                desktop_bot.wait(500)
                                desktop_bot.type_keys(['ctrl', 'home'])
                                desktop_bot.wait(500)
                                desktop_bot.type_keys(['shift', 'end'])
                                desktop_bot.wait(500)
                                desktop_bot.control_c()
                                num_comunicado = desktop_bot.get_clipboard()
                                print('17.1. COPIADO NUMERO DO COMUNICADO')

                                print('17.2  VARIAVEL AVANÇAR AGORA É 1')
                                desktop_bot.enter(wait=100, presses=8)
                                desktop_bot.control_c()
                                print('18.1 COPIADO NUMERO DO PROTOCOLO')
                                num_protocolo = desktop_bot.get_clipboard()
                                num_protocolo = str(num_protocolo)

                                bot_excel.add_row([data3, num_comunicado, num_protocolo, datahora])
                                bot_excel.write(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx")
                                print(f'19.1 PLANILHA SALVA {num_comunicado} / {num_protocolo}')

                                if not desktop_bot.find( "aba_tabela", matching=0.97, waiting_time=3000):
                                    not_found("aba_tabela")
                                desktop_bot.click()
                                desktop_bot.wait(1500)
                                print('20.1 ABA TABELA')
                                if not desktop_bot.find( "seta_selecao", matching=0.95, waiting_time=10000):
                                    not_found("seta_selecao")
                                desktop_bot.right_click()
                                print('21.1 SETA SELECAO')
                                desktop_bot.wait(2000)
                                print('22.1 IMPRIMINDO A ORDEM')
                                desktop_bot.type_down()
                                desktop_bot.type_down()
                                desktop_bot.enter()
                                desktop_bot.tab()
                                desktop_bot.type_key('m')
                                desktop_bot.tab()
                                desktop_bot.wait(1000)
                                print(f'23.1 GERANDO PDF DA ORDEM {num_comunicado}')
                                desktop_bot.type_keys(['alt', 'i'])

                                desktop_bot.wait(1500)
                                # SALVANDO O ARQUIVO PDF DA FOLHA DE ROSTO
                                desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{num_comunicado}')
                                desktop_bot.wait(2000)
                                desktop_bot.enter()
                                desktop_bot.wait(3000)
                                print(f'24.1 PDF SALVO')
                                if not desktop_bot.find( "aba_campos", matching=0.97, waiting_time=3000):
                                    not_found("aba_campos")
                                desktop_bot.click()
                                desktop_bot.wait(1500)
                                if not desktop_bot.find( "seta_avançar", matching=0.97, waiting_time=3000):
                                    not_found("seta_avançar")
                                desktop_bot.click()
                                print(f'25.1 AVANÇADO')
                            else:
                                pass
                        else:
                            print('PARA TUDO')
                            desktop_bot.enter(wait=100, presses=8)
                            desktop_bot.control_c()
                            num_protocolo = desktop_bot.get_clipboard()
                            num_protocolo = str(num_protocolo)
                            bot_excel.add_row([data3, num_comunicado, num_protocolo, datahora])
                            bot_excel.write(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx")
                            print('PLANILHA SALVA')

                            if not desktop_bot.find( "aba_tabela", matching=0.97, waiting_time=3000):
                                not_found("aba_tabela")
                            desktop_bot.click()
                            desktop_bot.wait(1500)

                            if not desktop_bot.find( "seta_selecao", matching=0.95, waiting_time=10000):
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

                            if not desktop_bot.find( "aba_campos", matching=0.97, waiting_time=3000):
                                not_found("aba_campos")
                            desktop_bot.click()
                            desktop_bot.wait(1500)
                            if not desktop_bot.find( "seta_avançar", matching=0.97, waiting_time=3000):
                                not_found("seta_avançar")
                            desktop_bot.click()

                desktop_bot.type_keys(['alt', 's'])
                desktop_bot.wait(5000)
                #desktop_bot.enter(wait=300)
                desktop_bot.key_esc()
                desktop_bot.wait(2000)

                desktop_bot.type_keys(['ctrl', 'r'])
                print('ACESSANDO RECEPCAO DE REGISTRO')
                desktop_bot.wait(2000)

                planilha = pd.read_excel(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx", 'Sheet1',
                                         keep_default_na=False)
                lista2 = len(planilha['COMUNICADO'])
                print(lista2)

                if lista2 >= 0:
                    print('passou do if lista2')
                    for x in range(lista2):
                        desktop_bot.wait(3000)
                        print('FOR PARA IMPRIMIR PROTOCOLO')
                        desktop_bot.type_keys(['alt', 'l'])
                        desktop_bot.wait(500)
                        desktop_bot.type_key('p')
                        desktop_bot.wait(1000)
                        desktop_bot.enter()
                        protocoloregister = str(planilha['PROTOCOLO'][x])
                        protocolocnib = str(planilha['COMUNICADO'][x])
                        desktop_bot.kb_type(text=protocoloregister)
                        desktop_bot.enter()

                        if not desktop_bot.find( "aba_campos_recepcao", matching=0.97, waiting_time=10000):
                            not_found("aba_campos_recepcao")
                        desktop_bot.click()

                        if not desktop_bot.find( 'campo_tipo', waiting_time=2000):
                            print('CAMPO TIPO PREENCHIDO, PROTOCOLO JA IMPRESSO')
                            if not desktop_bot.find( "aba_tabela_recepcao", matching=0.97, waiting_time=2000):
                                not_found("aba_tabela_recepcao")
                            print(f'Protocolo {protocoloregister} já impresso')
                            desktop_bot.click()

                            continue
                        else:
                            print(f'Protocolo {protocoloregister} para Impressao')
                            print('ALTERANDO O TIPO PARA JUDICIAL')
                            desktop_bot.type_keys(['alt', 'a'])
                            desktop_bot.wait(500)
                            desktop_bot.tab(presses=3)
                            desktop_bot.type_key('j')
                            desktop_bot.enter()
                            desktop_bot.tab(presses=2)
                            desktop_bot.type_keys(['alt', 's'])
                            desktop_bot.enter()
                            desktop_bot.wait(1000)

                            if not desktop_bot.find( "bt_imprimir", matching=0.97, waiting_time=10000):
                                not_found("bt_imprimir")
                            desktop_bot.click()

                            print('IMPRIMINDO')
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
                            desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{protocolocnib}_{protocoloregister}')
                            desktop_bot.wait(1000)
                            desktop_bot.enter()
                            print('PROTOCOLO SALVO') # AQUI ENCERRA A IMPRESSAO DOS PROTOCOLOS
                desktop_bot.key_esc()
                desktop_bot.wait(1000)
                desktop_bot.alt_f4()
                #ACESSAR SITE

                # Uncomment to set the WebDriver path
                webbot.driver_path = r"D:\RPA\cnib\indisponibilidade\geckodriver.exe"

                # Setting the extension path
                extension_path = r"D:\RPA\cnib\indisponibilidade\webpki-firefox-ext.xpi"

                # Instalação da Extensão
                webbot.install_firefox_extension(extension=extension_path)
                webbot.browse("https://www.indisponibilidade.org.br/autenticacao/")
                webbot.maximize_window()
                desktop_bot.click_at(x=1763, y=808)
                webbot.wait(5000)

                if webbot.find_element('/html/body/section/article', By.XPATH, waiting_time=5000):
                    fechar = webbot.find_element('/html/body/section/article/a', By.XPATH)
                    fechar.click()
                    webbot.wait(2000)
                else:
                    webbot.wait(2000)
                    pass

                bt_certificado = webbot.find_element('cnib-autenticador', By.CLASS_NAME)
                bt_certificado.click()
                webbot.wait(1500)

                selecionar_certificado = webbot.find_element(
                    '#cnib-auth-dropdown > optgroup:nth-child(1) > option:nth-child(2)', By.CSS_SELECTOR)

                selecionar_certificado.click()
                webbot.wait(1500)

                bt_autenticar = webbot.find_element(
                    '.cnib-auth-visible > div:nth-child(1) > div:nth-child(2) > button:nth-child(4)', By.CSS_SELECTOR)
                bt_autenticar.click()
                webbot.wait(2500)

                desktop_bot.click_at(x=949, y=539)
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
                planilha = BotExcelPlugin().read(r"D:\RPA\cnib\indisponibilidade\protocolos.xlsx").set_nan_as(value='')
                desktop_bot = DesktopBot()
                desktop_bot.wait(1000)
                dados = planilha.as_list()[1:]

                for index, dados in enumerate(dados):
                    prot_cnib = dados[1]
                    prot_cnib = str(prot_cnib)

                    verificar = dados[4]
                    if verificar == '':

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
                        desktop_bot.wait(2000)
                        desktop_bot.type_keys(['alt','n'])
                        desktop_bot.wait(500)
                        desktop_bot.paste(rf'\\safira\REGISTRO\CNIB\{data3}\{prot_cnib}_SITE')
                        desktop_bot.wait(2000)
                        desktop_bot.type_keys(['alt','l'])
                        print('ARQUIVO SALVO')
                        altera = planilha.set_cell(column='E', row=2 + index, value='x')
                        # planilha.remove_row(row=1,sheet='Sheet1')
                        desktop_bot.wait(1500)
                        planilha.write(r'D:\RPA\cnib\indisponibilidade\protocolos.xlsx')
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
                    else:
                        desktop_bot.wait(2000)
                        print('NAO TEM MAIS PARA IMPORTAR')
                        continue
                print('SALVANDO PLANILHA APOS O ELSE')
                planilha.write(r'D:\RPA\cnib\indisponibilidade\protocolos.xlsx')
                desktop_bot.wait(2000)
                webbot.stop_browser()
                print('FECHAR NAVEGADOR')
                desktop_bot.alt_f4()
                print('FECHAR REGISTER')


                #ENVIAR EMAIL DE ORDENS PROTOCOLADAS
                print('ABRIR WEBMAIL')
                desktop_bot.wait(2000)
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
                    desktop_bot.wait(1000)
                    desktop_bot.control_a()
                    desktop_bot.wait(1000)
                    desktop_bot.type_keys(['alt', 'a'])
                    logging.info('ANEXADO ARQUIVOS')

                    while desktop_bot.find('aguardar_anexo',waiting_time=2000):
                        print('AGUARDANDO CARREGAMENTO DOS ANEXOS')
                        desktop_bot.wait(2000)
                    else:
                        pass


                    print('ANEXOS CARREGADOS')
                    # Botão Enviar
                    enviar = webbot.find_element(selector='#rcmbtn115')
                    enviar.click()
                    webbot.wait(2500)
                    # webbot.stop_browser()

        else:
            print('ENVIAR EMAIL - SEM ORDEM PROTOCOLADA')
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
                    text=f'ORDENS ENVIADAS',
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
        webbot.stop_browser()

        #desktop_bot.wait(5000)
        #print('BACKUP PASTAS')
        backup_pastas()
        #desktop_bot.wait(5000)
        #print('BACKUP PASTAS')
        #arquivo_excel()

        # Enviar Msg Via Telegram
        try:
            response = telegram.send_message(
                text=f'BACKUPS REALIZADOS',
                group=grupo_telegram,
                username=[username_telegram]
            )
            print('BACKUP XML')


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

