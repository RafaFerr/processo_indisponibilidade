import bot
from botcity.core import Backend

from botcity.core import DesktopBot


from botcity.plugins.excel import BotExcelPlugin
def main():


    planilha = BotExcelPlugin().read(r"C:\RPA\Indisponibilidade\protocolos.xlsx").set_nan_as(value='')
    desktop_bot = DesktopBot()
    desktop_bot.wait(1000)
    dados = planilha.as_list()
    print(dados)

    for index, dados in enumerate(dados):
        data3 = '25-09-2023'
        prot = dados[2]
        prot_cnib = dados[1]
        prot = str(prot)
        prot_cnib = str(prot_cnib)

        desktop_bot.type_keys(['alt', 'l'])
        desktop_bot.type_key('p')
        desktop_bot.wait(1000)
        desktop_bot.enter()
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
        desktop_bot.wait(2000)

    






def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()




