import botcity.web
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
    # maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    # execution = maestro.get_execution()

    # print(f"Task ID is: {execution.task_id}")
    # print(f"Task Parameters are: {execution.parameters}")




    desktop_bot = DesktopBot()
    desktop_bot.wait(2000)
    '''
    desktop_bot.type_keys(['alt', 'c'])
    desktop_bot.wait(1000)
    desktop_bot.type_keys(['alt', 'n'])
    desktop_bot.paste(r'C:\RPA\Indisponibilidade\ordens')
    desktop_bot.type_keys(['alt' ,'o'])
    desktop_bot.wait(1000)
    desktop_bot.shift_tab()
    desktop_bot.type_key('o')
    desktop_bot.type_keys(['alt', 'o'])
    desktop_bot.wait(500)

    if desktop_bot.find( "atencao", matching=0.97, waiting_time=50000):
        not_found("atencao")
    desktop_bot.enter()
    desktop_bot.wait(5000)
    desktop_bot.type_keys(['alt', 'p', 'r'])
    desktop_bot.wait(500)
    desktop_bot.type_keys(['alt', 's'])
    desktop_bot.wait(7000)

    desktop_bot.type_keys(['alt', 'c'])
    desktop_bot.type_keys(['alt', 'n'])
    desktop_bot.shift_tab()
    desktop_bot.type_key('o')
    desktop_bot.delete()
    desktop_bot.key_esc()
    desktop_bot.type_keys(['alt', 's'])
    desktop_bot.key_esc()
    desktop_bot.alt_f4()
    '''

    #ele = desktop_bot.get_element_coords('teste', matching=0.97)
    #print(ele)
    '''
    if desktop_bot.find( 'protocolar',waiting_time=5000):
        desktop_bot.click_at(x=1000, y=425)
        desktop_bot.wait(1000)
        desktop_bot.click_at(x=1000, y=425)
        desktop_bot.control_c()
        num_comunicado = desktop_bot.get_clipboard()
        print(num_comunicado)
        desktop_bot.wait(2000)
        desktop_bot.right_click_at(x=650, y=425)
        desktop_bot.wait(2000)
        desktop_bot.type_key('i')
        desktop_bot.wait(1000)
        desktop_bot.tab()
        desktop_bot.type_key('m')
        desktop_bot.tab()
        desktop_bot.wait(500)
        desktop_bot.type_keys(['alt','i'])
        desktop_bot.paste(f'C:\RPA\Indisponibilidade\ordens\impresso\{num_comunicado}')
        desktop_bot.wait(2000)
        desktop_bot.enter()
        desktop_bot.wait(3000)
        desktop_bot.right_click_at(x=650, y=425)
        desktop_bot.wait(1000)
        desktop_bot.type_key('p')
        desktop_bot.wait(3000)
        '''
    desktop_bot.wait(2000)
    if not desktop_bot.find( "buscar_isencao", matching=0.97, waiting_time=10000):
        not_found("buscar_isencao")
    desktop_bot.click_relative(149, -29)
    desktop_bot.wait(1500)
    desktop_bot.click(clicks=36,interval_between_clicks=2000)
    input()
    desktop_bot.wait(1000)


    
    #if not desktop_bot.find( "buscar_isencao", matching=0.97, waiting_time=10000):
    #    not_found("buscar_isencao")
    #desktop_bot.click_relative(147, -29)
    

    
        
        





def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()


