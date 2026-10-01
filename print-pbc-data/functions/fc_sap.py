from typing import Tuple
from botcity.core import DesktopBot
import win32com.client as win32
from functions.fc_botcity_maestro import BotcityMaestroConnection
from functions.sap_dialog_watcher import (
    start_sap_error_dialog_watcher,
    stop_sap_error_dialog_watcher,
)


def _on_sap_dialog_dismissed(tipo: str, texto: str) -> None:
    try:
        BotcityMaestroConnection().log_sap_dialog_dismissed(tipo, texto)
    except Exception:
        t = (texto or "")[:500]
        print(f"[SAP dialog watcher] {tipo}: {t}")


def login(sapLanguage:str , sso:bool, sapInstance:str = "", maestroCredentialsLabel:str="") -> Tuple[DesktopBot, object]:
    ''' DESCRIPTION

    ***** ENGLISH *****
    Function to login in SAP

    Args:
        maestro: The maestro instance that was created at the beginning of the method will be used to obtain the Sharepoint connection credentials
        maestroCredentialsLabel: Label name in the maestro credentials to obtain the technical user's login and password to authenticate in SAP
        sapInstance: Name of the SAP instance that the bot should connect to
        sapLanguage: Language to be used Ex: EN, PT, ES

    Returns:
        SAPGUI.exe instance and SAP session that is logged in

    
    ***** PORTUGUÊS *****
    Função para login no SAP

    Args:
        maestro: A instância do maestro que foi criada no início do método será usada para obter as credenciais de conexão do Sharepoint
        maestroCredentialsLabel: Nome do rótulo nas credenciais do maestro para obter o login e a senha do usuário técnico para autenticação no SAP
        sapInstance: Nome da instância SAP à qual o bot deve se conectar
        sapLanguage: Idioma a ser usado

    Returns:
        Instância SAPGUI.exe e sessão SAP que está logada


    **** ESPAÑOL ****
    Función para iniciar sesión en SAP

    Args:
        maestro: La instancia de Maestro creada al inicio del método se usará para obtener las credenciales de conexión de SharePoint.
        maestroCredentialsLabel: Nombre de la etiqueta en las credenciales de Maestro para obtener el nombre de usuario y la contraseña del usuario técnico para autenticarse en SAP.
        sapInstance: Nombre de la instancia de SAP a la que el bot debe conectarse.
        sapLanguage: Idioma a utilizar.

    Returns:
        Instancia SAPGUI.exe y sesión de SAP que está iniciada
    '''
    start_sap_error_dialog_watcher(on_dismiss=_on_sap_dialog_dismissed)

    bot = DesktopBot()

    sapLogonFilePath = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    bot.execute(sapLogonFilePath)
    bot.wait(5000)

    SapGuiAuto = win32.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    
    if sso:
        if sapInstance == "":
            maestro_conn = BotcityMaestroConnection()
            credentials = maestro_conn.get_credentials(maestroCredentialsLabel)["instance"]

        connection = application.OpenConnection(sapInstance, True)
        session = connection.Children(0)

    else:
        maestro_conn = BotcityMaestroConnection()
        credentials = maestro_conn.get_credentials(maestroCredentialsLabel)
        sapLogin = credentials["login"]
        sapPassword = credentials["password"]
        sapInstance = credentials["instance"]

        connection = application.OpenConnection(sapInstance, True)
        session = connection.Children(0)        

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = sapLogin
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = sapPassword
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = sapLanguage
        session.findById("wnd[0]").sendVKey(0)

    #Check window for a multiple logins
    try:
        session.findById('wnd[1]/usr/radMULTI_LOGON_OPT2').select()
        session.findById('wnd[1]/usr/radMULTI_LOGON_OPT2').setFocus()
        session.findById('wnd[1]/tbar[0]/btn[0]').press()
    except:
        pass

    #Check window for a Copyright message
    try:
        session.findById('wnd[1]/tbar[0]/btn[0]').press()
    except:
        pass        

    return bot, session

def logout():
    stop_sap_error_dialog_watcher()
    bot = DesktopBot()
    process = bot.find_process("saplogon.exe")
    if process != None:
        bot.terminate_process(process)