from botcity import maestro
from botcity.core import DesktopBot
from datetime import datetime
from botcity.maestro import *
import win32com.client as win32
import os
import time
from functions.fc_office365 import fc_email
from functions.fc_office365 import fc_sharepoint
from functions.fc_office365 import fc_teams
from functions.fc_office365 import fc_teams_enumMessageType
from functions import fc_utils
from functions import fc_sap
from functions.fc_utils  import now
from functions.fc_execution_log import ExecutionLogConnection
from functions.fc_botcity_maestro import BotcityMaestroConnection

def main():
    bot = DesktopBot() 
    try:
        
        maestro_conn = BotcityMaestroConnection()
        log_conn = ExecutionLogConnection()
        log_conn.start_task()

        #Obtem a pasta temp - Use essa pasta pra salvar arquivos temporários
        #Get the temp folder - Use this folder to save temporary files
        #Obtenga la carpeta temp - use esta carpeta para guardar archivos temporales
        temp_folder_path = fc_utils.getTempFolderAbsolutePath()
        file_name = now().strftime("%Y-%m-%d-%H-%M-%S") + ".xlsx"

        #SAP Login********************************
        #****************ATTENTION****************
        #Change here the SAP Instance name and inform login, password, language if needed
        #Altere aqui o nome da instância SAP e informe o login, senha e idioma, se necessário
        ##Cambie aquí el nombre de la instancia de SAP e informe el nombre de usuario, la contraseña y el idioma si es necesario

        bot, session = fc_sap.login(sapLanguage="EN", 
                                   sso=False, 
                                   maestroCredentialsLabel="MaestroSAPScriptDemo")
        
        #********************************SAP Login
        
        #SAP Script*********************************
        #**********INCLUA AQUI O SEU SCRIPT**********
        #**********INCLUDE YOUR SCRIPT HERE**********
        #**********INCLUYE TU GUIÓN AQUÍ*************

        #Exemple
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nva05"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ctxtVBCOM-VKORG").text = "MW09"
        session.findById("wnd[1]/usr/ctxtVBCOM-VKORG").caretPosition = 4
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = 2
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectAll()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = temp_folder_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[0]").press()       

        #*********************************SAP Script

        #Encerra o Excel
        #Close Excel
        #Cierra Excel.
        try:
            os.system("taskkill /f /im excel.exe")
        except:
            pass                


        bot.wait(5000)
        log_conn.add_success()

        lst_files = []
        lst_files.append(f"{temp_folder_path}\\{file_name}")

        #******************** SHAREPOINT ********************
        #Opcionalmente faz o upload do arquivo para o Sharepoint
        #Opcionalmente, suba el archivo a SharePoint.
        #Optionally, upload the file to SharePoint.
        #Acesse fc_sharepoint (fc_office365.py --> class fc_sharepoint --> def __init__) para definir a URL e pasta Sharepoint
        #Acceda a fc_sharepoint (fc_office365.py --> class fc_sharepoint --> def __init__) para definir la URL y la carpeta de SharePoint.
        #Access fc_sharepoint (fc_office365.py --> class fc_sharepoint --> def __init__) to define the SharePoint URL and folder.
        sp = fc_sharepoint()
        sp.upload(lst_files)

        #Opcionalmente a URL e Pasta podem ser enviadas como parâmetro
        #Opcionalmente, la URL y la carpeta pueden enviarse como parámetros.
        #Optionally, the URL and folder can be sent as parameters.
        #sp.upload(file_paths=lst_files, 
        #          sharepoint_url="https://bayergroup.sharepoint.com/sites/028933",
        #          sharepoint_folder="Shared Documents/RPA/Upload")
        
        #******************** E-MAIL ********************
        #Opcionalmente faz o envio do arquivo por email
        #Opcionalmente, envíe el archivo por correo electrónico.
        #Optionally, send the file by email.
        
        #Para habilitar o envio de email veja mais detalhes em #https://docs.int.bayer.com/apserpaframework/functions/office365/#2-e-mail-fc_emailsendmail
        #Para habilitar el envío de correos electrónicos, consulte más detalles en https://docs.int.bayer.com/apserpaframework/es/functions/office365/#2-correo-fc_emailsendmail
        #To enable email sending, see more details at https://docs.int.bayer.com/apserpaframework/en/functions/office365/#2-email-fc_emailsendmail
        fc_email.sendMail(to="to@bayer.com",
                          subject="Test Template Python Title",
                          message="Test Template Python message body", 
                          lstFile=lst_files
        )


        #******************** TEAMS ********************
        #Opcionalmente pode enviar uma mensagem via TEAMS
        #Opcionalmente, envíe el una messagem vai TEAM.
        #Optionally, send a message using TEAMS.
        fc_teams.sendCardChatBot(messageTo="to@bayer.com", 
                                 messageType=fc_teams_enumMessageType.Info, 
                                 messageTitle="Message Title Test", 
                                 messageBody="Message Body Test"
        )



        #******************** MAESTRO ********************
        #Finaliza a execução no Maestro - OBRIGATÓRIO
        #Finalizar la actuación en el Maestro - OBLIGATORIO
        #Finalize the performance at the Maestro - MANDATORY
        log_conn.finish_task(message="Processo executado com sucesso")
        

    except Exception as ex:
        #Finalizando a tarefa com erro no orquestrador - OBRIGATÓRIO em caso de erro
        #Ending the task with an error in the orchestrator - MANDATORY in case of error
        #Finalizar la tarea con un error en el orquestador - OBLIGATORIO en caso de error
        log_conn.add_error()
        log_conn.finish_task(message=f"Ocorreu um erro no processamento - {str(ex)}")

    finally:
        fc_sap.logout()  

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
    