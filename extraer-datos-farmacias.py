import os, os.path
from os import walk
import os
from getpass import getuser
from datetime import date
from datetime import datetime
import win32com.client as win32
import pythoncom
import win32com.client
import sys
import subprocess
import time
import openpyxl
from openpyxl import load_workbook

def ingresarsap(usuario_sap,contrasena_sap):
    try:

        pythoncom.CoInitialize()
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(2)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection("QAS", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        time.sleep(1)
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario_sap
        time.sleep(0.3)
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = contrasena_sap
        time.sleep(0.3)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)
    except:
        print(sys.exc_info()[0] + "Usuario y/o contraseña incorrecto. Vuelva a intentar.")

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


def meteteensap(sesionsap, nro_afiliado):
    
    pythoncom.CoInitialize()

    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return
    connection = application.Children(0)

    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return

    session = connection.Children(sesionsap)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
    
    try:
        print(f"** Consultando afiliado: {nro_afiliado}")
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nz_sd_estados_hist"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtSO_VKORG-LOW").text = "SC10"
        session.findById("wnd[0]/usr/ctxtSO_VDATU-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtSO_AFILI-LOW").text = f"{nro_afiliado}"
        session.findById("wnd[0]/usr/ctxtSO_AFILI-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtSO_AFILI-LOW").caretPosition = 8
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(32,"TEXT")
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 28
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "32"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        
        # CAMPOS QUE NECESITAMOS EXTRAER.
        cont = 0
        try:
            for i in range(0, 4000):
                dispone = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(i, "DESTINA")
                nombre_destinatario = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(i, "DESTINA_NAME")
                domicilio = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(i, "DOMICILIO")
                localidad = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(i, "LOCALIDAD")
                cont += 1
        except:
            print(f"Entrando en except de sap. {str(cont - 1)}")
            dispone = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(cont - 1, "DESTINA")
            nombre_destinatario = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(cont - 1, "DESTINA_NAME")
            domicilio = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(cont - 1, "DOMICILIO")
            localidad = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(cont - 1, "LOCALIDAD")
        
        
        print(f"Dispones: {dispone}")
        print(f"Nombre de farmacia: {nombre_destinatario}")
        print(f"Domicilio: {domicilio}")
        print(f"Localidad: {localidad}")
        print("")

        return dispone, nombre_destinatario, domicilio, localidad
    except:
        print("Error en el script.. Programa finalizado!")


def completar_datos_excel(ruta_excel):

    excel = load_workbook(ruta_excel)
    excel = load_workbook(ruta_excel, data_only=True)
    hoja = excel["Padron Nico"]
    try:
        for i in range(2, 675):
            print(f"**** {hoja[f'A{i}'].value}****")
            if hoja[f"A{i}"].value != None and hoja[f"D{i}"].value == "NO":
                nro_afiliado = hoja[f"A{i}"].value
                datos_farma = meteteensap(0, f"{nro_afiliado}")

                dispone = datos_farma[0]
                farmacia = datos_farma[1]
                domicilio = datos_farma[2]
                localidad = datos_farma[3]

                hoja[f"D{i}"].value = str(farmacia)
                hoja[f"E{i}"].value = str(dispone)
                hoja[f"F{i}"].value = str(domicilio)
                hoja[f"G{i}"].value = str(localidad)
            else:
                print("Nada")
                continue
    except Exception as e:
        print("Algo salio mal en el excel.", e)
    finally:
        excel.save(ruta_excel)
        excel.close()

ruta_excel = "C:/Users/aalarcon/Desktop/OyP/CARGA AUTOMATICA EUROSISTEMAS/P2.xlsx"
completar_datos_excel(ruta_excel)    
