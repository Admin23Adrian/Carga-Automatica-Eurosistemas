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


def meteteensap(sesionsap, clase_de_pedido, canal, sector, solicitante, dispone, fecha_entrega, lista_medicamentos, centro, cantidades, convenio, turno, almacen, afiliado_sap, codificacion_afiliado, nombre, apellido, dni_afiliado, datos_farmacia, sexo):
    
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
        ##arranque
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA01"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = clase_de_pedido
        session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = "SC10"
        session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = str(canal)
        session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = str(sector)
        session.findById("wnd[0]").sendVKey(0)
        # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = ped_ext # Columna E EUROSISTEMAS.
        # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = "" # Columna E Euro
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = solicitante
        
        # v=str(farmacia[0])
        # v1=v[0]
        # if v1=="S":
        #     session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = "84005541"
        # else:
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = dispone
        
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 9
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").text = fecha_entrega
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()
        
        for i in range(0, len(lista_medicamentos)):
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," + str(i) + "]").text = lista_medicamentos[i]
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2," + str(i) + "]").text = cantidades[i]
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[12," + str(i) + "]").text = centro
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[12," + str(i) + "]").setFocus()
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[12," + str(i) + "]").caretPosition = 0
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
        try:
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
        except:
            session.findById("wnd[0]").sendVKey(0)
        try:
            session.findById("wnd[1]").sendVKey(0)
        except:
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08").select()
            # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").key = "ZR"
            
            ## PRD
            # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").text = dispone
            
            #---------------------
            ## QAS
            # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").text = dispone
            # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").setFocus()
            # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").caretPosition = 8
            #---------------------
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").Select()
            time.sleep(1)

            ## PRD
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZCONVENIO").text = convenio
            
            ## QAS
            # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZCONVENIO").text = convenio
            
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").text = turno
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").setFocus()
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").caretPosition = 3
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(3)
            ped = session.findById("wnd[0]/sbar").text
            ped_final = ped[18:25]

                
        session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZSD_TOMA"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").text = ped_final
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").setFocus()
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").currentCellColumn = "BSTKD"
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton("FN_MODPED")
        
        for i in range(0, len(lista_medicamentos)):
            session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/subSUBS_TRAB:ZDMSD_TOMA_PEDIDO:0111/tblZDMSD_TOMA_PEDIDOTC_CARRITO/ctxtGS_CARRITO-LGORT[5," + str(i) + "]").text = almacen
            # session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/subSUBS_TRAB:ZDMSD_TOMA_PEDIDO:0111/tblZDMSD_TOMA_PEDIDOTC_CARRITO/txtGS_CARRITO-COD_EXTERNO[15," + str(i) + "]").text = nrobono[i]
            # session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/subSUBS_TRAB:ZDMSD_TOMA_PEDIDO:0111/tblZDMSD_TOMA_PEDIDOTC_CARRITO/txtGS_CARRITO-COD_EXTERNO[15," + str(i) + "]").setFocus()
            # session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/subSUBS_TRAB:ZDMSD_TOMA_PEDIDO:0111/tblZDMSD_TOMA_PEDIDOTC_CARRITO/txtGS_CARRITO-COD_EXTERNO[15," + str(i) + "]").caretPosition = 4

        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK").key = "PZ"
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT").select()

        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-DISPONE_ID").text = dispone
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-DISPONE_ID").setFocus()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-DISPONE_ID").caretPosition = 8
        session.findById("wnd[0]").sendVKey(0)
        
        # SE CREA EL AFILIADO SI NOS VIENE COMO NO CARGADO DESDE E EXCEL.
        if afiliado_sap == "no_cargado" and solicitante != "40000001":
            try:
                session.findById("wnd[0]").sendVKey(4)
                session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[7,24]").text = codificacion_afiliado
                session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[7,24]").setFocus()
                session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[7,24]").caretPosition = 14
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]/usr/lbl[29,3]").setFocus()
                session.findById("wnd[1]/usr/lbl[29,3]").caretPosition = 11
                session.findById("wnd[1]").sendVKey(2)
            except:
                session.findById("wnd[1]/tbar[0]/btn[12]").press()
                session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CREAAFI").press()
                session.findById("wnd[1]/usr/chkS_NORMALIZA").selected = "false"
                session.findById("wnd[1]/usr/txtGS_AFILIADOS-NRO_AFILIADO").text = codificacion_afiliado #CODIGO AFIL: EJ: MFALU09876543 | M231292
                session.findById("wnd[1]/usr/txtGS_AFILIADOS-NOMBRE").text = nombre
                session.findById("wnd[1]/usr/txtGS_AFILIADOS-APELLIDO").text = apellido
                # session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-TIPO_ID").text = "86"
                session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-FECHA_NAC").text =""
                session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-CALLE").text = "SD"
                session.findById("wnd[1]/usr/txtGS_AFILIADOS-NUMERO").text = "SD"
                session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-LOCALIDAD").text = "SD"
                session.findById("wnd[1]/usr/txtGS_AFILIADOS-CODIGO_POSTAL").text = "1234"
                session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-PROVINCIA").text = "11"
                # session.findById("wnd[1]/usr/txtGS_AFILIADOS-NUMERO_ID").text = dni_afiliado
                session.findById("wnd[1]/usr/txtGS_AFILIADOS-NUMERO_ID").caretPosition = 8
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[2]/usr/btnBUTTON_1").press()
                session.findById("wnd[1]/tbar[0]/btn[11]").press()     
        
        elif afiliado_sap == "no_cargado" and solicitante == "40000001":
            # DETERMINAR SEXO DE AFILIADO: se toma la cadena de codificacion de afiliado y se extrae
            # el primer caracter, en este caso para eurosistema deberia ser F o M.
            # sexo = codificacion_afiliado[0]
            apodo = codificacion_afiliado[0:5]
            try:
                session.findById("wnd[0]").sendVKey(4)
                session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[7,24]").text = codificacion_afiliado
                session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[7,24]").setFocus()
                session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[7,24]").caretPosition = 13
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]/usr/lbl[29,3]").setFocus()
                session.findById("wnd[1]/usr/lbl[29,3]").caretPosition = 11
                session.findById("wnd[1]").sendVKey(2)
            except:
                try:
                    session.findById("wnd[1]/tbar[0]/btn[12]").press()
                    session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CREAAFI").press()
                    # Check de ventana para Femenino o Masculino.
                    if sexo == "F":
                        session.findById("wnd[1]/usr/radRB_F").select()
                    elif sexo == "M":
                        session.findById("wnd[1]/usr/radRB_M").select()

                    session.findById("wnd[1]/usr/chkS_NORMALIZA").selected = "false"
                    session.findById("wnd[1]/usr/txtGS_AFILIADOS-NRO_AFILIADO").text = codificacion_afiliado
                    session.findById("wnd[1]/usr/txtGS_AFILIADOS-NOMBRE").text = nombre
                    session.findById("wnd[1]/usr/txtGS_AFILIADOS-APELLIDO").text = apellido
                    session.findById("wnd[1]/usr/txtGS_AFILIADOS-APODO").text = apodo
                    # session.findById("wnd[1]/usr/txtGS_AFILIADOS-NUMERO_ID").text = dni_afiliado
                    # session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-TIPO_ID").text = "96"
                    session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-FECHA_NAC").text = "11112000"
                    session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-CALLE").text = "SD"
                    session.findById("wnd[1]/usr/txtGS_AFILIADOS-NUMERO").text = "SD"
                    session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-LOCALIDAD").text = "SD"
                    session.findById("wnd[1]/usr/txtGS_AFILIADOS-CODIGO_POSTAL").text = "1234"
                    session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-PROVINCIA").text = "11"
                    session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-PROVINCIA").setFocus()
                    session.findById("wnd[1]/usr/ctxtGS_AFILIADOS-PROVINCIA").caretPosition = 2
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    session.findById("wnd[2]/usr/btnBUTTON_1").press()
                except:
                    print(f"No se pudo crear el afiliado: {codificacion_afiliado}")

        else:
            session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-AFIL_NRO").text = afiliado_sap # Codigo afiliado SAP
        
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        
        if dispone == "84005541":
            session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/txtGS_ENTREGA-OBSERV_INT").text = datos_farmacia
        
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/txtGS_ENTREGA-OBSERV_INT").setFocus()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/txtGS_ENTREGA-OBSERV_INT").caretPosition = 10
         
        try:
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except:    
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        
        try:
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except:
            try:
                ##sincartel11:
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
            except:
                ##sinvalidacion:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            session.findById("wnd[0]/tbar[1]/btn[7]").press()
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
        except:
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
            return ped_final
        return ped_final
    
    except:
        time.sleep(3)
        return pedido
        # session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        # session.findById("wnd[0]").sendVKey(0)
    
