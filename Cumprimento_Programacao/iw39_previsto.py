import win32com.client
import subprocess
import time
import os
import pandas as pd
import pyautogui
from datetime import datetime
from pathlib import Path

def saplogin(system_name="SBP - ERP ECC - Produção"):
    try:
        # Step 1: Launch SAP Logon
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(10)  # Wait for SAP Logon to open

        # Step 2: Get SAP GUI Scripting Engine
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
        except Exception as e:
            print(f"Error: Unable to connect to SAP GUI. {e}")
            return None

        # Step 3: Open SAP Connection
        try:
            connection = application.OpenConnection(system_name, True)
            time.sleep(5)  # Wait for SAP login screen
        except Exception as e:
            print(f"Error: Unable to open connection {system_name}. {e}")
            return None

        # Step 4: Get SAP Session
        try:
            session = connection.Children(0)
        except Exception as e:
            print(f"Error: Unable to start SAP session. {e}")
            return None

        # Step 5: Enter Login Credentials
        username = "ISAACKO"  # Change as needed
        password = "gmE@8rs.fG@RSx"  # Change as needed

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]").sendVKey(0)

        print("SAP Login Successful!")
        return session  # Return the session object

    except Exception as e:
        print(f"Unexpected Error: {e}")
        return None

def extract_ORDEM(session, dir_sm):
    try:
        # Define file paths 
        planilha01 = f"Planilha 01 - OM's {revisao}"
        path_planilha01 = os.path.join(dir_sm, planilha01)

        #share_point_path_planilha01 = r"https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/99%20-%20Indicadores%20Industriais/01%20-%20Indicadores%20Ribas/02%20-%20Painel%20Extração%20SAP/"

        # Access IW39 transaction - ORDEM 
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nIW39"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATUV").text = ""
        session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "2298"
        session.findById("wnd[0]/usr/ctxtREVNR-LOW").text = revisao
        session.findById("wnd[0]/usr/ctxtVARIANT").text = "/2298-PROGOM"
        session.findById("wnd[0]").sendVKey(8)
    
        # Export as Excel (.xlsx) - ORDEM
        session.findById("wnd[0]").sendVKey(16)
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus()
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        try:        
            # Connect to the running Excel instance
            excel = win32com.client.Dispatch("Excel.Application") 
            excel.Visible = True  # Set to False if you want to run it in the background
            wb = excel.ActiveWorkbook # Get the active workbook

            # Save the file in .xlsx format
            wb.SaveAs(path_planilha01, FileFormat=51)  # 51 = .xlsx format
            print(f"Excel file saved successfully: {path_planilha01}")

            # Close the workbook (without prompt)
            wb.Close(SaveChanges=False)

            # Quit Excel application
            excel.Quit()

            # Release COM objects
            #del ws, wb, excel

        except Exception as e:
            print(f"Error saving Excel file: {e}")

    except Exception as e:
        print(f"Error extracting IW39 data: {e}")

def extract_OPERACAO(session, dir_sm):
    try:
        # Define file paths 
        planilha02 = f"Planilha 02 - OP's {revisao}"
        path_planilha02 = os.path.join(dir_sm, planilha02)

        #arquivo_xlsx = r"https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/99%20-%20Indicadores%20Industriais/01%20-%20Indicadores%20Ribas/02%20-%20Painel%20Extração%20SAP/2298 - IW39.xlsx"

    # Access IW39 transaction - OPERAÇÃO from ORDEM
        session.findById("wnd[0]").maximize()
        # Select all entries in the table/grid
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        grid.setCurrentCell(-1, "")  # Selects entire table
        grid.selectAll()
        # Press VKey 43 (Typically "Copy to Clipboard" or "Export")
        session.findById("wnd[0]").sendVKey(43)

        # Open menu option
        session.findById("wnd[0]/mbar/menu[5]/menu[2]/menu[1]").select()

        # Handle pop-up window for table selection
        popup_grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
        popup_grid.currentCellRow = -1
        popup_grid.selectColumn("VARIANT")
        popup_grid.pressColumnHeader("VARIANT")

        # Open search context menu
        popup_grid.contextMenu()
        popup_grid.selectContextMenuItem("&FIND")

        # Enter search value ("2298") in search pop-up
        search_field = session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE")
        search_field.Text = "2298"
        search_field.caretPosition = 4  # Position cursor at the end
        session.findById("wnd[2]/tbar[0]/btn[0]").press()  # Press "Search"
        
        # Close search window
        session.findById("wnd[2]").close()

        # Clear previous selection
        popup_grid.clearSelection()
        
        # Select specific row (e.g., row 576)
        popup_grid.selectedRows = "576"
        
        # Click on the selected row
        popup_grid.clickCurrentCell()

        # Export as Excel (.xlsx) - OPERACAO
        session.findById("wnd[0]").sendVKey(16)
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus()
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]").sendVKey(0)
        

        try:        
            # Connect to the running Excel instance
            excel = win32com.client.Dispatch("Excel.Application") 
            excel.Visible = True  # Set to False if you want to run it in the background
            wb = excel.ActiveWorkbook # Get the active workbook

            # Save the file in .xlsx format
            wb.SaveAs(path_planilha02, FileFormat=51)  # 51 = .xlsx format
            print(f"Excel file saved successfully: {path_planilha02}")

            # Close the workbook (without prompt)
            wb.Close(SaveChanges=False)

            # Quit Excel application
            excel.Quit()

            # Release COM objects
            #del ws, wb, excel

        except Exception as e:
            print(f"Error saving Excel file: {e}")
    except Exception as e:
        print(f"Error extracting IW39 data: {e}")

# Main Execution

# Get current date and time
now = datetime.now()

# Get date data
week_number = now.isocalendar()[1]  # Week number of the year
month_number = now.strftime("%m")  # Month as a number (01-12)
last_two_digits_year = now.strftime("%y")  # Last two digits of the year

# Construct REVISÃO code
revisao = f"SM{week_number}{month_number}{last_two_digits_year}"

pasta_cumprimento_programacao = r"C:\Users\isaacko\OneDrive - Suzano S A\Documentos\Cumprimento_da_Programacao"
#Create dir
dir_sm = Path(f"{pasta_cumprimento_programacao}\\{revisao}")
dir_sm.mkdir(parents=True, exist_ok=True) # Creates the directory if it doesn't exist
print(f"Directory '{dir_sm}' created successfully!")

sap_session = saplogin("SBP - ERP ECC - Produção")

if sap_session:
    extract_ORDEM(sap_session, dir_sm)
    extract_OPERACAO(sap_session, dir_sm)

time.sleep(5)        
sap_session.findById("wnd[0]").close()
pyautogui.hotkey('alt', 'f4')

