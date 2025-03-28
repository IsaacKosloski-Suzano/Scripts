import os
import time
import win32com.client
import pyautogui
import subprocess
import getpass
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# SharePoint credentials
SHAREPOINT_SITE_URL = "https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao"
CLIENT_ID = "isaacko"
CLIENT_SECRET = "gmE@8rs.fG@RSx"

# The raw SharePoint URL with encoded characters
raw_url = "https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2FProjetoCerrado%2DConfiabilidadeInovao%2FDocumentos%20Compartilhados%2F99%20%2D%20Indicadores%20Industriais%2F01%20%2D%20Indicadores%20Ribas%2F04%20%2D%20Hist%C3%B3rico%20Prog%20Semanal"

# Extract the relative path after "id="
split_url = raw_url.split("id=")[-1].split("&")[0]
decoded_path = unquote(split_url)  # Decode URL-encoded characters

# Remove unnecessary "/sites/ProjetoCerrado-ConfiabilidadeInovao"
relative_folder_path = decoded_path.replace("/sites/ProjetoCerrado-ConfiabilidadeInovao/", "")

# Authenticate with SharePoint
ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(ClientCredential(CLIENT_ID, CLIENT_SECRET))


def create_sharepoint_folder(folder_name):
    """
    Creates a folder in SharePoint if it does not already exist.
    """
    folder_url = f"{relative_folder_path}/{folder_name}"
    
    try:
        # Try to get the folder
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        ctx.load(folder)
        ctx.execute_query()
        print(f"Folder '{folder_name}' already exists in SharePoint.")
    
    except Exception as e:
        # If folder does not exist, create it
        try:
            root_folder = ctx.web.get_folder_by_server_relative_url(relative_folder_path)
            root_folder.folders.add(folder_name)
            ctx.execute_query()
            print(f"Folder '{folder_name}' created successfully in SharePoint.")
        except Exception as create_error:
            print(f"Error creating SharePoint folder '{folder_name}': {create_error}")


def upload_file_to_sharepoint(local_file_path, sharepoint_folder_path):
    """
    Uploads a file to the specified SharePoint folder.
    """
    file_name = os.path.basename(local_file_path)
    target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder_path)
    
    with open(local_file_path, "rb") as file_content:
        target_folder.upload_file(file_name, file_content)
        ctx.execute_query()
    
    print(f"File '{file_name}' uploaded successfully to SharePoint folder: {sharepoint_folder_path}")


def saplogin(system_name="SBP - ERP ECC - Produção"):
    """
    Automates SAP login using SAP GUI Scripting.
    """
    try:
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(10)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(system_name, True)
        time.sleep(5)
        session = connection.Children(0)

        username = "ISAACKO"
        password = "gmE@8rs.fG@RSx"

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]").sendVKey(0)

        print("SAP Login Successful!")
        return session
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

    return path_planilha01

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

    return path_planilha02


# Main Execution

# Get current date and time
now = datetime.now()

# Get date data
week_number = now.isocalendar()[1]  # Week number of the year
month_number = now.strftime("%m")  # Month as a number (01-12)
last_two_digits_year = now.strftime("%y")  # Last two digits of the year

# Construct REVISÃO code
revisao = f"SM{week_number}{month_number}{last_two_digits_year}"

# Get os user
user = getpass.getuser()
path_download = f"C:\Users\{user}\Downloads"

#Create dir
dir_sm = Path(f"{path_download}\\{revisao}")
dir_sm.mkdir(parents=True, exist_ok=True) # Creates the directory if it doesn't exist
print(f"Directory '{dir_sm}' created successfully!")

# Create SharePoint folder for the revision
create_sharepoint_folder(revisao)

sap_session = saplogin("SBP - ERP ECC - Produção")

if sap_session:
    file_ordem = extract_ORDEM(sap_session, dir_sm)
    file_operacao = extract_OPERACAO(sap_session, dir_sm)

    # Upload extracted files to SharePoint
    upload_file_to_sharepoint(revisao, file_ordem)
    upload_file_to_sharepoint(revisao, file_operacao)

    # Close SAP session
    time.sleep(5)
    sap_session.findById("wnd[0]").close()
    pyautogui.hotkey('alt', 'f4')
