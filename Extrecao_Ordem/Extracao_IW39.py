import os
import time
import pandas as pd
from pywinauto import Application
from pywinauto.keyboard import send_keys

# Define the download folder
download_folder = os.path.join(os.environ["USERPROFILE"], "Downloads")

# SAP GUI Automation
# Connect to SAP GUI
app = Application(backend="win32").connect(title_re=".*SAP.*")
session = app.window(title_re=".*SAP.*")

# Maximize the window and navigate to IW39
session.type_keys("%n")  # Alt + n to focus on the transaction code field
session.type_keys("/nIW39{ENTER}")
time.sleep(2)  # Wait for the screen to load

# Input parameters
session.type_keys("01.01.2024{TAB}")  # Start date
session.type_keys("{TAB}")  # Leave end date blank
session.type_keys("2298{TAB}")  # Plant code
session.type_keys("/indiciw38{TAB}")  # Variant
session.type_keys("{F8}")  # Execute the transaction

# Export to Excel
session.type_keys("%fe")  # Alt + f, then e to open the export menu
time.sleep(1)
session.type_keys("{DOWN 2}{ENTER}")  # Select Excel format
time.sleep(1)
session.type_keys(download_folder + "\\2298 - IW39.xls{ENTER}")  # Save file
time.sleep(5)  # Wait for the export to complete

# Excel File Processing
input_file = os.path.join(download_folder, "2298 - IW39.xls")
output_file = "https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/99%20-%20Indicadores%20Industriais/01%20-%20Indicadores%20Ribas/02%20-%20Painel%20Extra%C3%A7%C3%A3o%20SAP/2298%20-%20IW39.xlsx"

# Read the Excel file
df = pd.read_excel(input_file, header=None)

# Delete unnecessary rows and columns
df = df.drop([0, 1, 2])  # Delete rows 1-3
df = df.drop(1)  # Delete row 2
df = df.drop(df.columns[0], axis=1)  # Delete column A

# Save as .xlsx
df.to_excel(output_file, index=False, header=False)

print("Process completed successfully!")