from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# === SharePoint Credentials ===
USERNAME = "isaacko@suzano.com.br"
PASSWORD = "gmE@8rs.fG@RSx"
SHAREPOINT_FOLDER_URL = "https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2FProjetoCerrado%2DConfiabilidadeInovao%2FDocumentos%20Compartilhados%2F99%20%2D%20Indicadores%20Industriais%2F01%20%2D%20Indicadores%20Ribas%2F04%20%2D%20Hist%C3%B3rico%20Prog%20Semanal&viewid=1234ddc1%2D8ffa%2D4e4f%2D8244%2D9cbabb46f807&e=5%3A4b6052593f5645109b75df767c65ecd8&sharingv2=true&fromShare=true&at=9&CT=1743014215520&OR=OWA%2DNT%2DMail&CID=5404928a%2D779d%2Da2de%2D2f86%2D231233fec94c&FolderCTID=0x0120004BC1EE6B50F0FC4581C3ADD013BED0FB"
NEW_FOLDER_NAME = "Automated_Folder"

# === Set Up Selenium WebDriver ===
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Open browser maximized
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)  # Wait up to 10 seconds for elements to load

try:
    # === Step 1: Open Chrome and Access SharePoint ===
    print("🌐 Accessing SharePoint...")
    driver.get(SHAREPOINT_FOLDER_URL)
    
    # === Step 2: Log In (If Required) ===
    try:
        email_input = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
        email_input.send_keys(USERNAME)
        email_input.send_keys(Keys.RETURN)
        time.sleep(3)

        password_input = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
        password_input.send_keys(PASSWORD)
        password_input.send_keys(Keys.RETURN)
        time.sleep(5)
    except:
        print("🔓 Already logged in or no login required.")

    # === Step 3: Wait for Page to Load ===
    time.sleep(5)  # Allow full page load

    # === Step 4: Click "New" Button ===
    print("📁 Creating New Folder...")
    new_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Novo')]")))
    new_button.click()
    time.sleep(2)

    # === Step 5: Click "Folder" Option ===
    folder_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Pasta')]")))
    folder_option.click()
    time.sleep(2)

    # === Step 6: Enter Folder Name and Confirm ===
    name_input = driver.switch_to.active_element  # Select active input field
    name_input.send_keys(NEW_FOLDER_NAME)
    name_input.send_keys(Keys.RETURN)
    time.sleep(5)

    print(f"✅ Folder '{NEW_FOLDER_NAME}' created successfully!")

except Exception as e:
    print("❌ Error:", e)

finally:
    time.sleep(5)
    driver.quit()
    print("🔚 Browser closed.")
