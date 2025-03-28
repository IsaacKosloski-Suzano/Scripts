from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui as pag

# Set login data
EMAIL = r"isaacko@suzano.com.br"
PASSWORD = r"gmE@8rs.fG@RSx"

# Set the url to the Sharepoint folder
url_historico_prog_sem = r'https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2FProjetoCerrado%2DConfiabilidadeInovao%2FDocumentos%20Compartilhados%2F99%20%2D%20Indicadores%20Industriais%2F01%20%2D%20Indicadores%20Ribas%2F04%20%2D%20Hist%C3%B3rico%20Prog%20Semanal&viewid=1234ddc1%2D8ffa%2D4e4f%2D8244%2D9cbabb46f807'

# Set the XPpath to the new folder
new_button_xPath = r'//*[@id="appRoot"]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div/div[1]/div[1]/div/div/span[1]/button[1]'
new_folder_xPath = r'//*[@id="command-bar-menu-id"]/div/ul/li[1]/button/div/span[2]'
new_folder_color_xPath = r'//*[@id="swatchColorPicker384-grey-8"]/span/svg'
folder_name_input_xPath = r'//*[@id="textField4"]'
create_button_xPath = r'/html/body/div[8]/div[2]/div/div[2]/div[3]/button[1]'
create_button_selector = r'body > div.fui-FluentProvider1.___1uxlrft.f1euv43f.f15twtuk.f1vgc2s3.f1e31b4d.f494woh > div.fui-DialogSurface.rsmdyd3.main-213 > div > div.fui-DialogBody.r1h3qql9.ms-Dialog-main > div.fui-DialogActions.rhfpeu0.ms-Dialog-actions.___pj8zjl0.f1a7i8kp.fd46tj4.fsyjsko.f1f41i0t.f1jaqex3.f2ao6jk > button.fui-Button.r1alrhcs.ms-Button.root-140.ms-Button--primary.ms-ButtonShim--primary.___1ozmuvc.f4wwz4m.f1p3nwhy.f11589ue.f1q5o8ev.f1pdflbu.f1phragk.f15831mw.f1s2uweq.fr80ssc.f1ukrpxl.fecsdlb.f1vjjmo9.fnp9lpt.f1h0usnq.fs4ktlq.f16h9ulv.fx2bmrt.f1d6v5y2.f1rirnrt.f1uu00uk.fkvaka8.f1ux7til.f9a0qzu.f1lkg8j3.fkc42ay.fq7113v.ff1wgvm.fiob0tu.f1j6scgf.f1x4h75k.f4xjyn1.fbgcvur.f1ks1yx8.f1o6qegi.fcnxywj.fmxjhhp.f9ddjv3.f17t0x8g.f194v5ow.f1qgg65p.fk7jm04.fhgccpy.f32wu9k.fu5nqqq.f13prjl2.f1czftr5.f1nl83rv.f12k37oa.fr96u23.f10pi13n.f14t3ns0.f1k6fduh.f17mccla.f1xqy1su.f1ewtqcl.f1g0x7ka.f1gbmcue.f1qch9an.f1rh9g5y.f2cu5sd.figsok6.f1vdwobg.f15fhrkf.f1gl3qri.fwzdp62.f492aqz.f1eeugec.f141lptb.fx7fgq1.f1xbyz6e.f1bq85fa.fudw3ie.f15sz26m.f1ct40e8.f12btegu.fhcypdg.f1bzmdyv.fnn62qi.f1xvby9z.fimu2fn.f4l4af4.fdc9mbu.ffomqby.f1f9t5xu.f1d2rq10.f1jhnwmq.f1xxju6a.f1j2v30v.fv161b.f11gj9w9.f1qioim3.f7kir35.fpf1hop.f1h3a8gf.f15lhznk.fjlw612.fqwrxe1.ficcvdo.f1o3bz38.f1s8xf6m.f179hqoo.f1gox7ek.f1jlt9y9.f1lyr7jz.fy9n62g.fvazur4'

# Set up Chrome driver
service = Service(r"C:\Users\isaacko\chromedriver-win64\chromedriver.exe")  # Replace with your path
driver = webdriver.Chrome(service=service)

# Open SharePoint login page
driver.get("https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/")

# Wait for the email input field
wait = WebDriverWait(driver, 30)
email_input = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
email_input.send_keys(EMAIL)  # Replace with your email
email_input.send_keys(Keys.RETURN)

time.sleep(2)  # Adjust as needed

# Enter password (wait until it's available)
password_input = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
password_input.send_keys(PASSWORD)  # Replace with your password (Not recommended in plaintext)
password_input.send_keys(Keys.RETURN)

time.sleep(25)  # Adjust for MFA or redirection

# Navigate to the target folder
driver.get(url_historico_prog_sem)

# Wait and click the "New Folder" button
# Wait until the page is fully loaded
WebDriverWait(driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')

# ✅ Step 1: Click the "New" button
new_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, new_button_xPath))
)
new_button.click()

# ✅ Step 2: Wait for the "Folder" option to appear and click it
folder_option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, new_folder_xPath))
)
folder_option.click()

"""# ✅ Step 3: Wait for the "Color" option to appear and click it
color_option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, new_folder_color_xPath))
)
color_option.click()"""

# ✅ Step 3: Find and enter folder name
folder_name_input = wait.until(EC.element_to_be_clickable((By.XPATH, folder_name_input_xPath)))
folder_name_input.click()  # Ensure focus before typing
folder_name_input.send_keys("MyNewFolder")
folder_name_input.send_keys(Keys.RETURN)

# ✅ Step 4: Click "Create" button
create_button = wait.until(EC.element_to_be_clickable((By.XPATH, create_button_xPath)))
create_button.click()

print("✅ Folder created successfully!")

# Close browser after some time
time.sleep(15)
driver.quit()
