from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# Specify the path to ChromeDriver (update the path accordingly)
CHROME_DRIVER_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Configure Chrome options
options = Options()
options.add_argument("--start-maximized")

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(CHROME_DRIVER_PATH))
driver.get("https://www.deepseek.com")

# Open the URL
url = "https://www.deepseek.com"
driver.get(url)

# Verify page load
print("Opened:", driver.current_url)
print("Title:", driver.title)

# Wait & close
import time
time.sleep(3)
driver.quit()
print("Browser closed.")