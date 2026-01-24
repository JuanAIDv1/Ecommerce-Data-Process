import json
import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
load_dotenv("C:/Ingestador/configs/Login.env")
print(os.environ.get("DROPI_USER"))
print(os.environ.get("DROPI_PASS"))
USER = os.getenv("DROPI_USER")
PASS = os.getenv("DROPI_PASS")
TIENDAS = "DropiPro"
COOKIES_PATH = f"C:/Ingestador/cookies/dropipro/{TIENDAS}.json"
URL_LOGIN = "https://dropipro.com"
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
for tienda in TIENDAS:
    print(f"🔐 Iniciando sesión en tienda: {tienda}")
try:
    driver.get(URL_LOGIN)
    wait = WebDriverWait(driver, 15)
    username_input = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="input-username"]'))
    )
    username_input.send_keys (USER)
    password_input = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="input-password"]'))
    )
    password_input.send_keys(PASS)
    password_input.send_keys(Keys.RETURN)
    try:
        wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="side-menu"]/li[7]/a'))
        )
    except TimeoutException:
        print("⚠️ No se detectó login exitoso, revisa tus credenciales o la página.")
    cookies = driver.get_cookies()
    os.makedirs(os.path.dirname(COOKIES_PATH), exist_ok=True)
    with open(COOKIES_PATH, "w") as f:
        json.dump(cookies, f)
    print(f"✅ Cookies guardadas en {COOKIES_PATH}")
finally:
    driver.quit()
