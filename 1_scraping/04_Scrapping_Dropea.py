import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
EMAIL = "paulapamp@icloud.com"
PASSWORD = "*****"
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
try:
    driver.get("https://app.dropea.com/#/auth/login")
    WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='email']"))
    ).send_keys(EMAIL)
    WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='password']"))
    ).send_keys(PASSWORD)
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button#kt_sign_in_submit"))
    ).click()
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div#kt_app_content_container"))
    )
    print("[login] Sesión iniciada correctamente ✅")
    pedido_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[routerlink='/orders']"))
    )
    ActionChains(driver).move_to_element(pedido_btn).click().perform()
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "app-order"))
    )
    print("[nav] Entró al menú 'Pedidos' correctamente ✅")
    fecha_filtro_btn = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[1]/app-date-picker-order/div[2]/div[1]/div/div/a/span/i/span[2]"))
    )
    time.sleep(2)
    driver.execute_script("arguments[0].click();", fecha_filtro_btn)
    time.sleep(1)
    fecha_creacion_btn = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[1]/app-date-picker-order/div[2]/div[1]/div/div/div/div[1]/a/span[2]"))
    )
    driver.execute_script("arguments[0].click();", fecha_creacion_btn)
    print("[filtros] Fecha de Creación seleccionada ✅")
    time.sleep(1)
    fecha_inicio_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[1]/app-date-picker-order/div[2]/div[2]/div/input"))
    )
    fecha_inicio_input.clear()
    fecha_inicio_input.send_keys("01/05/2025")
    print("[filtros] Fecha inicio puesta: 01/05/2025 ✅")
    fecha_hasta_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[1]/app-date-picker-order/div[2]/div[3]/div/input"))
    )
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    fecha_hasta_input.clear()
    fecha_hasta_input.send_keys(fecha_hoy)
    print(f"[filtros] Fecha hasta puesta: {fecha_hoy} ✅")
    estado_btn = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[2]/app-orderstate/ng-multiselect-dropdown//div/div[2]/ul[1]/li[1]/div"))
    )
    time.sleep(2)
    driver.execute_script("arguments[0].click();", estado_btn)
    time.sleep(1)
    estado_caret = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[2]/app-orderstate/ng-multiselect-dropdown/div/div[1]/span/span[2]/span"))
    )
    driver.execute_script("arguments[0].click();", estado_caret)
    time.sleep(1)
    seleccionar_todos = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[1]/div[1]/div[2]/app-orderstate/ng-multiselect-dropdown/div/div[2]/ul[1]/li[1]/div"))
    )
    driver.execute_script("arguments[0].click();", seleccionar_todos)
    print("[filtros] Todos los estados seleccionados ✅")
    filtrar_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[2]/div/button[3]"))
    )
    driver.execute_script("arguments[0].click();", filtrar_btn)
    print("[filtros] Botón 'Filtrar' pulsado ✅")
    time.sleep(10)
    export_btn = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='kt_app_content_container']/app-order/div[1]/app-filters/div/div/form/div[2]/div/button[1]"))
    )
    driver.execute_script("arguments[0].click();", export_btn)
    print("[export] Botón 'Exportar' pulsado ✅")
    time.sleep(15)
    downloads_path = os.path.join(os.environ["USERPROFILE"], "Downloads")
    output_folder = r"C:\Ingestador\output"
    final_filename = "pedidos_dropea.xlsx"
    final_filepath = os.path.join(output_folder, final_filename)
    downloaded_file = None
    timeout = 60
    poll_time = 1
    start_time = time.time()
    while (time.time() - start_time) < timeout:
        xlsx_files = [f for f in os.listdir(downloads_path) if f.lower().endswith(".xlsx")]
        if xlsx_files:
            xlsx_files.sort(key=lambda f: os.path.getmtime(os.path.join(downloads_path, f)), reverse=True)
            downloaded_file = os.path.join(downloads_path, xlsx_files[0])
            if not downloaded_file.endswith(".crdownload"):
                break
        time.sleep(poll_time)
    if not downloaded_file:
        raise Exception("No se detectó el archivo exportado en la carpeta Descargas.")
    print("[export] Archivo descargado exitosamente desde Dropea ✅")
    os.makedirs(output_folder, exist_ok=True)
    if os.path.exists(final_filepath):
        os.remove(final_filepath)
    shutil.move(downloaded_file, final_filepath)
    print(f"[export] Archivo guardado correctamente en: {final_filepath} ✅")
    zone_identifier = final_filepath + ":Zone.Identifier"
    if os.path.exists(zone_identifier):
        os.remove(zone_identifier)
        print("[export] Archivo desbloqueado ✅")
finally:
    driver.quit()
