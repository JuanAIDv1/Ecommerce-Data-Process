import os
import json
import time
import re
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
load_dotenv("C:/Ingestador/configs/Login.env")
USER = os.getenv("DROPI_USER")
PASS = os.getenv("DROPI_PASS")
print("Login Exitoso")
print(f"Usuario: {USER}")
print(f"Contraseña: {PASS}")
COOKIES_PATH = "C:/Ingestador/cookies/dropipro/dropipro.json"
URL_LOGIN = "https://dropipro.com"
OUTPUT_FILE = "C:/Ingestador/output/pedidos_dropipro.xlsx"
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)
def login():
    driver.get(URL_LOGIN)
    try:
        username_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-username"]')))
        username_input.send_keys(USER)
        password_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-password"]')))
        password_input.send_keys(PASS)
        password_input.submit()
        time.sleep(5)
        with open(COOKIES_PATH, "w") as f:
            json.dump(driver.get_cookies(), f)
        print("✅ Login exitoso y cookies guardadas")
    except Exception as e:
        print("❌ Error en login:", e)
def load_cookies():
    driver.get(URL_LOGIN)
    if os.path.exists(COOKIES_PATH):
        with open(COOKIES_PATH, "r") as f:
            cookies = json.load(f)
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.refresh()
        print("✅ Sesión restaurada con cookies")
    else:
        login()
def click_con_xpath(xpath, driver, wait, retries=3):
    for attempt in range(retries):
        try:
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            driver.execute_script("arguments[0].click();", element)
            time.sleep(1)
            return True
        except Exception:
            time.sleep(1)
    print(f"⚠️ Error al hacer click en {xpath}")
    return False
def safe_get(xpath, attr="value", text=False):
    """Intenta obtener un valor de un elemento. Si no existe, retorna ''."""
    try:
        elem = driver.find_element(By.XPATH, xpath)
        if text:
            return elem.text.strip()
        return elem.get_attribute(attr) or ""
    except NoSuchElementException:
        print(f"   ⚠️ Campo no encontrado: {xpath}")
        return ""
def scrape():
    pedidos_menu_xpath = '//*[@id="side-menu"]/li[7]/a'
    click_con_xpath(pedidos_menu_xpath, driver, wait)
    click_con_xpath('//*[@id="status_select"]', driver, wait)
    click_con_xpath('//*[@id="status_select"]/option[1]', driver, wait)
    print("✅ Estado de pedido seleccionado: TODOS")
    click_con_xpath('//*[@id="store_select"]', driver, wait)
    click_con_xpath('//*[@id="store_select"]/option[1]', driver, wait)
    print("✅ Tienda seleccionada: Todas")
    time.sleep(3)
    ultima_pagina = int(driver.find_element(By.XPATH, '//ul[@class="pagination"]/li[last()-1]/a').text)
    print(f"📄 Total de páginas: {ultima_pagina}")
    columnas = [
        "Tienda", "Numero de Pedido", "Nombre completo", "Fecha de Pedido",
        "Dirección", "Dirección 2", "Teléfono", "Email", "Ciudad", "Provincia",
        "Código Postal", "Método de envío", "Productos", "Cantidades", "Coste productos (sin IVA)",
        "Coste logística", "Coste total (IVA incl)", "Importe reembolso",
        "Beneficio Dropshipper", "Código de seguimiento",
        "ID de pedido externo", "ID externo", "Estado"
    ]
    df = pd.DataFrame(columns=columnas)
    for pagina in range(ultima_pagina, 0, -1):
        print(f"➡️ Procesando página {pagina}...")
        pagina_xpath = f'//ul[@class="pagination"]/li/a[text()="{pagina}"]'
        click_con_xpath(pagina_xpath, driver, wait)
        time.sleep(2)
        rows = driver.find_elements(By.XPATH, '//table/tbody/tr')
        for i in range(len(rows)-1, -1, -1):
            try:
                rows = driver.find_elements(By.XPATH, '//table/tbody/tr')
                row = rows[i]
                fecha_desde_tabla = ""
                try:
                    tds = row.find_elements(By.TAG_NAME, "td")
                    for td in tds:
                        text = td.text.strip()
                        if not text:
                            continue
                        if re.search(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', text) or re.search(r'\b\d{4}-\d{2}-\d{2}\b', text):
                            fecha_desde_tabla = text
                            break
                except Exception as e:
                    fecha_desde_tabla = ""
                    print(f"⚠️ Error extrayendo fecha desde la fila: {e}")
                print(f"   📅 Fecha (desde tabla): {fecha_desde_tabla}")
                estado_desde_tabla = ""
                try:
                    estado_btn = row.find_element(By.XPATH, './/button[contains(@class,"btn-rounded")]')
                    estado_desde_tabla = estado_btn.text.strip()
                except Exception as e:
                    print(f"⚠️ No se pudo extraer estado desde la tabla: {e}")
                editar_btns = row.find_elements(By.XPATH, './/a[contains(@aria-label,"Editar pedido")]')
                if not editar_btns:
                    print("⚠️ No hay botón Editar para este pedido, se salta")
                    continue
                driver.execute_script("arguments[0].click();", editar_btns[0])
                time.sleep(2)
                productos_list = []
                total_cantidades = 0
                try:
                    product_rows = driver.find_elements(By.XPATH, '//*[@id="productsTable"]/tr')
                    for prod_row in product_rows:
                        try:
                            prod_input = prod_row.find_element(By.XPATH, './td[1]/input')
                            cant_input = prod_row.find_element(By.XPATH, './td[2]/input')
                            productos_list.append(prod_input.get_attribute('value'))
                            valor_cant = cant_input.get_attribute('value')
                            if valor_cant:
                                v = valor_cant.strip().replace(',', '.')
                                try:
                                    total_cantidades += int(float(v))
                                except:
                                    pass
                        except Exception as e:
                            print(f"⚠️ Error leyendo fila de producto: {e}")
                    productos_str = " | ".join(productos_list)
                except Exception as e:
                    productos_str = ""
                    total_cantidades = 0
                    print(f"⚠️ Error extrayendo productos del pedido: {e}")
                try:
                    numero_raw = row.find_element(By.XPATH, './td[2]').text.strip()
                except Exception:
                    numero_raw = safe_get('//*[@id="page-topbar"]/div/div[1]/div[2]/h4', text=True)
                match = re.search(r'\d+', numero_raw)
                numero_pedido_tabla = match.group(0) if match else numero_raw
                pedido = {
                    "Tienda": safe_get('//*[@id="input-store"]'),
                    "Numero de Pedido": numero_pedido_tabla,
                    "Nombre completo": safe_get('//*[@id="input-name"]'),
                    "Fecha de Pedido": fecha_desde_tabla,
                    "Dirección": safe_get('//*[@id="input-address"]'),
                    "Dirección 2": safe_get('//*[@id="input-address_2"]'),
                    "Teléfono": safe_get('//*[@id="input-phone"]'),
                    "Email": safe_get('//*[@id="input-email"]'),
                    "Ciudad": safe_get('//*[@id="input-city"]'),
                    "Provincia": safe_get('//*[@id="input-province"]'),
                    "Código Postal": safe_get('//*[@id="input-zip"]'),
                    "Método de envío": safe_get('//*[@id="input-shipping_method"]'),
                    "Productos": productos_str,
                    "Cantidades": total_cantidades,
                    "Coste productos (sin IVA)": safe_get('//*[@id="products_cost"]', text=True),
                    "Coste logística": safe_get('//*[@id="shipping_cost"]', text=True),
                    "Coste total (IVA incl)": safe_get('//*[@id="total_cost"]', text=True),
                    "Importe reembolso": safe_get('//*[@id="order_total_amount"]', text=True),
                    "Beneficio Dropshipper": safe_get('//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div[1]/div[1]/div[2]/div/table/tbody[2]/tr[5]/td[2]/button', text=True),
                    "Código de seguimiento": safe_get('//*[@id="input-tracking_code"]'),
                    "ID de pedido externo": safe_get('//*[@id="input-original_order_id"]'),
                    "ID externo": safe_get('//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div[2]/div/div[2]/div[17]/input'),
                    "Estado": estado_desde_tabla
                }
                df = pd.concat([df, pd.DataFrame([pedido])], ignore_index=True)
                print(f"   ✅ Pedido {pedido['Numero de Pedido']} extraído")
                driver.back()
                time.sleep(2)
            except StaleElementReferenceException:
                print("⚠️ Elemento obsoleto, se reintenta")
                continue
            except Exception as e:
                print(f"⚠️ Error procesando pedido: {e}")
                driver.back()
                time.sleep(1)
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"💾 Archivo guardado en {OUTPUT_FILE}")
try:
    load_cookies()
    scrape()
finally:
    driver.quit()
