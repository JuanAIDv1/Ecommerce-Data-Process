import os
import json
import time
import unicodedata
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
ENV_PATH = r"C:/Ingestador/configs/Login.env"
COOKIES_PATH = r"C:/Ingestador/cookies/dropiro.json"
EXCEL_PATH = r"C:/Ingestador/output/pedidos_dropipro.xlsx"
LOGIN_URL = "https://dropipro.com"
USER_FALLBACK = "starttok.y@gmail.com"
PWD_FALLBACK = "Proyectoy25"
COLUMNS_EXPECTED = [
    "Tienda", "Numero de Pedido", "Nombre completo", "Fecha de Pedido",
    "Dirección", "Dirección 2", "Teléfono", "Email", "Ciudad", "Provincia",
    "Código Postal", "Método de envío", "Productos", "Cantidades",
    "Coste productos (sin IVA)", "Coste logística", "Coste total (IVA incl)",
    "Importe reembolso", "Beneficio Dropshipper", "Código de seguimiento",
    "ID de pedido externo", "ID externo", "Estado"
]
ESTADOS_INICIALES = [
    "Nuevo", "Pendiente de confirmación", "Confirmado",
    "Carrito Abandonado", "Preparado"
]
ESTADOS_OPERACIONALES = [
    "Enviado", "En Ruta", "Incidencia", "Pedido Aplazado",
    "Rehusado", "Duplicado", "No Confirmable"
]
XPATH_SEARCH = '//*[@id="searchForm"]/div/div[4]/input'
XPATH_TABLE_STATUS = '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div/div[2]/div[1]/table/tbody/tr[1]/td[2]/button'
XPATH_EDIT_BTN = '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div/div[2]/div[1]/table/tbody/tr[1]/td[8]/a[1]'
XPATH_PRODUCTS_ROWS = '//*[@id="productsTable"]/tr'
FIELDS_XPATH = {
    "Tienda": '//*[@id="input-store"]',
    "Nombre completo": '//*[@id="input-name"]',
    "Dirección": '//*[@id="input-address"]',
    "Dirección 2": '//*[@id="input-address_2"]',
    "Teléfono": '//*[@id="input-phone"]',
    "Email": '//*[@id="input-email"]',
    "Ciudad": '//*[@id="input-city"]',
    "Provincia": '//*[@id="input-province"]',
    "Código Postal": '//*[@id="input-zip"]',
    "Método de envío": '//*[@id="input-shipping_method"]',
    "Coste productos (sin IVA)": '//*[@id="products_cost"]',
    "Coste logística": '//*[@id="shipping_cost"]',
    "Coste total (IVA incl)": '//*[@id="total_cost"]',
    "Importe reembolso": '//*[@id="order_total_amount"]',
    "Beneficio Dropshipper": '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div[1]/div[1]/div[2]/div/table/tbody[2]/tr[5]/td[2]/button',
    "Código de seguimiento": '//*[@id="input-tracking_code"]',
    "ID de pedido externo": '//*[@id="input-original_order_id"]',
    "ID externo": '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div[2]/div/div[2]/div[17]/input'
}
def norm(text: str) -> str:
    if text is None:
        return ""
    s = str(text).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()
def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)
    return driver, wait
def get_credentials():
    if os.path.exists(ENV_PATH):
        load_dotenv(ENV_PATH)
        user = os.getenv("DROPI_USER")
        pwd = os.getenv("DROPI_PASS")
        if user and pwd:
            return user, pwd
    return USER_FALLBACK, PWD_FALLBACK
def save_excel_safe(df: pd.DataFrame, path: str):
    """Guardar Excel, si está en uso crear copia con timestamp."""
    try:
        df.to_excel(path, index=False)
        print(f"💾 Excel guardado: {path}")
    except PermissionError:
        alt = path.replace(".xlsx", f"_{int(time.time())}.xlsx")
        try:
            df.to_excel(alt, index=False)
            print(f"⚠️ Excel estaba en uso. Guardado en copia: {alt}")
        except Exception as e:
            print(f"❌ No se pudo guardar ni en copia: {e}")
def safe_get(driver, xpath, text=False):
    """Obtener texto o value del elemento, tolerante a errores."""
    try:
        el = driver.find_element(By.XPATH, xpath)
        if text:
            return el.text.strip()
        return (el.get_attribute("value") or "").strip()
    except Exception:
        return ""
def try_load_cookies(driver, wait):
    if not os.path.exists(COOKIES_PATH):
        return False
    try:
        driver.get(LOGIN_URL)
        with open(COOKIES_PATH, "r", encoding="utf-8") as f:
            cookies = json.load(f)
        for c in cookies:
            try:
                driver.add_cookie(c)
            except Exception:
                continue
        driver.get("https://dropipro.com/admin/index")
        time.sleep(2)
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="side-menu"]/li[7]/a/span')))
            print("✅ Sesión restaurada con cookies.")
            return True
        except TimeoutException:
            print("⚠️ Cookies inválidas o sesión expirada.")
            return False
    except Exception as e:
        print(f"⚠️ Error cargando cookies: {e}")
        return False
        if "/login" in driver.current_url:
            return False
        print("✅ Sesión restaurada con cookies.")
        return True
    except Exception as e:
        print(f"⚠️ No se pudo restaurar sesión con cookies: {e}")
        return False
def login_with_credentials(driver, wait, user, pwd):
    driver.get(LOGIN_URL)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-username"]')))
        driver.find_element(By.XPATH, '//*[@id="input-username"]').clear()
        driver.find_element(By.XPATH, '//*[@id="input-username"]').send_keys(user)
        driver.find_element(By.XPATH, '//*[@id="input-password"]').clear()
        driver.find_element(By.XPATH, '//*[@id="input-password"]').send_keys(pwd)
        btn = driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/div/div[2]/div[1]/div/div[2]/form/div[4]/button')
        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        driver.execute_script("arguments[0].click();", btn)
        wait.until(lambda d: "/login" not in d.current_url)
        os.makedirs(os.path.dirname(COOKIES_PATH), exist_ok=True)
        with open(COOKIES_PATH, "w", encoding="utf-8") as f:
            json.dump(driver.get_cookies(), f)
        print("✅ Login correcto (credenciales).")
        return True
    except Exception as e:
        print(f"❌ Error en login: {e}")
        return False
def navigate_to_pedidos(driver, wait):
    try:
        menu = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="side-menu"]/li[7]/a/span')))
        driver.execute_script("arguments[0].click();", menu)
        wait.until(EC.presence_of_element_located((By.XPATH, '//table')))
        print("✅ Navegado a Pedidos -> TODOS")
    except Exception as e:
        print(f"❌ No se pudo navegar a Pedidos: {e}")
        raise
def seleccionar_todos_estados_y_tiendas(driver, wait):
    try:
        estados = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="status_select"]')))
        try:
            estados.find_element(By.XPATH, './/option[@value=""]').click()
        except Exception:
            driver.execute_script("arguments[0].value = '';", estados)
        print("✅ 'Todos' en estados seleccionado")
    except Exception as e:
        print(f"⚠️ No se pudo seleccionar estados: {e}")
    try:
        tiendas = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="store_select"]')))
        try:
            tiendas.find_element(By.XPATH, './/option[@value=""]').click()
        except Exception:
            driver.execute_script("arguments[0].value = '';", tiendas)
        print("✅ 'Todas' las tiendas seleccionado")
    except Exception as e:
        print(f"⚠️ No se pudo seleccionar tiendas: {e}")
def load_and_filter_excel():
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ No existe el archivo Excel en: {EXCEL_PATH}")
        return None, None
    try:
        df = pd.read_excel(EXCEL_PATH, dtype=str)
        for col in COLUMNS_EXPECTED:
            if col not in df.columns:
                print(f"❌ Columna esperada no encontrada en Excel: '{col}'")
                return None, None
        df = df.fillna("")
        estados_norm = {norm(s) for s in ESTADOS_INICIALES}
        mask = df["Estado"].apply(lambda s: norm(s) in estados_norm)
        df_filtrado = df[mask].copy()
        if df_filtrado.empty:
            print("ℹ️ No se encontraron órdenes en estados iniciales en el Excel.")
        else:
            print(f"ℹ️ {len(df_filtrado)} órdenes en estados iniciales encontradas.")
        return df, df_filtrado
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return None, None
def get_table_estado_after_search(driver, wait):
    try:
        el = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_TABLE_STATUS)))
        return el.text.strip()
    except Exception:
        return ""
def click_edit_for_first_row(driver, wait):
    try:
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_EDIT_BTN)))
        driver.execute_script("arguments[0].click();", btn)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-name"]')))
        return True
    except Exception:
        return False
def scrape_fields_from_edit(driver, wait):
    """Extrae todos los FIELDS_XPATH + Productos/Cantidades."""
    result = {}
    try:
        prod_rows = driver.find_elements(By.XPATH, XPATH_PRODUCTS_ROWS)
        productos = []
        cantidades = []
        for r in prod_rows:
            try:
                name_el = r.find_element(By.XPATH, './td[1]/input')
                qty_el = r.find_element(By.XPATH, './td[2]/input')
                name = (name_el.get_attribute("value") or "").strip()
                qty = (qty_el.get_attribute("value") or "").strip()
                if name:
                    productos.append(name)
                if qty:
                    cantidades.append(qty)
            except Exception:
                continue
        result["Productos"] = " | ".join(productos)
        total_cant = 0
        is_num = True
        for q in cantidades:
            try:
                total_cant += int(float(q.replace(',', '.')))
            except Exception:
                is_num = False
                break
        result["Cantidades"] = total_cant if is_num else " | ".join(cantidades)
    except Exception:
        result["Productos"] = ""
        result["Cantidades"] = ""
    for col, xpath in FIELDS_XPATH.items():
        if col in ("Productos", "Cantidades"):
            continue
        v = safe_get(driver, xpath, text=False)
        if not v:
            v = safe_get(driver, xpath, text=True)
        result[col] = v or ""
    return result
def process_each_order(driver, wait, df, df_filtrado):
    """Itera pedidos en iniciales: actualiza Estado si cambió y rellena campos vacíos."""
    if df is None or df_filtrado is None:
        return df
    seleccionar_todos_estados_y_tiendas(driver, wait)
    time.sleep(1)
    for idx, row in df_filtrado.iterrows():
        pedido_id = str(row["Numero de Pedido"]).strip().replace("…", "")
        estado_excel = str(row["Estado"]).strip()
        if not pedido_id:
            print(f"⚠️ Fila {idx} sin Numero de Pedido, se omite.")
            continue
        print(f"🔍 (F1) Buscando pedido {pedido_id} ...")
        try:
            search_box = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_SEARCH)))
            try:
                search_box.clear()
            except Exception:
                driver.execute_script("arguments[0].value = '';", search_box)
            search_box.send_keys(pedido_id)
            search_box.send_keys(Keys.ENTER)
        except Exception as e:
            print(f"❌ No se pudo usar la barra de búsqueda para {pedido_id}: {e}")
            continue
        time.sleep(2)
        estado_dropi = get_table_estado_after_search(driver, wait)
        if not estado_dropi:
            print(f"❌ Pedido {pedido_id}: no apareció en la vista tras búsqueda.")
            try:
                sb = driver.find_element(By.XPATH, XPATH_SEARCH)
                driver.execute_script("arguments[0].value = '';", sb)
            except Exception:
                pass
            continue
        print(f"   Estado DropiPro = '{estado_dropi}'  |  Estado Excel = '{estado_excel}'")
        if norm(estado_dropi) == norm(estado_excel):
            print(f"   ✅ Pedido {pedido_id}: estado coincide. Se omite scraping de campos.")
            try:
                sb = driver.find_element(By.XPATH, XPATH_SEARCH)
                driver.execute_script("arguments[0].value = '';", sb)
            except Exception:
                pass
            continue
        opened = click_edit_for_first_row(driver, wait)
        if not opened:
            print(f"❌ Pedido {pedido_id}: no se pudo abrir 'Editar'. Se omite.")
            try:
                sb = driver.find_element(By.XPATH, XPATH_SEARCH)
                driver.execute_script("arguments[0].value = '';", sb)
            except Exception:
                pass
            continue
        scraped = scrape_fields_from_edit(driver, wait)
        updated_any = False
        if norm(estado_dropi) != norm(estado_excel):
            df.at[idx, "Estado"] = estado_dropi
            updated_any = True
            print(f"   📝 Estado actualizado: {estado_excel} -> {estado_dropi}")
        for col in COLUMNS_EXPECTED:
            if col not in df.columns:
                continue
            if col == "Estado":
                continue
            current = str(df.at[idx, col]).strip()
            scraped_val = scraped.get(col, "")
            if (current == "" or current.lower() == "nan") and scraped_val:
                df.at[idx, col] = scraped_val
                updated_any = True
                print(f"   📝 Rellenado '{col}' = {scraped_val}")
        try:
            driver.back()
            wait.until(EC.presence_of_element_located((By.XPATH, '//table')))
            time.sleep(1)
        except Exception:
            try:
                driver.refresh()
                wait.until(EC.presence_of_element_located((By.XPATH, '//table')))
            except Exception:
                pass
        try:
            sb = driver.find_element(By.XPATH, XPATH_SEARCH)
            driver.execute_script("arguments[0].value = '';", sb)
        except Exception:
            pass
        if updated_any:
            print(f"   ✅ Pedido {pedido_id}: actualizados campos en memoria.")
        else:
            print(f"   ℹ️ Pedido {pedido_id}: no se encontraron campos vacíos para rellenar.")
    return df
def validar_estados_operacionales(driver, wait):
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ No existe el archivo Excel en: {EXCEL_PATH}")
        return
    try:
        df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
    except Exception as e:
        print(f"❌ Error leyendo Excel en fase 2: {e}")
        return
    estados_norm = {norm(s) for s in ESTADOS_OPERACIONALES}
    mask = df["Estado"].apply(lambda s: norm(s) in estados_norm)
    df_filtrado = df[mask].copy()
    if df_filtrado.empty:
        print("ℹ️ No se encontraron órdenes en estados operacionales.")
        return
    print(f"ℹ️ {len(df_filtrado)} órdenes en estados operacionales encontradas.")
    seleccionar_todos_estados_y_tiendas(driver, wait)
    time.sleep(1)
    cambios = 0
    for idx, row in df_filtrado.iterrows():
        pedido_id = str(row["Numero de Pedido"]).strip().replace("…", "")
        estado_excel = str(row["Estado"]).strip()
        if not pedido_id:
            print(f"⚠️ Fila {idx} sin Numero de Pedido, se salta.")
            continue
        print(f"🔍 (F2) Revisando pedido {pedido_id} ...")
        try:
            search_box = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_SEARCH)))
            try:
                search_box.clear()
            except Exception:
                driver.execute_script("arguments[0].value = '';", search_box)
            search_box.send_keys(pedido_id)
            search_box.send_keys(Keys.ENTER)
        except Exception as e:
            print(f"❌ No se pudo usar la barra de búsqueda (Fase2) para {pedido_id}: {e}")
            continue
        time.sleep(2)
        estado_dropi = get_table_estado_after_search(driver, wait)
        if not estado_dropi:
            print(f"   ❌ Pedido {pedido_id}: no apareció tras búsqueda (F2).")
            try:
                sb = driver.find_element(By.XPATH, XPATH_SEARCH)
                driver.execute_script("arguments[0].value = '';", sb)
            except Exception:
                pass
            continue
        print(f"   Estado DropiPro = '{estado_dropi}'  |  Estado Excel = '{estado_excel}'")
        if norm(estado_dropi) != norm(estado_excel):
            df.at[idx, "Estado"] = estado_dropi
            cambios += 1
            print(f"   📝 Actualizado 'Estado' -> {estado_dropi}")
        try:
            sb = driver.find_element(By.XPATH, XPATH_SEARCH)
            driver.execute_script("arguments[0].value = '';", sb)
        except Exception:
            pass
    save_excel_safe(df, EXCEL_PATH)
    print(f"✅ Fase 2 completada. Cambios en estados operacionales: {cambios}")
def main():
    print("🚀 Iniciando DropiPro Bot (fase 1 + fase 2)...")
    driver, wait = setup_driver()
    try:
        if not try_load_cookies(driver, wait):
            user, pwd = get_credentials()
            ok = login_with_credentials(driver, wait, user, pwd)
            if not ok:
                print("❌ No fue posible iniciar sesión. Termina ejecución.")
                return
        navigate_to_pedidos(driver, wait)
        df, df_filtrado = load_and_filter_excel()
        if df is None:
            print("❌ Error leyendo Excel. Termina ejecución.")
            return
        if df_filtrado is None or df_filtrado.empty:
            print("ℹ️ No hay órdenes en estados iniciales. Se procede a Fase 2.")
        else:
            print("🚀 Ejecutando Fase 1: completar campos vacíos y actualizar estados (iniciales)...")
            df_updated = process_each_order(driver, wait, df, df_filtrado)
            save_excel_safe(df_updated, EXCEL_PATH)
        print("\n🚀 Iniciando Fase 2: validar estados operacionales...")
        validar_estados_operacionales(driver, wait)
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        print("🛑 Proceso finalizado.")
if __name__ == "__main__":
    main()
