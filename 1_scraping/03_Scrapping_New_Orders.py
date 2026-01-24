import os
import re
import json
import time
import unicodedata
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
)
from webdriver_manager.chrome import ChromeDriverManager
URL_BASE = "https://dropipro.com/"
COOKIES_PATH = r"C:\Ingestador\cookies\dropipro\dropipro.json"
ENV_PATH = r"C:\Ingestador\configs\Login.env"
EXCEL_PATH = r"C:\Ingestador\output\pedidos_dropipro.xlsx"
USER_FALLBACK = "***********@*********.com"
PWD_FALLBACK = "*********"
XPATH_SIDE_MENU_PEDIDOS_TODOS = '//*[@id="side-menu"]/li[7]/a/span'
XPATH_TABLE_ROWS = '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div/div[2]/div[1]/table/tbody/tr'
XPATH_ROW_NUMBER = lambda r: f'({XPATH_TABLE_ROWS})[{r}]/td[1]'
XPATH_ROW_DATE = lambda r: f'({XPATH_TABLE_ROWS})[{r}]/td[5]'
XPATH_ROW_STATUS = lambda r: f'({XPATH_TABLE_ROWS})[{r}]/td[2]/button'
XPATH_EDIT_IN_ROW = lambda r: f'({XPATH_TABLE_ROWS})[{r}]/td[8]/a[1]'
XPATH_PAGINATOR_NEXT = '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/nav/ul/li[15]/a'
XPATH_PAGINATOR_PREV = '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/nav/ul/li[1]/a'
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
    "ID externo": '//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div[2]/div/div[2]/div[17]/input',
    "Productos_rows": '//*[@id="productsTable"]/tr/td[1]/input',
    "Cantidades_rows": '//*[@id="productsTable"]/tr/td[2]/input',
}
COLUMNS_EXPECTED = [
    "Tienda", "Numero de Pedido", "Nombre completo", "Fecha de Pedido",
    "Dirección", "Dirección 2", "Teléfono", "Email", "Ciudad", "Provincia",
    "Código Postal", "Método de envío", "Productos", "Cantidades", "Coste productos (sin IVA)",
    "Coste logística", "Coste total (IVA incl)", "Importe reembolso",
    "Beneficio Dropshipper", "Código de seguimiento",
    "ID de pedido externo", "ID externo", "Estado"
]
def norm(text):
    if text is None:
        return ""
    s = str(text).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s
def limpiar_numero_pedido(raw):
    """Extrae el número de pedido como entero desde un string (depura '…' y caracteres)"""
    if raw is None:
        return None
    s = re.search(r"\d{5,}", str(raw))
    if s:
        try:
            return int(s.group())
        except:
            return None
    return None
def save_excel_safe(df, path):
    """Intenta sobrescribir; si está en uso, guarda copia con timestamp."""
    try:
        df.to_excel(path, index=False)
        print(f"💾 Excel guardado: {path}")
    except PermissionError:
        alt = path.replace(".xlsx", f"_{int(time.time())}.xlsx")
        try:
            df.to_excel(alt, index=False)
            print(f"⚠️ Archivo estaba en uso. Guardado en copia: {alt}")
        except Exception as e:
            print(f"❌ No se pudo guardar ni en copia: {e}")
def safe_get_text_or_value(driver, xpath, wait=None, only_text=False, retries=3, delay=0.5):
    """Intenta obtener text/valor de un elemento por xpath de forma robusta."""
    for attempt in range(retries):
        try:
            el = driver.find_element(By.XPATH, xpath)
            tag = el.tag_name.lower() if el is not None else ""
            if only_text or tag in ("button", "span", "div", "label", "p"):
                return el.text.strip()
            val = el.get_attribute("value")
            if val is None:
                return el.text.strip()
            return str(val).strip()
        except (StaleElementReferenceException, NoSuchElementException):
            time.sleep(delay)
            continue
        except Exception:
            time.sleep(delay)
            continue
    return ""
def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 18)
    return driver, wait
driver, wait = setup_driver()
def login(driver, wait):
    driver.get(URL_BASE)
    time.sleep(1)
    if os.path.exists(COOKIES_PATH):
        try:
            with open(COOKIES_PATH, "r", encoding="utf-8") as f:
                cookies = json.load(f)
            for c in cookies:
                try:
                    driver.add_cookie(c)
                except Exception:
                    continue
            driver.refresh()
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, XPATH_SIDE_MENU_PEDIDOS_TODOS)))
                print("✅ Sesión restaurada con cookies.")
                return True
            except TimeoutException:
                pass
        except Exception as e:
            print(f"⚠️ Error cargando cookies: {e}")
    try:
        if os.path.exists(ENV_PATH):
            load_dotenv(ENV_PATH)
            user = os.getenv("DROPI_USER")
            pwd = os.getenv("DROPI_PASS")
        else:
            user = None
            pwd = None
    except Exception:
        user = None
        pwd = None
    if user and pwd:
        try:
            driver.get(URL_BASE)
            el_user = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-username"]')))
            el_pass = driver.find_element(By.XPATH, '//*[@id="input-password"]')
            el_user.clear()
            el_user.send_keys(user)
            el_pass.clear()
            el_pass.send_keys(pwd)
            try:
                btn = driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/div/div[2]/div[1]/div/div[2]/form/div[4]/button')
                driver.execute_script("arguments[0].click();", btn)
            except Exception:
                el_pass.send_keys(Keys.ENTER)
            wait.until(EC.presence_of_element_located((By.XPATH, XPATH_SIDE_MENU_PEDIDOS_TODOS)))
            try:
                os.makedirs(os.path.dirname(COOKIES_PATH), exist_ok=True)
                with open(COOKIES_PATH, "w", encoding="utf-8") as f:
                    json.dump(driver.get_cookies(), f)
            except Exception:
                pass
            print("✅ Sesión iniciada exitosamente con archivo .env.")
            return True
        except Exception as e:
            print(f"⚠️ No se pudo iniciar sesión con archivo .env: {e}")
    try:
        driver.get(URL_BASE)
        el_user = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-username"]')))
        el_pass = driver.find_element(By.XPATH, '//*[@id="input-password"]')
        el_user.clear()
        el_user.send_keys(USER_FALLBACK)
        el_pass.clear()
        el_pass.send_keys(PWD_FALLBACK)
        try:
            btn = driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/div/div[2]/div[1]/div/div[2]/form/div[4]/button')
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            el_pass.send_keys(Keys.ENTER)
        wait.until(EC.presence_of_element_located((By.XPATH, XPATH_SIDE_MENU_PEDIDOS_TODOS)))
        try:
            os.makedirs(os.path.dirname(COOKIES_PATH), exist_ok=True)
            with open(COOKIES_PATH, "w", encoding="utf-8") as f:
                json.dump(driver.get_cookies(), f)
        except Exception:
            pass
        print("✅ Sesión iniciada exitosamente con credenciales fallback.")
        return True
    except Exception as e:
        print(f"❌ Error crítico: no se pudo iniciar sesión: {e}")
        return False
def load_excel_and_get_last(path):
    if not os.path.exists(path):
        print(f"❌ No existe el archivo Excel: {path}")
        return None, None
    try:
        df = pd.read_excel(path, dtype=str)
        for c in COLUMNS_EXPECTED:
            if c not in df.columns:
                print(f"⚠️ Columna esperada no encontrada en Excel: '{c}'. Continuo pero ten cuidado.")
        if df.shape[1] < 2:
            print("❌ Excel no tiene columna B con numero de pedido.")
            return df, None
        raw_series = df.iloc[:, 1].fillna("").astype(str)
        numeros = raw_series.apply(lambda x: limpiar_numero_pedido(x))
        if numeros.dropna().empty:
            return df, None
        ultimo = int(numeros.dropna().astype(int).max())
        return df, ultimo
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return None, None
def scrape_one_row(driver, wait, row_index):
    """
    row_index: 1-based index of the row in the visible table.
    Devuelve diccionario con claves exactamente iguales a COLUMNS_EXPECTED.
    """
    base = {}
    try:
        row_num_xpath = XPATH_ROW_NUMBER(row_index)
        raw_num = safe_get_text_or_value(driver, row_num_xpath, wait=wait, only_text=True)
        numero = limpiar_numero_pedido(raw_num)
        base["Numero de Pedido"] = str(numero) if numero is not None else ""
    except Exception as e:
        print(f"⚠️ No se pudo leer Numero de Pedido en fila {row_index}: {e}")
        base["Numero de Pedido"] = ""
    try:
        raw_date = safe_get_text_or_value(driver, XPATH_ROW_DATE(row_index), wait=wait, only_text=True)
        base["Fecha de Pedido"] = raw_date
    except Exception:
        base["Fecha de Pedido"] = ""
    try:
        raw_status = safe_get_text_or_value(driver, XPATH_ROW_STATUS(row_index), wait=wait, only_text=True)
        base["Estado"] = raw_status
    except Exception:
        base["Estado"] = ""
    try:
        edit_xpath = XPATH_EDIT_IN_ROW(row_index)
        btn = driver.find_element(By.XPATH, edit_xpath)
        driver.execute_script("arguments[0].click();", btn)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-name"]')))
        time.sleep(0.25)
    except Exception as e:
        print(f"❌ No se pudo abrir 'Editar' para la fila {row_index}: {e}")
        for col in COLUMNS_EXPECTED:
            if col not in base:
                base[col] = ""
        return base
    for col in COLUMNS_EXPECTED:
        if col == "Numero de Pedido" or col == "Fecha de Pedido" or col == "Estado":
            continue
        elif col == "Productos":
            try:
                prods = driver.find_elements(By.XPATH, FIELDS_XPATH["Productos_rows"])
                productos = []
                for p in prods:
                    val = (p.get_attribute("value") or "").strip()
                    if val:
                        productos.append(val)
                base["Productos"] = " | ".join(productos)
                if not productos:
                    print(f"   ℹ️ (Editar) campo 'Productos' vacío para {base.get('Numero de Pedido','?')}")
            except Exception:
                base["Productos"] = ""
                print(f"   ⚠️ No se encontró Productos (editar) para {base.get('Numero de Pedido','?')}")
        elif col == "Cantidades":
            try:
                qtys = driver.find_elements(By.XPATH, FIELDS_XPATH["Cantidades_rows"])
                cantidades_list = []
                total = 0
                any_num = False
                for q in qtys:
                    raw = (q.get_attribute("value") or "").strip()
                    if raw:
                        cantidades_list.append(raw)
                        try:
                            num = int(float(raw.replace(',', '.')))
                            total += num
                            any_num = True
                        except:
                            pass
                base["Cantidades"] = str(total) if any_num else " | ".join(cantidades_list)
                if not cantidades_list:
                    print(f"   ℹ️ (Editar) campo 'Cantidades' vacío para {base.get('Numero de Pedido','?')}")
            except Exception:
                base["Cantidades"] = ""
                print(f"   ⚠️ No se encontró Cantidades (editar) para {base.get('Numero de Pedido','?')}")
        else:
            xpath_field = FIELDS_XPATH.get(col)
            if xpath_field:
                val = safe_get_text_or_value(driver, xpath_field, wait=wait, only_text=False)
                if val == "":
                    val = safe_get_text_or_value(driver, xpath_field, wait=wait, only_text=True)
                base[col] = val
                if val == "":
                    print(f"   ℹ️ (Editar) no se encontró valor para campo '{col}' pedido {base.get('Numero de Pedido','?')}")
            else:
                base[col] = ""
    try:
        driver.back()
        wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_TABLE_ROWS)))
        time.sleep(0.2)
    except Exception:
        try:
            driver.refresh()
            wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_TABLE_ROWS)))
            time.sleep(0.2)
        except Exception:
            pass
    return base
def fase3_main():
    print("🚀 Iniciando Fase 3: Scraping de pedidos posteriores al último en Excel...")
    ok = login(driver, wait)
    if not ok:
        print("❌ No se pudo iniciar sesión. Termina ejecución.")
        driver.quit()
        return
    try:
        menu = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_SIDE_MENU_PEDIDOS_TODOS)))
        driver.execute_script("arguments[0].click();", menu)
        wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_TABLE_ROWS)))
        time.sleep(0.5)
        print("✅ Navegación a Pedidos -> TODOS completada.")
    except Exception as e:
        print(f"❌ No se pudo navegar a Pedidos -> TODOS: {e}")
        try:
            driver.get(URL_BASE + "app/orders")
            wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_TABLE_ROWS)))
            time.sleep(0.5)
            print("✅ Navegación alternativa a Pedidos -> TODOS completada.")
        except Exception as e2:
            print(f"❌ No se pudo acceder a la lista de pedidos: {e2}")
            driver.quit()
            return
    df, ultimo_pedido = load_excel_and_get_last(EXCEL_PATH)
    if df is None:
        print("❌ Error cargando Excel. Termina Fase 3.")
        driver.quit()
        return
    if ultimo_pedido is None:
        print("⚠️ No se detectó último pedido (columna B vacía o no numérica). Se considerará como 0 y se scrapeará todo lo que sea numérico.")
        ultimo_pedido = 0
    print(f"📌 Último pedido en Excel: {ultimo_pedido}")
    nuevos = []
    pagina = 1
    encontrado = False
    max_pages_guard = 500
    while not encontrado and pagina <= max_pages_guard:
        print(f"🔎 Revisando página {pagina} en DropiPro...")
        try:
            wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_TABLE_ROWS)))
        except TimeoutException:
            print("⚠️ Timeout esperando las filas de la tabla en la página actual.")
        time.sleep(0.5)
        try:
            filas = driver.find_elements(By.XPATH, XPATH_TABLE_ROWS)
        except Exception as e:
            print(f"❌ Error obteniendo filas en página {pagina}: {e}")
            filas = []
        print(f"   ➤ pedidos en página {pagina}: {len(filas)}")
        for idx in range(1, len(filas) + 1):
            raw_num = safe_get_text_or_value(driver, XPATH_ROW_NUMBER(idx), wait=wait, only_text=True)
            num = limpiar_numero_pedido(raw_num)
            if num is None:
                continue
            if num == ultimo_pedido:
                print(f"✅ Alcanzado pedido existente en Excel: {num} -> detener búsqueda adicional.")
                encontrado = True
                break
            if num > ultimo_pedido:
                print(f"   ➕ Pedido nuevo detectado: {num} (scrapeando...)")
                try:
                    new_row = scrape_one_row(driver, wait, idx)
                    if not new_row.get("Numero de Pedido"):
                        new_row["Numero de Pedido"] = str(num)
                    nuevos.append(new_row)
                except Exception as e:
                    print(f"❌ Error scrapeando pedido {num}: {e}")
                    continue
            else:
                pass
        if not encontrado:
            try:
                next_el = driver.find_element(By.XPATH, XPATH_PAGINATOR_NEXT)
                driver.execute_script("arguments[0].click();", next_el)
                pagina += 1
                time.sleep(1.2)
                continue
            except NoSuchElementException:
                print("⚠️ No hay más páginas y no se encontró el último pedido registrado en Excel.")
                break
            except Exception as e:
                print(f"⚠️ Error al avanzar de página: {e}")
                break
    if nuevos:
        print(f"✅ Se encontraron {len(nuevos)} pedidos nuevos. Lista (números): {[n.get('Numero de Pedido') for n in nuevos]}")
        try:
            original_df = pd.read_excel(EXCEL_PATH, dtype=str)
        except Exception:
            original_df = pd.DataFrame(columns=COLUMNS_EXPECTED)
        rows_to_append = []
        for item in nuevos:
            row = {col: "" for col in original_df.columns}
            row["Tienda"] = item.get("Tienda", "")
            row["Numero de Pedido"] = item.get("Numero de Pedido", "")
            row["Nombre completo"] = item.get("Nombre completo", "")
            row["Fecha de Pedido"] = item.get("Fecha de Pedido", "")
            row["Dirección"] = item.get("Dirección", "")
            row["Dirección 2"] = item.get("Dirección 2", "")
            row["Teléfono"] = item.get("Teléfono", "")
            row["Email"] = item.get("Email", "")
            row["Ciudad"] = item.get("Ciudad", "")
            row["Provincia"] = item.get("Provincia", "")
            row["Código Postal"] = item.get("Código Postal", "")
            row["Método de envío"] = item.get("Método de envío", "")
            row["Productos"] = item.get("Productos", "")
            row["Cantidades"] = item.get("Cantidades", "")
            row["Coste productos (sin IVA)"] = item.get("Coste productos (sin IVA)", "")
            row["Coste logística"] = item.get("Coste logística", "")
            row["Coste total (IVA incl)"] = item.get("Coste total (IVA incl)", "")
            row["Importe reembolso"] = item.get("Importe reembolso", "")
            row["Beneficio Dropshipper"] = item.get("Beneficio Dropshipper", "")
            row["Código de seguimiento"] = item.get("Código de seguimiento", "")
            row["ID de pedido externo"] = item.get("ID de pedido externo", "")
            row["ID externo"] = item.get("ID externo", "")
            row["Estado"] = item.get("Estado", "")
            rows_to_append.append(row)
        appended_df = pd.concat([original_df, pd.DataFrame(rows_to_append)], ignore_index=True, sort=False)
        if "Numero de Pedido" in appended_df.columns:
            appended_df["_num_for_sort"] = appended_df["Numero de Pedido"].apply(lambda x: limpiar_numero_pedido(x) if pd.notna(x) else None)
            appended_df = appended_df.sort_values(by=["_num_for_sort"], na_position="first").drop(columns=["_num_for_sort"]).reset_index(drop=True)
        else:
            print("⚠️ 'Numero de Pedido' no encontrada en el Excel final - no se ordenará.")
        save_excel_safe(appended_df, EXCEL_PATH)
        print("✅ Fase 3 completada: pedidos nuevos agregados y Excel guardado.")
    else:
        print("ℹ️ No se encontraron pedidos nuevos posteriores al último del Excel.")
    try:
        driver.quit()
    except:
        pass
if __name__ == "__main__":
    fase3_main()
