from selenium import webdriver
import os
import pickle
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time

REFERENCIAS = ['2382501XJ0128S0001MM', '8970002WJ9187B0001BO']
URL = "https://www1.sedecatastro.gob.es/Cartografia/mapa.aspx?buscar=S"
XPATH_CUADRO_BUSQUEDA = "/html/body/form/fieldset/div[2]/div/div/div[4]/div/div[1]/div/div/div[1]/fieldset/div[1]/div[2]/input"
XPATH_BOTON_DATOS = '//*[@id="ctl00_Contenido_btnDatos"]'
XPATH_BOTON_CARTOGRAFIA = '//*[@id="ctl00_Contenido_btnNuevaCartografia"]'
COOKIES_PICKLE = "cookies_catastro.pkl"

DOWNLOAD_DIR = os.path.abspath("descargas_temp")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def save_cookies(driver, filename):
    with open(filename, "wb") as f:
        pickle.dump(driver.get_cookies(), f)

def load_cookies(driver, filename):
    with open(filename, "rb") as f:
        cookies = pickle.load(f)
        for cookie in cookies:
            if isinstance(cookie.get('expiry', None), float):
                cookie['expiry'] = int(cookie['expiry'])
            driver.add_cookie(cookie)

chrome_options = Options()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-features=VizDisplayCompositor")
chrome_options.add_argument("--remote-debugging-port=9222")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 2)

driver.get(URL)
if os.path.exists(COOKIES_PICKLE):
    load_cookies(driver, COOKIES_PICKLE)
    driver.refresh()
    # Comprobar si el botón de aceptar cookies sigue visible tras cargar las cookies
    try:
        wait_cookie = WebDriverWait(driver, 3)
        accept_btn = wait_cookie.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[aria-label="allow cookies"]'))
        )
        accept_btn.click()
        save_cookies(driver, COOKIES_PICKLE)  # Guardar nuevo pickle si fue necesario aceptar
    except (TimeoutException, NoSuchElementException):
        pass
else:
    try:
        wait_cookie = WebDriverWait(driver, 3)
        accept_btn = wait_cookie.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[aria-label="allow cookies"]'))
        )
        accept_btn.click()
        save_cookies(driver, COOKIES_PICKLE)
    except (TimeoutException, NoSuchElementException):
        pass

for idx, ref in enumerate(REFERENCIAS, start=1):
    resultados = []
    carpeta_ref = os.path.abspath(ref)
    os.makedirs(carpeta_ref, exist_ok=True)

    try:
        driver.get(URL)
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            driver.switch_to.frame(iframes[0])

        search_box = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_CUADRO_BUSQUEDA)))
        driver.execute_script("arguments[0].value = arguments[1];", search_box, ref)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", search_box)
        search_box.send_keys('\n')

        boton_cartografia = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_BOTON_CARTOGRAFIA)))
        boton_cartografia.click()
        try:
            driver.switch_to.default_content()  # Salir del iframe antes de buscar el botón de capas
            boton_capas = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnCapasC"]')))
            boton_capas.click()
            try:
                boton_PNOE = wait.until(EC.element_to_be_clickable((By.ID, 'aPNOA')))
                boton_PNOE.click()
                boton_mapa = wait.until(EC.element_to_be_clickable((By.ID, 'IBImprimir')))
                boton_mapa.click()

                boton_imprimir = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_Contenido_bImprimir")))
                boton_imprimir.click()

                # Esperar a que se inicie la descarga del PDF
            except Exception as e_btn:
                print(f"No se pudo pulsar el botón Imprimir: {e_btn}")
            except Exception as e_inner:
                print(f"No se pudo pulsar el botón PNOA o el botón Mapa: {e_inner}")
        except Exception as e:
            print(f"No se pudo pulsar el botón de capas cartográficas: {e}")
        time.sleep(0.5)
    except Exception as e:
        print(f"Error procesando referencia {ref}: {e}")
        