# -*- coding: utf-8 -*-
import os
import pickle
import shutil
import re
import requests
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

REFERENCIAS = ['2382501XJ0128S0001MM', '8970002WJ9187B0001BO']
URL = "https://www1.sedecatastro.gob.es/Cartografia/mapa.aspx?buscar=S"
XPATH_CUADRO_BUSQUEDA = "/html/body/form/fieldset/div[2]/div/div/div[4]/div/div[1]/div/div/div[1]/fieldset/div[1]/div[2]/input"
XPATH_BOTON_DATOS = '//*[@id="ctl00_Contenido_btnDatos"]'
XPATH_BOTON_CARTOGRAFIA = '//*[@id="ctl00_Contenido_btnNuevaCartografia"]'
COOKIES_PICKLE = "cookies_catastro.pkl"

DOWNLOAD_DIR = os.path.abspath("descargas_temp")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def limpiar_descargas(download_dir):
    """Elimina todos los archivos de la carpeta de descargas temporal"""
    for f in os.listdir(download_dir):
        file_path = os.path.join(download_dir, f)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception:
            pass

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
chrome_options.add_argument("--window-position=0,0")         # Arriba a la izquierda de la pantalla
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-features=VizDisplayCompositor")
chrome_options.add_argument("--remote-debugging-port=9222")
chrome_options.add_argument("--disable-notifications")



driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 3)

driver.get(URL)
if os.path.exists(COOKIES_PICKLE):
    load_cookies(driver, COOKIES_PICKLE)
    driver.refresh()
    try:
        wait_cookie = WebDriverWait(driver, 1.5)
        accept_btn = wait_cookie.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[aria-label="allow cookies"]'))
        )
        accept_btn.click()
        save_cookies(driver, COOKIES_PICKLE)
    except (TimeoutException, NoSuchElementException):
        pass
else:
    try:
        wait_cookie = WebDriverWait(driver, 1.5)
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

        # --- DATOS CATASTRALES ---
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_CUADRO_BUSQUEDA)))
        driver.execute_script("arguments[0].value = arguments[1];", search_box, ref)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", search_box)
        search_box.send_keys('\n')

        boton_datos = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_BOTON_DATOS)))
        boton_datos.click()
        time.sleep(0.5)
        clase = driver.execute_script('return document.querySelector("#ctl00_Contenido_tblInmueble > div:nth-child(3) > div > span > label")?.innerText || "NA";')
        uso = driver.execute_script('return document.querySelector("#ctl00_Contenido_tblInmueble > div:nth-child(4) > div > span > label")?.innerText || "NA";')
        superficie = driver.execute_script('return document.querySelector("#ctl00_Contenido_tblInmueble > div:nth-child(5) > div > span > label")?.innerText || "NA";')
        anio_construccion = driver.execute_script('return document.querySelector("#ctl00_Contenido_tblInmueble > div:nth-child(6) > div > span > label")?.innerText || "NA";')
        direccion = driver.execute_script('return document.querySelector("#ctl00_Contenido_tblInmueble > div:nth-child(2) > div > span > label")?.innerText || "NA";')
        resultados.append({
            'Referencia': ref,
            'Clase': clase,
            'Uso': uso,
            'Superficie': superficie,
            'Año_Construccion': anio_construccion,
            'Dirección': direccion
        })

        # --- PDF Croquis y Datos ---
        limpiar_descargas(DOWNLOAD_DIR)
        enlace = wait.until(EC.element_to_be_clickable((By.ID, "BImpCroquisYDatos")))
        enlace.click()
        time.sleep(2)

        # Espera a que aparezca el PDF en la carpeta de descargas
        pdf_file = None
        for _ in range(20):
            files = [f for f in os.listdir(DOWNLOAD_DIR) if f.lower().endswith(".pdf")]
            if files:
                pdf_file = files[0]
                break
            time.sleep(1)
        if pdf_file:
            src = os.path.join(DOWNLOAD_DIR, pdf_file)
            dst = os.path.join(carpeta_ref, "croquis_y_datos.pdf")
            shutil.move(src, dst)
            print(f"✅ [{idx}/{len(REFERENCIAS)}] PDF Croquis y Datos guardado.")
        else:
            print(f"❌ [{idx}/{len(REFERENCIAS)}] No se descargó el PDF Croquis y Datos.")

        # --- Guardar TXT con datos ---
        with open(os.path.join(carpeta_ref, "resultados_catastro.txt"), "w", encoding="utf-8") as f:
            for item in resultados:
                f.write(
                    f"Referencia: {item['Referencia']}\n"
                    f"Clase: {item['Clase']}\n"
                    f"Uso: {item['Uso']}\n"
                    f"Superficie: {item['Superficie']}\n"
                    f"Año_Construccion: {item['Año_Construccion']}\n"
                    f"Dirección: {item['Dirección']}\n"
                    "--------------------------\n"
                )

        # --- Descargar foto de fachada JUSTO AQUÍ ---
        del_val = mun_val = None
        for _ in range(10):
            enlaces = driver.find_elements(By.TAG_NAME, "a")
            for a in enlaces:
                href = a.get_attribute("href")
                if href and "OVCConCiud.aspx" in href and "del=" in href and "mun=" in href:
                    match = re.search(r'del=(\d+)&mun=(\d+)', href)
                    if match:
                        del_val = match.group(1)
                        mun_val = match.group(2)
                        print(f"del: {del_val}, mun: {mun_val}")
                        break
            if del_val and mun_val:
                break
            time.sleep(0.5)

        if del_val and mun_val:
            url_foto = f"https://ovc.catastro.meh.es/OVCServWeb/OVCWcfLibres/OVCFotoFachada.svc/RecuperarFotoFachadaGet?del={del_val}&mun={mun_val}&ReferenciaCatastral={ref}"
            foto_path = os.path.join(carpeta_ref, "foto_fachada.jpg")
            try:
                resp = requests.get(url_foto, timeout=10)
                if resp.status_code == 200 and resp.content:
                    with open(foto_path, "wb") as f:
                        f.write(resp.content)
                    print(f"✅ Foto fachada guardada para referencia: {ref}")
                else:
                    print(f"❌ No se pudo descargar la foto de fachada para referencia: {ref}")
            except Exception as e:
                print(f"❌ Error descargando foto fachada: {e}")
        else:
            print(f"❌ No se pudo extraer del y mun para referencia: {ref}")

        # --- PDF CARTOGRAFÍA (MAPA/IMPRIMIR) ---
        driver.switch_to.default_content()
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
        driver.switch_to.default_content()
        boton_capas = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnCapasC"]')))
        boton_capas.click()
        time.sleep(0.5)
        boton_PNOE = wait.until(EC.element_to_be_clickable((By.ID, 'aPNOA')))
        boton_PNOE.click()
        time.sleep(0.1)
        boton_mapa = wait.until(EC.element_to_be_clickable((By.ID, 'IBImprimir')))
        boton_mapa.click()
        boton_imprimir = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_Contenido_bImprimir")))

        limpiar_descargas(DOWNLOAD_DIR)  # <<-- LIMPIA ANTES DE HACER LA SEGUNDA DESCARGA
        boton_imprimir.click()
        # Espera a que aparezca el nuevo PDF de cartografía
        pdf_cartografia = None
        for _ in range(20):
            files = [f for f in os.listdir(DOWNLOAD_DIR) if f.lower().endswith(".pdf")]
            if files:
                pdf_cartografia = files[0]
                break
            time.sleep(1)
        if pdf_cartografia:
            src = os.path.join(DOWNLOAD_DIR, pdf_cartografia)
            dst = os.path.join(carpeta_ref, "cartografia.pdf")
            shutil.move(src, dst)
            print(f"✅ [{idx}/{len(REFERENCIAS)}] PDF Cartografía guardado.")
        else:
            print(f"❌ [{idx}/{len(REFERENCIAS)}] No se descargó el PDF de cartografía.")

        driver.switch_to.default_content()

    except Exception as e:
        print(f"❌ [{idx}/{len(REFERENCIAS)}] Error en referencia {ref}: {e}")
        driver.switch_to.default_content()
        continue

driver.quit()

