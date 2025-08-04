import xlwings as xw
import pandas as pd
import os
from scraper_catastro import scrape_catastro
import ctypes
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def completar_tabla_excel():
    # Limpieza previa de Chromes zombie (Windows)
    os.system("taskkill /f /im chromedriver.exe >nul 2>&1")
    os.system("taskkill /f /im chrome.exe >nul 2>&1")

    wb = xw.Book.caller()
    wb.app.api.WindowState = -4140  # Minimiza Excel

    sht = wb.sheets["CEE"]
    tabla = sht.api.ListObjects("Tabla1")
    data_range = tabla.DataBodyRange
    nrows = data_range.Rows.Count
    ncols = data_range.Columns.Count

    header_range = tabla.HeaderRowRange
    headers = [cell.Value for cell in header_range.Columns]
    values = data_range.Value
    if nrows == 1:
        values = [values]

    df = pd.DataFrame(values, columns=headers)

    pendientes = df[
        (df["REF CATASTRAL"].notnull()) &
        (df["REF CATASTRAL"].astype(str).str.strip() != '') &
        ((df["DIRECCION"].isnull()) | (df["DIRECCION"] == ""))
    ]

    excel_path = wb.fullname
    base_dir = os.path.dirname(excel_path)

    chrome_options = Options()
    prefs = {
        "download.default_directory": os.path.abspath("descargas_temp"),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-position=0,0")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-features=VizDisplayCompositor")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(options=chrome_options)

    try:
        referencias_defectuosas = []

        # PRIMER BUCLE: SIEMPRE se scrapean datos de cada referencia
        for idx, fila in pendientes.iterrows():
            referencia = str(fila["REF CATASTRAL"]).strip()
            if not referencia or referencia.lower() in ("nan", "none"):
                continue

            id_centro = str(fila["ID CENTRO"]).strip()
            centro = str(fila["CENTRO"]).strip()
            nombre_carpeta = f"{id_centro}_{centro}"
            carpeta = os.path.join(base_dir, nombre_carpeta)

            foto_path = os.path.join(carpeta, "foto_fachada.jpg")
            plano_path = os.path.join(carpeta, "plano.pdf")
            doc_path = os.path.join(carpeta, f"{referencia}.pdf")

            foto_ok = os.path.isfile(foto_path)
            plano_ok = os.path.isfile(plano_path)
            doc_ok = os.path.isfile(doc_path)

            try:
                # SIEMPRE scrapea datos; solo descarga archivos si faltan
                datos_lista = scrape_catastro(
                    referencia, nombre_carpeta, base_dir, driver,
                    descargar_foto=not foto_ok,
                    descargar_plano=not plano_ok,
                    nombre_pdf=doc_path if not doc_ok else None
                )
                if datos_lista and isinstance(datos_lista, list) and len(datos_lista) > 0:
                    datos = datos_lista[0]
                    for k, v in datos.items():
                        if k in df.columns:
                            df.at[idx, k] = v
            except Exception as e:
                print(f"❌ Error en referencia {referencia}: {e}")

            # Actualiza flags
            df.at[idx, "PLANO CATASTRO"] = 1 if os.path.isfile(plano_path) else 0
            df.at[idx, "DOC CATASTRO"] = 1 if os.path.isfile(doc_path) else 0
            df.at[idx, "FOTO EDIF"] = 1 if os.path.isfile(foto_path) else 0

            # Defectuosas solo si falta algo
            if not (os.path.isfile(foto_path) and os.path.isfile(plano_path) and os.path.isfile(doc_path)):
                referencias_defectuosas.append((idx, referencia, nombre_carpeta))

        # SEGUNDO BUCLE: reintentos si sigue fallando
        for idx, referencia, nombre_carpeta in referencias_defectuosas:
            carpeta = os.path.join(base_dir, nombre_carpeta)
            foto_path = os.path.join(carpeta, "foto_fachada.jpg")
            plano_path = os.path.join(carpeta, "plano.pdf")
            doc_path = os.path.join(carpeta, f"{referencia}.pdf")
            foto_ok = os.path.isfile(foto_path)
            plano_ok = os.path.isfile(plano_path)
            doc_ok = os.path.isfile(doc_path)

            if not (foto_ok and plano_ok and doc_ok):
                print(f"Reintentando referencia defectuosa: {referencia}")
                try:
                    datos_lista = scrape_catastro(
                        referencia, nombre_carpeta, base_dir, driver,
                        descargar_foto=not foto_ok,
                        descargar_plano=not plano_ok,
                        nombre_pdf=doc_path if not doc_ok else None
                    )
                    if datos_lista and isinstance(datos_lista, list) and len(datos_lista) > 0:
                        datos = datos_lista[0]
                        for k, v in datos.items():
                            if k in df.columns:
                                df.at[idx, k] = v
                except Exception as e:
                    print(f"❌ Reintento falló para referencia {referencia}: {e}")

                # Flags
                df.at[idx, "PLANO CATASTRO"] = 1 if os.path.isfile(plano_path) else 0
                df.at[idx, "DOC CATASTRO"] = 1 if os.path.isfile(doc_path) else 0
                df.at[idx, "FOTO EDIF"] = 1 if os.path.isfile(foto_path) else 0

        # BUCLE FINAL: Subrayado
        for idx, fila in df.iterrows():
            id_centro = str(fila["ID CENTRO"]).strip()
            centro = str(fila["CENTRO"]).strip()
            referencia = str(fila["REF CATASTRAL"]).strip()
            nombre_carpeta = f"{id_centro}_{centro}"
            carpeta = os.path.join(base_dir, nombre_carpeta)
            foto_path = os.path.join(carpeta, "foto_fachada.jpg")
            plano_path = os.path.join(carpeta, "plano.pdf")
            doc_path = os.path.join(carpeta, f"{referencia}.pdf")
            fila_excel = data_range.Rows[idx+1]
            if not (os.path.isfile(foto_path) and os.path.isfile(plano_path) and os.path.isfile(doc_path)):
                fila_excel.Interior.Color = 255
            else:
                fila_excel.Interior.ColorIndex = -4142

        sht.range(data_range.Address).value = df.values.tolist()

        ctypes.windll.user32.MessageBoxW(0, "¡Catastro actualizado!", "Catastro", 0)
        wb.app.api.WindowState = -4137

    finally:
        driver.quit()