import requests
import pandas as pd
import xlwings as xw
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import unidecode

def formatear_texto(texto):
    texto = texto.strip().lower()
    texto = unidecode.unidecode(texto)
    texto = texto.replace(" ", "-")
    return texto

def obtener_datos_mes(ciudad, anio, mes):
    url = f"https://www.sunrise-and-sunset.com/es/sun/espana/{ciudad}/{anio}/{mes}"
    datos = []
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        soup = BeautifulSoup(r.content, "html.parser")
        tabla = soup.find("table", class_="table")
        if not tabla:
            return []
        filas = tabla.find_all("tr")[1:]
        for fila in filas:
            cols = [c.text.strip() for c in fila.find_all("td")]
            if len(cols) >= 3:
                try:
                    salida = datetime.strptime(cols[1], "%H:%M")
                    puesta = datetime.strptime(cols[2], "%H:%M")
                    duracion = puesta - salida
                    if duracion < timedelta(0):
                        duracion += timedelta(days=1)
                    duracion_str = f"{duracion.seconds // 3600}:{(duracion.seconds // 60) % 60:02}"
                except:
                    duracion_str = ""
                datos.append({
                    "Fecha": cols[0],
                    "Salida del sol": cols[1],
                    "Puesta del sol": cols[2],
                    "Duración": duracion_str
                })
        return datos
    except:
        return []

def procesar_amaneceres():
    wb = xw.Book.caller()
    ws_ListCP = wb.sheets['ListCP']
    ws_Tele = wb.sheets['Tele']

    ciudad_input = ws_ListCP.range("C5").value
    if not ciudad_input:
        ws_Tele.range("AB7").value = "⚠️ Ciudad no especificada"
        return

    ciudad_formateada = formatear_texto(ciudad_input)
    anio_actual = datetime.now().year
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]

    # Estado de la web
    try:
        test_url = "https://www.sunrise-and-sunset.com/es"
        r = requests.get(test_url, timeout=10)
        r.raise_for_status()
        estado_web = "🟢 Página disponible"
    except Exception as e:
        estado_web = f"🔴 Página caída: {e}"
    ws_Tele.range("AB7").value = estado_web

    if "🔴" in estado_web:
        return

    all_data = []
    for mes in meses:
        all_data.extend(obtener_datos_mes(ciudad_formateada, anio_actual, mes))

    if not all_data:
        ws_Tele.range("AB7").value = "❌ No se pudieron obtener datos"
        return

    df = pd.DataFrame(all_data)
    ws_Tele.range("AA8").options(index=False, header=True).value = df

# Para pruebas fuera de Excel
if __name__ == "__main__":
    xw.Book.set_mock_caller(r"C:\Users\indiva\Desktop\excel_automatizar\TEC-ECO_UTIEL_V2.xlsm")
    procesar_amaneceres()
