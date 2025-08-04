import requests
import utm
import xlwings as xw
import os
import folium
from folium.plugins import MarkerCluster
from xlwings.utils import rgb_to_int
from matplotlib import cm
import matplotlib.colors as mcolors

API_KEY = "AIzaSyCFnkUNVcfnuNxxkzx8men5p0uE7cIyoGs"

def utm_to_address(easting, northing, zone_number, zone_letter):
    lat, lon = utm.to_latlon(easting, northing, zone_number, zone_letter)
    url = f"https://maps.googleapis.com/maps/api/geocode/json?latlng={lat},{lon}&key={API_KEY}"
    response = requests.get(url)
    data = response.json()

    if data['status'] != 'OK':
        return {"Error": data['status']}

    address = data['results'][0]['formatted_address']
    municipio = None
    for component in data['results'][0]['address_components']:
        if "locality" in component['types']:
            municipio = component['long_name']
            break

    if "plus_code" in data:
        compound_code = data['plus_code'].get('compound_code')
        if compound_code:
            entidad_menor = compound_code.split(' ', 1)[1].split(',')[0]
            municipio = entidad_menor

    gmaps_link = f"https://www.google.com/maps?q={lat},{lon}"

    return {
        "Latitud": lat,
        "Longitud": lon,
        "Dirección completa": address,
        "Municipio": municipio,
        "Google Maps": gmaps_link,
        "UTM": f"{easting}, {northing}"
    }

def get_zone_letter(municipio_nombre):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={municipio_nombre}&key={API_KEY}"
    response = requests.get(url)
    data = response.json()

    if data['status'] != 'OK' or not data['results']:
        return None, None, None, None

    location = data['results'][0]['geometry']['location']
    lat, lon = location['lat'], location['lng']
    _, _, zone_number, zone_letter = utm.from_latlon(lat, lon)
    return zone_number, zone_letter, lat, lon

def procesar_lista_utm():
    wb = xw.Book.caller()
    ws_ListCP = wb.sheets['ListCP']
    ws_CM = wb.sheets['CM']
    excel_path = wb.fullname
    carpeta = os.path.dirname(excel_path)

    municipio_ref = ws_ListCP.range('C5').value
    if not municipio_ref:
        ws_CM.range('M1').value = "❌ Error: Celda C5 vacía"
        return

    zone_number, zone_letter, centro_lat, centro_lon = get_zone_letter(municipio_ref)
    if not zone_number:
        ws_CM.range('M1').value = "❌ Error: No se pudo deducir zona UTM"
        return

    coor_range = ws_ListCP.range('F5').expand('down').value
    calles_range = ws_CM.range('K2').expand('down').value
    resultados = []
    datos_para_mapa = []
    lineas_txt = []

    for idx, (coor_texto, calle) in enumerate(zip(coor_range, calles_range), start=2):
        try:
            coor_parts = coor_texto.replace(" ", "").split(',')
            if len(coor_parts) >= 4:
                easting = float(coor_parts[0] + '.' + coor_parts[1])
                northing = float(coor_parts[2] + '.' + coor_parts[3])
            else:
                raise ValueError("Coordenada mal formada")

            datos = utm_to_address(easting, northing, zone_number, zone_letter)
            municipio = datos.get("Municipio", "❌ Error")
            resultados.append([municipio])

            if "Latitud" in datos and "Longitud" in datos:
                datos_para_mapa.append(datos)
                lineas_txt.append(f"{calle} | {datos['UTM']} -> {datos['Google Maps']}")

        except Exception as e:
            resultados.append(["❌ Error"])
            lineas_txt.append(f"{calle} | {coor_texto} -> ❌ Error")

    # Escribir municipios en Excel
    ws_CM.range('M1').value = "NÚCLEO POBLACIONAL"
    ws_CM.range('M2').value = resultados
    ws_ListCP.range('K1').value = "✅ NÚCLEO POBLACIONAL actualizado"

    # Colorear celdas en Excel con paleta predefinida
    municipios_unicos = list({r[0] for r in resultados if r[0] != "❌ Error"})
    cmap = cm.get_cmap('tab10', len(municipios_unicos))
    color_map = {m: mcolors.to_hex(cmap(i)) for i, m in enumerate(municipios_unicos)}

    for i, celda in enumerate(ws_CM.range('M2').expand('down')):
        nombre = celda.value
        if nombre in color_map:
            hex_color = color_map[nombre].lstrip("#")
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            celda.color = (r, g, b)

    # Crear carpeta de salida
    subcarpeta = os.path.join(carpeta, "complementos_buscar_núcleos")
    os.makedirs(subcarpeta, exist_ok=True)

    # Crear archivo TXT
    txt_path = os.path.join(subcarpeta, "coordenadas_googlemaps.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("CALLE | COORDENADAS -> ENLACE GOOGLE MAPS\n")
        for linea in lineas_txt:
            f.write(linea + "\n")

    # Crear mapa HTML
    mapa = folium.Map(location=[centro_lat, centro_lon], zoom_start=12)
    marker_cluster = MarkerCluster().add_to(mapa)
    for dato in datos_para_mapa:
        color = color_map.get(dato["Municipio"], "#000000")
        folium.Marker(
            location=[dato["Latitud"], dato["Longitud"]],
            popup=f"<b>{dato['Municipio']}</b><br>{dato['Dirección completa']}",
            icon=folium.Icon(color='white', icon='info-sign', prefix='glyphicon', icon_color=color)
        ).add_to(marker_cluster)

    # Añadir leyenda
    leyenda_html = """
    <div style='position: fixed; bottom: 30px; left: 30px; width: 220px; background-color: white; 
         z-index:9999; padding: 10px; box-shadow: 0 0 15px rgba(0,0,0,0.3); font-size:14px;'>
    <b>Leyenda - Municipios</b><br>
    """
    for municipio, color in color_map.items():
        leyenda_html += f"<i style='background:{color};width:12px;height:12px;display:inline-block;margin-right:6px;'></i>{municipio}<br>"
    leyenda_html += "</div>"
    mapa.get_root().html.add_child(folium.Element(leyenda_html))
    mapa.fit_bounds([(d["Latitud"], d["Longitud"]) for d in datos_para_mapa])
    mapa_path = os.path.join(subcarpeta, "mapa_municipios.html")
    mapa.save(mapa_path)

if __name__ == "__main__":
    xw.Book.set_mock_caller(r"C:\\Users\\indiva\\Desktop\\excel_automatizar\\TEC-ECO_UTIEL_V2.xlsm")
    procesar_lista_utm()
