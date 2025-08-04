import xlwings as xw
import pythoncom
import time
import os
import subprocess
import re
from pathlib import Path

def obtener_valores_tabla_dinamica(sheet, nombre_tabla):
    pt = sheet.api.PivotTables(nombre_tabla)
    data_body_range = pt.TableRange1
    valores = []
    for i in range(2, data_body_range.Rows.Count + 1):
        celda = data_body_range.Cells(i, 1)
        texto = celda.Text.strip()
        if texto and texto.lower() != "total general":
            valores.append(texto)
    return valores

def aplicar_filtros_iterativos(nombre_archivo, hoja_slicers="Filtros_G", hoja_tabla="Filtros_G", nombre_tabla="TablaDinámica12"):
    pythoncom.CoInitialize()
    app = xw.App(visible=False)
    wb = app.books.open(nombre_archivo)
    hoja = wb.sheets[hoja_tabla]
    valores = obtener_valores_tabla_dinamica(hoja, nombre_tabla)
    
    slicer_estado = {}

    # 1️⃣ Inicializar todos los slicers en ZZZ_NULO
    for slicer_cache in wb.api.SlicerCaches:
        try:
            slicers_validos = [
                s for s in slicer_cache.Slicers
                if s.Shape.TopLeftCell.Worksheet.Name == hoja_slicers
            ]
            if not slicers_validos:
                continue

            items = slicer_cache.SlicerItems
            nombre_a_idx = {
                items.Item(i).Name.strip().lower(): i
                for i in range(1, items.Count + 1)
            }

            slicer_name = slicer_cache.Name
            idx_nulo = nombre_a_idx.get("zzz_nulo")
            if idx_nulo is None:
                print(f"⚠️ Slicer '{slicer_name}' no tiene 'ZZZ_NULO'. Se omite.")
                continue

            for i in range(1, items.Count + 1):
                items.Item(i).Selected = (i == idx_nulo)

            slicer_estado[slicer_name] = "zzz_nulo"
            print(f"[{slicer_name}] Inicializado → ZZZ_NULO")

        except Exception as e:
            print(f"⚠️ Error inicializando slicer '{slicer_cache.Name}': {e}")

    # 2️⃣ Iterar los valores de la tabla
    nombre_excel = Path(nombre_archivo).stem
    
    for valor in valores:
        filtro_objetivo = valor.strip().lower()
        print(f"\n🎯 Aplicando filtro: {valor}")

        for slicer_cache in wb.api.SlicerCaches:
            try:
                slicers_validos = [
                    s for s in slicer_cache.Slicers
                    if s.Shape.TopLeftCell.Worksheet.Name == hoja_slicers
                ]
                if not slicers_validos:
                    continue

                items = slicer_cache.SlicerItems
                nombre_a_idx = {
                    items.Item(i).Name.strip().lower(): i
                    for i in range(1, items.Count + 1)
                }

                slicer_name = slicer_cache.Name
                idx_nulo = nombre_a_idx.get("zzz_nulo")
                idx_nuevo = nombre_a_idx.get(filtro_objetivo)
                valor_anterior = slicer_estado.get(slicer_name, "zzz_nulo")
                idx_anterior = nombre_a_idx.get(valor_anterior)

                if idx_nuevo is not None:
                    # Activar el nuevo primero
                    items.Item(idx_nuevo).Selected = True
                    # Desactivar el anterior solo si es distinto
                    if idx_anterior is not None and idx_anterior != idx_nuevo:
                        items.Item(idx_anterior).Selected = False
                    slicer_estado[slicer_name] = filtro_objetivo
                    print(f"[{slicer_name}] {valor.upper()} activado")
                else:
                    print(f"[{slicer_name}] {valor.upper()} NO está → se activa ZZZ_NULO")
                    # Activar ZZZ_NULO
                    if idx_nulo is not None:
                        items.Item(idx_nulo).Selected = True
                        # Desactivar anterior si era distinto
                        if idx_anterior is not None and idx_anterior != idx_nulo:
                            items.Item(idx_anterior).Selected = False
                        slicer_estado[slicer_name] = "zzz_nulo"

            except Exception as e:
                print(f"⚠️ Error en slicer '{slicer_cache.Name}': {e}")

        # 💾 Exportar el Excel completo con el filtro aplicado
        print(f"📄 Guardando Excel con filtro: {valor}")
        base_path = Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\FRONT") / nombre_excel / valor.strip()
        base_path.mkdir(parents=True, exist_ok=True)
        ruta_excel_filtrado = base_path / f"{valor.strip()}.xlsx"
        
        # Guardar el archivo completo con el filtro aplicado
        wb.save(str(ruta_excel_filtrado))
        print(f"✅ Excel guardado: {ruta_excel_filtrado}")
        
        # Pausa mínima antes del siguiente filtro
        time.sleep(0.3)

    print(f"\n🏁 PROCESAMIENTO COMPLETADO:")
    print(f"   📊 Total filtros procesados: {len(valores)}")
    print(f"   📁 Archivos Excel guardados en: Z:\\DOCUMENTACION TRABAJO\\CARPETAS PERSONAL\\SO\\FRONT\\{nombre_excel}")
    print(f"   ✅ Todos los archivos Excel filtrados creados correctamente")
    
    wb.save()
    app.visible = False
    
    print(f"\n💡 Revisa la carpeta de salida para ver todos los archivos Excel generados:")
    print(f"   📂 Z:\\DOCUMENTACION TRABAJO\\CARPETAS PERSONAL\\SO\\FRONT\\{nombre_excel}")
    print("=" * 80)

if __name__ == "__main__":
    ruta_excel = r"C:\Users\indiva\Desktop\Copia de Copia de ANALISIS AUD-ENER_V26_NULO.xlsx"
    aplicar_filtros_iterativos(ruta_excel)
