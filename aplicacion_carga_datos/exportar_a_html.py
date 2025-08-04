# exportar_a_html.py
import xlwings as xw
import pythoncom
import sys
from pathlib import Path
import re
import os

def actualizar_tablas_dinamicas(wb, hojas_disponibles):
    """Actualiza las tablas dinámicas de BALANCE y E_MAEs"""
    try:
        # Actualizar TablaDinámica5 en hoja BALANCE
        if "BALANCE" in hojas_disponibles:
            hoja_balance = wb.sheets["BALANCE"]
            print("🔄 Actualizando TablaDinámica5 en hoja BALANCE...")
            try:
                tabla_balance = hoja_balance.api.PivotTables("TablaDinámica5")
                tabla_balance.RefreshTable()
                print("✅ TablaDinámica5 en BALANCE actualizada")
            except Exception as e:
                print(f"⚠️ No se pudo actualizar TablaDinámica5 en BALANCE: {e}")
        
        # Actualizar TablaDinámica5 en hoja E_MAEs
        if "E_MAEs" in hojas_disponibles:
            hoja_emaes = wb.sheets["E_MAEs"]
            print("🔄 Actualizando TablaDinámica5 en hoja E_MAEs...")
            try:
                tabla_emaes = hoja_emaes.api.PivotTables("TablaDinámica5")
                tabla_emaes.RefreshTable()
                print("✅ TablaDinámica5 en E_MAEs actualizada")
            except Exception as e:
                print(f"⚠️ No se pudo actualizar TablaDinámica5 en E_MAEs: {e}")
                
    except Exception as e:
        print(f"⚠️ Error general actualizando tablas dinámicas: {e}")

def exportar_emaes_completo(wb, carpeta_filtro, filtro_sanitizado):
    """Exporta E_MAEs con todas las combinaciones de MAEs por grupos"""
    if "E_MAEs" not in {hoja.name.strip() for hoja in wb.sheets}:
        print("⚠️ Hoja E_MAEs no encontrada, omitiendo exportación compleja")
        return
    
    hoja_emaes = wb.sheets["E_MAEs"]
    print(f"\n🔧 Iniciando exportación compleja de E_MAEs por grupos...")
    
    # Actualizar tablas dinámicas antes de exportar E_MAEs
    hojas_disponibles = {hoja.name.strip() for hoja in wb.sheets}
    print("🔄 Actualizando tablas dinámicas antes de exportar E_MAEs...")
    actualizar_tablas_dinamicas(wb, hojas_disponibles)
    
    # Crear carpeta E_MAEs dentro del filtro
    carpeta_emaes = carpeta_filtro / "E_MAEs"
    carpeta_emaes.mkdir(parents=True, exist_ok=True)
    
    try:
        # Obtener slicers de GRUPO MAE y N° MAE
        slicer_grupo_mae = None
        slicer_numero_mae = None
        
        for slicer_cache in wb.api.SlicerCaches:
            for slicer in slicer_cache.Slicers:
                if slicer.Shape.TopLeftCell.Worksheet.Name == "E_MAEs":
                    if "GRUPO" in slicer_cache.Name.upper() and "MAE" in slicer_cache.Name.upper():
                        slicer_grupo_mae = slicer_cache
                        print(f"🎯 Encontrado slicer GRUPO MAE: {slicer_cache.Name}")
                    elif "MAE" in slicer_cache.Name.upper() and "GRUPO" not in slicer_cache.Name.upper():
                        slicer_numero_mae = slicer_cache
                        print(f"🎯 Encontrado slicer N° MAE: {slicer_cache.Name}")
        
        if not slicer_grupo_mae or not slicer_numero_mae:
            print("⚠️ No se encontraron slicers de GRUPO MAE o N° MAE en E_MAEs")
            return
        
        # Obtener todos los grupos MAE disponibles
        grupos_mae = []
        items_grupo = slicer_grupo_mae.SlicerItems
        for i in range(1, items_grupo.Count + 1):
            grupo_name = items_grupo.Item(i).Name.strip()
            if grupo_name and grupo_name.lower() != "total general":
                grupos_mae.append((grupo_name, i))
        
        print(f"📋 Grupos MAE encontrados: {len(grupos_mae)}")
        for grupo_name, _ in grupos_mae:
            print(f"   - {grupo_name}")
        
        total_exportaciones = 0
        
        # 🔄 PROCESAR CADA GRUPO MAE
        for grupo_name, grupo_idx in grupos_mae:
            print(f"\n{'='*50}")
            print(f"🎯 PROCESANDO GRUPO: {grupo_name}")
            print(f"{'='*50}")
            
            # 1️⃣ ACTIVAR SOLO ESTE GRUPO MAE
            for i in range(1, items_grupo.Count + 1):
                items_grupo.Item(i).Selected = (i == grupo_idx)
            
            print(f"✅ Grupo activado: {grupo_name}")
            
            # Crear carpeta para este grupo
            grupo_sanitizado = re.sub(r'[\\/*?:"<>|]', "_", grupo_name)
            carpeta_grupo = carpeta_emaes / grupo_sanitizado
            carpeta_grupo.mkdir(parents=True, exist_ok=True)
            
            # 2️⃣ ACTIVAR TODOS LOS N°MAE DE ESTE GRUPO
            items_numero = slicer_numero_mae.SlicerItems
            for i in range(1, items_numero.Count + 1):
                items_numero.Item(i).Selected = True
            
            # Exportar con TODOS los MAEs del grupo
            carpeta_todos = carpeta_grupo / "TODOS"
            carpeta_todos.mkdir(parents=True, exist_ok=True)
            ruta_html_todos = carpeta_todos / f"{filtro_sanitizado}-E_MAEs-{grupo_sanitizado}-TODOS.htm"
            
            print(f"📊 Exportando TODOS los MAEs del grupo {grupo_name}...")
            exportar_hoja_individual(wb, "E_MAEs", str(ruta_html_todos))
            print(f"✅ Exportado: TODOS del grupo {grupo_name}")
            total_exportaciones += 1
            
            # 3️⃣ OBTENER MAEs INDIVIDUALES DE ESTE GRUPO
            # Necesitamos ver qué MAEs están disponibles cuando este grupo está activo
            print(f"� Obteniendo MAEs individuales del grupo {grupo_name}...")
            
            # Obtener MAEs disponibles (los que están visibles con este grupo activo)
            maes_del_grupo = []
            for i in range(1, items_numero.Count + 1):
                mae_name = items_numero.Item(i).Name.strip()
                if mae_name and mae_name.lower() != "total general":
                    # Verificar si este MAE es visible/relevante para este grupo
                    # Por ahora tomamos todos, pero podrías filtrar según lógica específica
                    maes_del_grupo.append((mae_name, i))
            
            print(f"📋 MAEs en grupo {grupo_name}: {len(maes_del_grupo)}")
            
            # 4️⃣ EXPORTAR CADA MAE INDIVIDUAL
            for mae_name, mae_idx in maes_del_grupo:
                print(f"🎯 Procesando MAE individual: {mae_name}")
                
                # Activar solo este MAE (filosofía: primero activar, luego desactivar otros)
                items_numero.Item(mae_idx).Selected = True
                
                # Desactivar todos los demás MAEs
                for i in range(1, items_numero.Count + 1):
                    if i != mae_idx:
                        items_numero.Item(i).Selected = False
                
                # Crear carpeta para este MAE
                mae_sanitizado = re.sub(r'[\\/*?:"<>|]', "_", mae_name)
                carpeta_mae = carpeta_grupo / mae_sanitizado
                carpeta_mae.mkdir(parents=True, exist_ok=True)
                
                # Exportar este MAE individual
                ruta_html_mae = carpeta_mae / f"{filtro_sanitizado}-E_MAEs-{grupo_sanitizado}-{mae_sanitizado}.htm"
                exportar_hoja_individual(wb, "E_MAEs", str(ruta_html_mae))
                print(f"✅ MAE exportado: {mae_name}")
                total_exportaciones += 1
            
            print(f"🏁 Grupo {grupo_name} completado: {len(maes_del_grupo) + 1} exportaciones")
        
        # 5️⃣ RESTAURAR ESTADO (todos los grupos y todos los MAEs activos)
        for i in range(1, items_grupo.Count + 1):
            items_grupo.Item(i).Selected = True
        for i in range(1, items_numero.Count + 1):
            items_numero.Item(i).Selected = True
        
        print(f"\n🏁 EXPORTACIÓN COMPLEJA DE E_MAEs COMPLETADA:")
        print(f"   📊 Total grupos procesados: {len(grupos_mae)}")
        print(f"   📊 Total exportaciones: {total_exportaciones}")
        print(f"   📁 Carpeta base: {carpeta_emaes}")
        
    except Exception as e:
        print(f"❌ Error en exportación compleja de E_MAEs: {e}")

def exportar_hoja_individual(wb, nombre_hoja, ruta_html):
    """Exporta una hoja individual a HTML"""
    try:
        macro_name = f"ExportarTemp_{nombre_hoja}_{hash(ruta_html) % 10000}"
        macro_vba = f'''
        Sub {macro_name}()
            Dim ruta As String
            ruta = "{ruta_html.replace("\\", "\\\\")}"
           
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
           
            ThisWorkbook.Sheets("{nombre_hoja}").Select
            ActiveWorkbook.PublishObjects.Add( _
                SourceType:=xlSourceSheet, _
                Filename:=ruta, _
                Sheet:="{nombre_hoja}", _
                HtmlType:=xlHtmlStatic).Publish True
           
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
        End Sub
        '''
        
        wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(macro_vba)
        wb.macro(macro_name)()
        
    except Exception as e:
        print(f"❌ Error exportando {nombre_hoja}: {e}")

def exportar_hojas_html(ruta_excel, filtro):
    print(f"\n🚀 Exportando HTML para: {filtro}")
    pythoncom.CoInitialize()
   
    filtro_sanitizado = re.sub(r'[\\/*?:"<>|]', "_", filtro.strip())
    
    # Obtener el nombre base de la carpeta que estamos procesando
    ruta_excel_path = Path(ruta_excel)
    nombre_carpeta_base = ruta_excel_path.parent.parent.name  # Obtiene "Nombre de la carpeta padre"
    
    # Crear carpeta de imágenes separada: imagenes_[nombre_carpeta_base]
    carpeta_imagenes = Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\FRONT") / f"imagenes_{nombre_carpeta_base}"
    carpeta_filtro = carpeta_imagenes / filtro_sanitizado  # Carpeta del filtro específico
    carpeta_filtro.mkdir(parents=True, exist_ok=True)
   
    app = xw.App(visible=False)
    app.display_alerts = False
    wb = app.books.open(ruta_excel)

    # Obtener nombres de hojas una sola vez y crear conjunto para búsquedas rápidas
    hojas_disponibles = {hoja.name.strip() for hoja in wb.sheets}
    print(f"📋 Hojas disponibles en el archivo: {', '.join(sorted(hojas_disponibles))}")

    # Hojas específicas a exportar (sin E_MAEs que se maneja por separado)
    hojas_a_exportar = ["T_CENT", "T_DEPEN", "T_ACOM", "T_SistCC", "T_Clima", "T_EqHoriz", "T_Eleva", "T_OtrosEq", "T_Ilum", "BALANCE"]
   
    # Filtrar hojas válidas usando intersección de conjuntos (más eficiente)
    hojas_validas = [hoja for hoja in hojas_a_exportar if hoja in hojas_disponibles]
    
    print(f"📋 Hojas estándar a exportar: {', '.join(hojas_validas)} ({len(hojas_validas)} hojas)")
    
    # 🎯 EXPORTACIÓN COMPLEJA DE E_MAEs
    if "E_MAEs" in hojas_disponibles:
        exportar_emaes_completo(wb, carpeta_filtro, filtro_sanitizado)
   
    # 📊 EXPORTAR HOJAS ESTÁNDAR
    print(f"\n🌐 Exportando {len(hojas_validas)} hojas estándar...")
    
    # Actualizar tablas dinámicas antes de exportar hojas estándar (especialmente BALANCE)
    if any(hoja in hojas_validas for hoja in ["BALANCE"]):
        print("🔄 Actualizando tablas dinámicas antes de exportar hojas estándar...")
        actualizar_tablas_dinamicas(wb, hojas_disponibles)
    
    # Crear carpeta HOJAS para las hojas estándar
    carpeta_hojas = carpeta_filtro / "HOJAS"
    carpeta_hojas.mkdir(parents=True, exist_ok=True)
    
    for i, nombre_hoja in enumerate(hojas_validas):
        # Crear carpeta para esta hoja dentro de HOJAS
        carpeta_hoja = carpeta_hojas / nombre_hoja
        carpeta_hoja.mkdir(parents=True, exist_ok=True)
       
        # Ruta del archivo HTML
        nombre_html = f"{filtro_sanitizado}-{nombre_hoja}.htm"
        ruta_html = carpeta_hoja / nombre_html
       
        print(f"🌐 Exportando hoja '{nombre_hoja}' a {ruta_html}")
       
        exportar_hoja_individual(wb, nombre_hoja, str(ruta_html))
        print(f"✅ Exportado: {nombre_hoja} → {nombre_html}")

    wb.close()
    app.quit()
    
    # Calcular total de exportaciones
    total_exportaciones = len(hojas_validas)
    if "E_MAEs" in hojas_disponibles:
        # Contar exportaciones de E_MAEs (estimación)
        total_exportaciones += 1  # Se contará exactamente en la función de E_MAEs
    
    print(f"🏁 Finalizado para filtro: {filtro}")
    print(f"📊 Hojas estándar exportadas: {len(hojas_validas)}")
    if "E_MAEs" in hojas_disponibles:
        print(f"🎯 E_MAEs exportado con múltiples combinaciones")
    print(f"📁 Archivos HTML guardados en: {carpeta_filtro}")
    print(f"   📂 Hojas estándar: {carpeta_filtro / 'HOJAS'}")
    if "E_MAEs" in hojas_disponibles:
        print(f"   🎯 E_MAEs: {carpeta_filtro / 'E_MAEs'}")
    print("-" * 40)

def procesar_carpeta_completa(ruta_excel_original=None):
    """Procesa todos los archivos Excel en la carpeta base generada a partir del archivo original"""
    if ruta_excel_original:
        # Usar el nombre del archivo seleccionado por el usuario
        nombre_excel = Path(ruta_excel_original).stem
        carpeta_base = Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\FRONT") / nombre_excel
    else:
        # Fallback para compatibilidad con versiones anteriores
        carpeta_base = Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\FRONT\Copia de Copia de ANALISIS AUD-ENER_V26_NULO")
    
    if not carpeta_base.exists():
        print(f"❌ La carpeta no existe: {carpeta_base}")
        return
    
    print(f"🔍 Buscando archivos Excel en: {carpeta_base}")
    
    # Buscar todas las subcarpetas que contienen archivos .xlsx
    archivos_encontrados = []
    for subcarpeta in carpeta_base.iterdir():
        if subcarpeta.is_dir():
            for archivo in subcarpeta.glob("*.xlsx"):
                # El nombre del filtro es el nombre de la subcarpeta
                filtro = subcarpeta.name
                archivos_encontrados.append((str(archivo), filtro))
    
    if not archivos_encontrados:
        print("❌ No se encontraron archivos Excel (.xlsx) en las subcarpetas")
        return
    
    print(f"📊 Encontrados {len(archivos_encontrados)} archivos Excel para procesar:")
    for i, (ruta, filtro) in enumerate(archivos_encontrados, 1):
        print(f"   {i}. {filtro} → {Path(ruta).name}")
    
    # Procesar cada archivo
    for i, (ruta_excel, filtro) in enumerate(archivos_encontrados, 1):
        print(f"\n{'='*60}")
        print(f"🚀 Procesando archivo {i}/{len(archivos_encontrados)}: {filtro}")
        print(f"{'='*60}")
        
        try:
            exportar_hojas_html(ruta_excel, filtro)
        except Exception as e:
            print(f"❌ Error procesando {filtro}: {e}")
    
    print(f"\n🏁 PROCESAMIENTO COMPLETO:")
    print(f"   📊 Total archivos procesados: {len(archivos_encontrados)}")
    print(f"   📂 Carpeta Excel: {carpeta_base}")
    print(f"   🖼️  Carpeta HTML: Z:\\DOCUMENTACION TRABAJO\\CARPETAS PERSONAL\\SO\\FRONT\\imagenes_{carpeta_base.name}")
    print("=" * 80)

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        # Modo tradicional: recibe ruta y filtro como argumentos
        ruta_excel = sys.argv[1]
        filtro = sys.argv[2]
        exportar_hojas_html(ruta_excel, filtro)
    elif len(sys.argv) == 2:
        # Modo con archivo específico: procesa carpeta basada en el archivo
        ruta_excel_original = sys.argv[1]
        print(f"🔄 Modo automático para archivo: {Path(ruta_excel_original).name}")
        procesar_carpeta_completa(ruta_excel_original)
    else:
        # Modo automático: procesa toda la carpeta (legacy)
        print("🔄 Modo automático: procesando carpeta legacy...")
        procesar_carpeta_completa()

