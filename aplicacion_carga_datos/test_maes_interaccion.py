# test_maes_interaccion.py
# Prueba de interacción con slicers de E_MAEs para verificar la lógica

import xlwings as xw
import pythoncom
import re
from pathlib import Path

def test_interaccion_maes(nombre_archivo):
    """Prueba la interacción con los slicers de E_MAEs sin exportar HTML"""
    print("🧪 INICIANDO PRUEBA DE INTERACCIÓN E_MAEs")
    print("=" * 60)
    
    pythoncom.CoInitialize()
    app = xw.App(visible=True)  # Visible para poder ver lo que pasa
    wb = app.books.open(nombre_archivo)
    
    # Verificar que existe la hoja E_MAEs
    hojas_disponibles = {hoja.name.strip() for hoja in wb.sheets}
    if "E_MAEs" not in hojas_disponibles:
        print("❌ Hoja E_MAEs no encontrada")
        wb.close()
        app.quit()
        return
    
    print("✅ Hoja E_MAEs encontrada")
    hoja_emaes = wb.sheets["E_MAEs"]
    
    try:
        # Buscar los slicers
        slicer_grupo_mae = None
        slicer_numero_mae = None
        
        print("\n🔍 Buscando slicers...")
        
        for slicer_cache in wb.api.SlicerCaches:
            for slicer in slicer_cache.Slicers:
                if slicer.Shape.TopLeftCell.Worksheet.Name == "E_MAEs":
                    print(f"   📋 Encontrado slicer: {slicer_cache.Name}")
                    
                    if "GRUPO" in slicer_cache.Name.upper() and "MAE" in slicer_cache.Name.upper():
                        slicer_grupo_mae = slicer_cache
                        print(f"   🎯 → Es GRUPO MAE: {slicer_cache.Name}")
                    elif "MAE" in slicer_cache.Name.upper() and "GRUPO" not in slicer_cache.Name.upper():
                        slicer_numero_mae = slicer_cache
                        print(f"   🔢 → Es N° MAE: {slicer_cache.Name}")
        
        if not slicer_grupo_mae:
            print("❌ No se encontró slicer de GRUPO MAE")
            return
        if not slicer_numero_mae:
            print("❌ No se encontró slicer de N° MAE")
            return
        
        print(f"\n✅ Slicers encontrados:")
        print(f"   🎯 GRUPO MAE: {slicer_grupo_mae.Name}")
        print(f"   🔢 N° MAE: {slicer_numero_mae.Name}")
        
        # Obtener todos los grupos MAE disponibles
        print(f"\n📋 GRUPOS MAE DISPONIBLES:")
        grupos_mae = []
        items_grupo = slicer_grupo_mae.SlicerItems
        for i in range(1, items_grupo.Count + 1):
            grupo_name = items_grupo.Item(i).Name.strip()
            if grupo_name and grupo_name.lower() != "total general":
                grupos_mae.append((grupo_name, i))
                print(f"   {i}. {grupo_name}")
        
        # Obtener todos los N° MAE disponibles
        print(f"\n🔢 N° MAE DISPONIBLES:")
        items_numero = slicer_numero_mae.SlicerItems
        todos_maes = []
        for i in range(1, items_numero.Count + 1):
            mae_name = items_numero.Item(i).Name.strip()
            if mae_name and mae_name.lower() != "total general":
                todos_maes.append((mae_name, i))
                print(f"   {i}. {mae_name}")
        
        print(f"\n🎯 INICIANDO SIMULACIÓN DE EXPORTACIÓN...")
        print(f"   📊 Total grupos: {len(grupos_mae)}")
        print(f"   🔢 Total MAEs: {len(todos_maes)}")
        
        # Simular el proceso por cada grupo
        total_simulaciones = 0
        
        for idx_grupo, (grupo_name, grupo_idx) in enumerate(grupos_mae, 1):
            print(f"\n{'='*50}")
            print(f"🎯 GRUPO {idx_grupo}/{len(grupos_mae)}: {grupo_name}")
            print(f"{'='*50}")
            
            # 1️⃣ ACTIVAR SOLO ESTE GRUPO
            print(f"🔄 Activando solo el grupo: {grupo_name}")
            for i in range(1, items_grupo.Count + 1):
                items_grupo.Item(i).Selected = (i == grupo_idx)
            
            print(f"✅ Grupo activado: {grupo_name}")
            
            # Esperar un momento para que Excel procese
            import time
            time.sleep(1)
            
            # 2️⃣ ACTIVAR TODOS LOS N°MAE (para la exportación "TODOS")
            print(f"📊 Activando TODOS los N° MAE...")
            for i in range(1, items_numero.Count + 1):
                items_numero.Item(i).Selected = True
            
            print(f"✅ SIMULACIÓN: Exportaría 'TODOS' del grupo {grupo_name}")
            total_simulaciones += 1
            
            # 3️⃣ PROCESAR CADA MAE INDIVIDUAL
            print(f"🔄 Procesando MAEs individuales...")
            
            maes_del_grupo = 0
            for mae_name, mae_idx in todos_maes:
                # Activar solo este MAE
                items_numero.Item(mae_idx).Selected = True
                
                # Desactivar todos los demás
                for i in range(1, items_numero.Count + 1):
                    if i != mae_idx:
                        items_numero.Item(i).Selected = False
                
                print(f"   🎯 SIMULACIÓN: Exportaría MAE individual '{mae_name}' del grupo '{grupo_name}'")
                total_simulaciones += 1
                maes_del_grupo += 1
                
                # Pequeña pausa
                time.sleep(0.5)
            
            print(f"📊 Grupo {grupo_name} completado:")
            print(f"   ✅ 1 exportación 'TODOS'")
            print(f"   ✅ {maes_del_grupo} exportaciones individuales")
            print(f"   📊 Total del grupo: {maes_del_grupo + 1}")
        
        # 4️⃣ RESTAURAR ESTADO FINAL
        print(f"\n🔄 Restaurando estado final (todos activos)...")
        for i in range(1, items_grupo.Count + 1):
            items_grupo.Item(i).Selected = True
        for i in range(1, items_numero.Count + 1):
            items_numero.Item(i).Selected = True
        
        print(f"\n✅ Estado restaurado")
        
        print(f"\n🏁 SIMULACIÓN COMPLETADA:")
        print(f"   📊 Total grupos procesados: {len(grupos_mae)}")
        print(f"   🔢 Total MAEs: {len(todos_maes)}")
        print(f"   📁 Total simulaciones de exportación: {total_simulaciones}")
        print(f"   💡 Estructura que se crearía:")
        
        for grupo_name, _ in grupos_mae:
            grupo_sanitizado = re.sub(r'[\\/*?:"<>|]', "_", grupo_name)
            print(f"      📁 E_MAEs/{grupo_sanitizado}/")
            print(f"         📁 TODOS/")
            for mae_name, _ in todos_maes:
                mae_sanitizado = re.sub(r'[\\/*?:"<>|]', "_", mae_name)
                print(f"         📁 {mae_sanitizado}/")
        
        print(f"\n🎯 ¿La lógica parece correcta? (Y/N)")
        respuesta = input().strip().upper()
        
        if respuesta == 'Y':
            print("✅ ¡Perfecto! La lógica es correcta.")
        else:
            print("⚠️ Revisa los pasos y ajusta según sea necesario.")
            
    except Exception as e:
        print(f"❌ Error durante la prueba: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        print(f"\n🔚 Cerrando archivo...")
        wb.save()  # Opcional: guardar cambios
        wb.close()
        app.quit()
        print("✅ Archivo cerrado")

if __name__ == "__main__":
    # Cambia esta ruta por la de tu archivo Excel de prueba
    ruta_excel = input("📁 Introduce la ruta del archivo Excel: ").strip().strip('"')
    
    if not Path(ruta_excel).exists():
        print("❌ El archivo no existe")
    else:
        test_interaccion_maes(ruta_excel)
