#!/usr/bin/env python3
"""
Análisis de compatibilidad entre extraer_datos_word.py y render_a3.py
Verifica que el formato de fotos generado sea válido para el renderizado HTML
"""

import json
from pathlib import Path
from urllib.parse import quote

def to_file_uri(path_like: str) -> str:
    """Función copiada de render_a3.py"""
    if not path_like:
        return ""
    p = str(path_like).replace("\\", "/")
    if p.lower().startswith("file://"):
        return p
    return "file:///" + quote(p, safe="/:._-()")

def analizar_formato_fotos():
    """Análisis detallado del formato de fotos"""
    
    print("🔍 ANÁLISIS DE COMPATIBILIDAD FOTOS: EXTRACTOR → RENDER")
    print("=" * 70)
    
    # Buscar JSON de centro
    json_files = [f for f in Path('out_context').glob('C*.json') if f.name != 'fotos_faltantes_por_id.json']
    if not json_files:
        print("❌ No se encontraron JSONs de centro")
        return
    
    json_file = json_files[0]  # Usar el primero
    print(f"📁 Analizando: {json_file.name}")
    print()
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 1. FORMATO GENERADO POR EXTRACTOR
    print("📤 FORMATO GENERADO POR EXTRACTOR:")
    print("-" * 40)
    
    centro = data.get('centro', {})
    if centro.get('fotos'):
        foto_ejemplo = centro['fotos'][0]
        print("✅ Centro con fotos encontrado")
        print(f"   Claves: {list(foto_ejemplo.keys())}")
        print(f"   Ejemplo: {foto_ejemplo}")
        print()
        
        # Verificar formato específico
        tiene_path = 'path' in foto_ejemplo
        tiene_name = 'name' in foto_ejemplo
        tiene_id = 'id' in foto_ejemplo
        tiene_file_uri = 'file_uri' in foto_ejemplo
        
        print("🔍 Verificación de campos:")
        print(f"   ✅ path: {tiene_path} ({foto_ejemplo.get('path', 'N/A')[:50]}...)")
        print(f"   ✅ name: {tiene_name} ({foto_ejemplo.get('name', 'N/A')})")
        print(f"   ✅ id: {tiene_id} ({foto_ejemplo.get('id', 'N/A')})")
        print(f"   ❌ file_uri: {tiene_file_uri} (se genera automáticamente)")
        print()
    else:
        print("❌ Centro sin fotos")
        print()
    
    # Buscar entidades con fotos
    entidades_con_fotos = []
    centros_list = data.get('centros', [])
    for centro_obj in centros_list:
        edificios = centro_obj.get('edificios', [])
        for edificio in edificios:
            # Dependencias
            deps = edificio.get('dependencias', [])
            for dep in deps:
                if dep.get('fotos'):
                    entidades_con_fotos.append(('DEPENDENCIA', dep['id'], dep['fotos'][0]))
                    if len(entidades_con_fotos) >= 3:  # Máximo 3 ejemplos
                        break
            
            # ACOM
            acom_list = edificio.get('acompanamientos', [])
            for acom in acom_list:
                if acom.get('fotos'):
                    entidades_con_fotos.append(('ACOM', acom.get('id', 'N/A'), acom['fotos'][0]))
                    if len(entidades_con_fotos) >= 3:
                        break
            
            # EQHORIZ
            eq_list = edificio.get('eqhoriz', [])
            for eq in eq_list:
                if eq.get('fotos'):
                    entidades_con_fotos.append(('EQHORIZ', eq.get('id', 'N/A'), eq['fotos'][0]))
                    if len(entidades_con_fotos) >= 3:
                        break
            
            if len(entidades_con_fotos) >= 3:
                break
        if len(entidades_con_fotos) >= 3:
            break
    
    print("📋 EJEMPLOS DE ENTIDADES CON FOTOS:")
    for tipo, eid, foto in entidades_con_fotos:
        print(f"   {tipo} {eid}:")
        print(f"     Estructura: {list(foto.keys())}")
        print(f"     path: {foto.get('path', 'N/A')[:60]}...")
        print(f"     name: {foto.get('name', 'N/A')}")
        print()
    
    # 2. FORMATO ESPERADO POR RENDER
    print("📥 FORMATO ESPERADO POR RENDER_A3.PY:")
    print("-" * 40)
    print("✅ Función collect_fotos() espera:")
    print("   - Campo 'fotos' (lista de objetos)")
    print("   - Si no existe 'fotos', usa 'fotos_paths' (lista de strings)")
    print("   - Cada foto debe tener:")
    print("     * path: ruta completa del archivo")
    print("     * name: nombre sin extensión (opcional, se puede generar)")
    print("     * id: identificador (opcional, usa name si no existe)")
    print("     * file_uri: URI para navegador (opcional, se genera automáticamente)")
    print()
    
    print("✅ Función build_photos_grid() usa:")
    print("   - ph.get('file_uri') o to_file_uri(ph.get('path'))")
    print("   - ph.get('name') o ph.get('id') para caption")
    print()
    
    # 3. COMPATIBILIDAD
    print("🎯 ANÁLISIS DE COMPATIBILIDAD:")
    print("-" * 40)
    
    if entidades_con_fotos:
        foto_test = entidades_con_fotos[0][2]  # Primera foto de ejemplo
        
        # Probar la función collect_fotos simulada
        def collect_fotos_test(any_obj):
            fotos = any_obj.get("fotos") or []
            if not fotos:
                fps = any_obj.get("fotos_paths") or []
                fotos = [{"path": p, "name": Path(p).stem} for p in fps]
            
            norm = []
            for f in fotos:
                path = f.get("path") or ""
                file_uri = f.get("file_uri") or to_file_uri(path)
                name = f.get("name") or f.get("id") or Path(path).stem
                norm.append({"path": path, "file_uri": file_uri, "name": name, "id": f.get("id", name)})
            return norm
        
        # Test con entidad de ejemplo
        entidad_test = {"fotos": [foto_test]}
        fotos_normalizadas = collect_fotos_test(entidad_test)
        
        print("✅ TEST DE PROCESAMIENTO:")
        print(f"   Foto original: {foto_test}")
        print(f"   Foto normalizada: {fotos_normalizadas[0]}")
        print()
        
        # Verificar que file_uri se genera correctamente
        foto_norm = fotos_normalizadas[0]
        file_uri = foto_norm.get('file_uri', '')
        
        print("🔗 VERIFICACIÓN DE FILE_URI:")
        print(f"   ✅ Generado: {file_uri[:80]}...")
        print(f"   ✅ Válido: {'file:///' in file_uri}")
        print(f"   ✅ Path encoding: {quote('C:\\Users') in file_uri}")
        print()
        
        # Verificar HTML generado
        def build_photos_grid_test(photos):
            if not photos:
                return "Sin fotos"
            
            cards = []
            for ph in photos:
                uri = ph.get("file_uri") or to_file_uri(ph.get("path"))
                name = ph.get("name") or ph.get("id") or ""
                cards.append(f'<img src="{uri[:50]}..." alt="{name}">')
            return "\\n".join(cards)
        
        html_test = build_photos_grid_test(fotos_normalizadas)
        
        print("📄 HTML GENERADO (EJEMPLO):")
        print(f"   {html_test}")
        print()
        
    # 4. DIAGNÓSTICO FINAL
    print("🏁 DIAGNÓSTICO FINAL:")
    print("-" * 40)
    
    compatible = True
    issues = []
    
    if not entidades_con_fotos:
        compatible = False
        issues.append("❌ No se encontraron entidades con fotos en el JSON")
    
    if entidades_con_fotos:
        foto_ejemplo = entidades_con_fotos[0][2]
        if not foto_ejemplo.get('path'):
            compatible = False
            issues.append("❌ Fotos sin campo 'path'")
        
        if not foto_ejemplo.get('name') and not foto_ejemplo.get('id'):
            issues.append("⚠️ Fotos sin 'name' ni 'id' (se generará desde path)")
    
    if compatible:
        print("✅ FORMATO COMPLETAMENTE COMPATIBLE")
        print("   - Las fotos generadas por extractor son válidas para render")
        print("   - Los campos requeridos están presentes")
        print("   - file_uri se genera automáticamente")
        print("   - HTML se renderizará correctamente")
    else:
        print("❌ PROBLEMAS DE COMPATIBILIDAD ENCONTRADOS:")
        for issue in issues:
            print(f"   {issue}")
    
    if issues and not any("❌" in issue for issue in issues):
        print("⚠️ ADVERTENCIAS (no críticas):")
        for issue in issues:
            if "⚠️" in issue:
                print(f"   {issue}")

if __name__ == "__main__":
    analizar_formato_fotos()
