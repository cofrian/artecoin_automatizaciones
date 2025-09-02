#!/usr/bin/env python3
"""
Debug avanzado para el matching CDRO
"""

import re

def debug_matching_completo():
    """Test completo del sistema de matching"""
    
    # Caso problemático
    tipo = "ACOM"
    ident = "C0011E001"
    foto_name = "CDRO_PPAL_FD0001"
    
    print("🔍 DEBUG COMPLETO DEL MATCHING")
    print("=" * 60)
    print(f"📋 Tipo entidad: {tipo}")
    print(f"🆔 ID entidad: {ident}")  
    print(f"📷 Nombre foto: {foto_name}")
    print()
    
    # Simular el proceso completo
    foto_clean = foto_name.upper()
    
    # Remover prefijos
    prefixes_to_remove = ["FOTO_", "IMG_", "IMAGE_"]
    for prefix in prefixes_to_remove:
        if foto_clean.startswith(prefix):
            foto_clean = foto_clean[len(prefix):]
            break
    
    print(f"🧹 Foto limpia: {foto_clean}")
    
    # MAPEO INTELIGENTE POR TIPOS DE ENTIDAD
    PATRON_MAPEO_TIPOS = {
        "ACOM": ["CDRO", "CUADRO", "ELECT", "ACOM"],
        "EQHORIZ": ["QH", "BOMB", "B00", "BOMBA", "EQUIPO"],
        "SISTCC": ["QG", "CALEF", "BOMB", "SIST", "SISTEMA", "B00"],
        "CLIMA": ["QE", "QI", "CLIMA", "CLIM", "AC", "HVAC"],
        "ILUM": ["I00", "ILUM", "LUZ", "LAMP", "LED"],
        "ENVOL": ["CR", "ENVOL", "CERR", "FACH", "VENT"],
        "ELEVA": ["QV", "ELEV", "ASCEN", "MONTAC"],
        "OTROSEQ": ["OTROS", "EQUIP", "MAQUIN"],
        "DEPENDENCIA": ["D00", "DEP", "SALA", "AULA"],
        "EDIFICIO": ["E00", "EDIF", "BLOQ", "NAVE"],
        "CENTRO": ["C00", "CENT", "FC"]
    }
    
    # Extraer primera parte
    foto_parts = foto_clean.split('_')
    foto_id_part = foto_parts[0] if foto_parts else foto_clean
    print(f"🎯 Foto ID part: '{foto_id_part}'")
    print()
    
    # EJECUTAR TODOS LOS MÉTODOS
    incluir_foto = False
    metodo_usado = None
    
    print("🚀 EJECUTANDO MÉTODOS DE MATCHING:")
    print("-" * 40)
    
    # MÉTODO 1: Matching exacto tradicional
    print(f"1️⃣ MÉTODO 1: Matching exacto")
    print(f"   Buscar '{foto_id_part}' en '{ident.upper()}'")
    if foto_id_part in ident.upper():
        incluir_foto = True
        metodo_usado = "MÉTODO 1"
        print(f"   ✅ MATCH encontrado!")
    else:
        print(f"   ❌ No match")
    print()
    
    # MÉTODO 2: Matching por patrones de tipo de entidad
    if not incluir_foto and tipo in PATRON_MAPEO_TIPOS:
        patrones_tipo = PATRON_MAPEO_TIPOS[tipo]
        print(f"2️⃣ MÉTODO 2: Patrones para tipo {tipo}")
        print(f"   Patrones: {patrones_tipo}")
        print(f"   Probar con foto: '{foto_clean}'")
        
        for i, patron in enumerate(patrones_tipo):
            startswith = foto_clean.startswith(patron)
            contains = patron in foto_clean
            print(f"   [{i+1}] Patrón '{patron}':")
            print(f"       - startswith('{patron}'): {startswith}")
            print(f"       - '{patron}' in foto: {contains}")
            
            if startswith or contains:
                incluir_foto = True
                metodo_usado = f"MÉTODO 2 (patrón '{patron}')"
                print(f"       ✅ MATCH!")
                break
            else:
                print(f"       ❌ No match")
        print()
    
    # MÉTODO 3: Matching específico para equipos con numeración
    if not incluir_foto:
        print(f"3️⃣ MÉTODO 3: Equipos con numeración")
        if re.match(r'^[A-Z]{1,2}\d+', foto_id_part):
            letra_match = re.match(r'^([A-Z]{1,2})', foto_id_part)
            if letra_match:
                prefijo_foto = letra_match.group(1)
                print(f"   Prefijo extraído: '{prefijo_foto}'")
                print(f"   Buscar '{prefijo_foto}' en '{ident.upper()}'")
                if prefijo_foto in ident.upper():
                    incluir_foto = True
                    metodo_usado = f"MÉTODO 3 (prefijo '{prefijo_foto}')"
                    print(f"   ✅ MATCH!")
                else:
                    print(f"   ❌ No match")
            else:
                print(f"   ❌ No se pudo extraer prefijo")
        else:
            print(f"   ❌ No es patrón de equipo con numeración")
        
        # Bombas/equipos B
        if foto_id_part.startswith('B00') or foto_id_part.startswith('BOMB'):
            print(f"   Foto es bomba/equipo B")
            if tipo in ["EQHORIZ", "SISTCC", "ACOM"]:
                incluir_foto = True
                metodo_usado = "MÉTODO 3 (bomba para tipo compatible)"
                print(f"   ✅ MATCH! (bomba para {tipo})")
            else:
                print(f"   ❌ Tipo {tipo} no compatible con bombas")
        print()
    
    # MÉTODO 4: Centro
    if tipo == "CENTRO" and re.match(r'^C0*\d+$', foto_id_part):
        incluir_foto = True
        metodo_usado = "MÉTODO 4 (centro)"
        print(f"4️⃣ MÉTODO 4: ✅ MATCH (centro)")
    
    # MÉTODO 5: Dependencia/Edificio por números
    if not incluir_foto and tipo in ["DEPENDENCIA", "EDIFICIO"]:
        print(f"5️⃣ MÉTODO 5: Números dependencia/edificio")
        numeros_entidad = re.findall(r'\d+', ident)
        numeros_foto = re.findall(r'\d+', foto_id_part)
        print(f"   Números entidad: {numeros_entidad}")
        print(f"   Números foto: {numeros_foto}")
        
        coincidencias = [num for num in numeros_foto if num in numeros_entidad]
        if coincidencias:
            incluir_foto = True
            metodo_usado = f"MÉTODO 5 (números: {coincidencias})"
            print(f"   ✅ MATCH! Números coincidentes: {coincidencias}")
        else:
            print(f"   ❌ No hay números coincidentes")
        print()
    
    # MÉTODO 6: Proximidad equipos
    if not incluir_foto and tipo in ["EQHORIZ", "SISTCC", "CLIMA"]:
        print(f"6️⃣ MÉTODO 6: Proximidad equipos para {tipo}")
        equipos_generales = ["BOMBA", "BOMB", "EQUIP", "MAQUIN", "MOTOR", "PANEL"]
        matches = [eq for eq in equipos_generales if eq in foto_clean]
        if matches:
            incluir_foto = True
            metodo_usado = f"MÉTODO 6 (equipos: {matches})"
            print(f"   ✅ MATCH! Equipos encontrados: {matches}")
        else:
            print(f"   ❌ No hay equipos generales en foto")
        print()
    
    # MÉTODO 7: Cuadros eléctricos ACOM
    if not incluir_foto and tipo == "ACOM":
        print(f"7️⃣ MÉTODO 7: Cuadros eléctricos para ACOM")
        cuadros_electricos = ["CDRO", "CUADRO", "ELECT", "PANEL", "ARMARIO"]
        print(f"   Patrones cuadros: {cuadros_electricos}")
        print(f"   Foto limpia: '{foto_clean}'")
        
        matches = []
        for cuadro in cuadros_electricos:
            if cuadro in foto_clean:
                matches.append(cuadro)
                print(f"   ✅ '{cuadro}' encontrado en foto")
            else:
                print(f"   ❌ '{cuadro}' no encontrado")
        
        if matches:
            incluir_foto = True
            metodo_usado = f"MÉTODO 7 (cuadros: {matches})"
            print(f"   ✅ MATCH! Cuadros encontrados: {matches}")
        else:
            print(f"   ❌ No hay cuadros eléctricos en foto")
        print()
    
    print("🎯 RESULTADO FINAL:")
    print("=" * 40)
    if incluir_foto:
        print(f"✅ INCLUIR FOTO")
        print(f"🔧 Método usado: {metodo_usado}")
    else:
        print(f"❌ FILTRAR FOTO") 
        print(f"😞 Ningún método hizo match")
    
    return incluir_foto, metodo_usado

if __name__ == "__main__":
    debug_matching_completo()
