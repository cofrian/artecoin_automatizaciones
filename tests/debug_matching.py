#!/usr/bin/env python3
"""
Script para debuggear el matching de fotos CDRO con entidades ACOM
"""

import re

def test_matching_logic():
    """Testear la lógica de matching con el caso problemático"""
    
    # Caso problemático real
    tipo = "ACOM"
    ident = "C0011E001" 
    foto_name = "CDRO_PPAL_FD0001"
    
    print("🐛 DEBUG DEL MATCHING INTELIGENTE")
    print("=" * 50)
    print(f"Tipo: {tipo}")
    print(f"Entidad ID: {ident}")
    print(f"Foto: {foto_name}")
    print()
    
    # Lógica igual que en el código
    foto_clean = foto_name.upper()
    
    # Remover prefijos comunes
    prefixes_to_remove = ["FOTO_", "IMG_", "IMAGE_"]
    for prefix in prefixes_to_remove:
        if foto_clean.startswith(prefix):
            foto_clean = foto_clean[len(prefix):]
            break
    
    print(f"Foto limpia: {foto_clean}")
    
    # MAPEO INTELIGENTE
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
    print(f"Foto ID part: {foto_id_part}")
    
    # MÉTODO 1: Matching exacto
    incluir_foto = False
    if foto_id_part in ident.upper():
        incluir_foto = True
        print("✅ MÉTODO 1: Match exacto")
    else:
        print(f"❌ MÉTODO 1: '{foto_id_part}' no está en '{ident.upper()}'")
    
    # MÉTODO 2: Matching por patrones de tipo
    if not incluir_foto and tipo in PATRON_MAPEO_TIPOS:
        patrones_tipo = PATRON_MAPEO_TIPOS[tipo]
        print(f"🔍 MÉTODO 2: Probando patrones para {tipo}: {patrones_tipo}")
        
        for patron in patrones_tipo:
            if foto_clean.startswith(patron) or patron in foto_clean:
                incluir_foto = True
                print(f"✅ MÉTODO 2: Match con patrón '{patron}'")
                break
            else:
                print(f"  ❌ Patrón '{patron}': no match con '{foto_clean}'")
    
    if not incluir_foto:
        print("❌ MÉTODO 2: No match")
    
    # MÉTODO 3: Matching específico para equipos
    if not incluir_foto:
        print("🔍 MÉTODO 3: Checking equipos...")
        if re.match(r'^[A-Z]{1,2}\d+', foto_id_part):
            letra_match = re.match(r'^([A-Z]{1,2})', foto_id_part)
            if letra_match:
                prefijo_foto = letra_match.group(1)
                print(f"  Prefijo extraído: {prefijo_foto}")
                if prefijo_foto in ident.upper():
                    incluir_foto = True
                    print("✅ MÉTODO 3: Match por prefijo")
                else:
                    print(f"  ❌ '{prefijo_foto}' no está en '{ident.upper()}'")
        
        if foto_id_part.startswith('B00') or foto_id_part.startswith('BOMB'):
            if tipo in ["EQHORIZ", "SISTCC", "ACOM"]:
                incluir_foto = True
                print("✅ MÉTODO 3: Match por bomba")
    
    # MÉTODO 7: Cuadros eléctricos específico
    if not incluir_foto and tipo == "ACOM":
        cuadros_electricos = ["CDRO", "CUADRO", "ELECT", "PANEL", "ARMARIO"]
        print(f"🔍 MÉTODO 7: Probando cuadros eléctricos: {cuadros_electricos}")
        
        for cuadro in cuadros_electricos:
            if cuadro in foto_clean:
                incluir_foto = True
                print(f"✅ MÉTODO 7: Match con '{cuadro}' en '{foto_clean}'")
                break
            else:
                print(f"  ❌ '{cuadro}' no está en '{foto_clean}'")
    
    print()
    print(f"🎯 RESULTADO FINAL: {'✅ INCLUIR' if incluir_foto else '❌ FILTRAR'}")
    
    return incluir_foto

if __name__ == "__main__":
    test_matching_logic()
