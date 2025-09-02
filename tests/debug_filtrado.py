#!/usr/bin/env python3
"""
Script para analizar el problema de filtrado de fotos
"""
import re
from pathlib import Path

# Datos del reporte TEST_FOTOS
fotos_fallback = [
    ("CENTRO", "C0007", ["C001_FC0001"]),
    ("EDIFICIO", "C0007E001", ["E001_FE0001"]),
    ("DEPENDENCIA", "C0007E001D0001", ["D0001_FD0001"]),
    ("DEPENDENCIA", "C0007E001D0002", ["D0002_FD0001"]),
    ("CLIMA", "C0007E001D0001QE001", ["QE001_FQE0001", "QE001_FQE0002"]),
    ("CLIMA", "C0007E001D0001QE004", ["QE004_FQE0001", "QE004_FQE0002"]),
    ("EQHORIZ", "C0007E001D0002QH0002", ["QH0002_FD0001"])
]

# Fotos disponibles en carpeta (muestra)
fotos_disponibles = [
    "CR00001_FF0001", "CR00002_FVN0001", "D0001_FD0001", "D0002_FD0001", 
    "E001_FE0001", "QE001_FQE0001", "QE001_FQE0002", "QE004_FQE0001", 
    "QH0002_FD0001", "I00001_FI001", "C001_FC0001"
]

def test_filtrado(tipo, ident, foto_name):
    """Simula el filtrado del código actual"""
    
    # Aplicar filtrado solo si NO viene del Excel (simulamos que todas son adicionales)
    foto_clean = foto_name.upper()
    
    # Remover prefijos comunes
    prefixes_to_remove = ["FOTO_", "IMG_", "IMAGE_"]
    for prefix in prefixes_to_remove:
        if foto_clean.startswith(prefix):
            foto_clean = foto_clean[len(prefix):]
            break
    
    # Extraer la primera parte del nombre de foto (antes del primer _)
    foto_parts = foto_clean.split('_')
    foto_id_part = foto_parts[0] if foto_parts else foto_clean
    
    # Verificar si el ID de la foto está contenido en el ID de la entidad
    incluir_foto = False
    if foto_id_part in ident.upper():
        incluir_foto = True
    
    # EXCEPCIÓN ESPECIAL para fotos de centro
    if tipo == "CENTRO" and re.match(r'^C0*\d+$', foto_id_part):
        incluir_foto = True
    
    return incluir_foto, foto_id_part

print("🔍 ANÁLISIS DE FILTRADO DE FOTOS")
print("="*60)

for tipo, ident, fotos_esperadas in fotos_fallback:
    print(f"\n📋 {tipo}:{ident}")
    
    for foto_name in fotos_esperadas:
        incluir, foto_id_part = test_filtrado(tipo, ident, foto_name)
        status = "✅ PASA" if incluir else "❌ FILTRADA"
        print(f"  {foto_name} → parte: '{foto_id_part}' → {status}")

print(f"\n{'='*60}")
print("🔍 PRUEBA CON TODAS LAS FOTOS DISPONIBLES")
print("="*60)

# Verificar todas las fotos contra todas las entidades
for foto_name in fotos_disponibles:
    print(f"\n📷 Foto: {foto_name}")
    foto_matcheada = False
    
    for tipo, ident, _ in fotos_fallback:
        incluir, foto_id_part = test_filtrado(tipo, ident, foto_name)
        if incluir:
            print(f"   ✅ Matches {tipo}:{ident} (parte: '{foto_id_part}')")
            foto_matcheada = True
    
    if not foto_matcheada:
        foto_parts = foto_name.upper().split('_')
        foto_id_part = foto_parts[0] if foto_parts else foto_name
        print(f"   ❌ NO MATCH (parte: '{foto_id_part}' no encontrada en ninguna entidad)")
