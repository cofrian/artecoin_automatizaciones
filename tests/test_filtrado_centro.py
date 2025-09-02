#!/usr/bin/env python3
"""
Script de prueba para verificar la corrección del filtrado de fotos por centro.
"""

import sys
import os
from pathlib import Path

# Agregar el directorio interfaz al PATH para importar el módulo
interfaz_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "interfaz")
sys.path.insert(0, interfaz_dir)

# Importar la función corregida
from extraer_datos_word import _fallback_candidates_optimized

def test_filtrado_centro():
    """Prueba simple para verificar que el filtrado por centro funciona."""
    
    # Simular datos de prueba
    ident = "C0007E001D0007QI007"  # ID de ejemplo para centro C0007
    tipo = "CLIMA"
    
    # Simular índice de fotos con rutas de diferentes centros
    path_to_info = {
        Path("C:/Desktop/SONINGEO/1_CONSULTA 1/C0007_TEST/Referencias/QI007_FQI0001.jpg"): {
            "stem": "QI007_FQI0001",
            "realpath": "C:/Desktop/SONINGEO/1_CONSULTA 1/C0007_TEST/Referencias/QI007_FQI0001.jpg"
        },
        Path("C:/Desktop/SONINGEO/1_CONSULTA 1/C0049_COLEGIO/Referencias/QI007_FQI0001.jpg"): {
            "stem": "QI007_FQI0001", 
            "realpath": "C:/Desktop/SONINGEO/1_CONSULTA 1/C0049_COLEGIO/Referencias/QI007_FQI0001.jpg"
        },
        Path("C:/Desktop/SONINGEO/2_CONSULTA 2/C0016_COLEGIO/Referencias/QI007_FQI0001.jpg"): {
            "stem": "QI007_FQI0001",
            "realpath": "C:/Desktop/SONINGEO/2_CONSULTA 2/C0016_COLEGIO/Referencias/QI007_FQI0001.jpg"
        }
    }
    
    # Simular normalized_index
    normalized_index = {
        "qi007_fqi0001": list(path_to_info.keys())
    }
    
    # Simular exact_index (vacío para esta prueba)
    exact_index = {}
    
    # Ejecutar la función
    print(f"🔍 Probando filtrado para: {ident} (tipo: {tipo})")
    print(f"📁 Fotos disponibles: {len(path_to_info)}")
    
    resultado = _fallback_candidates_optimized(
        ident=ident,
        exact_index=exact_index, 
        normalized_index=normalized_index,
        path_to_info=path_to_info,
        max_photos=6,
        tipo=tipo
    )
    
    print(f"✅ Fotos seleccionadas: {len(resultado)}")
    for foto in resultado:
        print(f"   📸 {foto}")
    
    # Verificar que solo se seleccionó la foto del centro correcto (C0007)
    center_found = any("C0007" in str(foto).upper() for foto in resultado)
    other_centers = any(center in str(foto).upper() for foto in resultado for center in ["C0049", "C0016"])
    
    print(f"\n📊 Resultados:")
    print(f"   Centro C0007 encontrado: {'✅' if center_found else '❌'}")
    print(f"   Otros centros incluidos: {'❌' if other_centers else '✅'}")
    
    if center_found and not other_centers:
        print("🎉 ¡CORRECCIÓN EXITOSA! Solo se incluyen fotos del centro correcto.")
    else:
        print("⚠️  La corrección necesita ajustes.")

if __name__ == "__main__":
    test_filtrado_centro()
