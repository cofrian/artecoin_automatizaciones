#!/usr/bin/env python3
"""
Script para analizar todos los archivos TEST_FOTOS y identificar patrones de problemas
"""

import os
import re
from pathlib import Path
from collections import defaultdict, Counter

def analizar_test_fotos():
    """Analiza todos los archivos TEST_FOTOS para identificar patrones"""
    
    out_dir = Path("out_context")
    test_files = list(out_dir.glob("TEST_FOTOS_*.txt"))
    
    print(f"📊 ANÁLISIS DE {len(test_files)} CENTROS")
    print("="*80)
    
    # Contadores globales
    entidades_sin_fotos = []
    fotos_sin_usar = []
    fotos_filtradas = []
    fotos_faltantes_excel = []
    tipos_fotos_problematicas = Counter()
    patrones_nombres_sin_usar = Counter()
    patrones_filtradas = Counter()
    
    for test_file in sorted(test_files):
        centro_id = test_file.stem.replace("TEST_FOTOS_", "")
        
        try:
            with open(test_file, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            # [1] Fotos declaradas en Excel pero NO encontradas
            seccion_1 = re.search(r'\[1\] Declaradas en Excel PERO NO encontradas en disco:(.*?)\[2\]', contenido, re.DOTALL)
            if seccion_1:
                lineas_faltantes = [l.strip() for l in seccion_1.group(1).split('\n') if l.strip() and not l.strip() == '(ninguna)']
                for linea in lineas_faltantes:
                    if '→' in linea:
                        fotos_faltantes_excel.append((centro_id, linea.strip()))
            
            # [4] Entidades con 0 fotos
            seccion_4 = re.search(r'\[4\] Entidades con 0 fotos después de Excel \+ secuenciales \+ fallback:(.*?)\[5\]', contenido, re.DOTALL)
            if seccion_4:
                lineas_sin_fotos = [l.strip() for l in seccion_4.group(1).split('\n') if l.strip() and l.strip().startswith('- ')]
                for linea in lineas_sin_fotos:
                    entidad = linea.replace('- ', '')
                    entidades_sin_fotos.append((centro_id, entidad))
            
            # [5] Fotos sin usar
            seccion_5 = re.search(r'\[5\] Ficheros en carpeta SIN usar:(.*?)\[6\]', contenido, re.DOTALL)
            if seccion_5:
                lineas_sin_usar = [l.strip() for l in seccion_5.group(1).split('\n') if l.strip() and l.strip().startswith('- ')]
                for linea in lineas_sin_usar:
                    ruta_foto = linea.replace('- ', '')
                    nombre_foto = Path(ruta_foto).stem
                    fotos_sin_usar.append((centro_id, nombre_foto))
                    
                    # Analizar patrones de nombres sin usar
                    patron = nombre_foto.split('_')[0] if '_' in nombre_foto else nombre_foto
                    patrones_nombres_sin_usar[patron] += 1
            
            # [8] Fotos filtradas
            seccion_8 = re.search(r'\[8\] Fotos filtradas por restricción universal.*?:', contenido, re.DOTALL)
            if seccion_8:
                resto_contenido = contenido[seccion_8.end():]
                lineas_filtradas = [l.strip() for l in resto_contenido.split('\n') if l.strip() and '|' in l]
                for linea in lineas_filtradas:
                    fotos_filtradas.append((centro_id, linea.strip()))
                    
                    # Extraer patrón de foto filtrada
                    if 'Foto:' in linea and 'ID parte' in linea:
                        patron_match = re.search(r"Foto: (\w+)_", linea)
                        if patron_match:
                            patron = patron_match.group(1)
                            patrones_filtradas[patron] += 1
            
        except Exception as e:
            print(f"❌ Error procesando {test_file}: {e}")
    
    # RESULTADOS
    print(f"\n🔍 RESUMEN DE PROBLEMAS ENCONTRADOS:")
    print(f"📊 Entidades sin fotos: {len(entidades_sin_fotos)}")
    print(f"📸 Fotos sin usar: {len(fotos_sin_usar)}")
    print(f"🚫 Fotos filtradas: {len(fotos_filtradas)}")
    print(f"❌ Fotos faltantes del Excel: {len(fotos_faltantes_excel)}")
    
    print(f"\n📋 TOP 10 PATRONES DE FOTOS SIN USAR:")
    for patron, count in patrones_nombres_sin_usar.most_common(10):
        print(f"  {patron}: {count} fotos")
    
    print(f"\n🚫 TOP 10 PATRONES DE FOTOS FILTRADAS:")
    for patron, count in patrones_filtradas.most_common(10):
        print(f"  {patron}: {count} fotos")
    
    # Análisis por tipo de entidad sin fotos
    tipos_sin_fotos = Counter()
    for centro_id, entidad in entidades_sin_fotos:
        tipo = entidad.split(':')[0] if ':' in entidad else 'DESCONOCIDO'
        tipos_sin_fotos[tipo] += 1
    
    print(f"\n📊 TIPOS DE ENTIDADES SIN FOTOS:")
    for tipo, count in tipos_sin_fotos.most_common():
        print(f"  {tipo}: {count} entidades")
    
    # Casos específicos problemáticos
    print(f"\n🔍 CASOS ESPECÍFICOS PROBLEMÁTICOS:")
    
    print(f"\n❌ FOTOS FALTANTES DEL EXCEL (primeras 10):")
    for i, (centro_id, linea) in enumerate(fotos_faltantes_excel[:10]):
        print(f"  {centro_id}: {linea}")
    
    print(f"\n📸 FOTOS SIN USAR (primeras 10 por patrón):")
    for patron in ["B001", "B002", "CDRO", "QH", "CR"]:
        ejemplos = [(c, f) for c, f in fotos_sin_usar if f.startswith(patron)][:3]
        if ejemplos:
            print(f"  Patrón {patron}:")
            for centro_id, foto in ejemplos:
                print(f"    {centro_id}: {foto}")
    
    print(f"\n🚫 FOTOS FILTRADAS (primeras 10):")
    for i, (centro_id, linea) in enumerate(fotos_filtradas[:10]):
        print(f"  {centro_id}: {linea}")
    
    return {
        'entidades_sin_fotos': entidades_sin_fotos,
        'fotos_sin_usar': fotos_sin_usar,
        'fotos_filtradas': fotos_filtradas,
        'fotos_faltantes_excel': fotos_faltantes_excel,
        'patrones_nombres_sin_usar': patrones_nombres_sin_usar,
        'patrones_filtradas': patrones_filtradas
    }

if __name__ == "__main__":
    resultados = analizar_test_fotos()
