#!/usr/bin/env python3
"""
Script para calcular estadísticas de cobertura de fotos basándose en el JSON ya generado.
"""
import json
import os
from pathlib import Path
from collections import defaultdict

def calcular_estadisticas_cobertura(json_path, carpeta_fotos):
    """Calcula estadísticas de cobertura de fotos"""
    
    # Cargar JSON
    print(f"📖 Leyendo JSON: {Path(json_path).name}")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Contar fotos totales en carpeta
    print(f"📁 Escaneando carpeta: {carpeta_fotos}")
    fotos_carpeta = []
    carpeta_referencias = Path(carpeta_fotos)
    if carpeta_referencias.exists():
        for ext in ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.gif']:
            fotos_carpeta.extend(carpeta_referencias.rglob(ext))
    
    total_fotos_carpeta = len(fotos_carpeta)
    nombres_fotos_carpeta = {foto.stem for foto in fotos_carpeta}
    
    # Extraer fotos usadas del JSON
    fotos_usadas = set()
    fotos_por_entidad = defaultdict(int)
    
    def procesar_entidad(entidad, tipo, entidad_id=""):
        """Procesa una entidad para extraer fotos"""
        if isinstance(entidad, dict):
            fotos_nombres = entidad.get('fotos_nombres', [])
            if fotos_nombres:
                for foto_nombre in fotos_nombres:
                    fotos_usadas.add(foto_nombre)
                    fotos_por_entidad[f"{tipo}:{entidad_id}"] += 1
    
    # Procesar centro
    centro = data.get("centro", {})
    procesar_entidad(centro, "CENTRO", centro.get("id", ""))
    
    # Procesar edificios y sus elementos
    edificios = data.get("edificios", [])
    for edificio in edificios:
        edificio_id = edificio.get("id", "")
        procesar_entidad(edificio, "EDIFICIO", edificio_id)
        
        # Dependencias
        for dep in edificio.get("dependencias", []):
            procesar_entidad(dep, "DEPENDENCIA", dep.get("id", ""))
        
        # Otros elementos
        for acom in edificio.get("acom", []):
            procesar_entidad(acom, "ACOM", acom.get("id", ""))
        
        for envol in edificio.get("envolventes", []):
            procesar_entidad(envol, "ENVOL", envol.get("id", ""))
            
        for clima in edificio.get("equipos_clima", []):
            procesar_entidad(clima, "CLIMA", clima.get("id", ""))
            
        for horiz in edificio.get("equipos_horiz", []):
            procesar_entidad(horiz, "EQHORIZ", horiz.get("id", ""))
            
        for eleva in edificio.get("elevadores", []):
            procesar_entidad(eleva, "ELEVA", eleva.get("id", ""))
            
        for otros in edificio.get("otros_equipos", []):
            procesar_entidad(otros, "OTROSEQ", otros.get("id", ""))
            
        for ilum in edificio.get("iluminacion", []):
            procesar_entidad(ilum, "ILUM", ilum.get("id", ""))
    
    # Calcular estadísticas
    fotos_usadas_count = len(fotos_usadas)
    porcentaje = (fotos_usadas_count / total_fotos_carpeta * 100) if total_fotos_carpeta > 0 else 0
    fotos_no_usadas = total_fotos_carpeta - fotos_usadas_count
    
    # Mostrar resultados
    print(f"\n{'='*80}")
    print(f"📊 ESTADÍSTICAS DE COBERTURA - {centro.get('nombre', 'CENTRO')}")
    print(f"{'='*80}")
    print(f"📁 Carpeta Referencias: {Path(carpeta_fotos).name}")
    print(f"📷 Total fotos en carpeta: {total_fotos_carpeta}")
    print(f"✅ Fotos incluidas en JSON: {fotos_usadas_count}")
    print(f"📈 PORCENTAJE DE COBERTURA: {porcentaje:.1f}%")
    print(f"🚫 Fotos no usadas: {fotos_no_usadas}")
    print(f"{'='*80}")
    
    # Detalle por tipo de entidad
    if fotos_por_entidad:
        print("📋 DETALLE POR ENTIDAD:")
        for entidad, count in sorted(fotos_por_entidad.items()):
            if count > 0:
                print(f"   {entidad}: {count} fotos")
    
    # Fotos no utilizadas
    fotos_no_utilizadas = nombres_fotos_carpeta - fotos_usadas
    if fotos_no_utilizadas:
        print(f"\n🚫 FOTOS NO UTILIZADAS ({len(fotos_no_utilizadas)}):")
        for foto in sorted(fotos_no_utilizadas):
            print(f"   - {foto}")
    
    print(f"{'='*80}\n")
    
    return {
        'total_fotos': total_fotos_carpeta,
        'fotos_usadas': fotos_usadas_count,
        'porcentaje': porcentaje,
        'fotos_por_entidad': dict(fotos_por_entidad),
        'fotos_no_utilizadas': list(fotos_no_utilizadas)
    }

if __name__ == "__main__":
    json_path = "test_output/C0007_AYUNTAMIENTO.json"
    carpeta_fotos = r"C:\Users\IGP\Desktop\02_ENTREGA SONINGEO\1_CONSULTA 1\C0007_AYUNTAMIENTO\Referencias"
    
    if not os.path.exists(json_path):
        print(f"❌ Error: No se encuentra el archivo JSON {json_path}")
        exit(1)
    
    if not os.path.exists(carpeta_fotos):
        print(f"❌ Error: No se encuentra la carpeta de fotos {carpeta_fotos}")
        exit(1)
    
    stats = calcular_estadisticas_cobertura(json_path, carpeta_fotos)
