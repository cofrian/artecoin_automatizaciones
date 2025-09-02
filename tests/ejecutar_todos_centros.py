#!/usr/bin/env python3
"""
Script para ejecutar extracción para todos los centros y generar análisis de resultados
"""
import os
import sys
import subprocess
import time
from pathlib import Path

def ejecutar_extraccion_todos():
    """Ejecuta la extracción para todos los centros"""
    
    # Parámetros
    xlsx_path = r"Y:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\github_app\artecoin_automatizaciones\excel\proyecto\ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1_V20.xlsx"
    fotos_root = r"C:\Users\IGP\Desktop\02_ENTREGA SONINGEO\1_CONSULTA 1"
    script_path = r"Y:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\github_app\artecoin_automatizaciones\interfaz\extraer_datos_word.py"
    
    print("🚀 EJECUTANDO EXTRACCIÓN PARA TODOS LOS CENTROS")
    print("=" * 60)
    print(f"📊 Excel: {Path(xlsx_path).name}")
    print(f"📁 Fotos: {Path(fotos_root).name}")
    print()
    
    # Comando
    cmd = [
        "python", script_path,
        "--xlsx", xlsx_path,
        "--fotos-root", fotos_root
    ]
    
    # Ejecutar con redirección de salida
    print("⏳ Iniciando extracción (esto puede tomar varios minutos)...")
    start_time = time.time()
    
    try:
        # Ejecutar con salida limitada
        result = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            cwd=r"Y:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\github_app\artecoin_automatizaciones",
            timeout=1800  # 30 minutos máximo
        )
        
        end_time = time.time()
        duration = end_time - start_time
        
        print(f"✅ Proceso completado en {duration:.1f} segundos")
        
        # Mostrar solo las líneas importantes de la salida
        if result.stdout:
            lines = result.stdout.split('\n')
            important_lines = [line for line in lines if any(keyword in line.lower() for keyword in [
                'procesando centro', 'centro:', 'completado', 'error', 'warning', 'finalizado'
            ])]
            
            if important_lines:
                print("\n📋 RESUMEN DE EJECUCIÓN:")
                print("-" * 40)
                for line in important_lines[-20:]:  # Últimas 20 líneas importantes
                    if line.strip():
                        print(f"  {line}")
        
        if result.stderr:
            print(f"\n⚠️ Errores: {result.stderr[:500]}...")
            
        return result.returncode == 0
        
    except subprocess.TimeoutExpired:
        print("❌ El proceso ha tardado demasiado (>30 min)")
        return False
    except Exception as e:
        print(f"❌ Error ejecutando: {e}")
        return False

def analizar_resultados():
    """Analiza los archivos TEST_FOTOS generados"""
    
    print("\n🔍 ANALIZANDO RESULTADOS DE MATCHING DE FOTOS")
    print("=" * 60)
    
    # Buscar archivos TEST_FOTOS
    interfaz_dir = Path(r"Y:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\github_app\artecoin_automatizaciones\interfaz")
    test_files = list(interfaz_dir.glob("TEST_FOTOS_*.txt"))
    
    if not test_files:
        print("❌ No se encontraron archivos TEST_FOTOS_*.txt")
        return
    
    print(f"📁 Encontrados {len(test_files)} archivos de test")
    
    # Contadores globales
    total_entidades = 0
    total_sin_fotos = 0
    total_fotos_no_usadas = 0
    problemas_por_tipo = {}
    
    for test_file in sorted(test_files):
        centro = test_file.stem.replace("TEST_FOTOS_", "")
        
        try:
            with open(test_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Contar entidades sin fotos
            sin_fotos = content.count("❌ SIN FOTOS:")
            fotos_no_usadas = content.count("📷 FOTOS NO USADAS:")
            entidades = content.count("Entidades encontradas:")
            
            total_entidades += entidades
            total_sin_fotos += sin_fotos
            total_fotos_no_usadas += fotos_no_usadas
            
            # Extraer tipos problemáticos
            lines = content.split('\n')
            for line in lines:
                if "❌ SIN FOTOS:" in line and "(" in line:
                    tipo = line.split("(")[0].replace("❌ SIN FOTOS:", "").strip()
                    if tipo:
                        problemas_por_tipo[tipo] = problemas_por_tipo.get(tipo, 0) + 1
            
            print(f"  {centro}: {sin_fotos} sin fotos, {fotos_no_usadas} fotos no usadas")
            
        except Exception as e:
            print(f"❌ Error leyendo {test_file}: {e}")
    
    # Resumen global
    print(f"\n📊 RESUMEN GLOBAL:")
    print(f"  📋 Total entidades: {total_entidades}")
    print(f"  ❌ Sin fotos: {total_sin_fotos}")
    print(f"  📷 Fotos no usadas: {total_fotos_no_usadas}")
    
    if total_entidades > 0:
        coverage = ((total_entidades - total_sin_fotos) / total_entidades) * 100
        print(f"  ✅ Cobertura: {coverage:.1f}%")
    
    # Top tipos problemáticos
    if problemas_por_tipo:
        print(f"\n🔥 TOP TIPOS SIN FOTOS:")
        for tipo, count in sorted(problemas_por_tipo.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"  {tipo}: {count}")

if __name__ == "__main__":
    print("🎯 SCRIPT DE EJECUCIÓN Y ANÁLISIS COMPLETO")
    print("=" * 60)
    
    # Ejecutar extracción
    success = ejecutar_extraccion_todos()
    
    if success:
        print("✅ Extracción completada exitosamente")
        
        # Esperar un momento para que se escriban los archivos
        time.sleep(2)
        
        # Analizar resultados
        analizar_resultados()
        
    else:
        print("❌ La extracción falló")
    
    print("\n🏁 Proceso finalizado")
