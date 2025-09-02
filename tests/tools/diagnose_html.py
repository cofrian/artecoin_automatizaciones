# -*- coding: utf-8 -*-
"""
Herramienta de diagnóstico para archivos HTML problemáticos.
Analiza archivos HTML que fallan en la conversión a PDF para identificar problemas comunes.
"""

import os
import sys
from pathlib import Path
from urllib.parse import quote
import argparse

def analyze_html_file(html_path: str) -> dict:
    """Analiza un archivo HTML y reporta posibles problemas"""
    
    result = {
        'file_path': html_path,
        'exists': False,
        'size': 0,
        'encoding_ok': False,
        'basic_structure': False,
        'has_images': False,
        'has_scripts': False,
        'potential_issues': [],
        'content_preview': '',
        'file_uri': ''
    }
    
    html_file = Path(html_path)
    
    # 1. Verificar que el archivo existe
    if not html_file.exists():
        result['potential_issues'].append("El archivo no existe")
        return result
    
    result['exists'] = True
    
    # 2. Verificar tamaño del archivo
    try:
        file_stat = html_file.stat()
        result['size'] = file_stat.st_size
        
        if file_stat.st_size == 0:
            result['potential_issues'].append("El archivo está vacío (0 bytes)")
            return result
            
        if file_stat.st_size > 10 * 1024 * 1024:  # 10MB
            result['potential_issues'].append(f"Archivo muy grande ({file_stat.st_size / (1024*1024):.1f}MB)")
            
    except Exception as e:
        result['potential_issues'].append(f"Error al obtener información del archivo: {e}")
        return result
    
    # 3. Generar URI del archivo para Playwright
    try:
        abs_path = html_file.resolve()
        uri_path = str(abs_path).replace("\\", "/")
        result['file_uri'] = "file:///" + quote(uri_path, safe="/:._-()")
    except Exception as e:
        result['potential_issues'].append(f"Error generando URI: {e}")
    
    # 4. Intentar leer el contenido
    encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
    content = None
    
    for encoding in encodings_to_try:
        try:
            with open(html_file, 'r', encoding=encoding, errors='replace') as f:
                content = f.read()
                result['encoding_ok'] = True
                break
        except Exception as e:
            continue
    
    if content is None:
        result['potential_issues'].append("No se pudo leer el archivo con ninguna codificación")
        return result
    
    # 5. Análisis del contenido
    content_lower = content.lower()
    
    # Vista previa del contenido (primeros 500 caracteres)
    result['content_preview'] = content[:500] + ("..." if len(content) > 500 else "")
    
    # Verificar estructura HTML básica
    html_tags = ['<html', '<head', '<body', '<!doctype']
    if any(tag in content_lower for tag in html_tags):
        result['basic_structure'] = True
    else:
        result['potential_issues'].append("No contiene estructura HTML básica")
    
    # Verificar elementos problemáticos
    if '<img' in content_lower:
        result['has_images'] = True
        # Contar imágenes
        img_count = content_lower.count('<img')
        if img_count > 20:
            result['potential_issues'].append(f"Muchas imágenes ({img_count}), puede causar problemas de memoria")
    
    if '<script' in content_lower:
        result['has_scripts'] = True
        result['potential_issues'].append("Contiene JavaScript - puede causar problemas en la conversión")
    
    # Verificar caracteres problemáticos
    if '\x00' in content:
        result['potential_issues'].append("Contiene caracteres nulos (\\x00)")
    
    # Verificar rutas de imágenes rotas
    if 'src="' in content:
        import re
        src_patterns = re.findall(r'src=["\']([^"\']+)["\']', content, re.IGNORECASE)
        broken_images = []
        for src in src_patterns[:10]:  # Revisar solo las primeras 10
            if src.startswith(('http://', 'https://')):
                continue  # Skip URLs externas
            if src.startswith('data:'):
                continue  # Skip data URIs
                
            # Verificar rutas relativas
            if not os.path.isabs(src):
                img_path = html_file.parent / src
            else:
                img_path = Path(src)
                
            if not img_path.exists():
                broken_images.append(src)
        
        if broken_images:
            result['potential_issues'].append(f"Imágenes no encontradas: {broken_images[:3]}")
    
    # Verificar tamaño del HTML
    if len(content) > 1024 * 1024:  # 1MB
        result['potential_issues'].append(f"HTML muy grande ({len(content) / 1024:.0f}KB)")
    
    # Verificar elementos que pueden causar problemas en Playwright
    problematic_elements = [
        ('iframe', 'Contiene iframes'),
        ('embed', 'Contiene elementos embed'),
        ('object', 'Contiene elementos object'),
        ('video', 'Contiene elementos de video'),
        ('audio', 'Contiene elementos de audio')
    ]
    
    for element, message in problematic_elements:
        if f'<{element}' in content_lower:
            result['potential_issues'].append(message)
    
    return result

def format_analysis_report(analysis: dict) -> str:
    """Formatea el análisis en un reporte legible"""
    
    report = []
    report.append("=" * 80)
    report.append(f"DIAGNÓSTICO HTML: {analysis['file_path']}")
    report.append("=" * 80)
    
    # Estado básico
    report.append(f"✅ Archivo existe: {'Sí' if analysis['exists'] else 'No'}")
    if analysis['exists']:
        report.append(f"📊 Tamaño: {analysis['size']:,} bytes")
        report.append(f"🔤 Codificación: {'OK' if analysis['encoding_ok'] else 'ERROR'}")
        report.append(f"🏗️ Estructura HTML: {'Válida' if analysis['basic_structure'] else 'Inválida'}")
        report.append(f"🖼️ Contiene imágenes: {'Sí' if analysis['has_images'] else 'No'}")
        report.append(f"⚡ Contiene JavaScript: {'Sí' if analysis['has_scripts'] else 'No'}")
    
    # URI para Playwright
    if analysis['file_uri']:
        report.append(f"🔗 URI generada: {analysis['file_uri']}")
    
    # Problemas encontrados
    if analysis['potential_issues']:
        report.append("\n⚠️ PROBLEMAS DETECTADOS:")
        for i, issue in enumerate(analysis['potential_issues'], 1):
            report.append(f"   {i}. {issue}")
    else:
        report.append("\n✅ No se detectaron problemas evidentes")
    
    # Vista previa del contenido
    if analysis['content_preview']:
        report.append("\n📝 VISTA PREVIA DEL CONTENIDO:")
        report.append("-" * 40)
        report.append(analysis['content_preview'])
        report.append("-" * 40)
    
    return "\n".join(report)

def main():
    parser = argparse.ArgumentParser(description='Diagnosticar archivos HTML problemáticos')
    parser.add_argument('html_file', help='Ruta al archivo HTML a analizar')
    parser.add_argument('--output', '-o', help='Archivo de salida para el reporte (opcional)')
    
    args = parser.parse_args()
    
    # Analizar el archivo
    print("🔍 Analizando archivo HTML...")
    analysis = analyze_html_file(args.html_file)
    
    # Generar reporte
    report = format_analysis_report(analysis)
    
    # Mostrar en consola
    print(report)
    
    # Guardar en archivo si se especifica
    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(report)
            print(f"\n💾 Reporte guardado en: {args.output}")
        except Exception as e:
            print(f"\n❌ Error guardando reporte: {e}")
    
    # Sugerencias de solución
    print("\n" + "=" * 80)
    print("🔧 SUGERENCIAS DE SOLUCIÓN:")
    print("=" * 80)
    
    if not analysis['exists']:
        print("• Verificar que la ruta del archivo sea correcta")
        print("• Verificar que el archivo no haya sido movido o eliminado")
    elif not analysis['encoding_ok']:
        print("• El archivo puede estar corrupto o tener codificación inválida")
        print("• Intentar regenerar el archivo HTML")
    elif not analysis['basic_structure']:
        print("• El archivo no parece ser HTML válido")
        print("• Verificar el proceso de generación del archivo")
    elif analysis['potential_issues']:
        print("• Considerar regenerar el archivo con menos elementos problemáticos")
        print("• Verificar que todas las imágenes referenciadas existan")
        print("• Reducir la complejidad del HTML si es posible")
    else:
        print("• El archivo parece estar bien formado")
        print("• El problema puede ser específico de Playwright/Chromium")
        print("• Intentar con diferentes parámetros de conversión")

if __name__ == '__main__':
    main()
