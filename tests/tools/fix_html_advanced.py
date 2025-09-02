#!/usr/bin/env python3
"""
Herramienta avanzada para reparar archivos HTML con referencias de imágenes rotas.
Maneja correctamente URIs con codificación URL y valida la existencia de archivos.
"""

import re
import os
import sys
from urllib.parse import unquote
from pathlib import Path

def decode_file_uri(uri):
    """Decodifica una URI file:/// y devuelve la ruta del archivo."""
    if uri.startswith('file:///'):
        # Remover el prefijo file:///
        path = uri[8:]  # Saltar 'file:///'
        # Decodificar URL encoding (%20, %C3%8D, etc.)
        decoded_path = unquote(path)
        return decoded_path
    return uri

def find_images_in_html(html_content):
    """Encuentra todas las referencias de imágenes en el HTML."""
    # Patrón más completo para encontrar imágenes
    patterns = [
        r'<img[^>]+src=[\'"]([^\'"]+)[\'"]',
        r'background-image:\s*url\([\'"]?([^\'"]+)[\'"]?\)',
        r'content:[\'"][^\']*url\([\'"]?([^\'"]+)[\'"]?\)',
    ]
    
    images = []
    for pattern in patterns:
        matches = re.findall(pattern, html_content, re.IGNORECASE)
        images.extend(matches)
    
    return list(set(images))  # Eliminar duplicados

def check_image_exists(image_path):
    """Verifica si una imagen existe."""
    try:
        if image_path.startswith('file:///'):
            decoded_path = decode_file_uri(image_path)
            return os.path.exists(decoded_path)
        else:
            return os.path.exists(image_path)
    except Exception:
        return False

def fix_broken_images_in_html(html_content, html_file_path=None):
    """
    Remueve referencias a imágenes rotas del HTML.
    Retorna el HTML modificado y un resumen de cambios.
    """
    original_content = html_content
    images = find_images_in_html(html_content)
    
    removed_images = []
    kept_images = []
    
    print(f"🔍 Encontradas {len(images)} referencias de imágenes")
    
    for img_src in images:
        exists = check_image_exists(img_src)
        
        if exists:
            kept_images.append(img_src)
            print(f"✅ Conservada: {img_src}")
        else:
            removed_images.append(img_src)
            print(f"❌ Eliminando: {img_src}")
            
            # Diferentes estrategias para remover la imagen
            # 1. Remover tag img completo
            img_pattern = rf'<img[^>]*src=[\'\"]{re.escape(img_src)}[\'\"][^>]*/?>'
            html_content = re.sub(img_pattern, '', html_content, flags=re.IGNORECASE)
            
            # 2. Remover background-image CSS
            bg_pattern = rf'background-image:\s*url\([\'\"]{re.escape(img_src)}[\'\"]\)\s*;?'
            html_content = re.sub(bg_pattern, '', html_content, flags=re.IGNORECASE)
            
            # 3. Remover de estilos inline
            style_pattern = rf'style=[\'\"][^\'\"]*url\([\'\"]{re.escape(img_src)}[\'\"]\)[^\'\"]*[\'\"]\s*'
            html_content = re.sub(style_pattern, '', html_content, flags=re.IGNORECASE)
    
    # Limpiar espacios múltiples y líneas vacías resultantes
    html_content = re.sub(r'\n\s*\n', '\n', html_content)
    html_content = re.sub(r'  +', ' ', html_content)
    
    changes_made = len(removed_images) > 0
    
    summary = {
        'file': html_file_path or 'HTML content',
        'total_images': len(images),
        'kept_images': len(kept_images),
        'removed_images': len(removed_images),
        'removed_list': removed_images,
        'kept_list': kept_images,
        'changes_made': changes_made
    }
    
    return html_content, summary

def fix_html_file(file_path):
    """Repara un archivo HTML individual."""
    try:
        # Leer con múltiples codificaciones
        encodings = ['utf-8', 'latin1', 'cp1252']
        content = None
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            print(f"❌ No se pudo leer {file_path} con ninguna codificación")
            return False
        
        fixed_content, summary = fix_broken_images_in_html(content, file_path)
        
        if summary['changes_made']:
            # Crear backup
            backup_path = f"{file_path}.backup"
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"📄 Backup guardado: {backup_path}")
            
            # Guardar archivo reparado
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(fixed_content)
            
            print(f"✅ Archivo reparado: {file_path}")
            print(f"   📊 Imágenes: {summary['kept_images']} conservadas, {summary['removed_images']} eliminadas")
        else:
            print(f"ℹ️  No se necesitaron cambios: {file_path}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error procesando {file_path}: {e}")
        return False

def fix_directory(directory_path):
    """Repara todos los archivos HTML en un directorio."""
    html_files = []
    
    # Buscar archivos HTML recursivamente
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.lower().endswith('.html'):
                html_files.append(os.path.join(root, file))
    
    print(f"🔍 Encontrados {len(html_files)} archivos HTML en {directory_path}")
    
    success_count = 0
    for html_file in html_files:
        print(f"\n📝 Procesando: {html_file}")
        if fix_html_file(html_file):
            success_count += 1
    
    print(f"\n✅ Procesados exitosamente: {success_count}/{len(html_files)}")
    return success_count

def main():
    if len(sys.argv) < 2:
        print("Uso: python fix_html_advanced.py <archivo.html | directorio>")
        sys.exit(1)
    
    path = sys.argv[1]
    
    print("🔧 HERRAMIENTA AVANZADA DE REPARACIÓN HTML")
    print("=" * 60)
    
    if os.path.isfile(path):
        print(f"📁 Archivo: {path}\n")
        fix_html_file(path)
    elif os.path.isdir(path):
        print(f"📁 Directorio: {path}\n")
        fix_directory(path)
    else:
        print(f"❌ Error: '{path}' no existe o no es accesible")
        sys.exit(1)

if __name__ == "__main__":
    main()
