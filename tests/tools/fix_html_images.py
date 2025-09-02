#!/usr/bin/env python3
"""
Herramienta para convertir referencias absolutas de imágenes en HTML a referencias relativas.
Copia las imágenes referenciadas al directorio del HTML para evitar problemas de acceso.
"""

import re
import os
import sys
import shutil
from urllib.parse import unquote
from pathlib import Path

def decode_file_uri(uri):
    """Decodifica una URI file:/// y devuelve la ruta del archivo."""
    if uri.startswith('file:///'):
        path = uri[8:]  # Saltar 'file:///'
        decoded_path = unquote(path)
        return decoded_path
    return uri

def find_images_in_html(html_content):
    """Encuentra todas las referencias de imágenes file:/// en el HTML."""
    pattern = r'file:///[^\'"\s<>]+'
    matches = re.findall(pattern, html_content, re.IGNORECASE)
    return list(set(matches))

def copy_and_fix_images(html_content, html_file_path):
    """
    Copia imágenes referenciadas al directorio del HTML y actualiza las referencias.
    """
    if not html_file_path:
        return html_content, []
    
    html_dir = os.path.dirname(html_file_path)
    images_dir = os.path.join(html_dir, 'images')
    
    # Crear directorio de imágenes si no existe
    os.makedirs(images_dir, exist_ok=True)
    
    image_uris = find_images_in_html(html_content)
    copied_images = []
    
    print(f"🔍 Encontradas {len(image_uris)} referencias de imágenes file:///")
    
    for img_uri in image_uris:
        try:
            # Decodificar la URI
            source_path = decode_file_uri(img_uri)
            
            if os.path.exists(source_path):
                # Obtener el nombre del archivo
                filename = os.path.basename(source_path)
                dest_path = os.path.join(images_dir, filename)
                
                # Copiar la imagen
                shutil.copy2(source_path, dest_path)
                
                # Crear la nueva referencia relativa
                relative_path = f"images/{filename}"
                
                # Reemplazar en el HTML
                html_content = html_content.replace(img_uri, relative_path)
                
                copied_images.append({
                    'original': img_uri,
                    'source': source_path,
                    'destination': dest_path,
                    'relative': relative_path
                })
                
                print(f"✅ Copiada: {filename}")
                print(f"   📁 De: {source_path}")
                print(f"   📁 A:  {dest_path}")
                print(f"   🔗 Nueva referencia: {relative_path}")
            else:
                print(f"❌ No encontrada: {source_path}")
                
        except Exception as e:
            print(f"❌ Error procesando {img_uri}: {e}")
    
    return html_content, copied_images

def fix_html_file_with_images(file_path):
    """Repara un archivo HTML copiando imágenes y actualizando referencias."""
    try:
        # Leer el archivo
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
            print(f"❌ No se pudo leer {file_path}")
            return False
        
        # Copiar imágenes y actualizar referencias
        fixed_content, copied_images = copy_and_fix_images(content, file_path)
        
        if copied_images:
            # Crear backup
            backup_path = f"{file_path}.backup"
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"📄 Backup guardado: {backup_path}")
            
            # Guardar archivo actualizado
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(fixed_content)
            
            print(f"✅ Archivo actualizado: {file_path}")
            print(f"📊 Imágenes procesadas: {len(copied_images)}")
        else:
            print(f"ℹ️  No se encontraron imágenes file:/// para procesar")
        
        return True
        
    except Exception as e:
        print(f"❌ Error procesando {file_path}: {e}")
        return False

def fix_directory_with_images(directory_path):
    """Procesa todos los archivos HTML en un directorio."""
    html_files = []
    
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.lower().endswith('.html'):
                html_files.append(os.path.join(root, file))
    
    print(f"🔍 Encontrados {len(html_files)} archivos HTML en {directory_path}")
    
    success_count = 0
    total_images = 0
    
    for html_file in html_files:
        print(f"\n📝 Procesando: {html_file}")
        if fix_html_file_with_images(html_file):
            success_count += 1
    
    print(f"\n✅ Archivos procesados: {success_count}/{len(html_files)}")

def main():
    if len(sys.argv) < 2:
        print("Uso: python fix_html_images.py <archivo.html | directorio>")
        sys.exit(1)
    
    path = sys.argv[1]
    
    print("🖼️  HERRAMIENTA DE REPARACIÓN DE IMÁGENES HTML")
    print("=" * 60)
    print("Convierte referencias file:/// a referencias relativas")
    print("y copia las imágenes al directorio del HTML")
    print("=" * 60)
    
    if os.path.isfile(path):
        print(f"📁 Archivo: {path}\n")
        fix_html_file_with_images(path)
    elif os.path.isdir(path):
        print(f"📁 Directorio: {path}\n")
        fix_directory_with_images(path)
    else:
        print(f"❌ Error: '{path}' no existe")
        sys.exit(1)

if __name__ == "__main__":
    main()
