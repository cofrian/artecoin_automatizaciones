# -*- coding: utf-8 -*-
"""
Script para reparar archivos HTML con imágenes rotas.
Remueve o reemplaza referencias a imágenes que no existen.
"""

import os
import re
from pathlib import Path
from urllib.parse import unquote
import argparse
import shutil

def fix_broken_images_in_html(html_path: str, backup: bool = True) -> dict:
    """
    Repara un archivo HTML eliminando referencias a imágenes rotas.
    
    Returns:
        dict: Información sobre las reparaciones realizadas
    """
    
    result = {
        'file_path': html_path,
        'success': False,
        'backup_created': False,
        'images_removed': 0,
        'images_total': 0,
        'errors': []
    }
    
    html_file = Path(html_path)
    
    if not html_file.exists():
        result['errors'].append("El archivo no existe")
        return result
    
    try:
        # Crear backup si se solicita
        if backup:
            backup_path = html_file.with_suffix(f"{html_file.suffix}.backup")
            shutil.copy2(html_file, backup_path)
            result['backup_created'] = True
            print(f"✅ Backup creado: {backup_path}")
        
        # Leer contenido original
        with open(html_file, 'r', encoding='utf-8', errors='replace') as f:
            original_content = f.read()
        
        modified_content = original_content
        
        # Buscar todas las imágenes
        img_pattern = r'<img[^>]*src=["\']([^"\']+)["\'][^>]*>'
        images = re.findall(img_pattern, modified_content, re.IGNORECASE)
        result['images_total'] = len(images)
        
        print(f"🔍 Encontradas {len(images)} imágenes en el HTML")
        
        # Verificar cada imagen
        broken_img_tags = []
        for match in re.finditer(img_pattern, modified_content, re.IGNORECASE):
            img_tag = match.group(0)
            src = match.group(1)
            
            # Skip web URLs and data URIs
            if src.startswith(('http://', 'https://', 'data:')):
                continue
                
            # Check if image file exists
            img_exists = False
            if src.startswith('file:///'):
                # Extract path from file URI
                img_path = Path(unquote(src.replace('file:///', '')))
                img_exists = img_path.exists()
            elif not os.path.isabs(src):
                img_path = html_file.parent / src
                img_exists = img_path.exists()
            else:
                img_path = Path(src)
                img_exists = img_path.exists()
            
            if not img_exists:
                broken_img_tags.append((img_tag, src))
                print(f"❌ Imagen rota: {src[:60]}{'...' if len(src) > 60 else ''}")
        
        # Remover imágenes rotas
        if broken_img_tags:
            for img_tag, src in broken_img_tags:
                # Opción 1: Remover completamente la etiqueta img
                modified_content = modified_content.replace(img_tag, '')
                result['images_removed'] += 1
                
                # Opción 2 (alternativa): Reemplazar con placeholder
                # placeholder = f'<!-- Imagen no encontrada: {src[:50]}{"..." if len(src) > 50 else ""} -->'
                # modified_content = modified_content.replace(img_tag, placeholder)
            
            print(f"🔧 Removidas {result['images_removed']} imágenes rotas")
            
            # Guardar archivo reparado
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(modified_content)
                
            result['success'] = True
            print(f"✅ Archivo reparado: {html_file}")
        else:
            print("✅ No se encontraron imágenes rotas")
            result['success'] = True
            
    except Exception as e:
        result['errors'].append(f"Error reparando archivo: {e}")
        print(f"❌ Error: {e}")
    
    return result

def fix_html_directory(html_dir: str, pattern: str = "*.html", backup: bool = True) -> dict:
    """
    Repara todos los archivos HTML en un directorio.
    """
    
    html_path = Path(html_dir)
    if not html_path.exists():
        return {'error': 'El directorio no existe'}
    
    html_files = list(html_path.rglob(pattern))
    
    results = {
        'total_files': len(html_files),
        'repaired_files': 0,
        'failed_files': 0,
        'total_images_removed': 0,
        'files_with_errors': []
    }
    
    print(f"🔍 Encontrados {len(html_files)} archivos HTML para revisar")
    
    for html_file in html_files:
        print(f"\n📄 Procesando: {html_file.name}")
        
        result = fix_broken_images_in_html(str(html_file), backup=backup)
        
        if result['success']:
            results['repaired_files'] += 1
            results['total_images_removed'] += result['images_removed']
        else:
            results['failed_files'] += 1
            results['files_with_errors'].append((str(html_file), result['errors']))
    
    return results

def main():
    parser = argparse.ArgumentParser(description='Reparar archivos HTML con imágenes rotas')
    parser.add_argument('path', help='Ruta al archivo HTML o directorio')
    parser.add_argument('--no-backup', action='store_true', help='No crear backup antes de reparar')
    parser.add_argument('--pattern', default='*.html', help='Patrón para buscar archivos HTML (solo para directorios)')
    
    args = parser.parse_args()
    
    path = Path(args.path)
    
    if path.is_file():
        # Reparar archivo individual
        print(f"🔧 Reparando archivo: {path}")
        result = fix_broken_images_in_html(str(path), backup=not args.no_backup)
        
        if result['success']:
            print(f"\n✅ COMPLETADO:")
            print(f"   - Imágenes removidas: {result['images_removed']}")
            print(f"   - Total imágenes: {result['images_total']}")
            if result['backup_created']:
                print(f"   - Backup creado: {path}.backup")
        else:
            print(f"\n❌ FALLÓ:")
            for error in result['errors']:
                print(f"   - {error}")
                
    elif path.is_dir():
        # Reparar directorio
        print(f"🔧 Reparando directorio: {path}")
        results = fix_html_directory(str(path), args.pattern, backup=not args.no_backup)
        
        print(f"\n📊 RESUMEN:")
        print(f"   - Archivos procesados: {results['total_files']}")
        print(f"   - Archivos reparados: {results['repaired_files']}")
        print(f"   - Archivos con errores: {results['failed_files']}")
        print(f"   - Total imágenes removidas: {results['total_images_removed']}")
        
        if results['files_with_errors']:
            print(f"\n❌ ARCHIVOS CON ERRORES:")
            for file_path, errors in results['files_with_errors']:
                print(f"   - {file_path}")
                for error in errors:
                    print(f"     • {error}")
    else:
        print(f"❌ La ruta no existe: {path}")

if __name__ == '__main__':
    main()
