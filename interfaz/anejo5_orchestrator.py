#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
anejo5_orchestrator.py

Orquestador para la generación automática del Anejo 5.
Crea carpetas temporales, ejecuta los scripts necesarios y gestiona el flujo completo.
Se debe invocar desde la app con los parámetros necesarios.
"""
import os
import sys
import argparse
import subprocess
import shutil
from datetime import datetime
from pathlib import Path

# Configurar la consola de Windows para manejar UTF-8
if os.name == 'nt':  # Windows
    try:
        import locale
        # Intentar configurar UTF-8
        os.system('chcp 65001 > nul 2>&1')
        # Configurar variables de entorno
        os.environ['PYTHONIOENCODING'] = 'utf-8'
        os.environ['PYTHONUTF8'] = '1'
    except:
        pass

def create_temp_dirs(base_name: str = "temp_anejo5") -> dict:
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    dt = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_base = os.path.join(downloads, f"{base_name}_{dt}")
    dirs = {
        "base": temp_base,
        "json": os.path.join(temp_base, "json"),
        "html": os.path.join(temp_base, "html"),
        "pdf": os.path.join(temp_base, "pdf"),
    }
    for d in dirs.values():
        os.makedirs(d, exist_ok=True)
    return dirs

def copy_separated_jsons_to_base(json_dir: str):
    """
    Copia los JSONs separados de las carpetas de centro al nivel base
    para que render_a3.py pueda encontrarlos.
    """
    json_path = Path(json_dir)
    
    # Buscar carpetas de centro (formato: C####_NOMBRE)
    centro_dirs = [d for d in json_path.iterdir() if d.is_dir() and d.name.startswith('C')]
    
    for centro_dir in centro_dirs:
        print(f"[Anejo 5] Copiando JSONs separados de {centro_dir.name}")
        
        # Archivos JSON a copiar al nivel base
        json_files = [
            'centro.json', 'edificios.json', 'dependencias.json', 'envol.json',
            'acom.json', 'cc.json', 'clima.json', 'eqhoriz.json', 'eleva.json',
            'ilum.json', 'otroseq.json'
        ]
        
        for json_file in json_files:
            src_file = centro_dir / json_file
            if src_file.exists():
                dst_file = json_path / json_file
                shutil.copy2(src_file, dst_file)
                print(f"[Anejo 5] ✅ {json_file}")
            else:
                print(f"[Anejo 5] ⚠️  {json_file} no encontrado en {centro_dir.name}")

def main():
    parser = argparse.ArgumentParser(description="Orquestador automático para Anejo 5")
    parser.add_argument("--excel-dir", required=True)
    parser.add_argument("--photos-dir", required=True)
    parser.add_argument("--html-templates-dir", required=True)
    parser.add_argument("--caratulas-dir", required=True)
    parser.add_argument("--center", default=None)
    args = parser.parse_args()

    dirs = create_temp_dirs()
    print(f"[Anejo 5] Carpeta temporal creada: {dirs['base']}")

    # 1. extraer_datos_word.py - Manejar carpeta o archivo
    excel_path = Path(args.excel_dir)
    if excel_path.is_dir():
        # Buscar todos los .xlsx en la carpeta
        excel_files = sorted([f for f in excel_path.iterdir() if f.suffix.lower() == '.xlsx' and not f.name.startswith('~$')])
        if not excel_files:
            print(f"[ERROR] No se encontraron archivos .xlsx en {excel_path}")
            sys.exit(1)
        
        print(f"[Anejo 5] Encontrados {len(excel_files)} archivos Excel para procesar")
        for excel_file in excel_files:
            print(f"[Anejo 5] Procesando Excel: {excel_file.name}")
            cmd1 = [sys.executable, '-u', 'extraer_datos_word.py',
                    '--no-interactivo',
                    '--xlsx', str(excel_file),
                    '--fotos-root', args.photos_dir,
                    '--outdir', dirs['json'],  # Todos van a la misma carpeta
                    '--jsons-separados']  # Generar JSONs separados para render_a3.py
            if args.center:
                cmd1 += ['--centro', args.center]
            
            print(f"[Anejo 5] Ejecutando: {' '.join(cmd1)}")
            
            # Configurar entorno para UTF-8
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            env['PYTHONUTF8'] = '1'
            
            try:
                p1 = subprocess.run(cmd1, cwd=os.getcwd(), capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
                print(p1.stdout)
                if p1.stderr:
                    print(f"[STDERR] {p1.stderr}")
            except UnicodeDecodeError:
                # Fallback con codificación windows-1252
                p1 = subprocess.run(cmd1, cwd=os.getcwd(), capture_output=True, text=True, encoding='cp1252', errors='replace', env=env)
                print(p1.stdout)
                if p1.stderr:
                    print(f"[STDERR] {p1.stderr}")
            
            if p1.returncode != 0:
                print(f"[ERROR] extraer_datos_word.py falló para {excel_file.name} con código {p1.returncode}")
                sys.exit(1)
        
        # Copiar JSONs separados al nivel base para render_a3.py
        copy_separated_jsons_to_base(dirs['json'])
    else:
        # Archivo único (comportamiento original)
        cmd1 = [sys.executable, '-u', 'extraer_datos_word.py',
                '--no-interactivo',
                '--xlsx', args.excel_dir,
                '--fotos-root', args.photos_dir,
                '--outdir', dirs['json'],
                '--jsons-separados']  # Generar JSONs separados para render_a3.py
        if args.center:
            cmd1 += ['--centro', args.center]
        print(f"[Anejo 5] Ejecutando: {' '.join(cmd1)}")
        
        # Configurar entorno para UTF-8
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONUTF8'] = '1'
        
        try:
            p1 = subprocess.run(cmd1, cwd=os.getcwd(), capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
            print(p1.stdout)
            if p1.stderr:
                print(f"[STDERR] {p1.stderr}")
        except UnicodeDecodeError:
            # Fallback con codificación windows-1252
            p1 = subprocess.run(cmd1, cwd=os.getcwd(), capture_output=True, text=True, encoding='cp1252', errors='replace', env=env)
            print(p1.stdout)
            if p1.stderr:
                print(f"[STDERR] {p1.stderr}")
        
        if p1.returncode != 0:
            print(f"[ERROR] extraer_datos_word.py falló con código {p1.returncode}")
            sys.exit(1)
        
        # Copiar JSONs separados al nivel base para render_a3.py
        copy_separated_jsons_to_base(dirs['json'])

    # 2. render_a3.py
    cmd2 = [sys.executable, '-u', 'render_a3.py',
            '--data', dirs['json'],
            '--out', dirs['html'],
            '--tpl', args.html_templates_dir]
    print(f"[Anejo 5] Ejecutando: {' '.join(cmd2)}")
    
    # Configurar entorno para UTF-8
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'
    
    try:
        p2 = subprocess.run(cmd2, cwd=os.getcwd(), capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
        print(p2.stdout)
        if p2.stderr:
            print(f"[STDERR] {p2.stderr}")
    except UnicodeDecodeError:
        p2 = subprocess.run(cmd2, cwd=os.getcwd(), capture_output=True, text=True, encoding='cp1252', errors='replace', env=env)
        print(p2.stdout)
        if p2.stderr:
            print(f"[STDERR] {p2.stderr}")
    
    if p2.returncode != 0:
        print(f"[ERROR] render_a3.py falló con código {p2.returncode}")
        sys.exit(2)

    # 3. html2pdf_a3_fast.py
    cmd3 = [sys.executable, '-u', 'html2pdf_a3_fast.py',
            '--data', dirs['html'],
            '--out', dirs['pdf'],
            '--caratulas-dir', args.caratulas_dir]
    print(f"[Anejo 5] Ejecutando: {' '.join(cmd3)}")
    
    # Configurar entorno para UTF-8
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'
    
    try:
        p3 = subprocess.run(cmd3, cwd=os.getcwd(), capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
        print(p3.stdout)
        if p3.stderr:
            print(f"[STDERR] {p3.stderr}")
    except UnicodeDecodeError:
        p3 = subprocess.run(cmd3, cwd=os.getcwd(), capture_output=True, text=True, encoding='cp1252', errors='replace', env=env)
        print(p3.stdout)
        if p3.stderr:
            print(f"[STDERR] {p3.stderr}")
    
    if p3.returncode != 0:
        print(f"[ERROR] html2pdf_a3_fast.py falló con código {p3.returncode}")
        sys.exit(3)

    print(f"[Anejo 5] Proceso completado. PDFs en: {dirs['pdf']}")

if __name__ == "__main__":
    main()
