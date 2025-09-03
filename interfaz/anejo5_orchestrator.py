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
    parser.add_argument("--output-dir", default=None, help="Carpeta de salida para los anexos finales")
    parser.add_argument("--exclude-without-photos", action="store_true", help="Excluir elementos sin fotos del Anejo 5")
    args = parser.parse_args()
    print(f"DEBUG ORCHESTRATOR: args recibidos: {vars(args)}")
    print(f"DEBUG ORCHESTRATOR: exclude_without_photos = {args.exclude_without_photos}")

    # Obtener el directorio donde están los scripts
    script_dir = Path(__file__).parent
    extraer_datos_script = script_dir / 'extraer_datos_word.py'
    render_a3_script = script_dir / 'render_a3.py'
    html2pdf_script = script_dir / 'html2pdf_a3_fast.py'

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
            cmd1 = [sys.executable, '-u', str(extraer_datos_script),
                    '--no-interactivo',
                    '--xlsx', str(excel_file),
                    '--fotos-root', args.photos_dir,
                    '--outdir', dirs['json'],  # Todos van a la misma carpeta
                    '--jsons-separados']  # Generar JSONs separados para render_a3.py
            if args.center:
                cmd1 += ['--centro', args.center]
            
            print(f"[Anejo 5] Ejecutando: {' '.join(cmd1)}")
            print(f"[Anejo 5] Extracción de datos en tiempo real...")
            sys.stdout.flush()  # Forzar output inmediato
            
            # Configurar entorno para UTF-8
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            env['PYTHONUTF8'] = '1'
            
            try:
                # Ejecutar con output en tiempo real
                p1 = subprocess.Popen(
                    cmd1, 
                    cwd=os.getcwd(), 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.STDOUT,
                    text=True, 
                    encoding='utf-8', 
                    errors='replace', 
                    env=env,
                    bufsize=0  # Sin buffering para output inmediato
                )
                
                # Mostrar output línea por línea en tiempo real
                for line in p1.stdout:
                    clean_line = line.rstrip('\r\n')
                    if clean_line:  # Solo mostrar líneas no vacías
                        # Filtrar caracteres problemáticos
                        safe_line = clean_line.encode('ascii', errors='replace').decode('ascii')
                        print(safe_line)
                        sys.stdout.flush()  # Forzar output inmediato
                
                p1.wait()  # Esperar a que termine
                
                if p1.returncode != 0:
                    print(f"[ERROR] extraer_datos_word.py falló para {excel_file.name} con código {p1.returncode}")
                    sys.stdout.flush()
                    sys.exit(1)
                    
            except UnicodeDecodeError:
                print(f"[WARN] Problema de codificación con UTF-8, reintentando con cp1252...")
                sys.stdout.flush()
                
                p1 = subprocess.Popen(
                    cmd1, 
                    cwd=os.getcwd(), 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.STDOUT,
                    text=True, 
                    encoding='cp1252', 
                    errors='replace', 
                    env=env,
                    bufsize=0
                )
                
                for line in p1.stdout:
                    clean_line = line.rstrip('\r\n')
                    if clean_line:
                        safe_line = clean_line.encode('ascii', errors='replace').decode('ascii')
                        print(safe_line)
                        sys.stdout.flush()
                
                p1.wait()
                
                if p1.returncode != 0:
                    print(f"[ERROR] extraer_datos_word.py falló para {excel_file.name} con código {p1.returncode}")
                    sys.stdout.flush()
                    sys.exit(1)
        
        # Copiar JSONs separados al nivel base para render_a3.py
        copy_separated_jsons_to_base(dirs['json'])
    else:
        # Archivo único (comportamiento original)
        cmd1 = [sys.executable, '-u', str(extraer_datos_script),
                '--no-interactivo',
                '--xlsx', args.excel_dir,
                '--fotos-root', args.photos_dir,
                '--outdir', dirs['json'],
                '--jsons-separados']  # Generar JSONs separados para render_a3.py
        if args.center:
            cmd1 += ['--centro', args.center]
        print(f"[Anejo 5] Ejecutando: {' '.join(cmd1)}")
        print(f"[Anejo 5] Extracción de datos en tiempo real...")
        sys.stdout.flush()
        
        # Configurar entorno para UTF-8
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONUTF8'] = '1'
        
        try:
            # Ejecutar con output en tiempo real
            p1 = subprocess.Popen(
                cmd1, 
                cwd=os.getcwd(), 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT,
                text=True, 
                encoding='utf-8', 
                errors='replace', 
                env=env,
                bufsize=0
            )
            
            for line in p1.stdout:
                clean_line = line.rstrip('\r\n')
                if clean_line:
                    print(clean_line)
                    sys.stdout.flush()
            
            p1.wait()
            
        except UnicodeDecodeError:
            print(f"[WARN] Problema de codificación con UTF-8, reintentando con cp1252...")
            sys.stdout.flush()
            
            p1 = subprocess.Popen(
                cmd1, 
                cwd=os.getcwd(), 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT,
                text=True, 
                encoding='cp1252', 
                errors='replace', 
                env=env,
                bufsize=0
            )
            
            for line in p1.stdout:
                clean_line = line.rstrip('\r\n')
                if clean_line:
                    print(clean_line)
                    sys.stdout.flush()
            
            p1.wait()
        
        if p1.returncode != 0:
            print(f"[ERROR] extraer_datos_word.py falló con código {p1.returncode}")
            sys.stdout.flush()
            sys.exit(1)
        
        # Copiar JSONs separados al nivel base para render_a3.py
        copy_separated_jsons_to_base(dirs['json'])

    # 2. render_a3.py
    cmd2 = [sys.executable, '-u', str(render_a3_script),
            '--data', dirs['json'],
            '--out', dirs['html'],
            '--tpl', args.html_templates_dir]
    
    # Add photo filtering parameter if specified
    if args.exclude_without_photos:
        cmd2.append('--exclude-without-photos')
    print(f"[Anejo 5] Ejecutando: {' '.join(cmd2)}")
    print(f"[Anejo 5] Renderizado HTML en tiempo real...")
    sys.stdout.flush()
    
    # Configurar entorno para UTF-8
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'
    
    try:
        # Ejecutar con output en tiempo real
        p2 = subprocess.Popen(
            cmd2, 
            cwd=os.getcwd(), 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT,
            text=True, 
            encoding='utf-8', 
            errors='replace', 
            env=env,
            bufsize=0
        )
        
        for line in p2.stdout:
            clean_line = line.rstrip('\r\n')
            if clean_line:
                print(clean_line)
                sys.stdout.flush()
        
        p2.wait()
        
    except UnicodeDecodeError:
        print(f"[WARN] Problema de codificación con UTF-8, reintentando con cp1252...")
        sys.stdout.flush()
        
        p2 = subprocess.Popen(
            cmd2, 
            cwd=os.getcwd(), 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT,
            text=True, 
            encoding='cp1252', 
            errors='replace', 
            env=env,
            bufsize=0
        )
        
        for line in p2.stdout:
            clean_line = line.rstrip('\r\n')
            if clean_line:
                print(clean_line)
                sys.stdout.flush()
        
        p2.wait()
    
    if p2.returncode != 0:
        print(f"[ERROR] render_a3.py falló con código {p2.returncode}")
        sys.exit(2)

    # 3. html2pdf_a3_fast.py - configuración más conservadora para datasets grandes
    # Ajustar concurrencia basada en el número de archivos
    total_files = sum(len(list(Path(dirs['html']).rglob('*.html'))) for _ in [1])  # Contar archivos
    
    if total_files > 3000:
        concurrency = 4  # Muy conservador para datasets grandes
        merge_workers = 2
        print(f"[Anejo 5] Dataset grande detectado ({total_files} archivos) - usando configuración conservadora")
    elif total_files > 1000:
        concurrency = 6  # Moderado
        merge_workers = 3
        print(f"[Anejo 5] Dataset mediano detectado ({total_files} archivos) - usando configuración moderada")
    else:
        concurrency = 8  # Normal para datasets pequeños
        merge_workers = 4
        print(f"[Anejo 5] Dataset pequeño detectado ({total_files} archivos) - usando configuración normal")
    
    cmd3 = [sys.executable, '-u', str(html2pdf_script),
            '--data', dirs['html'],
            '--out', dirs['pdf'],
            '--concurrency', str(concurrency),
            '--merge-workers', str(merge_workers),
            '--wait', '900',  # Más tiempo de espera para archivos complejos
            '--caratulas-dir', args.caratulas_dir, '--port', '8800']
    print(f"[Anejo 5] Ejecutando con concurrency={concurrency}, merge-workers={merge_workers}: {' '.join(cmd3)}")
    print(f"[Anejo 5] Conversión HTML→PDF en tiempo real...")
    sys.stdout.flush()
    
    # Configurar entorno para UTF-8
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'
    
    try:
        # Ejecutar con output en tiempo real
        p3 = subprocess.Popen(
            cmd3, 
            cwd=os.getcwd(), 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT,
            text=True, 
            encoding='utf-8', 
            errors='replace', 
            env=env,
            bufsize=0
        )
        
        for line in p3.stdout:
            clean_line = line.rstrip('\r\n')
            if clean_line:
                print(clean_line)
                sys.stdout.flush()
        
        p3.wait()
        
    except UnicodeDecodeError:
        print(f"[WARN] Problema de codificación con UTF-8, reintentando con cp1252...")
        sys.stdout.flush()
        
        p3 = subprocess.Popen(
            cmd3, 
            cwd=os.getcwd(), 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT,
            text=True, 
            encoding='cp1252', 
            errors='replace', 
            env=env,
            bufsize=0
        )
        
        for line in p3.stdout:
            clean_line = line.rstrip('\r\n')
            if clean_line:
                print(clean_line)
                sys.stdout.flush()
        
        p3.wait()
    
    if p3.returncode != 0:
        print(f"[ERROR] html2pdf_a3_fast.py falló con código {p3.returncode}")
        sys.stdout.flush()
        sys.exit(3)

    print(f"[Anejo 5] Proceso completado. PDFs en: {dirs['pdf']}")
    sys.stdout.flush()

    # Si se especifica output_dir, copiar los archivos finales allí
    if args.output_dir:
        copy_files_to_output(dirs['pdf'], args.output_dir)

def copy_files_to_output(temp_pdf_dir: str, output_dir: str):
    """
    Copia SOLO el archivo final consolidado del Anejo 5 al directorio de salida especificado,
    organizándolo dentro de la carpeta de cada centro.
    Estructura final: output_dir/C####/05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf
    """
    temp_path = Path(temp_pdf_dir)
    output_path = Path(output_dir)
    
    print(f"[Anejo 5] Buscando archivo final en: {temp_path}")
    
    # Buscar SOLO el archivo final consolidado "05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf"
    final_pdf_files = list(temp_path.rglob("*05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf"))
    
    if not final_pdf_files:
        print("[WARN] No se encontró el archivo final 05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf")
        return
    
    copied_count = 0
    for pdf_file in final_pdf_files:
        # Extraer el ID del centro del directorio padre
        centro_id = None
        for part in pdf_file.parts:
            if part.startswith('C') and len(part) >= 5 and part[1:5].isdigit():
                centro_id = part
                break
        
        if not centro_id:
            # Si no se puede determinar el centro, usar un fallback
            centro_id = "C0000_SIN_CENTRO"
        
        # Crear la carpeta de destino para este centro
        centro_output_dir = output_path / centro_id
        centro_output_dir.mkdir(parents=True, exist_ok=True)
        
        # El archivo final se llamará igual pero estará en la carpeta del centro
        dest_file = centro_output_dir / pdf_file.name
        
        # Copiar el archivo
        try:
            shutil.copy2(pdf_file, dest_file)
            print(f"[Anejo 5] ✅ Archivo final copiado: {dest_file}")
            copied_count += 1
        except Exception as e:
            print(f"[ERROR] No se pudo copiar {pdf_file}: {e}")
    
    print(f"[Anejo 5] Copiados {copied_count} archivos finales del Anejo 5 a {output_path}")
    sys.stdout.flush()

if __name__ == "__main__":
    main()
