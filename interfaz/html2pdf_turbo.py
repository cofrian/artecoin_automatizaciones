#!/usr/bin/env python3
"""
Script de acceso rápido para conversión HTML → PDF ultra rápida.
Uso: python html2pdf_turbo.py <carpeta_html>
"""

import sys
import subprocess
from pathlib import Path

def main():
    if len(sys.argv) < 2:
        print("❌ Uso: python html2pdf_turbo.py <carpeta_html>")
        print("   Ejemplo: python html2pdf_turbo.py C:/temp/html")
        sys.exit(1)
    
    html_folder = Path(sys.argv[1])
    if not html_folder.exists():
        print(f"❌ Error: La carpeta {html_folder} no existe")
        sys.exit(1)
    
    # Detectar modo según argumentos adicionales
    modo = "ultra-fast"
    if len(sys.argv) > 2:
        if sys.argv[2] == "fast":
            modo = "fast"
        elif sys.argv[2] == "normal":
            modo = ""
    
    print(f"🚀 Iniciando conversión HTML → PDF en modo {modo.upper() if modo else 'NORMAL'}")
    print(f"📁 Carpeta: {html_folder}")
    
    # Construir comando
    script_path = Path(__file__).parent / "html2pdf_a3_fast.py"
    cmd = [
        sys.executable, str(script_path),
        str(html_folder),
        "--log-every", "5",  # Más logs para ver progreso
    ]
    
    if modo == "ultra-fast":
        cmd.extend(["--ultra-fast", "--block", "image,font,media,stylesheet"])
        print("⚡ Configuración: máxima velocidad, recursos bloqueados")
    elif modo == "fast":
        cmd.extend(["--fast", "--block", "image,media"])
        print("🔥 Configuración: alta velocidad, algunos recursos bloqueados")
    else:
        print("🐌 Configuración: velocidad normal, máxima compatibilidad")
    
    # Ejecutar
    try:
        result = subprocess.run(cmd, check=True)
        print("✅ ¡Conversión completada exitosamente!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error durante la conversión: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n⏹️  Conversión cancelada por el usuario")
        sys.exit(1)

if __name__ == "__main__":
    main()
