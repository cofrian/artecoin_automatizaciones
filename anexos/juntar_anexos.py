from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple
import unicodedata

try:
    from pypdf import PdfWriter
except ImportError:
    try:
        from PyPDF2 import PdfWriter  # type: ignore
    except ImportError:
        print("Error: Se necesita pypdf o PyPDF2 para unir PDFs.")
        print("Instala con: pip install pypdf")
        exit(1)

"""
Juntar Anexos
-------------
Recorre cada carpeta en word/anexos/ y junta todos los archivos PDF de anexos
en orden ascendente (Anexo 1, Anexo 2, ..., Anexo 7) en un solo archivo PDF
llamado Anexo_{id_centro}.pdf guardado en la misma carpeta.

Formato esperado de archivos:
- Anexo_1_*.pdf
- Anexo_2_*.pdf
- ...
- Anexo_7_*.pdf

El script detecta automáticamente el número de anexo del nombre del archivo.
"""

# --------------------------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------------------------

# Directorio base donde están las carpetas de cada centro
ANEXOS_DIR = Path(__file__).resolve().parent.parent / "word" / "anexos"

# Regex para extraer número de anexo del nombre del archivo
ANEXO_PATTERN = re.compile(r"^Anexo[_\s]+(\d+)[_\s].*\.pdf$", re.IGNORECASE)


# --------------------------------------------------------------------
# FUNCIONES
# --------------------------------------------------------------------


def clean_filename(filename: str) -> str:
    """Limpia el nombre del archivo eliminando caracteres no válidos."""
    invalid_chars = '<>:"|?*\\/""'
    cleaned = filename
    for char in invalid_chars:
        cleaned = cleaned.replace(char, "")

    # Normalizar caracteres con tilde
    cleaned = unicodedata.normalize("NFD", cleaned)
    cleaned = "".join(c for c in cleaned if unicodedata.category(c) != "Mn")

    # Limpiar espacios múltiples
    cleaned = re.sub(r"[_\s]+", "_", cleaned).strip()

    return cleaned


def find_anexo_pdfs(centro_dir: Path) -> List[Tuple[int, Path]]:
    """
    Busca archivos PDF de anexos en la carpeta del centro.
    Devuelve lista de tuplas (numero_anexo, ruta_pdf) ordenadas por número.
    """
    anexos: List[Tuple[int, Path]] = []

    for pdf_file in centro_dir.glob("*.pdf"):
        # Evitar procesar el archivo final ya generado
        if pdf_file.name.startswith("Anexo_") and not re.match(
            r"^Anexo_\d+_", pdf_file.name
        ):
            continue

        match = ANEXO_PATTERN.match(pdf_file.name)
        if match:
            numero_anexo = int(match.group(1))
            # Solo procesar anexos del 1 al 7
            if 1 <= numero_anexo <= 7:
                anexos.append((numero_anexo, pdf_file))

    # Ordenar por número de anexo
    anexos.sort(key=lambda x: x[0])
    return anexos


def merge_pdfs(output_pdf: Path, pdf_paths: List[Path]) -> None:
    """Une varios PDFs en uno solo usando pypdf."""
    if not pdf_paths:
        return

    writer = PdfWriter()

    for pdf_path in pdf_paths:
        try:
            writer.append(str(pdf_path))
        except Exception as e:
            print(f"   ! Error leyendo {pdf_path.name}: {e}")
            continue

    # Crear directorio de salida si no existe
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    try:
        with output_pdf.open("wb") as f:
            writer.write(f)
    except Exception as e:
        print(f"   ! Error escribiendo {output_pdf}: {e}")


def process_centro(centro_dir: Path) -> None:
    """Procesa una carpeta de centro, juntando todos sus anexos."""
    centro_id = centro_dir.name
    print(f"-> Procesando centro: {centro_id}")

    # Buscar archivos de anexos
    anexos = find_anexo_pdfs(centro_dir)

    if not anexos:
        print(f"   ! No se encontraron anexos en {centro_id}")
        return

    # Mostrar anexos encontrados
    anexos_numeros = [num for num, _ in anexos]
    print(f"   Anexos encontrados: {anexos_numeros}")

    # Crear archivo de salida
    output_file = centro_dir / f"Anexo_{centro_id}.pdf"

    # Obtener rutas de PDFs en orden
    pdf_paths = [path for _, path in anexos]

    # Unir PDFs
    try:
        merge_pdfs(output_file, pdf_paths)
        print(f"   ✓ Generado: {output_file.name} ({len(anexos)} anexos)")
    except Exception as e:
        print(f"   ! Error generando {output_file.name}: {e}")


def main():
    """Función principal que procesa todos los centros."""
    print("=== Juntador de Anexos ===")
    print(f"Buscando centros en: {ANEXOS_DIR}")

    if not ANEXOS_DIR.exists():
        print(f"Error: No existe el directorio {ANEXOS_DIR}")
        return

    # Buscar carpetas de centros (directorios que parecen ser IDs de centro)
    centros = []
    for item in ANEXOS_DIR.iterdir():
        if item.is_dir():
            # Filtrar carpetas que parecen ser IDs de centro (empiezan con C seguido de números)
            if re.match(r"^C\d+$", item.name) or not re.match(r"^[A-Z]", item.name):
                centros.append(item)

    if not centros:
        print("No se encontraron carpetas de centros.")
        return

    # Ordenar centros por nombre
    centros.sort(key=lambda x: x.name)

    print(f"Se encontraron {len(centros)} centros.")

    # Procesar cada centro
    procesados = 0
    for centro_dir in centros:
        try:
            process_centro(centro_dir)
            procesados += 1
        except Exception as e:
            print(f"   ! Error procesando {centro_dir.name}: {e}")

    # Resumen final
    print(f"\n{'=' * 60}")
    print("PROCESO COMPLETADO")
    print(f"{'=' * 60}")
    print(f"Centros procesados: {procesados}/{len(centros)}")
    print("Los archivos finales se encuentran en:")
    print(f"  {ANEXOS_DIR}")
    print("Formato: Anexo_[ID_CENTRO].pdf")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
