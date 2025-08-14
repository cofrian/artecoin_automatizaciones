from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple
from datetime import datetime
import unicodedata
import subprocess
import time
import win32com.client
import pythoncom
from pypdf import PdfWriter
from docxtpl import DocxTemplate
from concurrent.futures import ProcessPoolExecutor, as_completed  # MOVIDO ARRIBA

"""
Crear Anexo 6
-------------
Genera, para cada edificio (una carpeta por edificio), un único PDF que contiene:
1) La plantilla Word del Anexo 6 convertida a PDF.
2) Debajo, los PDF(s) de Certificados Energéticos del edificio, en orden E1, E2, E3...

Inspirado en el estilo y utilidades de crear_anexo_2.py (estructura, helpers, COM, prints).

Requisitos (Windows):
- pywin32  (para convertir Word a PDF):  pip install pywin32
- pypdf    (para unir PDFs):             pip install pypdf

Notas:
- No insertamos los PDF dentro del Word; convertimos la plantilla a PDF y luego
  unimos (merge) los PDFs resultantes (más robusto y rápido).
- Se ordenan certificados por el número 'E#' si existe; si no existe, se considera E1.
"""

# --------------------------------------------------------------------
# 1) CONFIGURACIÓN ----------------------------------------------------
# --------------------------------------------------------------------

# Ruta de la plantilla Word
TEMPLATE_DOCX = (
    Path(__file__).resolve().parent.parent
    / "word"
    / "anexos"
    / "Plantilla_Anexo_6.docx"
)

# Ruta raíz donde cada subcarpeta es un edificio y contiene sus PDFs de CEE
BUILDINGS_ROOT = Path(__file__).resolve().parent.parent / "CEE"

# Carpeta de salida (se crea si no existe)
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "word" / "anexos"


# --------------------------------------------------------------------
# 2) HELPERS ----------------------------------------------------------
# --------------------------------------------------------------------

# --------------------------------------------------------------------
# 2.1) Helpers para mes/año y reemplazo en Word ----------------------
# --------------------------------------------------------------------


def export_template_with_fields_to_pdf(
    template_docx: Path, out_pdf: Path, mes: str, anio: str
) -> None:
    """Rellena {{mes}} y {{anio}} con DocxTemplate (estilo Anexo 3), actualiza campos y exporta a PDF."""
    tmp_docx = out_pdf.with_suffix(".tmp.docx")
    doc = DocxTemplate(str(template_docx))
    doc.render({"mes": mes, "anio": anio})
    doc.save(str(tmp_docx))
    update_word_fields(str(tmp_docx))
    word_to_pdf(tmp_docx, out_pdf)
    try:
        tmp_docx.unlink()
    except Exception:
        pass


def update_word_fields(doc_path):
    """Actualiza campos del documento Word (estilo Anexo 3)."""
    try:
        pythoncom.CoInitialize()
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.ScreenUpdating = False
            word_app.DisplayAlerts = False
            if not Path(doc_path).exists():
                print(f"   ! El archivo no existe: {doc_path}")
                return
            try:
                doc = word_app.Documents.Open(
                    str(doc_path),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                )
            except Exception as e:
                print(f"   ! No se pudo abrir el documento: {e}")
                return
            try:
                if doc.Range().Fields.Count > 0:
                    doc.Range().Fields.Update()
                doc.Save()
                doc.Close(SaveChanges=False)
            except Exception as e:
                print(f"   ! Error al actualizar/guardar campos: {e}")
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            word_app.Quit()
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        print(f"   ! Error al actualizar campos: {e}")


def clean_filename(filename: str) -> str:
    """Limpia el nombre del archivo eliminando caracteres no válidos y tildes (Windows)."""
    invalid_chars = '<>:"|?*\\/“”'
    cleaned = filename
    for ch in invalid_chars:
        cleaned = cleaned.replace(ch, "")
    cleaned = unicodedata.normalize("NFD", cleaned)
    cleaned = "".join(c for c in cleaned if unicodedata.category(c) != "Mn")
    cleaned = re.sub(r"[_\s]+", " ", cleaned).strip()
    if len(cleaned) > 120:
        cleaned = cleaned[:120].strip()
    return cleaned


def cerrar_word_procesos() -> None:
    """Cierra Word para evitar bloqueos (similar a crear_anexo_2)."""
    try:
        try:
            pythoncom.CoInitialize()
            try:
                word_app = win32com.client.GetActiveObject("Word.Application")
                # Cerrar documentos sin guardar
                for doc in list(word_app.Documents):
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                word_app.Quit()
            except Exception:
                pass
            finally:
                pythoncom.CoUninitialize()
        except ImportError:
            pass

        subprocess.run(
            ["taskkill", "/F", "/IM", "winword.exe", "/T"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        time.sleep(1)
    except Exception:
        pass


def word_to_pdf(docx_path: Path, pdf_out_path: Path) -> None:
    """Convierte un DOCX a PDF usando Word (COM)."""
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(str(docx_path), ReadOnly=True, Visible=False)
        # 17 = wdExportFormatPDF
        doc.ExportAsFixedFormat(OutputFileName=str(pdf_out_path), ExportFormat=17)
        doc.Close(SaveChanges=False)
        word.Quit()
    finally:
        pythoncom.CoUninitialize()


ENDS_OK_RE = re.compile(r"(?i)CEE[_\s-]*ACTUAL\.pdf$")
E_NUM_RE = re.compile(r"(?i)[_\s-]E(\d+)[_\s-]CEE[_\s-]*ACTUAL\.pdf$")


def find_certificates(edificio_dir: Path) -> List[Tuple[int, Path]]:
    """
    Busca PDFs de certificados en la carpeta del edificio **y subcarpetas**.
    Formatos válidos (case-insensitive):
        {nombre}_E1_CEE_ACTUAL.pdf
        {nombre}_E2_CEE_ACTUAL.pdf
        ...
        {nombre}_CEE_ACTUAL.pdf  (si solo hay uno)
    Devuelve lista de tuplas (orden_E, ruta_pdf).
    """
    certs: List[Tuple[int, Path]] = []
    for pdf in edificio_dir.rglob("*.pdf"):
        name = pdf.name
        if ENDS_OK_RE.search(name):
            m = E_NUM_RE.search(name)
            order = int(m.group(1)) if m else 1
            certs.append((order, pdf))

    certs.sort(key=lambda x: (x[0], x[1].name.lower()))
    return certs


def merge_pdfs_fast(output_pdf: Path, pdf_paths: List[Path]) -> None:
    """
    Une PDFs prefiriendo pikepdf (rápido/robusto). Si no está disponible, usa pypdf (PdfWriter).
    """
    try:
        import pikepdf  # type: ignore

        with pikepdf.Pdf.new() as dst:
            for p in pdf_paths:
                with pikepdf.open(str(p)) as src:
                    dst.pages.extend(src.pages)
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            dst.save(str(output_pdf))
        return
    except Exception:
        # Fallback a pypdf
        merge_pdfs(output_pdf, pdf_paths)


def merge_pdfs(output_pdf: Path, pdf_paths: List[Path]) -> None:
    """Une varios PDFs en uno solo usando pypdf >= 5 (PdfWriter)."""
    writer = PdfWriter()
    for p in pdf_paths:
        writer.append(str(p))
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with output_pdf.open("wb") as f:
        writer.write(f)


def get_user_input():
    """Solicita al usuario el mes y año (estilo Anexo 3) con valores por defecto."""
    current_year = datetime.now().year
    current_month = datetime.now().month
    meses_espanol = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }
    while True:
        try:
            print(f"\nMes actual: {meses_espanol[current_month]} ({current_month})")
            mes_input = input(
                f"Ingrese el mes (1-12) [Enter para usar {current_month}]: "
            ).strip()
            mes_num = current_month if mes_input == "" else int(mes_input)
            if 1 <= mes_num <= 12:
                mes_nombre = meses_espanol[mes_num]
                break
            else:
                print("Error: El mes debe estar entre 1 y 12")
        except ValueError:
            print("Error: Por favor ingrese un número válido")
    while True:
        try:
            anio_input = input(
                f"Ingrese el año [Enter para usar {current_year}]: "
            ).strip()
            anio = current_year if anio_input == "" else int(anio_input)
            # CORREGIDO: validar contra el año actual
            if (current_year - 5) <= anio <= (current_year + 5):
                break
            else:
                print(
                    f"Error: El año debe estar entre {current_year - 5} y {current_year + 5}"
                )
        except ValueError:
            print("Error: Por favor ingrese un año válido")
    print("\nConfiguración seleccionada:")
    print(f"   Mes: {mes_nombre}")
    print(f"   Año: {anio}")
    return mes_nombre, anio


def main():
    print("=== Generador de Anexo 6 ===")
    cerrar_word_procesos()

    mes, anio = get_user_input()
    mes_render = mes
    # CORREGIDO: f-string con salto de línea al inicio
    print(f"\nBuscando edificios en: {BUILDINGS_ROOT}")

    if not BUILDINGS_ROOT.exists():
        print("No existe la carpeta de edificios indicada.")
        return

    edificios = sorted(
        [d for d in BUILDINGS_ROOT.iterdir() if d.is_dir()],
        key=lambda p: p.name.lower(),
    )
    if not edificios:
        print("No se encontraron carpetas de edificios.")
        return

    # Crear una carpeta temporal para la plantilla compartida
    temp_dir = Path(__file__).resolve().parent / "temp_anexo_6"
    temp_dir.mkdir(parents=True, exist_ok=True)
    print(f"Se encontraron {len(edificios)} edificios.")

    # Renderizar plantilla UNA vez y convertirla a PDF
    plantilla_pdf_compartida = (
        temp_dir / f"_Anexo_6__plantilla_{clean_filename(mes_render)}_{anio}.pdf"
    )
    try:
        export_template_with_fields_to_pdf(
            TEMPLATE_DOCX, plantilla_pdf_compartida, mes_render, str(anio)
        )
    except Exception as e:
        print(f"Error preparando plantilla PDF: {e}")
        return

    carpetas_sin_certs = []
    resultados = []

    # Paralelizar por edificio
    with ProcessPoolExecutor() as ex:
        futs = {
            ex.submit(
                _worker_procesar_edificio,
                str(ed),
                str(plantilla_pdf_compartida),
            ): ed
            for ed in edificios
        }
        for fut in as_completed(futs):
            edificio = futs[fut]
            try:
                nombre, ok, msg = fut.result()
                if ok:
                    print(f"  -> Generado: {msg}")
                else:
                    if "Sin certificados" in msg:
                        print(f"  ! {nombre}: {msg}")
                        carpetas_sin_certs.append(str(edificio))
                    else:
                        print(f"  ! {nombre}: {msg}")
                resultados.append((nombre, ok, msg))
            except Exception as e:
                print(f"  ! {edificio.name}: Error inesperado en worker: {e}")
                resultados.append((edificio.name, False, f"Error worker: {e}"))

    # Limpiar plantilla temporal
    try:
        plantilla_pdf_compartida.unlink()
        temp_dir.rmdir()
    except Exception:
        pass

    # Resumen
    # CORREGIDO: f-string con salto de línea al inicio
    print("\nResumen:")
    ok_count = sum(1 for _, ok, _ in resultados if ok)
    print(f"  Generados OK: {ok_count}/{len(edificios)}")
    if carpetas_sin_certs:
        print("  Carpetas sin certificados:")
        for c in carpetas_sin_certs:
            print(f"   - {c}")
    # CORREGIDO: f-string con salto de línea al inicio
    print("\nArchivos generados en las carpetas respectivas de cada edificio")


def _worker_procesar_edificio(
    edificio_dir_str: str, plantilla_pdf_str: str
) -> tuple[str, bool, str]:
    """
    Worker de proceso: une plantilla + certificados del edificio.
    Devuelve (nombre_edificio, ok, mensaje).
    """
    try:
        edificio_dir = Path(edificio_dir_str)
        plantilla_pdf = Path(plantilla_pdf_str)
        certificados = find_certificates(edificio_dir)
        if not certificados:
            return (edificio_dir.name, False, "Sin certificados")

        # Extraer ID CENTRO del nombre de la carpeta (formato: Cxxxx_NOMBRE)
        nombre_carpeta = edificio_dir.name
        id_centro_match = re.match(r"^(C\d+)_", nombre_carpeta)
        if id_centro_match:
            id_centro = id_centro_match.group(1)  # Ej: "C0001"
        else:
            # Fallback si no se encuentra el patrón
            id_centro = clean_filename(nombre_carpeta)

        # Limpiar nombre del edificio removiendo patrón "Cxxxx_" del inicio
        nombre_limpio = re.sub(r"^C\d+_", "", nombre_carpeta)
        nombre_base = clean_filename(nombre_limpio)

        # Guardar el PDF en la carpeta word/anexos/{id_centro}/
        output_center_dir = OUTPUT_DIR / id_centro
        output_center_dir.mkdir(parents=True, exist_ok=True)
        salida_pdf = output_center_dir / f"Anexo_6_{nombre_base}.pdf"
        pdfs_a_unir = [plantilla_pdf] + [p for _, p in certificados]
        merge_pdfs_fast(salida_pdf, pdfs_a_unir)
        return (edificio_dir.name, True, f"{id_centro}/{salida_pdf.name}")
    except Exception as e:
        return (edificio_dir.name, False, f"Error: {e}")


if __name__ == "__main__":
    main()
