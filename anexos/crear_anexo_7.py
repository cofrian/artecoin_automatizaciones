from __future__ import annotations

import re
from pathlib import Path
import tempfile
import shutil
from typing import List, Tuple

# from datetime import datetime
import unicodedata
import subprocess
import time
import win32com.client
import pythoncom
from pypdf import PdfWriter
from concurrent.futures import ProcessPoolExecutor, as_completed
# from docxtpl import DocxTemplate


# Silenciar avisos verbosos de pypdf (duplicados /PageMode, etc.)
import logging

logging.getLogger("pypdf").setLevel(logging.ERROR)
"""
Crear Anexo 7
-------------
Genera, para cada edificio (una carpeta por edificio), un único PDF que contiene:
1) La plantilla Word del Anexo 7 convertida a PDF.
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
    / "Plantilla_Anexo_7.docx"
)

# Ruta raíz donde cada subcarpeta es un edificio y contiene sus PDFs de CEE
BUILDINGS_ROOT = Path(__file__).resolve().parent.parent / "PLANOS"

# Carpeta de salida (se crea si no existe)
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "word" / "anexos"


# --------------------------------------------------------------------
# 2) HELPERS ----------------------------------------------------------
# --------------------------------------------------------------------

# --------------------------------------------------------------------
# 2.1) Helpers para mes/año y reemplazo en Word ----------------------
# --------------------------------------------------------------------


# def export_template_with_fields_to_pdf(
#     template_docx: Path, out_pdf: Path, mes: str, anio: str
# ) -> None:
#     """Rellena {{mes}} y {{anio}} con DocxTemplate (estilo Anexo 3), actualiza campos y exporta a PDF."""
#     tmp_docx = out_pdf.with_suffix(".tmp.docx")
#     doc = DocxTemplate(str(template_docx))
#     doc.render({"mes": mes, "anio": anio})
#     doc.save(str(tmp_docx))
#     update_word_fields(str(tmp_docx))
#     word_to_pdf(tmp_docx, out_pdf)
#     try:
#         tmp_docx.unlink()
#     except Exception:
#         pass


# def update_word_fields(doc_path):
#     """Actualiza campos del documento Word (estilo Anexo 3)."""
#     try:
#         pythoncom.CoInitialize()
#         try:
#             word_app = win32com.client.Dispatch("Word.Application")
#             word_app.Visible = False
#             word_app.ScreenUpdating = False
#             word_app.DisplayAlerts = False
#             if not Path(doc_path).exists():
#                 print(f"   ! El archivo no existe: {doc_path}")
#                 return
#             try:
#                 doc = word_app.Documents.Open(
#                     str(doc_path),
#                     ConfirmConversions=False,
#                     ReadOnly=False,
#                     AddToRecentFiles=False,
#                     Visible=False,
#                 )
#             except Exception as e:
#                 print(f"   ! No se pudo abrir el documento: {e}")
#                 return
#             try:
#                 if doc.Range().Fields.Count > 0:
#                     doc.Range().Fields.Update()
#                 doc.Save()
#                 doc.Close(SaveChanges=False)
#             except Exception as e:
#                 print(f"   ! Error al actualizar/guardar campos: {e}")
#                 try:
#                     doc.Close(SaveChanges=False)
#                 except Exception:
#                     pass
#             word_app.Quit()
#         finally:
#             pythoncom.CoUninitialize()
#     except Exception as e:
#         print(f"   ! Error al actualizar campos: {e}")


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


def clean_building_name(building_name: str) -> str:
    """Elimina el prefijo 'Cxxxx_' del nombre del edificio y limpia caracteres no válidos."""
    # Remover prefijo Cxxxx_ (donde xxxx son números)
    cleaned = re.sub(r"^C\d+_", "", building_name)
    return clean_filename(cleaned)


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


def find_certificates(edificio_dir: Path) -> List[Tuple[int, Path]]:
    """
    Busca PDFs de PLANOS en la carpeta del edificio **y subcarpetas**.
    Formato esperado (case-insensitive):
        C{1234}_{A}_E{12}_{texto}.pdf
      - {1234}: cualquier número
      - {A}: una letra (cualquier letra)
      - {12}: cualquier número (normalmente 00, 01, 02, ...)
    Devuelve lista de tuplas (orden_E, ruta_pdf) ordenadas por E de menor a mayor.
    """
    certs: List[Tuple[int, Path]] = []

    # Regex flexible para el patrón indicado. Ejemplos válidos:
    #   C0003_A_E00_PlantaBaja.pdf
    #   c1234_b_e2_alzado.pdf
    # Notas:
    #   - Ignora mayúsculas/minúsculas
    #   - Exige guiones bajos como en el patrón
    PLANOS_RE = re.compile(r"(?i)^C\d+_[A-Z]_E(\d+)_.*\.pdf$")

    for pdf in edificio_dir.rglob("*.pdf"):
        m = PLANOS_RE.match(pdf.name)
        if m:
            orden = int(m.group(1))  # '00' -> 0, '01' -> 1, '12' -> 12
            certs.append((orden, pdf))

    # Ordenar por E (numérico) y por nombre para estabilidad
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
        merge_pdfs(output_pdf, pdf_paths)


def _qpdf_sanitize(src: Path, dst: Path) -> bool:
    """Intenta reescribir un PDF con qpdf para sanear diccionarios duplicados (como /PageMode)."""
    try:
        result = subprocess.run(
            ["qpdf", "--linearize", str(src), str(dst)],
            capture_output=True,
            text=True,
            check=True,
        )
        return dst.exists() and result.returncode == 0
    except Exception:
        return False


def merge_pdfs(output_pdf: Path, pdf_paths: List[Path]) -> None:
    """Une varios PDFs en uno solo usando pypdf >= 5 (PdfWriter) con saneado previo vía qpdf si está disponible."""
    # Preparar carpeta temporal para PDFs saneados
    tmp_dir = Path(tempfile.mkdtemp(prefix="sanpdf_"))
    sanitized_paths: List[Path] = []
    try:
        for p in pdf_paths:
            # Intentar sanear cada entrada con qpdf; si falla, usar original
            dst = tmp_dir / p.name
            if _qpdf_sanitize(p, dst):
                sanitized_paths.append(dst)
            else:
                sanitized_paths.append(p)

        writer = PdfWriter()
        for sp in sanitized_paths:
            writer.append(str(sp))
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        with output_pdf.open("wb") as f:
            writer.write(f)
    finally:
        # Limpiar temporales
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


def _worker_procesar_edificio(
    edificio_dir_str: str, plantilla_pdf_str: str, output_dir_str: str
) -> tuple[str, bool, str]:
    try:
        edificio_dir = Path(edificio_dir_str)
        plantilla_pdf = Path(plantilla_pdf_str)
        output_base_dir = Path(output_dir_str)

        certificados = find_certificates(edificio_dir)
        if not certificados:
            return (edificio_dir.name, False, "Sin planos")

        # Limpiar el nombre del edificio removiendo Cxxxx_
        nombre_limpio = clean_building_name(edificio_dir.name)

        # Crear directorio de salida: word/anexos/{nombre_edificio_limpio}/
        edificio_output_dir = output_base_dir / nombre_limpio
        edificio_output_dir.mkdir(parents=True, exist_ok=True)

        salida_pdf = edificio_output_dir / f"Anexo_7_{nombre_limpio}.pdf"
        pdfs_a_unir = [plantilla_pdf] + [p for _, p in certificados]
        merge_pdfs_fast(salida_pdf, pdfs_a_unir)
        return (edificio_dir.name, True, str(salida_pdf))
    except Exception as e:
        return (edificio_dir.name, False, f"Error: {e}")


# Comentado temporalmente: no se solicita mes/año ni se insertan variables en Word.
# def get_user_input():
#     """Solicita al usuario el mes y año (estilo Anexo 3) con valores por defecto."""
#     current_year = datetime.now().year
#     current_month = datetime.now().month
#     meses_espanol = {
#         1: "Enero",
#         2: "Febrero",
#         3: "Marzo",
#         4: "Abril",
#         5: "Mayo",
#         6: "Junio",
#         7: "Julio",
#         8: "Agosto",
#         9: "Septiembre",
#         10: "Octubre",
#         11: "Noviembre",
#         12: "Diciembre",
#     }
#     while True:
#         try:
#             print(f"\nMes actual: {meses_espanol[current_month]} ({current_month})")
#             mes_input = input(
#                 f"Ingrese el mes (1-12) [Enter para usar {current_month}]: "
#             ).strip()
#             mes_num = current_month if mes_input == "" else int(mes_input)
#             if 1 <= mes_num <= 12:
#                 mes_nombre = meses_espanol[mes_num]
#                 break
#             else:
#                 print("Error: El mes debe estar entre 1 y 12")
#         except ValueError:
#             print("Error: Por favor ingrese un número válido")
#     while True:
#         try:
#             anio_input = input(
#                 f"Ingrese el año [Enter para usar {current_year}]: "
#             ).strip()
#             anio = current_year if anio_input == "" else int(anio_input)
#             if anio - 5 <= anio <= anio + 5:
#                 break
#             else:
#                 print(f"Error: El año debe estar entre {anio - 5} y {anio + 5}")
#         except ValueError:
#             print("Error: Por favor ingrese un año válido")
#     print("\nConfiguración seleccionada:")
#     print(f"   Mes: {mes_nombre}")
#     print(f"   Año: {anio}")
#     return mes_nombre, anio


def main():
    print("=== Generador de Anexo 7 ===")
    cerrar_word_procesos()
    print(f"Buscando edificios en: {BUILDINGS_ROOT}")

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

    print(f"Se encontraron {len(edificios)} edificios.")

    # Convertir plantilla UNA vez a PDF en ubicación temporal
    import tempfile

    with tempfile.TemporaryDirectory() as temp_dir:
        plantilla_pdf_compartida = Path(temp_dir) / "_Anexo_7__plantilla.pdf"
        try:
            word_to_pdf(TEMPLATE_DOCX, plantilla_pdf_compartida)
        except Exception as e:
            print(f"Error preparando plantilla PDF: {e}")
            return

        carpetas_sin_planos = []
        resultados = []

        with ProcessPoolExecutor() as ex:
            futs = {
                ex.submit(
                    _worker_procesar_edificio,
                    str(ed),
                    str(plantilla_pdf_compartida),
                    str(OUTPUT_DIR),  # Pasar el directorio base de salida
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
                        if "Sin planos" in msg:
                            print(f"  ! {nombre}: {msg}")
                            carpetas_sin_planos.append(str(edificio))
                        else:
                            print(f"  ! {nombre}: {msg}")
                    resultados.append((nombre, ok, msg))
                except Exception as e:
                    print(f"  ! {edificio.name}: Error inesperado en worker: {e}")
                    resultados.append((edificio.name, False, f"Error worker: {e}"))

    print("Resumen:")
    ok_count = sum(1 for _, ok, _ in resultados if ok)
    print(f"  Generados OK: {ok_count}/{len(edificios)}")
    if carpetas_sin_planos:
        print("  Carpetas sin planos:")
        for c in carpetas_sin_planos:
            print(f"   - {c}")
    print(f"Archivos guardados en: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
