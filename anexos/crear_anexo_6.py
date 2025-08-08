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
OUTPUT_DIR = Path(__file__).resolve().parent / "salida_anexo_6"


# --------------------------------------------------------------------
# 2) HELPERS ----------------------------------------------------------
# --------------------------------------------------------------------

# --------------------------------------------------------------------
# 2.1) Helpers para mes/año y reemplazo en Word ----------------------
# --------------------------------------------------------------------


def normalizar_mes(mes_input: str) -> str:
    """Convierte el mes ingresado (número o texto) a nombre en minúsculas en español."""
    meses = {
        "1": "enero",
        "01": "enero",
        "enero": "enero",
        "2": "febrero",
        "02": "febrero",
        "febrero": "febrero",
        "3": "marzo",
        "03": "marzo",
        "marzo": "marzo",
        "4": "abril",
        "04": "abril",
        "abril": "abril",
        "5": "mayo",
        "05": "mayo",
        "mayo": "mayo",
        "6": "junio",
        "06": "junio",
        "junio": "junio",
        "7": "julio",
        "07": "julio",
        "julio": "julio",
        "8": "agosto",
        "08": "agosto",
        "agosto": "agosto",
        "9": "septiembre",
        "09": "septiembre",
        "septiembre": "septiembre",
        "10": "octubre",
        "octubre": "octubre",
        "11": "noviembre",
        "noviembre": "noviembre",
        "12": "diciembre",
        "diciembre": "diciembre",
    }
    key = mes_input.strip().lower()
    if key not in meses:
        raise ValueError(f"Mes no reconocido: {mes_input}")
    return meses[key]


def pedir_mes_anio() -> tuple[str, str]:
    """Pregunta por consola el mes y año. Devuelve (mes_en_minusculas, anio_YYYY)."""
    while True:
        mes_in = input(
            "Introduce el MES (número 1-12 o nombre en español, p.ej. 'agosto'): "
        ).strip()
        try:
            mes = normalizar_mes(mes_in)
            break
        except Exception as e:
            print(f"  Valor inválido: {e}. Intenta de nuevo.")
    while True:
        anio = input("Introduce el AÑO (formato YYYY, p.ej. 2025): ").strip()
        if re.fullmatch(r"\\d{4}", anio):
            break
        print("  Año inválido. Debe tener 4 dígitos (YYYY).")
    return mes, anio


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
    ends_ok = re.compile(r"(?i)CEE[_\s-]*ACTUAL\.pdf$")
    e_pattern = re.compile(r"(?i)[_\s-]E(\d+)[_\s-]CEE[_\s-]*ACTUAL\.pdf$")
    for pdf in edificio_dir.rglob("*.pdf"):
        name = pdf.name
        if ends_ok.search(name):
            m = e_pattern.search(name)
            order = int(m.group(1)) if m else 1
            certs.append((order, pdf))

    certs.sort(key=lambda x: (x[0], x[1].name.lower()))
    return certs


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
            if anio - 5 <= anio <= anio + 5:
                break
            else:
                print(f"Error: El año debe estar entre {anio - 5} y {anio + 5}")
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

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Se encontraron {len(edificios)} edificios.")

    for edificio in edificios:
        print(f"\nProcesando: {edificio.name}")
        certificados = find_certificates(edificio)
        if not certificados:
            print(
                "  ! No se encontraron certificados energéticos (CEE_ACTUAL) en esta carpeta."
            )
            continue

        try:
            nombre_base = clean_filename(edificio.name)
            pdf_tmp_plantilla = OUTPUT_DIR / f"_{nombre_base}_Anexo_6_tmp.pdf"
            pdf_salida = OUTPUT_DIR / f"{nombre_base}_Anexo_6.pdf"

            # 1) Generar PDF de la plantilla con mes/año
            export_template_with_fields_to_pdf(
                TEMPLATE_DOCX, pdf_tmp_plantilla, mes_render, str(anio)
            )

            # 2) Unir plantilla + certificados E1..En
            pdfs_a_unir = [pdf_tmp_plantilla] + [p for _, p in certificados]
            merge_pdfs(pdf_salida, pdfs_a_unir)

            print(f"  -> Generado: {pdf_salida}")
        except Exception as e:
            print(f"  ! Error procesando '{edificio.name}': {e}")
        finally:
            try:
                if pdf_tmp_plantilla.exists():
                    pdf_tmp_plantilla.unlink()
            except Exception:
                pass

    print(f"\nTerminado. Archivos en: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
