# -*- coding: utf-8 -*-
"""
anexos_creator (refactor SOLID + Clean Architecture)

Este script es una refactorización del original, orientada a SOLID y arquitectura limpia,
y compatible con la GUI de `app.py`. Mantiene la funcionalidad actual (Anexo 2 y 3) y
prepara una base para escalar a Anexos 1..7.

Ejecución desde línea de comandos (igual que usa la GUI):
    python anexos_creator.py --excel-dir "C:\\ruta\\excel" --word-dir "C:\\ruta\\plantillas_word" --mode all
    python anexos_creator.py --excel-dir "C:\\ruta\\excel" --word-dir "C:\\ruta\\plantillas_word" --mode single --anexo 3
    # Flags opcionales:
    # --html-templates-dir, --photos-dir, --output-dir

Requisitos:
    - Windows (usa Word COM via pywin32)
    - Python 3.9+ recomendado
    - pip install: pandas, docxtpl, pypdf, pywin32
"""

from __future__ import annotations

import argparse
import logging
import queue
import re
import subprocess
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Protocol, Sequence, Tuple

import pandas as pd
import pythoncom
import win32com.client as win32_client  # type: ignore
from docxtpl import DocxTemplate  # type: ignore
from pypdf import PdfReader, PdfWriter  # type: ignore
from win32com.client import CDispatch  # type: ignore


# =====================================================================================
# Configuración / utilidades
# =====================================================================================

APP_LOGGER_NAME = "anexos_creator"
logger = logging.getLogger(APP_LOGGER_NAME)


def setup_logging(level: int = logging.INFO) -> None:
    """Configura logging sencillo por consola (stdout)."""
    handler = logging.StreamHandler(sys.stdout)
    fmt = logging.Formatter("%(message)s")
    handler.setFormatter(fmt)
    logger.setLevel(level)
    logger.handlers.clear()
    logger.addHandler(handler)


def ensure_utf8_console() -> None:
    """Asegura UTF-8 en consola (Windows)."""
    try:
        import io

        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )  # type: ignore[attr-defined]
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", errors="replace"
        )  # type: ignore[attr-defined]
    except Exception:
        pass


# =====================================================================================
# Domain models (dataclasses)
# =====================================================================================


@dataclass(frozen=True)
class RunConfig:
    excel_dir: Path
    word_dir: Optional[Path]
    mode: str  # "all" | "single"
    anexo: Optional[int]  # requerido si mode == "single"
    month: Optional[str] = None  # nombre del mes o número (1-12)
    year: Optional[int] = None  # año
    html_templates_dir: Optional[Path] = None
    photos_dir: Optional[Path] = None
    output_dir: Optional[Path] = None


@dataclass(frozen=True)
class OutputFile:
    docx_path: Path
    pdf_path: Optional[Path] = None


# =====================================================================================
# Ports (interfaces) - SOLID: DIP
# =====================================================================================


class TemplateProvider(Protocol):
    """Obtiene la plantilla DOCX como bytes para el anexo indicado."""

    def get_template(self, anexo_number: int) -> bytes: ...


class WordExporter(Protocol):
    """Operaciones con Word (COM)."""

    def close_word_processes(self) -> None: ...
    def remove_blank_pages_from_docx(self, docx_path: Path) -> int: ...
    def export_doc_to_pdf(self, doc: CDispatch, pdf_path: Path) -> None: ...
    def open_document(
        self, docx_path: Path, read_only: bool = False
    ) -> Tuple[CDispatch, CDispatch]: ...
    def update_toc(self, doc: CDispatch) -> None: ...
    def delete_pages(self, doc: CDispatch, pages_to_delete: Iterable[int]) -> None: ...
    def update_word_fields_bulk(self, doc_paths: List[str]) -> None: ...


class PdfInspector(Protocol):
    """Lectura/edición básica de PDF."""

    def remove_blank_pages_from_pdf(self, pdf_path: Path) -> None: ...
    def find_title_page_in_pdf(
        self, pdf_path: Path, title_text: str
    ) -> Optional[int]: ...
    def export_docx_to_temp_pdf(
        self, word: WordExporter, docx_path: Path, tmp_pdf: Path
    ) -> None: ...
    def read_total_pages(self, pdf_path: Path) -> int: ...
    def convert_docx_to_pdf_bulk(self, doc_paths: List[str]) -> List[str]: ...
    def remove_last_page_from_pdfs(self, pdf_paths: List[str]) -> None: ...


class ExcelRepository(Protocol):
    """Carga de hojas Excel limpias requeridas por anexos actuales."""

    def load_sheets_for_anexo2(self, excel_path: Path) -> Dict[str, pd.DataFrame]: ...
    def load_sheets_for_anexo3(self, excel_path: Path) -> Dict[str, pd.DataFrame]: ...
    def load_sheets_for_anexo4(self, excel_path: Path) -> Dict[str, pd.DataFrame]: ...
    def extract_unique_groups(
        self, group_column: str, tables: Dict[str, pd.DataFrame]
    ) -> Sequence[str]: ...
    def filter_tables_by_group(
        self, group_column: str, tables: Dict[str, pd.DataFrame], group: str
    ) -> List[pd.DataFrame]: ...
    def extract_center_id(self, df_group_list: List[pd.DataFrame]) -> str: ...
    def calculate_totals_by_center(
        self, complete_df: pd.DataFrame, df_group: pd.DataFrame
    ) -> Dict[str, str]: ...


class OutputPathBuilder(Protocol):
    """Genera rutas de salida consistentes y seguras."""

    def build_output_docx_path(
        self, config: RunConfig, center_id: str, filename: str
    ) -> Path: ...


# =====================================================================================
# Implementaciones (Adapters)
# =====================================================================================


class DefaultTemplateProvider:
    """
    Busca `Plantilla_Anexo_{n}.docx` en:
      1) --word-dir si se pasa
      2) ./word/anexos
      3) ../word/anexos
      4) carpeta del script
    """

    def __init__(self, word_dir: Optional[Path]) -> None:
        self.word_dir = word_dir

    def _candidate_paths(self, anexo_number: int) -> List[Path]:
        fname = f"Plantilla_Anexo_{anexo_number}.docx"
        here = Path(__file__).resolve().parent
        candidates = []
        if self.word_dir:
            candidates.append(self.word_dir / fname)
        candidates += [
            here / "word" / "anexos" / fname,
            here.parent / "word" / "anexos" / fname,
            here / fname,
        ]
        # Dedup manteniendo orden
        out: List[Path] = []
        seen = set()
        for p in candidates:
            if p not in seen:
                out.append(p)
                seen.add(p)
        return out

    def get_template(self, anexo_number: int) -> bytes:
        for path in self._candidate_paths(anexo_number):
            if path.is_file():
                logger.info(f"-> Usando plantilla: {path}")
                return path.read_bytes()
        raise FileNotFoundError(
            f"No se encontró la plantilla DOCX para Anexo {anexo_number}. "
            f"Coloca 'Plantilla_Anexo_{anexo_number}.docx' en --word-dir o en ./word/anexos"
        )


# ---------------- Word Services ----------------


class _ContentChecker:
    """Verifica si un rango de página tiene contenido relevante."""

    def __init__(self, doc: CDispatch):
        self.doc = doc

    def has_content(self, page_range: CDispatch) -> bool:
        try:
            return (
                self._has_text_content(page_range)
                or self._has_table_content(page_range)
                or self._has_inline_shapes_content(page_range)
                or self._has_shapes_content(page_range)
            )
        except Exception as e:
            logger.debug(f"Error checking page content: {e}")
            return True

    @staticmethod
    def _extract_meaningful_text(txt: str) -> str:
        if not txt:
            return ""
        return txt.replace("\r", "").replace("\n", "").replace("\t", "").strip()

    def _has_text_content(self, page_range: CDispatch) -> bool:
        try:
            text_content = page_range.Text.strip()
            return len(self._extract_meaningful_text(text_content)) > 0
        except Exception:
            return True

    def _has_table_content(self, page_range: CDispatch) -> bool:
        try:
            if page_range.Tables.Count == 0:
                return False
            table = page_range.Tables(1)
            table_text = table.Range.Text.strip()
            return len(self._extract_meaningful_text(table_text)) > 0
        except Exception:
            return False

    @staticmethod
    def _is_shape_in_range(shape: CDispatch, page_start: int, page_end: int) -> bool:
        try:
            return (
                hasattr(shape, "Anchor")
                and shape.Anchor.Start >= page_start
                and shape.Anchor.End <= page_end
            )
        except Exception:
            return False

    @staticmethod
    def _shape_has_text(shape: CDispatch) -> bool:
        try:
            has_text_frame = (
                hasattr(shape, "TextFrame")
                and hasattr(shape.TextFrame, "HasText")
                and shape.TextFrame.HasText
            )
            if not has_text_frame:
                return False
            shape_text = shape.TextFrame.TextRange.Text.strip()
            return bool(shape_text)
        except Exception:
            return False

    def _has_shapes_content(self, page_range: CDispatch) -> bool:
        try:
            page_start = page_range.Start
            page_end = page_range.End
            for shape in self.doc.Shapes:
                if self._is_shape_in_range(
                    shape, page_start, page_end
                ) and self._shape_has_text(shape):
                    return True
            return False
        except Exception:
            return False

    @staticmethod
    def _has_inline_shapes_content(page_range: CDispatch) -> bool:
        try:
            return page_range.InlineShapes.Count > 0
        except Exception:
            return False


class _WordPageManager:
    """Operaciones de página en Word."""

    WD_GO_TO_PAGE = 1
    WD_GO_TO_ABSOLUTE = 1
    WD_STORY = 6
    WD_STATISTIC_PAGES = 2

    def __init__(self, doc: CDispatch):
        self.doc = doc
        self.checker = _ContentChecker(doc)

    def _get_page_start(self, app: CDispatch, page_num: int) -> int:
        app.Selection.GoTo(
            What=self.WD_GO_TO_PAGE, Which=self.WD_GO_TO_ABSOLUTE, Count=page_num
        )
        return app.Selection.Range.Start

    def _get_page_end(self, app: CDispatch, page_num: int) -> int:
        total_pages = int(self.doc.ComputeStatistics(self.WD_STATISTIC_PAGES))
        if page_num < total_pages:
            app.Selection.GoTo(
                What=self.WD_GO_TO_PAGE,
                Which=self.WD_GO_TO_ABSOLUTE,
                Count=page_num + 1,
            )
            return app.Selection.Range.Start
        app.Selection.EndKey(Unit=self.WD_STORY)
        return app.Selection.Range.End

    def get_page_range(self, page_num: int) -> Optional[CDispatch]:
        try:
            app = self.doc.Application
            start = self._get_page_start(app, page_num)
            end = self._get_page_end(app, page_num)
            return self.doc.Range(Start=start, End=end) if end > start else None
        except Exception as e:
            logger.error(f"Error getting page range: {e}")
            return None

    def delete_page(self, page_num: int) -> bool:
        try:
            app = self.doc.Application
            total_pages = int(self.doc.ComputeStatistics(self.WD_STATISTIC_PAGES))
            if not (1 <= page_num <= total_pages):
                return False
            start = self._get_page_start(app, page_num)
            end = self._get_page_end(app, page_num)
            if end <= start:
                return False
            self.doc.Range(Start=start, End=end).Delete()
            return True
        except Exception as e:
            logger.error(f"Failed to delete page {page_num}: {e}")
            return False

    def is_page_blank(self, page_num: int) -> bool:
        try:
            page = self.get_page_range(page_num)
            if page is None:
                return False
            return not self.checker.has_content(page)
        except Exception as e:
            logger.error(f"Error blank check p{page_num}: {e}")
            return False


class DefaultWordExporter:
    """Implementación de WordExporter."""

    WD_EXPORT_FORMAT_PDF = 17
    WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
    WD_EXPORT_ALL_DOCUMENT = 0
    WD_EXPORT_DOCUMENT_CONTENT = 0
    WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1

    # Word field constants
    WD_FIELD_TOC = 13
    WD_FIELD_INDEX = 8
    WD_FIELD_TOA = 73

    def close_word_processes(self) -> None:
        """Cierra Word de forma segura. Si falla, fuerza cierre."""
        logger.info("Cerrando procesos de Word…")
        if not self._close_elegantly():
            self._force_close()
        time.sleep(0.3)
        logger.info("Word listo.")

    @staticmethod
    def _force_close() -> None:
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", "winword.exe", "/T"],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            if result.returncode not in (0, 128):
                logger.warning(f"taskkill => {result.returncode}: {result.stderr}")
        except Exception as e:
            logger.error(f"Error al forzar cierre Word: {e}")

    @staticmethod
    def _close_elegantly() -> bool:
        try:
            pythoncom.CoInitialize()
            try:
                word_app = win32_client.GetActiveObject("Word.Application")
                for i in range(word_app.Documents.Count, 0, -1):
                    try:
                        word_app.Documents(i).Close(SaveChanges=False)
                    except Exception:
                        pass
                word_app.Quit(SaveChanges=False)
                return True
            except Exception:
                return True  # no hay instancia activa
        except Exception:
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- core Word ops ---

    def open_document(
        self, docx_path: Path, read_only: bool = False
    ) -> Tuple[CDispatch, CDispatch]:
        pythoncom.CoInitialize()
        app: CDispatch = win32_client.Dispatch("Word.Application")
        app.Visible = False
        app.ScreenUpdating = False
        try:
            app.DisplayAlerts = (
                False  # algunos Word no lo permiten; si falla no es crítico
            )
        except Exception:
            pass
        doc: CDispatch = app.Documents.Open(
            str(docx_path),
            ConfirmConversions=False,
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Visible=False,
        )
        return app, doc

    def export_doc_to_pdf(self, doc: CDispatch, pdf_path: Path) -> None:
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_path),
            ExportFormat=self.WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=self.WD_EXPORT_OPTIMIZE_FOR_PRINT,
            Range=self.WD_EXPORT_ALL_DOCUMENT,
            Item=self.WD_EXPORT_DOCUMENT_CONTENT,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=self.WD_EXPORT_CREATE_HEADING_BOOKMARKS,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )

    @staticmethod
    def update_toc(doc: CDispatch) -> None:
        try:
            if doc.TablesOfContents.Count > 0:
                doc.TablesOfContents(1).Update()
        except Exception as e:
            logger.debug(f"Update TOC error: {e}")

    def delete_pages(self, doc: CDispatch, pages_to_delete: Iterable[int]) -> None:
        try:
            doc.Repaginate()
        except Exception:
            pass
        remover = _WordPageManager(doc)
        try:
            total = int(doc.ComputeStatistics(_WordPageManager.WD_STATISTIC_PAGES))
        except Exception:
            total = 10**6
        for p in sorted(set(pages_to_delete), reverse=True):
            if p <= total:
                remover.delete_page(p)

    def remove_blank_pages_from_docx(self, docx_path: Path) -> int:
        pages_removed = 0
        try:
            app, doc = self.open_document(docx_path, read_only=False)
            try:
                doc.Repaginate()
                total = int(doc.ComputeStatistics(_WordPageManager.WD_STATISTIC_PAGES))
                mgr = _WordPageManager(doc)
                for page in range(total, 0, -1):
                    if mgr.is_page_blank(page):
                        if mgr.delete_page(page):
                            pages_removed += 1
                doc.Repaginate()
                self.update_toc(doc)
                doc.Save()
            finally:
                doc.Close(SaveChanges=False)
                app.Quit()
        except Exception as e:
            logger.error(f"remove_blank_pages_from_docx error: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return pages_removed

    def update_word_fields_bulk(self, doc_paths: List[str]) -> None:
        """
        Actualiza campos en **lote** usando una sola instancia de Word.
        """
        if not doc_paths:
            return

        try:
            pythoncom.CoInitialize()
            word_app = win32_client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.ScreenUpdating = False
            try:
                word_app.DisplayAlerts = False
            except Exception:
                pass

            for doc_path in doc_paths:
                doc_path_str = str(doc_path)
                if not Path(doc_path_str).exists():
                    logger.warning(f"   ! No existe: {doc_path_str}")
                    continue

                try:
                    doc = word_app.Documents.Open(
                        doc_path_str,
                        ConfirmConversions=False,
                        ReadOnly=False,
                        AddToRecentFiles=False,
                        Visible=False,
                    )

                    # TOC: solo paginación (rápido) o actualización completa si se pide
                    try:
                        toc = doc.TablesOfContents(1)
                        toc.UpdatePageNumbers()
                    except Exception as e:
                        logger.debug(f"   ! Error al actualizar TOC en {Path(doc_path_str).name}: {e}")

                    # Guardar y cerrar
                    try:
                        doc.Save()
                        doc.Close(SaveChanges=False)
                        logger.debug(f"   ✓ Campos actualizados: {Path(doc_path_str).name}")
                    except Exception as e:
                        logger.warning(f"   ! Error al guardar/cerrar {Path(doc_path_str).name}: {e}")

                except Exception as e:
                    logger.warning(f"   ! No se pudo abrir: {Path(doc_path_str).name} -> {e}")
                    continue

            # Cerrar Word
            word_app.Quit()
            logger.info(f"   ✓ Campos actualizados en {len([p for p in doc_paths if Path(p).exists()])} documentos")

        except Exception as e:
            logger.error(f"   ! Error general en actualización de campos: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------------- PDF Services ----------------


class DefaultPdfInspector:
    TOC_KEYWORDS = ("indice", "índice", "INDICE", "ÍNDICE")

    def _normalize(self, s: Optional[str]) -> str:
        if not s:
            return ""
        s = unicodedata.normalize("NFD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").lower()
        return re.sub(r"\s+", " ", s).strip()

    def read_total_pages(self, pdf_path: Path) -> int:
        reader = PdfReader(str(pdf_path))
        return len(reader.pages)

    def export_docx_to_temp_pdf(
        self, word: WordExporter, docx_path: Path, tmp_pdf: Path
    ) -> None:
        app, doc = word.open_document(docx_path, read_only=False)
        try:
            try:
                doc.Repaginate()
            except Exception:
                pass
            word.export_doc_to_pdf(doc, tmp_pdf)
        finally:
            doc.Close(SaveChanges=False)
            app.Quit()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def find_title_page_in_pdf(self, pdf_path: Path, title_text: str) -> Optional[int]:
        reader = PdfReader(str(pdf_path))
        normalized_title = self._normalize(title_text)
        if not normalized_title:
            return None

        candidates: List[int] = []
        for idx in range(len(reader.pages)):
            text = reader.pages[idx].extract_text() or ""
            norm_page = self._normalize(text)
            if normalized_title in norm_page:
                candidates.append(idx + 1)  # 1-based

        if not candidates:
            return None

        def _is_toc(page_num: int) -> bool:
            text = reader.pages[page_num - 1].extract_text() or ""
            norm = self._normalize(text)
            return any(k.lower() in norm for k in self.TOC_KEYWORDS)

        non_toc = [p for p in candidates if not _is_toc(p)]
        return max(non_toc or candidates)

    def remove_blank_pages_from_pdf(self, pdf_path: Path) -> None:
        try:
            reader = PdfReader(str(pdf_path))
            n = len(reader.pages)
            if n <= 1:
                return
            keep: List[int] = []
            for i in range(n):
                text = (reader.pages[i].extract_text() or "").strip()
                if len(text) > 20 or len(text.split()) > 3:
                    keep.append(i)
            if len(keep) == n:
                return
            writer = PdfWriter()
            for i in keep:
                writer.add_page(reader.pages[i])
            tmp = pdf_path.with_suffix(".tmp")
            with open(tmp, "wb") as f:
                writer.write(f)
            tmp.replace(pdf_path)
        except Exception as e:
            logger.warning(f"remove_blank_pages_from_pdf error: {e}")

    def convert_docx_to_pdf_bulk(self, doc_paths: List[str]) -> List[str]:
        """
        Convierte documentos DOCX a PDF en lote usando una sola instancia de Word.
        """
        pdf_paths: List[str] = []
        
        try:
            pythoncom.CoInitialize()
            word_app = win32_client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.ScreenUpdating = False
            try:
                word_app.DisplayAlerts = False
            except Exception:
                pass

            for doc_path in doc_paths:
                doc_path_str = str(doc_path)
                if not Path(doc_path_str).exists():
                    logger.warning(f"   ! No existe: {doc_path_str}")
                    continue

                pdf_path = str(Path(doc_path_str).with_suffix(".pdf"))

                try:
                    doc = word_app.Documents.Open(
                        doc_path_str,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False,
                    )

                    # Exportar a PDF usando valores numéricos directos
                    doc.ExportAsFixedFormat(
                        OutputFileName=pdf_path,
                        ExportFormat=17,  # wdExportFormatPDF = 17
                        OpenAfterExport=False,
                        OptimizeFor=0,  # wdExportOptimizeForPrint = 0
                        BitmapMissingFonts=True,
                        DocStructureTags=True,
                        CreateBookmarks=1,  # wdExportCreateHeadingBookmarks = 1
                        UseISO19005_1=False,
                        Range=0,  # wdExportAllDocument = 0
                        Item=0,  # wdExportDocumentContent = 0
                        IncludeDocProps=True,
                        KeepIRM=True,
                    )

                    doc.Close(SaveChanges=False)
                    pdf_paths.append(pdf_path)
                    logger.info(f"   ✓ PDF generado: {Path(pdf_path).name}")

                except Exception as e:
                    logger.warning(f"   ! Error al convertir {Path(doc_path_str).name} a PDF: {e}")
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                    continue

            # Cerrar Word
            word_app.Quit()

        except Exception as e:
            logger.error(f"   ! Error general en conversión a PDF: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        return pdf_paths

    def remove_last_page_from_pdfs(self, pdf_paths: List[str]) -> None:
        """
        Elimina la última página de los archivos PDF proporcionados.
        """
        if not pdf_paths:
            return

        modified_count = 0

        for pdf_path in pdf_paths:
            try:
                # Leer el PDF
                reader = PdfReader(pdf_path)
                total_pages = len(reader.pages)

                if total_pages <= 1:
                    logger.info(
                        f"   ! {Path(pdf_path).name} tiene solo {total_pages} página(s), no se modifica"
                    )
                    continue

                # Crear nuevo PDF sin la última página
                writer = PdfWriter()
                for page_num in range(total_pages - 1):  # Excluir la última página
                    writer.add_page(reader.pages[page_num])

                # Escribir a archivo temporal
                temp_path = pdf_path + ".temp"
                with open(temp_path, "wb") as temp_file:
                    writer.write(temp_file)

                # Reemplazar el archivo original
                Path(temp_path).replace(pdf_path)

                logger.info(
                    f"   ✓ Última página eliminada de: {Path(pdf_path).name} ({total_pages} → {total_pages - 1} páginas)"
                )
                modified_count += 1

            except Exception as e:
                logger.warning(f"   ! Error al modificar {Path(pdf_path).name}: {e}")
                # Limpiar archivo temporal si existe
                temp_path = pdf_path + ".temp"
                if Path(temp_path).exists():
                    try:
                        Path(temp_path).unlink()
                    except Exception:
                        pass
                continue

        if modified_count > 0:
            logger.info(f"   ✓ Se modificaron {modified_count} archivos PDF correctamente")
        else:
            logger.warning("   ! No se pudo modificar ningún archivo PDF")


# ---------------- Excel Repository ----------------


class DefaultExcelRepository:
    """
    Implementa la lectura y limpieza de las hojas usadas por los diferentes anexos.
    """

    SHEETS_MAP: Dict[str, str] = {
        "Clima": "Sistemas de Climatización",
        "SistCC": "Sistemas de Calefacción",
        "Eleva": "Equipos Elevadores",
        "EqHoriz": "Equipos Horizontales",
        "Ilum": "Sistemas de Iluminación",
        "OtrosEq": "Otros Equipos",
        "Conta": "Conta",
        "Envol": "Envol"
    }

    @staticmethod
    def _get_sheets_for_anexo(anexo_number: int) -> Dict[str, str]:
        """Retorna las hojas específicas según el anexo."""
        if anexo_number == 2:
            return {"Conta": DefaultExcelRepository.SHEETS_MAP["Conta"]}
        elif anexo_number == 3:
            return {
                key: DefaultExcelRepository.SHEETS_MAP[key] 
                for key in ["Clima", "SistCC", "Eleva", "EqHoriz", "Ilum", "OtrosEq"]
            }
        elif anexo_number == 4:
            return {"Envol": DefaultExcelRepository.SHEETS_MAP["Envol"]}
        else:
            return DefaultExcelRepository.SHEETS_MAP

    @staticmethod
    def _delete_trash_rows(df: pd.DataFrame, col: str = "ID EDIFICIO") -> pd.DataFrame:
        if col in df.columns:
            for i in range(len(df) - 1, -1, -1):
                val = df[col].iloc[i]
                if pd.notna(val) and str(val).strip() not in ("", "0"):
                    return df.iloc[: i + 2].copy()
        return df.copy()

    @staticmethod
    def _round_numeric(df: pd.DataFrame) -> None:
        for col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col])
                if pd.api.types.is_float_dtype(df[col]):
                    df[col] = df[col].round(2)
            except Exception:
                pass

    def _load_sheets(self, excel_path: Path, sheets_map: Dict[str, str]) -> Dict[str, pd.DataFrame]:
        """Método genérico para cargar hojas específicas."""
        with pd.ExcelFile(excel_path) as xls:
            missing = [s for s in sheets_map if s not in xls.sheet_names]
            if missing:
                raise ValueError(f"Hojas faltantes en el Excel: {', '.join(missing)}")
            data: Dict[str, pd.DataFrame] = {}
            for sheet in sheets_map:
                logger.info(f"-> Procesando hoja: {sheet}")
                df = pd.read_excel(xls, sheet, header=0, dtype=str)
                df = self._delete_trash_rows(df).fillna("")
                self._round_numeric(df)
                data[sheet] = df
        return data

    def load_sheets_for_anexo2(self, excel_path: Path) -> Dict[str, pd.DataFrame]:
        """Carga solo la hoja 'Conta' para el Anexo 2."""
        sheets_map = self._get_sheets_for_anexo(2)
        return self._load_sheets(excel_path, sheets_map)

    def load_sheets_for_anexo3(self, excel_path: Path) -> Dict[str, pd.DataFrame]:
        """Carga las hojas específicas para el Anexo 3."""
        sheets_map = self._get_sheets_for_anexo(3)
        return self._load_sheets(excel_path, sheets_map)

    def load_sheets_for_anexo4(self, excel_path: Path) -> Dict[str, pd.DataFrame]:
        """Carga solo la hoja 'Envol' para el Anexo 4."""
        sheets_map = self._get_sheets_for_anexo(4)
        return self._load_sheets(excel_path, sheets_map)

    @staticmethod
    def extract_unique_groups(
        group_col: str, tables: Dict[str, pd.DataFrame]
    ) -> Sequence[str]:
        groups = set()
        for df in tables.values():
            if group_col in df.columns:
                values = df[group_col].dropna().unique()
                groups.update({str(v).strip() for v in values if str(v).strip()})
        return sorted(groups)

    @staticmethod
    def filter_tables_by_group_anexo2(
        group_col: str, tables: Dict[str, pd.DataFrame], group: str
    ) -> List[pd.DataFrame]:
        """Filtra tablas para Anexo 2 (solo Conta)."""
        return [tables["Conta"][tables["Conta"].get(group_col) == group].copy()]

    @staticmethod
    def filter_tables_by_group_anexo3(
        group_col: str, tables: Dict[str, pd.DataFrame], group: str
    ) -> List[pd.DataFrame]:
        """Filtra tablas para Anexo 3."""
        return [
            df[df.get(group_col) == group].copy()
            for df in (
                tables["Clima"],
                tables["SistCC"],
                tables["Eleva"],
                tables["EqHoriz"],
                tables["Ilum"],
                tables["OtrosEq"],
            )
        ]

    @staticmethod
    def filter_tables_by_group_anexo4(
        group_col: str, tables: Dict[str, pd.DataFrame], group: str
    ) -> List[pd.DataFrame]:
        """Filtra tablas para Anexo 4 (solo Envol)."""
        return [tables["Envol"][tables["Envol"].get(group_col) == group].copy()]

    @staticmethod
    def filter_tables_by_group(
        group_col: str, tables: Dict[str, pd.DataFrame], group: str
    ) -> List[pd.DataFrame]:
        """Método genérico que mantiene compatibilidad hacia atrás."""
        # Para compatibilidad, usar el comportamiento del Anexo 3
        available_keys = list(tables.keys())
        filtered_dfs = []
        for key in ["Clima", "SistCC", "Eleva", "EqHoriz", "Ilum", "OtrosEq"]:
            if key in available_keys:
                filtered_dfs.append(tables[key][tables[key].get(group_col) == group].copy())
        return filtered_dfs

    @staticmethod
    def extract_center_id(df_group_list: List[pd.DataFrame]) -> str:
        for dfg in df_group_list:
            if not dfg.empty:
                for _, row in dfg.iterrows():
                    if pd.notna(row.get("ID CENTRO")):
                        return str(row.get("ID CENTRO", ""))
        return ""

    @staticmethod
    def _is_valid_total_col(col: str, last_row: pd.Series) -> bool:
        return pd.notna(last_row[col]) and str(last_row[col]).strip() != ""

    def calculate_totals_by_center(
        self, complete_df: pd.DataFrame, df_group: pd.DataFrame
    ) -> Dict[str, str]:
        if complete_df.empty:
            return {}
        last_row = complete_df.iloc[-1]
        cols = [
            c for c in complete_df.columns[1:] if self._is_valid_total_col(c, last_row)
        ]
        key_label = "CENTRO" if "CENTRO" in complete_df.columns else "EDIFICIO"
        totals: Dict[str, str] = {key_label: "Total general"}
        for c in cols:
            nums = pd.to_numeric(df_group[c], errors="coerce").dropna()
            if nums.empty:
                totals[c] = ""
            else:
                s = nums.sum()
                totals[c] = (
                    str(int(round(s)))
                    if abs(s - round(s)) < 1e-6
                    else f"{s:.2f}".rstrip("0").rstrip(".")
                )
        return totals


# ---------------- Output path builder ----------------


class DefaultOutputPathBuilder:
    """Salida: {base}/{id_centro}/{filename}. Si no se indica --output-dir, usa ./word/anexos/."""

    def build_output_docx_path(
        self, config: RunConfig, center_id: str, filename: str
    ) -> Path:
        base = config.output_dir or (
            Path(__file__).resolve().parent / "word" / "anexos"
        )
        out_dir = base / center_id
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir / filename


# =====================================================================================
# Helpers de dominio
# =====================================================================================


def clean_name(filename: str) -> str:
    invalid = '<>:"|?*\\/“”'
    cleaned = filename
    for ch in invalid:
        cleaned = cleaned.replace(ch, "")
    cleaned = unicodedata.normalize("NFD", cleaned)
    cleaned = "".join(c for c in cleaned if unicodedata.category(c) != "Mn")
    cleaned = re.sub(r"[_\s]+", " ", cleaned).strip()
    return cleaned[:100].strip()


def request_month_and_year(
    config_month: Optional[str] = None,
    config_year: Optional[int] = None,
    stdin_queue: Optional["queue.Queue[str]"] = None,
) -> Tuple[str, int]:
    """Pide mes y año por stdin (compatible con la GUI) o usa valores de configuración."""

    def _input(prompt: str) -> str:
        try:
            return input(prompt)
        except EOFError:
            # si no hay stdin, usar valores por defecto
            return ""

    months = {
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
    now = datetime.now()

    # Si se proporcionan valores de configuración, usarlos
    if config_month is not None and config_year is not None:
        # Convertir mes si es número
        if config_month.isdigit():
            month_num = int(config_month)
            if 1 <= month_num <= 12:
                month_name = months[month_num]
            else:
                raise ValueError(
                    f"Mes inválido: {config_month}. Debe estar entre 1 y 12."
                )
        else:
            # Buscar mes por nombre
            month_name = config_month
            month_num = None
            for num, name in months.items():
                if name.lower() == config_month.lower():
                    month_num = num
                    break
            if month_num is None:
                raise ValueError(f"Nombre de mes inválido: {config_month}")

        year = config_year
        if not (now.year - 5 <= year <= now.year + 5):
            raise ValueError(
                f"Año inválido: {year}. Debe estar entre {now.year - 5} y {now.year + 5}."
            )

        logger.info(
            f"Usando configuración: {month_name} ({month_num if 'month_num' in locals() else 'N/A'}), {year}"
        )
        return month_name, year

    # MES - modo interactivo
    print(
        f"\nCONFIGURACIÓN DEL DOCUMENTO\nMes actual: {months[now.month]} ({now.month})"
    )
    while True:
        raw = _input(f"Ingrese el mes (1-12) [Enter para usar {now.month}]: ").strip()
        if not raw:
            m = now.month
        else:
            try:
                m = int(raw)
            except Exception:
                print("Error: Por favor ingrese un número válido")
                continue
        if 1 <= m <= 12:
            month_name = months[m]
            break
        print("Error: El mes debe estar entre 1 y 12")
    # AÑO - modo interactivo
    while True:
        raw = _input(f"Ingrese el año [Enter para usar {now.year}]: ").strip()
        y = now.year if not raw else int(raw)
        if now.year - 5 <= y <= now.year + 5:
            break
        print(f"Error: El año debe estar entre {now.year - 5} y {now.year + 5}")
    print(f"Selección: {month_name} ({m}), {y}\n")
    return month_name, y


# =====================================================================================
# Casos de uso (Use Cases)
# =====================================================================================


class AnexoGenerator(Protocol):
    """Contrato común de generadores de anexos."""

    anexo_number: int

    def generate(
        self,
        excel_path: Path,
        config_month: Optional[str] = None,
        config_year: Optional[int] = None,
    ) -> List[OutputFile]: ...


class Anexo3Generator:
    
    anexo_number = 3

    def __init__(
        self,
        templates: TemplateProvider,
        word: WordExporter,
        pdf: PdfInspector,
        excel: ExcelRepository,
        out: OutputPathBuilder,
        group_column: str = "CENTRO",
    ) -> None:
        self.templates = templates
        self.word = word
        self.pdf = pdf
        self.excel = excel
        self.out = out
        self.group_column = group_column

    def _export_and_prune_pdf(
        self,
        docx_path: Path,
        sections_empty_flags: Dict[str, bool],
        excel_titles: Dict[str, str],
    ) -> Optional[Path]:
        final_pdf = docx_path.with_suffix(".pdf")
        tmp_pdf = final_pdf.with_suffix(".tmp.pdf")
        # 1) DOCX -> PDF temporal
        self.pdf.export_docx_to_temp_pdf(self.word, docx_path, tmp_pdf)

        # 2) Detectar páginas a eliminar
        try:
            total_pages = self.pdf.read_total_pages(tmp_pdf)
        except Exception as e:
            logger.warning(f"   ! Error leyendo PDF temporal: {e}")
            try:
                final_pdf.unlink(missing_ok=True)  # type: ignore[attr-defined]
            except Exception:
                pass
            tmp_pdf.replace(final_pdf)
            return final_pdf

        pages_to_delete: set[int] = set()
        for key, is_empty in sections_empty_flags.items():
            if not is_empty:
                continue
            title_text = excel_titles.get(key)
            if not title_text:
                continue
            logger.info(f"   -> Buscando sección vacía: '{title_text}' (key: {key})")
            title_page_index = self.pdf.find_title_page_in_pdf(tmp_pdf, title_text)
            if title_page_index is not None:
                pages_to_delete.add(title_page_index)
                if title_page_index + 1 <= total_pages:
                    pages_to_delete.add(title_page_index + 1)

        logger.info(f"   -> Páginas a eliminar: {sorted(pages_to_delete)}")

        # 3) Si hay que borrar, hacerlo en el DOCX y reexportar el PDF
        if pages_to_delete:
            try:
                tmp_pdf.unlink(missing_ok=True)  # type: ignore[attr-defined]
            except Exception:
                pass
            app, doc = self.word.open_document(docx_path, read_only=False)
            try:
                self.word.delete_pages(doc, pages_to_delete)
                self.word.update_toc(doc)
                doc.Save()
                self.word.export_doc_to_pdf(doc, tmp_pdf)
            finally:
                doc.Close(SaveChanges=False)
                app.Quit()
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

            # quedar con PDF final sin páginas (casi) en blanco
            try:
                final_pdf.unlink(missing_ok=True)  # type: ignore[attr-defined]
            except Exception:
                pass
            try:
                self.pdf.remove_blank_pages_from_pdf(tmp_pdf)
            except Exception as e:
                logger.warning(f"   ! Limpieza PDF: {e}")
            tmp_pdf.replace(final_pdf)
            return final_pdf

        # 4) si no hay cambios
        try:
            final_pdf.unlink(missing_ok=True)  # type: ignore[attr-defined]
        except Exception:
            pass
        tmp_pdf.replace(final_pdf)
        try:
            self.pdf.remove_blank_pages_from_pdf(final_pdf)
        except Exception:
            pass
        return final_pdf

    def generate(
        self,
        excel_path: Path,
        config_month: Optional[str] = None,
        config_year: Optional[int] = None,
    ) -> List[OutputFile]:
        month_name, year = request_month_and_year(config_month, config_year)
        self.word.close_word_processes()

        tables = self.excel.load_sheets_for_anexo3(excel_path)
        centers = self.excel.extract_unique_groups(self.group_column, tables)
        logger.info(
            f"-> Se generarán documentos para {len(centers)} {self.group_column.lower()}s"
        )

        tpl_bytes = self.templates.get_template(self.anexo_number)

        outputs: List[OutputFile] = []
        for center in centers:
            (
                df_clima_grupo,
                df_sist_cc_grupo,
                df_eleva_grupo,
                df_eqhoriz_grupo,
                df_ilum_grupo,
                df_otros_eq_grupo,
            ) = self.excel.filter_tables_by_group_anexo3(self.group_column, tables, center)

            if all(
                len(d) == 0
                for d in (
                    df_clima_grupo,
                    df_sist_cc_grupo,
                    df_eleva_grupo,
                    df_eqhoriz_grupo,
                    df_ilum_grupo,
                    df_otros_eq_grupo,
                )
            ):
                continue

            totales_clima = self.excel.calculate_totals_by_center(
                tables["Clima"], df_clima_grupo
            )
            totales_sist_cc = self.excel.calculate_totals_by_center(
                tables["SistCC"], df_sist_cc_grupo
            )
            totales_eleva = self.excel.calculate_totals_by_center(
                tables["Eleva"], df_eleva_grupo
            )
            totales_eqhoriz = self.excel.calculate_totals_by_center(
                tables["EqHoriz"], df_eqhoriz_grupo
            )
            totales_ilum = self.excel.calculate_totals_by_center(
                tables["Ilum"], df_ilum_grupo
            )
            totales_otros_eq = self.excel.calculate_totals_by_center(
                tables["OtrosEq"], df_otros_eq_grupo
            )

            ctx = {
                "mes": month_name,
                "anio": year,
                "centro": center,
                "df_clima": df_clima_grupo.to_dict("records"),
                "df_sist_cc": df_sist_cc_grupo.to_dict("records"),
                "df_eleva": df_eleva_grupo.to_dict("records"),
                "df_eqhoriz": df_eqhoriz_grupo.to_dict("records"),
                "df_ilum": df_ilum_grupo.to_dict("records"),
                "df_otros_eq": df_otros_eq_grupo.to_dict("records"),
                "totales_clima": [totales_clima],
                "totales_sist_cc": [totales_sist_cc],
                "totales_eleva": [totales_eleva],
                "totales_eqhoriz": [totales_eqhoriz],
                "totales_ilum": [totales_ilum],
                "totales_otros_eq": [totales_otros_eq],
            }

            doc = DocxTemplate(BytesIO(tpl_bytes))
            doc.render(ctx)

            center_id = DefaultExcelRepository.extract_center_id(
                [
                    df_clima_grupo,
                    df_sist_cc_grupo,
                    df_eleva_grupo,
                    df_eqhoriz_grupo,
                    df_ilum_grupo,
                    df_otros_eq_grupo,
                ]
            )

            safe_center = clean_name(str(center))
            out_name = f"Anexo 3 {safe_center}.docx"
            out_path = self.out.build_output_docx_path(CONFIG, center_id, out_name)  # type: ignore[name-defined]
            doc.save(str(out_path))

            logger.info(f"   -> Eliminando páginas en blanco de {out_name}")
            try:
                removed = self.word.remove_blank_pages_from_docx(out_path)
                if removed > 0:
                    logger.info(f"   ✓ {removed} páginas en blanco eliminadas")
                else:
                    logger.info("   ✓ No se encontraron páginas en blanco")
            except Exception as e:
                logger.warning(f"   ! Error eliminando páginas en blanco: {e}")

            sections_empty = {
                "Clima": df_clima_grupo.empty,
                "SistCC": df_sist_cc_grupo.empty,
                "Eleva": df_eleva_grupo.empty,
                "EqHoriz": df_eqhoriz_grupo.empty,
                "Ilum": df_ilum_grupo.empty,
                "OtrosEq": df_otros_eq_grupo.empty,
            }

            pdf_path = None
            try:
                pdf_path = self._export_and_prune_pdf(
                    out_path, sections_empty, DefaultExcelRepository.SHEETS_MAP
                )
            except Exception as e:
                logger.warning(f"   ! No se pudo generar PDF limpio: {e}")

            logger.info(f"* Documento generado: {center_id}/{out_name}")
            outputs.append(OutputFile(docx_path=out_path, pdf_path=pdf_path))

        return outputs


class Anexo2Generator:
    """
    Implementación provisional del Anexo 2 basada en el código original.
    NOTA: La estructura real de datos puede diferir; este generador mantiene
    la lógica previa y deja listo el esqueleto para adaptar con precisión.
    """

    anexo_number = 2

    def __init__(
        self,
        templates: TemplateProvider,
        word: WordExporter,
        pdf: PdfInspector,
        excel: ExcelRepository,
        out: OutputPathBuilder,
        group_column: str = "CENTRO",
    ) -> None:
        self.templates = templates
        self.word = word
        self.pdf = pdf
        self.excel = excel
        self.out = out
        self.group_column = group_column

    def generate(
        self,
        excel_path: Path,
        config_month: Optional[str] = None,
        config_year: Optional[int] = None,
    ) -> List[OutputFile]:
        month_name, year = request_month_and_year(config_month, config_year)
        self.word.close_word_processes()

        # Cargar solo la hoja Conta para Anexo 2
        tables = self.excel.load_sheets_for_anexo2(excel_path)
        centers = self.excel.extract_unique_groups(self.group_column, tables)
        logger.info("* Datos cargados y limpiados\n-> Renderizando documentos…")

        tpl_bytes = self.templates.get_template(self.anexo_number)
        outputs: List[OutputFile] = []
        generated_docs: List[str] = []

        for center in centers:
            dfs = self.excel.filter_tables_by_group_anexo2(self.group_column, tables, center)
            df_conta = dfs[0]

            if df_conta.empty:
                continue

            ctx = {
                "mes": month_name,
                "anio": year,
                "centro": center,
                "df_conta": df_conta.to_dict("records"),
                "tipo_de_suministro": df_conta["SUMINISTRO"].unique().tolist()
                if "SUMINISTRO" in df_conta.columns
                else [],
            }

            doc = DocxTemplate(BytesIO(tpl_bytes))
            doc.render(ctx)

            center_id = self.excel.extract_center_id(dfs)
            safe_center = clean_name(center)
            out_name = f"Anexo 2 {safe_center}.docx"
            out_path = self.out.build_output_docx_path(CONFIG, center_id, out_name)  # type: ignore[name-defined]
            doc.save(str(out_path))

            generated_docs.append(str(out_path))
            outputs.append(OutputFile(docx_path=out_path, pdf_path=None))
            logger.info(f"* Documento generado: {center_id}/{out_name}")

        # Actualizar campos de Word en lote después de generar todos los documentos
        if generated_docs:
            logger.info("\nActualizando campos en lote (TOC: solo paginación)...")
            self.word.update_word_fields_bulk(generated_docs)

            # Procesar conversión a PDF en lote después de actualizar campos
            logger.info("\nConvirtiendo documentos a PDF...")
            pdf_files = self.pdf.convert_docx_to_pdf_bulk(generated_docs)

            if pdf_files:
                logger.info(f"   ✓ Se generaron {len(pdf_files)} archivos PDF correctamente")

                logger.info("\nEliminando última página de los PDFs...")
                self.pdf.remove_last_page_from_pdfs(pdf_files)

                # Actualizar outputs con las rutas de PDF
                pdf_dict = {str(Path(pdf).with_suffix(".docx")): Path(pdf) for pdf in pdf_files}
                for output in outputs:
                    if str(output.docx_path) in pdf_dict:
                        outputs[outputs.index(output)] = OutputFile(
                            docx_path=output.docx_path,
                            pdf_path=pdf_dict[str(output.docx_path)]
                        )
            else:
                logger.warning("   ! No se pudieron generar archivos PDF")

        return outputs
    
class Anexo4Generator:
    
    anexo_number = 4

    def __init__(
        self,
        templates: TemplateProvider,
        word: WordExporter,
        pdf: PdfInspector,
        excel: ExcelRepository,
        out: OutputPathBuilder,
        group_column: str = "CENTRO",
    ) -> None:
        self.templates = templates
        self.word = word
        self.pdf = pdf
        self.excel = excel
        self.out = out
        self.group_column = group_column

    def generate(
        self,
        excel_path: Path,
        config_month: Optional[str] = None,
        config_year: Optional[int] = None,
    ) -> List[OutputFile]:
        month_name, year = request_month_and_year(config_month, config_year)
        self.word.close_word_processes()

        tables = self.excel.load_sheets_for_anexo4(excel_path)
        centers = self.excel.extract_unique_groups(self.group_column, tables)
        logger.info(
            f"-> Se generarán documentos para {len(centers)} {self.group_column.lower()}s"
        )
        tpl_bytes = self.templates.get_template(self.anexo_number)
        
        generated_docs: List[str] = []
        outputs: List[OutputFile] = []
        
        for center in centers:
            df_envol_grupo = self.excel.filter_tables_by_group_anexo4(self.group_column, tables, center)[0]

            if df_envol_grupo.empty:
                continue

            totales_envol = self.excel.calculate_totals_by_center(
                tables["Envol"], df_envol_grupo
            )

            ctx = {
                "mes": month_name,
                "anio": year,
                "centro": center,
                "df_envol": df_envol_grupo.to_dict("records"),
                "totales_envol": [totales_envol],
            }

            doc = DocxTemplate(BytesIO(tpl_bytes))
            doc.render(ctx)

            center_id = DefaultExcelRepository.extract_center_id(
                [df_envol_grupo]
            )

            safe_center = clean_name(str(center))
            out_name = f"Anexo 4 {safe_center}.docx"
            out_path = self.out.build_output_docx_path(CONFIG, center_id, out_name)  # type: ignore[name-defined]
            doc.save(str(out_path))
            
            generated_docs.append(str(out_path))
            outputs.append(OutputFile(docx_path=out_path, pdf_path=None))
            logger.info(f"* Documento generado: {center_id}/{out_name}")
            
        # Procesar conversión a PDF en lote después de actualizar campos
        logger.info("\nConvirtiendo documentos a PDF...")
        pdf_files = self.pdf.convert_docx_to_pdf_bulk(generated_docs)

        if pdf_files:
            logger.info(f"   ✓ Se generaron {len(pdf_files)} archivos PDF correctamente")
            
        return outputs
            


# =====================================================================================
# Orquestador / Aplicación
# =====================================================================================


class AnexoFactory:
    """Registra e instancia generadores disponibles."""

    def __init__(
        self,
        templates: TemplateProvider,
        word: WordExporter,
        pdf: PdfInspector,
        excel: ExcelRepository,
        out: OutputPathBuilder,
    ) -> None:
        self._templates = templates
        self._word = word
        self._pdf = pdf
        self._excel = excel
        self._out = out

    def get(self, n: int) -> AnexoGenerator:
        if n == 3:
            return Anexo3Generator(
                self._templates, self._word, self._pdf, self._excel, self._out
            )
        if n == 2:
            return Anexo2Generator(self._templates, self._word, self._pdf, self._excel, self._out)
        if n == 4:
            return Anexo4Generator(self._templates, self._word, self._pdf, self._excel, self._out)
        raise NotImplementedError(f"Generador para Anexo {n} no implementado")



def run_application(config: RunConfig) -> int:
    global CONFIG  # usado por generadores para construir paths de salida
    CONFIG = config  # type: ignore[assignment]

    # Inyectar dependencias (adapters)
    templates = DefaultTemplateProvider(config.word_dir)
    word = DefaultWordExporter()
    pdf = DefaultPdfInspector()
    excel = DefaultExcelRepository()
    out = DefaultOutputPathBuilder()
    factory = AnexoFactory(templates, word, pdf, excel, out)

    # Validaciones mínimas (la GUI ya valida en capa anterior)
    if not config.excel_dir.is_dir():
        logger.error("La carpeta de Excel no existe o no es válida.")
        return 2
    excel_files = [
        p
        for p in config.excel_dir.iterdir()
        if p.suffix.lower() in (".xlsx", ".xlsm", ".xls")
    ]
    if not excel_files:
        logger.error("La carpeta de Excel no contiene archivos .xls/.xlsx/.xlsm.")
        return 2

    # Modo ejecución
    implemented = [2, 3]
    try:
        if config.mode == "all":
            target_anexos = implemented
        else:
            if config.anexo is None:
                logger.error("Debes indicar --anexo N cuando --mode single.")
                return 2
            target_anexos = [config.anexo]
        logger.info(f"Anexos seleccionados: {target_anexos}")
    except Exception as e:
        logger.error(f"Error interpretando modo/anexo: {e}")
        return 2

    # Procesar cada Excel encontrado
    for xfile in excel_files:
        logger.info(f"\n=== Procesando: {xfile.name} ===")
        for n in target_anexos:
            try:
                generator = factory.get(n)
            except NotImplementedError as e:
                logger.warning(str(e))
                continue
            try:
                generator.generate(xfile, config.month, config.year)
            except FileNotFoundError as e:
                logger.error(str(e))
                return 3
            except Exception as e:
                logger.error(f"Error generando Anexo {n}: {e}")
                # seguir con otros anexos/archivos

    logger.info("\n--- Proceso finalizado correctamente ---")
    return 0


# =====================================================================================
# CLI
# =====================================================================================


def parse_args(argv: Optional[Sequence[str]] = None) -> RunConfig:
    parser = argparse.ArgumentParser(
        description="Generador de Anexos (refactor SOLID + Clean Architecture)"
    )
    parser.add_argument(
        "--excel-dir", required=True, help="Carpeta que contiene los Excels de entrada"
    )
    parser.add_argument(
        "--word-dir", help="Carpeta donde están las plantillas de Word (DOCX). Opcional"
    )
    parser.add_argument(
        "--mode",
        choices=["all", "single"],
        required=True,
        help="Ejecutar todos o un anexo concreto",
    )
    parser.add_argument(
        "--anexo", type=int, help="Número de anexo cuando --mode single"
    )
    parser.add_argument(
        "--month",
        help="Mes para los documentos (nombre o número 1-12). Si no se proporciona, se solicita interactivamente",
    )
    parser.add_argument(
        "--year",
        type=int,
        help="Año para los documentos. Si no se proporciona, se solicita interactivamente",
    )
    parser.add_argument(
        "--html-templates-dir",
        help="Carpeta de plantillas HTML (reservado para anexos futuros)",
    )
    parser.add_argument(
        "--photos-dir", help="Carpeta de fotos (reservado para anexos futuros)"
    )
    parser.add_argument(
        "--output-dir",
        help="Carpeta de salida para anexos. Si no se indica, se usa ./word/anexos",
    )

    ns = parser.parse_args(argv)

    def _p(x: Optional[str]) -> Optional[Path]:  # helper
        return Path(x) if x else None

    return RunConfig(
        excel_dir=Path(ns.excel_dir),
        word_dir=_p(ns.word_dir),
        mode=ns.mode,
        anexo=ns.anexo,
        month=ns.month,
        year=ns.year,
        html_templates_dir=_p(ns.html_templates_dir),
        photos_dir=_p(ns.photos_dir),
        output_dir=_p(ns.output_dir),
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    ensure_utf8_console()
    setup_logging(logging.INFO)
    try:
        config = parse_args(argv)
        return run_application(config)
    except SystemExit as e:
        # argparse ya imprimió el error
        return int(e.code or 2)
    except Exception as e:
        logger.error(f"Error inesperado: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
