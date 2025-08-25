from __future__ import annotations
import pandas as pd
import unicodedata
import re
import subprocess
import time
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from dataclasses import dataclass
import logging
import win32com.client as win32_client
import pythoncom
from win32com.client import CDispatch
import argparse


class ContentChecker:
    """Unified checker for all types of page content in Word documents."""

    def __init__(self, doc: CDispatch):
        self.doc = doc

    def has_content(self, page_range: CDispatch) -> bool:
        """
        Check if page has any type of content.

        Returns True if page has:
        - Meaningful text content
        - Tables with content
        - Images or embedded objects
        - Shapes with text
        """
        try:
            return (
                self._has_text_content(page_range)
                or self._has_table_content(page_range)
                or self._has_inline_shapes_content(page_range)
                or self._has_shapes_content(page_range)
            )
        except Exception as e:
            logging.debug(f"Error checking page content: {e}")
            return True  # Assume content exists on error for safety

    def _has_text_content(self, page_range: CDispatch) -> bool:
        """Check if page has meaningful text content."""
        try:
            text_content = page_range.Text.strip()
            meaningful_text = self._extract_meaningful_text(text_content)
            return len(meaningful_text) > 0
        except Exception as e:
            logging.debug(f"Error checking text content: {e}")
            return True  # Assume content exists on error for safety

    def _has_table_content(self, page_range: CDispatch) -> bool:
        """Check if page has table content."""
        try:
            if page_range.Tables.Count == 0:
                return False

            table = page_range.Tables(1)  # Word uses 1-based indexing
            table_text = table.Range.Text.strip()
            meaningful_text = self._extract_meaningful_text(table_text)
            return len(meaningful_text) > 0
        except Exception as e:
            logging.debug(f"Error checking table content: {e}")
            return False

    def _has_inline_shapes_content(self, page_range: CDispatch) -> bool:
        """Check if page has inline shapes content."""
        try:
            return page_range.InlineShapes.Count > 0
        except Exception as e:
            logging.debug(f"Error checking inline shapes: {e}")
            return False

    def _has_shapes_content(self, page_range: CDispatch) -> bool:
        """Check if page has shapes with text content."""
        try:
            page_start = page_range.Start
            page_end = page_range.End

            for shape in self.doc.Shapes:
                if self._is_shape_in_page_range(shape, page_start, page_end):
                    if self._shape_has_text_content(shape):
                        return True
            return False
        except Exception as e:
            logging.debug(f"Error checking shapes content: {e}")
            return False

    @staticmethod
    def _extract_meaningful_text(text: str) -> str:
        """Extract meaningful text by removing whitespace characters."""
        if not text:
            return ""
        return text.replace("\r", "").replace("\n", "").replace("\t", "").strip()

    @staticmethod
    def _is_shape_in_page_range(
        shape: CDispatch, page_start: int, page_end: int
    ) -> bool:
        """Check if shape is within the page range."""
        try:
            return (
                hasattr(shape, "Anchor")
                and shape.Anchor.Start >= page_start
                and shape.Anchor.End <= page_end
            )
        except Exception:
            return False

    @staticmethod
    def _shape_has_text_content(shape: CDispatch) -> bool:
        """Check if shape has meaningful text content."""
        try:
            has_text_frame = (
                hasattr(shape, "TextFrame")
                and hasattr(shape.TextFrame, "HasText")
                and shape.TextFrame.HasText
            )

            if not has_text_frame:
                return False

            shape_text = shape.TextFrame.TextRange.Text.strip()
            return bool(shape_text and len(shape_text) > 0)
        except Exception:
            return False


class WordPageRemover:
    """Provides page functionality  for Word documents."""

    def __init__(self, doc: CDispatch):
        self.doc = doc
        self._constants = self._initialize_constants()
        self.content_checker = ContentChecker(doc)

    @staticmethod
    def _initialize_constants() -> dict[str, int]:
        """Initialize Word constants."""
        return {
            "WD_GO_TO_PAGE": 1,
            "WD_GO_TO_ABSOLUTE": 1,
            "WD_STORY": 6,
            "WD_STATISTIC_PAGES": 2,
            "WD_ACTIVE_END_PAGE_NUMBER": 3,
            "WD_EXPORT_FORMAT_PDF": 17,
            "WD_EXPORT_OPTIMIZE_FOR_PRINT": 0,
            "WD_EXPORT_ALL_DOCUMENT": 0,
            "WD_EXPORT_DOCUMENT_CONTENT": 0,
            "WD_EXPORT_CREATE_HEADING_BOOKMARKS": 1,
        }

    def get_page_range(self, page_num: int) -> CDispatch | None:
        """Get the range for a specific page."""
        try:
            app = self.doc.Application
            page_start = self._get_page_start(app, page_num)
            page_end = self._get_page_end(app, page_num)

            if page_end > page_start:
                return self.doc.Range(Start=page_start, End=page_end)
            return None
        except Exception as e:
            logging.error(f"Error getting page range for page {page_num}: {e}")
            return None

    def _get_page_start(self, app: CDispatch, page_num: int) -> int:
        """Get the start position of a page."""
        app.Selection.GoTo(
            What=self._constants["WD_GO_TO_PAGE"],
            Which=self._constants["WD_GO_TO_ABSOLUTE"],
            Count=page_num,
        )
        return app.Selection.Range.Start

    def _get_page_end(self, app: CDispatch, page_num: int) -> int:
        """Get the end position of a page."""
        total_pages = int(
            self.doc.ComputeStatistics(self._constants["WD_STATISTIC_PAGES"])
        )

        if page_num < total_pages:
            app.Selection.GoTo(
                What=self._constants["WD_GO_TO_PAGE"],
                Which=self._constants["WD_GO_TO_ABSOLUTE"],
                Count=page_num + 1,
            )
            return app.Selection.Range.Start
        else:
            app.Selection.EndKey(Unit=self._constants["WD_STORY"])
            return app.Selection.Range.End

    def delete_page(self, page_num: int) -> bool:
        """
        Delete a page from the Word document.

        Args:
            page_num: Page number to delete (1-based)

        Returns:
            bool: True if deletion was successful, False otherwise
        """
        try:
            page_range = self._get_precise_page_range_for_deletion(page_num)
            if page_range is None:
                logging.warning(
                    f"Could not get page range for deletion of page {page_num}"
                )
                return False

            page_range.Delete()
            logging.debug(f"Successfully deleted page {page_num}")
            return True
        except Exception as e:
            logging.error(f"Failed to delete page {page_num}: {e}")
            return False

    def _get_precise_page_range_for_deletion(self, page_num: int) -> CDispatch | None:
        """
        Get precise page range for deletion, handling edge cases.

        Args:
            page_num: Page number (1-based)

        Returns:
            CDispatch | None: Page range object or None if invalid
        """
        try:
            app = self.doc.Application
            total_pages = int(
                self.doc.ComputeStatistics(self._constants["WD_STATISTIC_PAGES"])
            )  # WD_STATISTIC_PAGES = 2

            if not self._is_valid_page_number(page_num, total_pages):
                return None

            start_pos = self._get_page_start(app, page_num)
            end_pos = self._get_page_end_for_deletion(app, page_num, total_pages)

            if end_pos <= start_pos:
                logging.warning(
                    f"Invalid range for page {page_num}: start={start_pos}, end={end_pos}"
                )
                return None

            return self.doc.Range(Start=start_pos, End=end_pos)
        except Exception as e:
            logging.error(
                f"Error getting page range for deletion of page {page_num}: {e}"
            )
            return None

    def _get_page_end_for_deletion(
        self, app: CDispatch, page_num: int, total_pages: int
    ) -> int:
        """Get the end position for page deletion, handling last page case."""
        if page_num < total_pages:
            app.Selection.GoTo(
                What=self._constants["WD_GO_TO_PAGE"],
                Which=self._constants["WD_GO_TO_ABSOLUTE"],
                Count=page_num + 1,
            )
            return app.Selection.Range.Start
        else:
            # For last page, go to end of document
            app.Selection.EndKey(Unit=self._constants["WD_STORY"])
            return app.Selection.Range.End

    @staticmethod
    def _is_valid_page_number(page_num: int, total_pages: int) -> bool:
        """Validate page number is within document bounds."""
        return 1 <= page_num <= total_pages

    def is_page_blank(self, page_num: int) -> bool:
        """
        Determine if a page is blank using multiple criteria.

        A page is considered blank if it has no:
        - Meaningful text content
        - Tables with content
        - Images or embedded objects
        - Shapes with text

        Args:
            page_num: Page number to check (1-based)

        Returns:
            True if page is blank, False otherwise
        """
        try:
            page_range = self.get_page_range(page_num)
            if page_range is None:
                logging.warning(f"Could not get page range for page {page_num}")
                return False  # Assume not blank if we can't check

            return not self.content_checker.has_content(page_range)
        except Exception as e:
            logging.error(f"Error checking if page {page_num} is blank: {e}")
            return False  # Assume not blank on error for safety


class WordManager:
    """Manages Word document operations."""

    def __init__(self, doc: CDispatch):
        self.doc = doc
        self.page_remover = WordPageRemover(doc)

    


@dataclass
class DfManager:
    def __init__(self):
        self._excel_sheets: list[str] = {
            "Clima": "Sistemas de Climatización",
            "SistCC": "Sistemas de Calefacción",
            "Eleva": "Equipos Elevadores",
            "EqHoriz": "Equipos Horizontales",
            "Ilum": "Sistemas de Iluminación",
            "OtrosEq": "Otros Equipos"
            # "Envol": "Envol",
            # "Conta": "Conta"
        }

    @staticmethod
    def _delete_trash_rows(
        df: pd.DataFrame, columna: str = "ID EDIFICIO"
    ) -> pd.DataFrame:
        """Elimina las filas vacías al final del DataFrame."""
        if columna in df.columns:
            # Encontrar la última fila con datos válidos
            for i in range(len(df) - 1, -1, -1):
                val = df[columna].iloc[i]

                value_is_not_na = pd.notna(val)
                value_is_non_empty = str(val).strip() != ""
                value_is_not_zero = str(val).strip() != "0"

                if value_is_not_na and value_is_non_empty and value_is_not_zero:
                    return df.iloc[: i + 2].copy()
        return df.copy()

    @staticmethod
    def clean_last_row(df: pd.DataFrame) -> pd.DataFrame:
        """
        Limpia la última fila, que se corresponde con la de totales
        """
        if df.empty:
            return df
        df2 = df.copy()
        mask = df2.iloc[-1:] == "Total"
        df2.iloc[-1:] = df2.iloc[-1:].mask(mask, pd.NA)
        return df2

    def load_and_clean_sheets(self, excel_path) -> dict[str, pd.DataFrame]:
        """
        Carga y limpia todas las hojas especificadas.
        """
        with pd.ExcelFile(excel_path) as xls:
            # Las claves del excel_sheets son los nombres reales de las hojas
            missing = [
                sheet for sheet in self._excel_sheets if sheet not in xls.sheet_names
            ]
            if missing:
                raise ValueError(f"Hojas faltantes en el Excel: {', '.join(missing)}")

            dataframes = {}
            for sheet_name in self._excel_sheets:
                print(
                    f"-> Procesando hoja: {sheet_name}"
                )  # key es el nombre real de la hoja
                df = pd.read_excel(
                    xls,
                    sheet_name,
                    header=0,
                    skiprows=None,
                    dtype=str,
                )
                df_cleaned = self._delete_trash_rows(df)

                # Redondear valores decimales a dos decimales
                self._round_decimal_values(df_cleaned)

                df_cleaned = df_cleaned.fillna("")

                dataframes[sheet_name] = df_cleaned

            return dataframes

    @staticmethod
    def _round_decimal_values(df_cleaned: pd.DataFrame) -> None:
        for col in df_cleaned.columns:
            # Intentar convertir a float, si falla, dejar como está
            try:
                df_cleaned[col] = pd.to_numeric(df_cleaned[col])
                if pd.api.types.is_float_dtype(df_cleaned[col]):
                    df_cleaned[col] = df_cleaned[col].round(2)
            except Exception:
                pass
            
    
    def calculate_totals_by_center(
        self, complete_df: pd.DataFrame, df_grupo: pd.DataFrame
    ) -> dict[str, str]:
        """Calcula totales por grupo usando la última fila como referencia.

        df_full: DataFrame completo (incluye fila final con totales globales precalculados).
        df_grupo: Subconjunto filtrado para un valor concreto de group_column.
        """
        if complete_df.empty:
            return {}

        last_row = complete_df.iloc[-1]
        columns_to_sum = [
            col
            for col in complete_df.columns[1:]
            if self._is_valid_column(col, last_row)
        ]

        # Clave descriptiva; si la plantilla aún espera 'EDIFICIO' la mantenemos.
        key_label = "CENTRO" if "CENTRO" in complete_df.columns else "EDIFICIO"
        aggregated_totals: dict[str, str] = {key_label: "Total general"}

        for col in columns_to_sum:
            numeric_col_values = pd.to_numeric(df_grupo[col], errors="coerce").dropna()
            
            if numeric_col_values.empty:
                aggregated_totals[col] = ""
                
            else:
                calculated_sum = numeric_col_values.sum()
                is_approximately_integer = (
                    abs(calculated_sum - round(calculated_sum)) < 1e-6
                )
                
                if is_approximately_integer:
                    aggregated_totals[col] = str(int(round(calculated_sum)))
                else:
                    aggregated_totals[col] = f"{calculated_sum:.2f}".rstrip("0").rstrip(
                        "."
                    )

        return aggregated_totals

    @staticmethod
    def _is_valid_column(col, last_row):
        return pd.notna(last_row[col]) and str(last_row[col]).strip() != ""


@dataclass
class AnexosCreator(DfManager):
    def __init__(self, cee_dir: Path, planos_dir: Path, templete_dir: Path):
        super().__init__()  
        self._constants = self._initialize_constants()
        self._doc_bytes_templete = Path(templete_dir).read_bytes()

    @staticmethod
    def _initialize_constants() -> dict[str, int]:
        """Initialize Word constants."""
        return {
            "WD_GO_TO_PAGE": 1,
            "WD_GO_TO_ABSOLUTE": 1,
            "WD_STORY": 6,
            "WD_STATISTIC_PAGES": 2,
            "WD_ACTIVE_END_PAGE_NUMBER": 3,
            "WD_EXPORT_FORMAT_PDF": 17,
            "WD_EXPORT_OPTIMIZE_FOR_PRINT": 0,
            "WD_EXPORT_ALL_DOCUMENT": 0,
            "WD_EXPORT_DOCUMENT_CONTENT": 0,
            "WD_EXPORT_CREATE_HEADING_BOOKMARKS": 1,
        }

    @staticmethod
    def clean_name(filename):
        """
        Limpia el nombre del archivo eliminando caracteres no válidos y
        tildes para Windows.
        """
        # Caracteres no válidos en Windows: < > : " | ? * \ /
        invalid_chars = '<>:"|?*\\/“”'
        cleaned = filename

        # Reemplazar caracteres no válidos
        for char in invalid_chars:
            cleaned = cleaned.replace(char, "")

        # Reemplazar letras con tilde por su versión sin tilde
        cleaned = unicodedata.normalize("NFD", cleaned)
        cleaned = "".join(c for c in cleaned if unicodedata.category(c) != "Mn")

        # Reemplazar múltiples espacios y guiones bajos consecutivos
        cleaned = re.sub(r"[_\s]+", " ", cleaned).strip()

        # Limitar la longitud (Windows tiene límite de 260 caracteres para la ruta completa)
        if len(cleaned) > 100:  # Dejamos margen para la ruta
            cleaned = cleaned[:100].strip()

        return cleaned



    def close_word_processes(self):
        """Cierra todos los procesos de Word para evitar conflictos de archivos."""
        logger = logging.getLogger(__name__)
        logger.info("Cerrando procesos de Word...")

        try:
            # Intentar cerrar Word de forma elegante si pywin32 está disponible
            closed_elegantly = self._close_open_word_documents()

            if not closed_elegantly:
                # Forzar cierre solo si el cierre elegante falló
                self._force_close_word_processes()

            # Breve pausa para asegurar que los procesos se han cerrado
            time.sleep(0.5)
            logger.info("Sistema listo para generar documentos")

        except Exception as e:
            logger.error(f"Error al cerrar Word: {e}")
            # Intentar forzar cierre como último recurso
            try:
                self._force_close_word_processes()
            except Exception as force_error:
                logger.error(f"Error al forzar cierre de Word: {force_error}")
                raise

    def _force_close_word_processes(self):
        """Fuerza el cierre de todos los procesos de Word."""
        logger = logging.getLogger(__name__)

        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", "winword.exe", "/T"],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,  # No lanzar excepción si returncode != 0
            )

            if result.returncode == 0:
                logger.info("Procesos de Word forzados a cerrar")
            elif result.returncode == 128:  # Process not found
                logger.debug("No había procesos de Word ejecutándose")
            else:
                logger.warning(
                    f"taskkill retornó código {result.returncode}: {result.stderr}"
                )

        except subprocess.TimeoutExpired:
            logger.error("Timeout al intentar cerrar Word forzosamente")
            raise
        except FileNotFoundError:
            logger.error("Comando taskkill no encontrado")
            raise
        except Exception as e:
            logger.error(f"Error inesperado al forzar cierre de Word: {e}")
            raise

    def _close_open_word_documents(self) -> bool:
        """
        Cierra documentos de Word abiertos de forma elegante.

        Returns:
            bool: True si se cerró exitosamente o no había documentos, False si falló
        """
        logger = logging.getLogger(__name__)

        try:
            pythoncom.CoInitialize()

            try:
                word_app = win32_client.GetActiveObject("Word.Application")
                number_of_open_docs = word_app.Documents.Count

                if number_of_open_docs > 0:
                    logger.info(
                        f"Cerrando {number_of_open_docs} documentos abiertos..."
                    )
                    # Cerrar documentos en orden inverso para evitar problemas de índices
                    for i in range(number_of_open_docs, 0, -1):
                        try:
                            doc = word_app.Documents(i)
                            doc.Close(
                                SaveChanges= False)
                        except Exception as e:
                            logger.warning(f"Error cerrando documento {i}: {e}")

                    logger.info("Documentos cerrados correctamente")

                # Cerrar la aplicación
                word_app.Quit(
                    SaveChanges=False
                )
                logger.info("Aplicación Word cerrada correctamente")
                return True

            except Exception as com_error:
                # Si no hay instancia activa o error COM, no es crítico
                logger.debug(f"No se pudo conectar a Word activo: {com_error}")
                return True  # No hay Word abierto, objetivo cumplido

        except Exception as e:
            logger.error(f"Error en cierre elegante de Word: {e}")
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.debug(f"Error al finalizar COM: {e}")

    def get_month_and_year(self) -> tuple[str, int]:
        """Solicita al usuario el mes y año para el documento."""
        print("\n" + "=" * 50)
        print("CONFIGURACIÓN DEL DOCUMENTO")
        print("=" * 50)

        # Solicitar mes
        month_number, month_name = self._request_month_selection()

        # Solicitar año
        year = self._request_year_selection()

        print("\nConfiguración seleccionada:")
        print(f"   Mes: {month_name} ({month_number})")
        print(f"   Año: {year}")
        print("=" * 50)

        return month_name, year

    @staticmethod
    def _request_year_selection():
        current_year = datetime.now().year

        while True:
            try:
                input_year = input(
                    f"Ingrese el año [Enter para usar {current_year}]: "
                ).strip()

                is_empty_input = input_year == ""
                if is_empty_input:
                    year = current_year
                else:
                    year = int(input_year)

                # Usar el año actual como referencia para el rango permitido
                is_year_within_allowed_limits = (
                    current_year - 5 <= year <= current_year + 5
                )
                if is_year_within_allowed_limits:
                    break
                else:
                    print(
                        f"Error: El año debe estar entre {current_year - 5} y {current_year + 5}"
                    )
            except ValueError:
                print("Error: Por favor ingrese un año válido")
        return year

    @staticmethod
    def _request_month_selection() -> tuple[int, str]:
        """Solicita al usuario el mes, con validación y opción de usar el mes actual."""

        current_month = datetime.now().month

        months_in_spanish = {
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
                print(
                    f"\nMes actual: {months_in_spanish[current_month]} ({current_month})"
                )
                input_month_number = input(
                    f"Ingrese el mes (1-12) [Enter para usar {current_month}]: "
                ).strip()

                is_empty_input = input_month_number == ""

                if is_empty_input:
                    month_num = current_month
                else:
                    month_num = int(input_month_number)

                is_valid_number_month = 1 <= month_num <= 12

                if is_valid_number_month:
                    month_name = months_in_spanish[month_num]
                    break
                else:
                    print("Error: El mes debe estar entre 1 y 12")

            except ValueError:
                print("Error: Por favor ingrese un número válido")

        return month_num, month_name

    @staticmethod
    def _find_page_number_by_position(doc, pos: int) -> int | None:
        """
        Devuelve el número de página (1-based) de la posición `pos` en el documento
        """
        WD_ACTIVE_END_PAGE_NUMBER = 3

        try:
            rng = doc.Range(Start=int(pos), End=int(pos) + 1)
            current_end_page_number = int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))

            return current_end_page_number

        except Exception:
            return None

    def _set_of_pages_ranging_from(self, doc, start: int, end: int) -> set[int]:
        """
        Devuelve el conjunto de páginas que cubre el rango [start, end)
        """
        try:
            start_page_number = self._find_page_number_by_position(doc, start)
            end_page_number = self._find_page_number_by_position(
                doc, max(start, end - 1)
            )

            has_invalid_page_references = (
                start_page_number is None or end_page_number is None
            )
            if has_invalid_page_references:
                return set()

            pages_in_range = set(
                range(int(start_page_number), int(end_page_number) + 1)
            )
            return pages_in_range

        except Exception:
            return set()

    def _add_toc_page(self, doc, pages, paragraph, style, txt):
        """
        Añade la página al conjunto si el estilo o el texto contiene "ÍNDICE".
        """
        # Verificar si el estilo contiene "toc"
        contains_toc_style = "Título TDC" in style

        # Verificar si el texto es exactamente "ÍNDICE" (con o sin acentos)
        normalized_txt = self._remove_accents(txt)
        contains_indice = normalized_txt == "indice" or txt == "índice"

        if contains_toc_style or contains_indice:
            page_number = self._find_page_number_by_position(doc, paragraph.Range.Start)

            page_number_exists = page_number is not None
            if page_number_exists:
                pages.add(int(page_number))

    @staticmethod
    def _remove_accents(string: str) -> str:
        try:
            normalized_string = "".join(
                character
                for character in unicodedata.normalize("NFD", string)
                if unicodedata.category(character) != "Mn"
            )
            return normalized_string

        except Exception:
            return string

    def _normalize_string(self, string: str) -> str:
        """
        Normaliza una cadena eliminando acentos, convirtiendo a minúsculas y
        reemplazando espacios múltiples por uno solo.
        """
        if string is None:
            return ""
        string = str(string)
        string = self._remove_accents(string)
        string = string.lower()
        string = re.sub(r"\s+", " ", string).strip()
        return string

    @staticmethod
    def _extract_total_pdf_pages(pdf_path):
        """
        Extrae las páginas de un PDF usando pypdf.
        Retorna una tupla (reader, num_of_pages) donde:
        - reader: objeto lector de PDF
        - num_of_pages: número total de páginas en el PDF
        """
        try:
            reader = PdfReader(str(pdf_path))
            num_of_pages = len(reader.pages)
            return reader, num_of_pages
        except Exception as e:
            raise RuntimeError(f"No se pudo abrir {pdf_path} con pypdf: {e}")

    @staticmethod
    def _get_text_from_pdf_on_page(reader, i):
        # i es 0-based
        try:
            page_text = reader.pages[i].extract_text()
            return page_text or ""
        except Exception:
            return ""

    @staticmethod
    def _is_toc_page(text_norm: str) -> bool:
        PDF_TOC_KEYWORDS = [
            "indice",
            "índice",
            "INDICE",
            "ÍNDICE",
        ]
        contains_toc_keyword = any(k in text_norm for k in PDF_TOC_KEYWORDS)
        return contains_toc_keyword

    def _find_title_page_in_pdf(self, pdf_path, title_text) -> int | None:
        """
        Encuentra la página que contiene el título especificado en un PDF.

        Args:
            pdf_path: Ruta al archivo PDF
            title_text: Texto del título a buscar

        Returns:
            Número de página (1-based) donde se encuentra el título, o None si no se encuentra
        """
        reader, number_of_pages = self._extract_total_pdf_pages(pdf_path)
        normalized_title = self._normalize_string(title_text)

        if not normalized_title:
            return None

        candidate_pages = self._find_candidate_pages(
            reader, number_of_pages, normalized_title
        )

        if not candidate_pages:
            return None

        best_candidate_page = self._select_best_candidate_page(reader, candidate_pages)
        return best_candidate_page

    def _find_candidate_pages(
        self, reader, number_of_pages: int, normalized_title: str
    ) -> list[int]:
        """
        Encuentra todas las páginas que contienen el título normalizado.

        Args:
            reader: Objeto lector de PDF
            number_of_pages: Número total de páginas
            normalized_title: Título normalizado a buscar

        Returns:
            Lista de números de página (1-based) que contienen el título
        """
        candidates = []

        for page_index in range(number_of_pages):
            page_text = self._get_text_from_pdf_on_page(reader, page_index)
            normalized_page_text = self._normalize_string(page_text)

            if normalized_page_text and normalized_title in normalized_page_text:
                candidates.append(page_index + 1)  # Convert to 1-based

        return candidates

    def _select_best_candidate_page(self, reader, candidate_pages: list[int]) -> int:
        """
        Selecciona la mejor página candidata, priorizando páginas que no sean TOC.

        Args:
            reader: Objeto lector de PDF
            candidate_pages: Lista de páginas candidatas (1-based)

        Returns:
            Número de página (1-based) seleccionada como la mejor opción
        """
        non_toc_pages = self._filter_non_toc_pages(reader, candidate_pages)

        if non_toc_pages:
            # Elegir la última ocurrencia fuera del TOC (más robusto)
            return max(non_toc_pages)

        # Si todas parecen TOC, elegir la última en general
        return max(candidate_pages)

    def _filter_non_toc_pages(self, reader, candidate_pages: list[int]) -> list[int]:
        """
        Filtra las páginas candidatas eliminando aquellas que parecen ser páginas de TOC.

        Args:
            reader: Objeto lector de PDF
            candidate_pages: Lista de páginas candidatas (1-based)

        Returns:
            Lista de páginas que NO son TOC
        """
        non_toc_pages = []

        for page_number in candidate_pages:
            page_text = self._get_text_from_pdf_on_page(
                reader, page_number - 1
            )  # Convert to 0-based
            normalized_page_text = self._normalize_string(page_text)

            if not self._is_toc_page(normalized_page_text):
                non_toc_pages.append(page_number)

        return non_toc_pages

    def remove_blank_pages_from_pdf(self, pdf_path: str) -> None:
        """
        Elimina páginas completamente en blanco de un PDF.
        Detecta páginas con muy poco texto o solo espacios en blanco.
        """
        try:
            reader, number_of_pages = self._extract_total_pdf_pages(pdf_path)

            if number_of_pages <= 1:
                return  # No procesar PDFs de una sola página o vacíos

            pages_to_keep = self.collect_significant_pages(reader, number_of_pages)

            # Si no hay cambios, no reescribir
            no_pages_to_remove = len(pages_to_keep) == number_of_pages
            if no_pages_to_remove:
                return

            # Reescribir PDF sin páginas en blanco
            self.write_pdf_without_blank_pages(pdf_path, reader, pages_to_keep)

        except Exception as e:
            print(f"   ! Error eliminando páginas en blanco: {e}")

    @staticmethod
    def write_pdf_without_blank_pages(pdf_path, reader, pages_to_keep):
        writer = PdfWriter()
        for page_index in pages_to_keep:
            writer.add_page(reader.pages[page_index])

        # Escribir a archivo temporal y luego reemplazar
        pdf_path_obj = Path(pdf_path)
        temp_path = pdf_path_obj.with_suffix(".tmp")
        
        with open(temp_path, "wb") as f:
            writer.write(f)

        # Reemplazar original
        temp_path.replace(pdf_path_obj) 

    def collect_significant_pages(self, reader, number_of_pages) -> list[int]:
        pages_to_keep = []
        for page_index in range(number_of_pages):
            text = self._get_text_from_pdf_on_page(reader, page_index).strip()

            # Considerar página no vacía si:
            # - Tiene más de 20 caracteres de texto
            # - O contiene palabras significativas (no solo espacios/puntuación)
            has_significant_text = len(text) > 20
            has_meaningful_word = text and len(text.split()) > 3
            if has_significant_text or has_meaningful_word:
                pages_to_keep.append(page_index)
        
        return pages_to_keep

    def _open_word_document(
        self, docx_path: str, read_only: bool = False
    ) -> tuple[CDispatch, CDispatch]:
        """
        Opens a Word document and returns the application and document objects.

        Args:
            docx_path: Path to the Word document
            read_only: Whether to open the document in read-only mode

        Returns:
            Tuple of (word_app, doc) COM objects

        Raises:
            Exception: If Word application cannot be created or document cannot be opened
        """
        pythoncom.CoInitialize()

        word_app: CDispatch = win32_client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.ScreenUpdating = False
        word_app.DisplayAlerts = False

        doc: CDispatch = word_app.Documents.Open(
            str(docx_path),
            ConfirmConversions=False,
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Visible=False,
        )

        return word_app, doc

    def remove_blank_pages_from_docx(self, docx_path: str) -> int:
        """
        Elimina páginas en blanco de un documento Word usando Word COM.
        Retorna el número de páginas eliminadas.
        """
        pages_removed = 0

        try:
            word_app, doc = self._open_word_document(docx_path, read_only=False)

            try:
                pages_removed = self._process_blank_pages_removal(doc)

                if pages_removed > 0:
                    self._finalize_document_after_page_removal(doc, pages_removed)

                doc.Close(SaveChanges=False)

            except Exception as e:
                logging.error(f"Error processing document: {e}")

            word_app.Quit()

        except Exception as e:
            logging.error(f"Error with Word COM: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        return pages_removed

    def _process_blank_pages_removal(self, doc: CDispatch) -> int:
        """Process the removal of blank pages from the document."""
        # Repaginar para asegurar conteo correcto
        doc.Repaginate()
        total_pages = int(doc.ComputeStatistics(2))  # wdStatisticPages = 2

        print(f"   -> Analizando {total_pages} páginas para eliminar páginas en blanco")

        page_processor = WordPageRemover(doc)
        pages_removed = 0

        # Trabajar de atrás hacia adelante para evitar problemas de renumeración
        for page_num in range(total_pages, 0, -1):
            try:
                if page_processor.is_page_blank(page_num):
                    print(f"   -> Eliminando página en blanco: {page_num}")
                    is_deleted_page = page_processor.delete_page(page_num)
                    if is_deleted_page:
                        pages_removed += 1
                    else:
                        print(f"   ! No se pudo eliminar página {page_num}")
            except Exception as e:
                print(f"   ! Error procesando página {page_num}: {e}")
                continue

        return pages_removed

    def _finalize_document_after_page_removal(
        self, doc: CDispatch, pages_removed: int
    ) -> None:
        """Finalize document after removing blank pages."""
        # Repaginar y actualizar campos después de eliminar páginas
        doc.Repaginate()

        # Actualizar TOC si existe
        try:
            if doc.TablesOfContents.Count > 0:
                self.update_toc(doc)
        except Exception as e:
            print(f"   ! Error actualizando TOC: {e}")

        # Guardar cambios
        doc.Save()
        print(f"   ✓ Eliminadas {pages_removed} páginas en blanco del documento")

    @staticmethod
    def update_toc(doc):
        doc.TablesOfContents(1).Update()

    def export_and_prune_pdf(
        self, docx_path, sections_empty_flags, excel_sheets, final_pdf_path=None
    ):
        """
        1) Exporta DOCX -> PDF temporal.
        2) Detecta páginas de títulos de secciones vacías en el PDF.
        3) Si hay páginas por eliminar: abre el DOCX, borra esas páginas, repagina y actualiza TOC.
        4) Exporta PDF final desde Word (índice correcto) y elimina páginas en blanco.
        """
        docx_path = Path(docx_path)
        if final_pdf_path is None:
            final_pdf_path = docx_path.with_suffix(".pdf")
        final_pdf_path = Path(final_pdf_path)
        tmp_pdf = final_pdf_path.parent / (final_pdf_path.stem + ".tmp.pdf")

        # 1) Exportar con Word a PDF temporal
        self.export_docx_to_temp_pdf(docx_path, tmp_pdf)

        # 2) Detectar páginas a eliminar leyendo el PDF
        try:
            _, num_of_pages = self._extract_total_pdf_pages(tmp_pdf)
        except Exception as e:
            print("   ! Error leyendo PDF temporal:", e)
            # Si falla, dejar el PDF exportado
            try:
                final_pdf_path.unlink(missing_ok=True)
            except Exception:
                pass
            tmp_pdf.replace(final_pdf_path)
            return str(final_pdf_path)

        pages_to_delete = self.find_empty_pages(docx_path, sections_empty_flags, excel_sheets, tmp_pdf, num_of_pages)

        print(f"   -> Total páginas marcadas para eliminar: {sorted(pages_to_delete)}")

        # 3) Si hay páginas que borrar, modificar DOCX y actualizar TOC
        if pages_to_delete:
            # Borramos el PDF temporal; vamos a regenerar desde Word
            try:
                tmp_pdf.unlink()
            except Exception:
                pass

            try:
                word_app, doc = self._open_word_document(docx_path, read_only=False)
                try:
                    doc.Repaginate()
                except Exception:
                    pass

                # Eliminar páginas de forma más simple: página por página en orden descendente
                self.remove_pages_from_doc(num_of_pages, pages_to_delete, doc)

                # Actualizar TOC
                try:
                    self.update_toc(doc)
                except Exception:
                    pass

                # Guardar DOCX actualizado
                try:
                    doc.Save()  # sobreescribe
                except Exception:
                    pass

                # Exportar PDF final con TOC actualizado
                try:
                    self.export_doc_to_pdf(tmp_pdf, doc)
                except Exception as e:
                    print("   ! Error exportando PDF final:", e)

                doc.Close(SaveChanges=False)
                word_app.Quit()
                print(f"   ✓ DOCX actualizado y PDF exportado: {final_pdf_path.name}")

                # Eliminar páginas completamente en blanco del PDF final
                try:
                    self.remove_blank_pages_from_pdf(str(tmp_pdf))
                except Exception as e:
                    print(f"   ! Limpieza de páginas en blanco falló: {e}")

                # Mover el archivo temporal al archivo final
                try:
                    final_pdf_path.unlink(missing_ok=True)
                except Exception:
                    pass
                tmp_pdf.replace(final_pdf_path)

                return str(final_pdf_path)
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        # 4) Si no hay páginas que borrar, conservar el PDF temporal como final
        try:
            final_pdf_path.unlink(missing_ok=True)
        except Exception:
            pass
        tmp_pdf.replace(final_pdf_path)
        print(f"   ✓ PDF generado (sin cambios): {final_pdf_path.name}")

        # Eliminar páginas completamente en blanco
        try:
            self.remove_blank_pages_from_pdf(str(final_pdf_path))
        except Exception as e:
            print(f"   ! Limpieza de páginas en blanco falló: {e}")

        return str(final_pdf_path)

    def remove_pages_from_doc(self, num_of_pages, pages_to_delete, doc):
        try:
            word_remover = WordPageRemover(doc)
                    # Obtener páginas totales del documento
            try:
                total_pages = int(
                            doc.ComputeStatistics(2)
                        )  # 2 = wdStatisticPages
            except Exception:
                total_pages = num_of_pages

            print(f"   -> Documento tiene {total_pages} páginas")

                    # Ordenar páginas a eliminar de mayor a menor para evitar problemas de renumeración
            pages_sorted = sorted(pages_to_delete, reverse=True)

            for page_num in pages_sorted:
                if page_num <= total_pages:
                    print(f"   -> Eliminando página {page_num}")
                    word_remover.delete_page(page_num)
                else:
                    print(
                                f"   ! Página {page_num} fuera de rango (total: {total_pages})"
                            )

        except Exception as e:
            print(f"   ! Error en eliminación de páginas: {e}")

    def find_empty_pages(self, docx_path, sections_empty_flags, excel_sheets, tmp_pdf, num_of_pages):
        pages_to_delete = set()
        for key, is_empty in sections_empty_flags.items():
            if not is_empty:
                continue
            title_text = excel_sheets.get(key)
            if not title_text:
                continue

            print(f"   -> Buscando sección vacía: '{title_text}' (key: {key})")
            title_page_index = self._find_title_page_in_pdf(tmp_pdf, title_text)
            if title_page_index is not None:
                print(
                    f"   -> Encontrada en página {title_page_index}, marcando para eliminar"
                )
                pages_to_delete.add(title_page_index)
                # Siempre intentar eliminar la página siguiente (tabla)
                if title_page_index + 1 <= num_of_pages:
                    pages_to_delete.add(title_page_index + 1)
                    print(f"   -> También página {title_page_index + 1} (tabla)")
            else:
                print(
                    f"   ! No se localizó el título '{title_text}' en el PDF: {docx_path.name}"
                )
                
        return pages_to_delete

    def export_docx_to_temp_pdf(self, docx_path, tmp_pdf):
        try:
            word_app, doc = self._open_word_document(docx_path, read_only=False)
            try:
                doc.Repaginate()
            except Exception:
                pass
            try:
                self.export_doc_to_pdf(tmp_pdf, doc)
            except Exception as e:
                print("   ! Error exportando a PDF temporal:", e)
            finally:
                doc.Close(SaveChanges=False)
                word_app.Quit()
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def export_doc_to_pdf(self, pdf_path, doc):
        doc.ExportAsFixedFormat(
                    OutputFileName=str(pdf_path),
                    ExportFormat=self._constants["WD_EXPORT_FORMAT_PDF"],
                    OpenAfterExport=False,
                    OptimizeFor=self._constants["WD_EXPORT_OPTIMIZE_FOR_PRINT"],
                    Range=self._constants["WD_EXPORT_ALL_DOCUMENT"],
                    Item=self._constants["WD_EXPORT_DOCUMENT_CONTENT"],
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=self._constants[
                        "WD_EXPORT_CREATE_HEADING_BOOKMARKS"
                    ],
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )

    def create_anexo_3(self, excel_path: str, group_column: str = "CENTRO") -> None:
        month_name, year = self.get_month_and_year()

        self.close_word_processes()

        all_dataframes = self.load_and_clean_sheets(excel_path)
        df_clima, df_sist_cc, df_eleva, df_eqhoriz, df_ilum, df_otros_eq = (
                all_dataframes.values()
            )
        total_groups: set[str] = self._extract_unique_groups(group_column, all_dataframes)
        print(
            f"-> Se generarán documentos para {len(total_groups)} {group_column.lower()}s"
        )

        for group in sorted(total_groups):
            dfs_filtered_by_group = self._filter_dataframes_by_group(
                group_column, all_dataframes, group
            )
            
            (
                df_clima_grupo,
                df_sist_cc_grupo,
                df_eleva_grupo,
                df_eqhoriz_grupo,
                df_ilum_grupo,
                df_otros_eq_grupo,
            ) = dfs_filtered_by_group

            # Saltar si todas las tablas están vacías para este centro
            all_dataframes_are_empty = all(
                len(df_) == 0
                for df_ in [
                    df_clima_grupo,
                    df_sist_cc_grupo,
                    df_eleva_grupo,
                    df_eqhoriz_grupo,
                    df_ilum_grupo,
                    df_otros_eq_grupo,
                ]
            )
            if all_dataframes_are_empty:
                continue

            # Calcular totales (reutilizamos función existente). Etiquetamos para indicar centro.
            totales_clima = self.calculate_totals_by_center(df_clima, df_clima_grupo)
            totales_sist_cc = self.calculate_totals_by_center(
                df_sist_cc, df_sist_cc_grupo
            )
            totales_eleva = self.calculate_totals_by_center(df_eleva, df_eleva_grupo)
            totales_eqhoriz = self.calculate_totals_by_center(
                df_eqhoriz, df_eqhoriz_grupo
            )
            totales_ilum = self.calculate_totals_by_center(df_ilum, df_ilum_grupo)
            totales_otros_eq = self.calculate_totals_by_center(
                df_otros_eq, df_otros_eq_grupo
            )

            context = {
                "mes": month_name,
                "anio": year,
                "centro": group,
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

            doc = DocxTemplate(BytesIO(self._doc_bytes_templete))
            doc.render(context)

            # Obtener ID CENTRO del primer registro válido de cualquier DataFrame
            center_id = ""
            center_id = self._extract_center_id(dfs_filtered_by_group)

            try:
                center_name = self.clean_name(str(group))
                output_filename = f"Anexo 3 {center_name}.docx"
                output_path = self._create_output_path(center_id, output_filename)
                doc.save(str(output_path))

                # Eliminar páginas en blanco del documento DOCX recién creado
                print(f"   -> Eliminando páginas en blanco de {output_filename}")
                try:
                    blank_pages_removed = self.remove_blank_pages_from_docx(
                        str(output_path)
                    )
                    if blank_pages_removed > 0:
                        print(
                            f"   ✓ {blank_pages_removed} páginas en blanco eliminadas de {output_filename}"
                        )
                    else:
                        print(
                            f"   ✓ No se encontraron páginas en blanco en {output_filename}"
                        )
                except Exception as e:
                    print(
                        f"   ! Error eliminando páginas en blanco de {output_filename}: {e}"
                    )

                # --- MODO PDF: exportar y eliminar páginas (título + tabla) de secciones vacías ---
                sections_empty = {
                    "Clima": df_clima_grupo.empty,
                    "SistCC": df_sist_cc_grupo.empty,
                    "Eleva": df_eleva_grupo.empty,
                    "EqHoriz": df_eqhoriz_grupo.empty,
                    "Ilum": df_ilum_grupo.empty,
                    "OtrosEq": df_otros_eq_grupo.empty,
                }

                print(
                    f"   -> Secciones vacías para {group}: {[k for k, v in sections_empty.items() if v]}"
                )

                try:
                    self.export_and_prune_pdf(
                        output_path, sections_empty, self._excel_sheets
                    )
                except Exception as e:
                    print(
                        f"   ! No se pudo generar PDF limpio en {output_filename}: {e}"
                    )

                print(f"* Documento generado: {center_id}/{output_filename}")
            except PermissionError as e:
                print(f"   ! Error de permisos con {output_filename}: {e}")
                print("   ! Saltando este archivo...")
                continue
            except Exception as e:
                print(f"   ! Error inesperado con {output_filename}: {e}")
                continue

    def _extract_center_id(self, df_group_list: list[pd.DataFrame]) -> str:
        """
        Extracts the 'ID CENTRO' from the first non-empty DataFrame in the list.
        """
        center_id = ""
        for df_centro in df_group_list:
            if not df_centro.empty:
                for _, row in df_centro.iterrows():
                    has_center_id = pd.notna(row.get("ID CENTRO"))
                        
                    if has_center_id:
                        center_id = str(row.get("ID CENTRO", ""))
                        break
                if center_id:
                    break
        return center_id

    def _create_output_path(self, id_centro, output_file):
        """
        Crea el directorio de salida y devuelve la ruta completa del archivo de salida.
        """
        base_dir = Path(__file__).resolve().parent
        output_dir = base_dir.parent / "word" / "anexos" / id_centro
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / output_file
        return output_path

    @staticmethod
    def _filter_dataframes_by_group(
        group_column: str, all_dataframes: dict[str, pd.DataFrame], group: str
    ) -> list[pd.DataFrame]:
        df_clima = all_dataframes["Clima"]
        df_sist_cc = all_dataframes["SistCC"]
        df_eleva = all_dataframes["Eleva"]
        df_eqhoriz = all_dataframes["EqHoriz"]
        df_ilum = all_dataframes["Ilum"]
        df_otros = all_dataframes["OtrosEq"]

        # Filtrar por centro (las filas totales globales no suelen tener CENTRO, por lo que no se incluyen)
        df_clima_grupo = df_clima[df_clima.get(group_column) == group].copy()
        df_sist_cc_grupo = df_sist_cc[df_sist_cc.get(group_column) == group].copy()
        df_eleva_grupo = df_eleva[df_eleva.get(group_column) == group].copy()
        df_eqhoriz_grupo = df_eqhoriz[df_eqhoriz.get(group_column) == group].copy()
        df_ilum_grupo = df_ilum[df_ilum.get(group_column) == group].copy()
        df_otros_eq_grupo = df_otros[df_otros.get(group_column) == group].copy()

        dfs_filtered_by_group = [
            df_clima_grupo,
            df_sist_cc_grupo,
            df_eleva_grupo,
            df_eqhoriz_grupo,
            df_ilum_grupo,
            df_otros_eq_grupo,
        ]
        return dfs_filtered_by_group

    @staticmethod
    def _extract_unique_groups(
        group_column: str, all_dataframes: dict[str, pd.DataFrame]
    ) -> set[str]:
        """
        Extracts a set of unique, non-empty group names from the specified column across multiple DataFrames.

        Args:
            group_column: The name of the column to extract unique values from.
            all_dataframes: A dictionary of DataFrames to process.

        Returns:
            A set of unique, non-empty group names.
        """
        unique_groups = set()

        for df in all_dataframes.values():
            if group_column in df.columns:
                group_values = df[group_column].dropna().unique()
                cleaned_values = {
                    str(value).strip() for value in group_values if str(value).strip()
                }
                unique_groups.update(cleaned_values)

        return unique_groups

    def create_anexo_2(self, excel_path: str, group_column: str = "CENTRO") -> None:
        month_name, year = self.get_month_and_year()

        self.close_word_processes()

        all_dataframes = self.load_and_clean_sheets(excel_path)

        total_groups: set[str] = self._extract_unique_groups(group_column, all_dataframes)

        print("* Datos cargados y limpiados")

        # Crear contexto para la plantilla
        print("-> Renderizando documentos...")

        generated_docs = []

        for group in sorted(total_groups):
            dfs_filtered_by_group = self._filter_dataframes_by_group(
                group_column, all_dataframes, group
            )
            
            df_conta_grupo = dfs_filtered_by_group[0]

            # Saltar si todas las tablas están vacías para este centro
            all_dataframes_are_empty = all(
                len(df_) == 0
                for df_ in [df_conta_grupo]
            )
            if all_dataframes_are_empty:
                continue

            context = {
                "mes": month_name,
                "anio": year,
                "centro": group,
                "df_conta": df_conta_grupo.to_dict("records"),
                "tipo_de_suministro": df_conta_grupo["SUMINISTRO"].unique(),
            }

            doc = DocxTemplate(BytesIO(self._doc_bytes_templete))
            doc.render(context)

            center_id = self._extract_center_id(dfs_filtered_by_group)
            
            try:
                # Crear nombre de archivo con el nombre del centro limpio
                nombre_centro = self.clean_name(group)
                output_filename = f"Anexo 2 {nombre_centro}.docx"
                output_path = self._create_output_path(center_id, output_filename)

                doc.save(str(output_path))
                # print(f"* Documento generado: {center_id}/{output_filename}")

                generated_docs.append(str(output_path))

            except PermissionError as e:
                print(f"   ! Error de permisos con {output_filename}: {e}")
                print("   ! Saltando este archivo...")
                continue
            except Exception as e:
                print(f"   ! Error inesperado con {output_filename}: {e}")
                continue

if __name__ == "__main__":
    anexos_creator = AnexosCreator('as', 'asd', r'C:\Users\ferma\Documents\repos\artecoin_automatizaciones\word\anexos\Plantilla_Anexo_3.docx')
    anexos_creator.create_anexo_3(r"C:\Users\ferma\Documents\repos\artecoin_automatizaciones\excel\proyecto\ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1.xlsx")
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel-dir", required=True)
    parser.add_argument("--word-dir")
    parser.add_argument("--mode", choices=["all", "single"], required=True)
    parser.add_argument("--anexo", type=int)
    args = parser.parse_args()
    