# -*- coding: utf-8 -*-
"""
script_carga_sonigeo.py

Carga datos en la hoja 'Cont' del libro de Artecoin desde una carpeta raíz
con subcarpetas por centro (formato: Cxxxx_NOMBRE/Cxxxx_NOMBRE.xlsx[x/m]).
Arquitectura limpia + SOLID, sin pandas. Diseñado para ejecutarse vía xlwings.

Punto de entrada recomendado desde Excel:
    RunPython("import script_carga_sonigeo; script_carga_sonigeo.main()")
"""
from __future__ import annotations

import json
import logging
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Iterable, Any, Set

import xlwings as xw


# --------------- Configuración de logging ---------------
logger = logging.getLogger("sonigeo_loader")
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter("%(levelname)s - %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)


# --------------- Utilidades de nombres / normalización ---------------
def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = _strip_accents(s)
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def first_nonempty(values: Iterable[str]) -> Optional[str]:
    for v in values:
        if v and str(v).strip():
            return str(v).strip()
    return None


# --------------- Modelos de datos ---------------
SHEETS_EXPECTED = [
    "DATOS_ELECTRICOS_EDIFICIOS",
    "CERRAMIENTOS",
    "CENTRO",
    "EDIFICIO",
    "DEPENDENCIA",
    "EQ HORIZ",
    "ILUM",
    "EQ GRUPBOM",
    "OTROS EQUIPOS",
    "EQ CLIMA EXT",
    "EQ CLIMA INT",
    "EQ GEN",
    "EQ ELEV"
]

# Mapeo por defecto hoja → tabla (puedes ajustarlo a tus nombres reales)
DEFAULT_SONIGEO_MAP = {
    "DATOS_ELECTRICOS_EDIFICIOS": "Tabla32",
    "CERRAMIENTOS": "Tabla37",
    "CENTRO": "Tabla34",
    "EDIFICIO": "Tabla35",
    "DEPENDENCIA": "Tabla36",
    "EQ HORIZ": "Tabla42",
    "ILUM": "Tabla45",
    "EQ GRUPBOM": "Tabla43",
    "OTROS EQUIPOS": "Tabla44",
    "EQ CLIMA EXT": "Tabla46",
    "EQ CLIMA INT": "Tabla40",
    "EQ GEN": "Tabla38",
    "EQ ELEV": "Tabla41"
}

CENTER_ID_COL_CANDIDATES_NORM = {
    "id centro", 
    "centro id", 
    "id_centro", 
    "centro_id",
    }


@dataclass
class SourceBlock:
    """Acumulación de filas por tipo de hoja de Sonigeo en memoria."""
    header_original: List[str]
    header_norm: List[str]
    rows: List[Dict[str, Any]] = field(default_factory=list)  # dict por 'col_norm' -> value


@dataclass
class WritePlan:
    """Plan de escritura por tabla destino."""
    table_name: str
    dest_headers: List[str]
    dest_headers_norm: List[str]
    protected_cols_norm: Set[str]
    formula_cols_indices: Set[int]
    id_center_col_idx: Optional[int]


# --------------- Infraestructura Excel (Gateway) ---------------
class ExcelGateway:
    def __init__(self, book: xw.Book):
        self.book = book
        self.app = book.app

    # ---- Helpers de Nombres definidos / FolderPicker ----
    def get_defined_text(self, name: str) -> Optional[str]:
        try:
            n = self.book.names[name]
        except KeyError:
            return None
        try:
            val = n.refers_to_range.value
            if isinstance(val, (int, float)):
                val = str(val)
            return (val or "").strip() or None
        except Exception:
            # Puede ser fórmula a texto -> intentar evaluar
            try:
                refers_to = n.refers_to or ""
                # = "texto"
                m = re.match(r'^="(.*)"$', refers_to)
                if m:
                    return m.group(1)
            except Exception:
                pass
        return None

    def choose_folder(self, title: str = "Seleccionar carpeta raíz Sonigeo") -> Optional[str]:
        # 4 = msoFileDialogFolderPicker
        fd = self.app.api.FileDialog(4)
        fd.Title = title
        fd.AllowMultiSelect = False
        if fd.Show() == -1:
            return fd.SelectedItems(1)
        return None

    # ---- Accesso a hoja y tablas ----
    def get_cont_sheet(self) -> xw.Sheet:
        try:
            return self.book.sheets["Cont"]
        except KeyError:
            raise RuntimeError("No se encuentra la hoja 'Cont' en el libro de Artecoin.")

    def list_tables_in_sheet(self, sheet: xw.Sheet) -> Dict[str, Any]:
        """Devuelve dict nombre_listobject(case-insens) -> api ListObject"""
        d: Dict[str, Any] = {}
        for lo in sheet.api.ListObjects:
            name = str(lo.Name)
            d[norm(name)] = lo
        return d

    def read_table_headers(self, lo) -> List[str]:
        rng = lo.HeaderRowRange
        vals = list(rng.Value[0]) if rng.Rows.Count == 1 else [c.Value for c in rng.Columns]
        return [str(v).strip() if v is not None else "" for v in vals]

    def is_total_shown(self, lo) -> bool:
        try:
            return bool(lo.ShowTotals)
        except Exception:
            return False

    def set_total_shown(self, lo, show: bool) -> None:
        try:
            lo.ShowTotals = bool(show)
        except Exception:
            pass

    def get_table_row_count(self, lo) -> int:
        try:
            dbr = lo.DataBodyRange
            if dbr is None:
                return 0
            return int(dbr.Rows.Count)
        except Exception:
            return 0

    def add_table_rows(self, lo, n_rows: int) -> None:
        for _ in range(n_rows):
            lo.ListRows.Add()

    def delete_table_rows_from_bottom(self, lo, n_rows: int) -> None:
        for _ in range(n_rows):
            cnt = self.get_table_row_count(lo)
            if cnt <= 0:
                break
            lo.ListRows(cnt).Delete()

    def write_column_values(self, lo, col_index_1based: int, values: List[Any]) -> None:
        """Escribe lista de valores en la columna col_index_1based del DataBodyRange."""
        dbr = lo.DataBodyRange
        if dbr is None:
            raise RuntimeError("La tabla no tiene DataBodyRange al escribir.")
        col_rng = dbr.Columns(col_index_1based)
        # Formar matriz 2D para Excel
        matrix = [[v] for v in values] if values else []
        if matrix:
            col_rng.Value = matrix

    def read_column_has_formula(self, lo, col_index_1based: int) -> bool:
        """Detecta si la columna tiene fórmula en alguna fila existente (si la hay)."""
        dbr = lo.DataBodyRange
        if dbr is None:
            return False
        try:
            col_rng = dbr.Columns(col_index_1based)
            # Revisar primera fila con datos
            cell = col_rng.Cells(1, 1)
            return bool(cell.HasFormula)
        except Exception:
            return False

    def read_column_values(self, lo, col_index_1based: int) -> List[Any]:
        dbr = lo.DataBodyRange
        if dbr is None:
            return []
        col_rng = dbr.Columns(col_index_1based)
        vals = col_rng.Value
        if isinstance(vals, tuple):
            vals = list(vals)
        # Convertir 2D -> 1D
        if vals and isinstance(vals[0], (list, tuple)):
            return [row[0] for row in vals]
        return vals or []


# --------------- Lector de orígenes Sonigeo ---------------
class SonigeoReader:
    def __init__(self, app: xw.App):
        self.app = app

    @staticmethod
    def infer_center_id_from_folder(folder: Path) -> Optional[str]:
        # Espera nombres tipo: C0007_AYUNTAMIENTO
        m = re.search(r"(C\d{4,})", folder.name, flags=re.IGNORECASE)
        if not m:
            return None
        return m.group(1).upper()

    def discover_centers(self, root: Path) -> List[Tuple[str, Path]]:
        """Devuelve lista de (center_id, path_xlsx)"""
        centers: List[Tuple[str, Path]] = []
        for sub in sorted(p for p in root.iterdir() if p.is_dir()):
            center_id = self.infer_center_id_from_folder(sub)
            if not center_id:
                logger.warning(f"   ! Carpeta ignorada (no ID centro): {sub}")
                continue
            # Buscar excel con el mismo nombre que la carpeta
            expected = sub / f"{sub.name}.xlsx"
            expectedm = sub / f"{sub.name}.xlsm"
            path = expected if expected.exists() else expectedm if expectedm.exists() else None
            if not path:
                # Si no coincide nombre exacto, coger el primer .xlsx/xlsm
                candidates = list(sub.glob("*.xls*"))
                if not candidates:
                    logger.warning(f"   ! No se encontró Excel para {center_id} en {sub}")
                    continue
                path = candidates[0]
            centers.append((center_id, path))
        return centers

    def read_center_book(self, path: Path) -> xw.Book:
        # Abrimos en modo solo lectura, sin actualización de vínculos
        return self.app.books.open(str(path), update_links=False, read_only=True)

    @staticmethod
    def _sheet_used_values(sh: xw.Sheet) -> Optional[List[List[Any]]]:
        try:
            ur = sh.used_range
            if ur is None or ur.value is None:
                return None
            vals = ur.value
            if vals is None:
                return None
            if not isinstance(vals, list):
                # Un único valor en una celda
                return [[vals]]
            # Asegurar matriz 2D
            if vals and not isinstance(vals[0], list):
                # Lista 1D -> 2D una columna
                return [[v] for v in vals]
            return vals
        except Exception:
            return None

    def read_sheets(self, book: xw.Book, center_id: str) -> Dict[str, SourceBlock]:
        """Lee todas las hojas esperadas y devuelve un dict hoja->SourceBlock."""
        res: Dict[str, SourceBlock] = {}
        for sheet_name in SHEETS_EXPECTED:
            try:
                sh = book.sheets[sheet_name]
            except Exception:
                logger.warning(f"   ! Hoja '{sheet_name}' no existe en {center_id}, se ignora.")
                continue
            matrix = self._sheet_used_values(sh)
            if not matrix or len(matrix) < 2:
                logger.warning(f"   ! Hoja '{sheet_name}' en {center_id} sin datos, se ignora.")
                continue

            header = [str(h or "").strip() for h in matrix[0]]
            header_norm = [norm(h) for h in header]

            # Construir filas como dict por col_norm -> value
            rows: List[Dict[str, Any]] = []
            for row in matrix[1:]:
                if not any(c is not None and str(c).strip() for c in row):
                    continue
                d: Dict[str, Any] = {}
                for i, colname_norm in enumerate(header_norm):
                    val = row[i] if i < len(row) else None
                    d[colname_norm] = val
                # Inyectar ID_CENTRO normalizado si luego lo necesita el destino
                d["id centro"] = center_id
                rows.append(d)

            if not rows:
                continue

            if sheet_name not in res:
                res[sheet_name] = SourceBlock(header_original=header, header_norm=header_norm, rows=rows)
            else:
                # Acumular (no debería repetirse hoja por centro, pero por robustez): unir filas
                res[sheet_name].rows.extend(rows)
        return res


# --------------- Column mapping ---------------
class ColumnMapper:
    @staticmethod
    def build_row_matrix_for_dest(dest_headers: List[str],
                                  dest_headers_norm: List[str],
                                  protected_cols_norm: Set[str],
                                  formula_cols_indices: Set[int],
                                  source_block: SourceBlock) -> List[List[Any]]:
        """Transforma los dicts de source (por col_norm) en listas alineadas a las columnas destino."""
        matrix: List[List[Any]] = []
        for row_dict in source_block.rows:
            out_row: List[Any] = []
            for j, dest_col_norm in enumerate(dest_headers_norm, start=1):
                if dest_col_norm in protected_cols_norm or j in formula_cols_indices:
                    out_row.append(None)
                    continue
                value = row_dict.get(dest_col_norm, None)
                out_row.append(value)
            matrix.append(out_row)
        return matrix

    @staticmethod
    def jaccard_similarity(a: Iterable[str], b: Iterable[str]) -> float:
        sa = set(a)
        sb = set(b)
        if not sa and not sb:
            return 1.0
        if not sa or not sb:
            return 0.0
        inter = len(sa & sb)
        union = len(sa | sb)
        return inter / max(1, union)


# --------------- Escritor de tablas Cont ---------------
class TableWriter:
    def __init__(self, gateway: ExcelGateway):
        self.gw = gateway

    def _detect_protected_columns(self) -> Set[str]:
        txt = self.gw.get_defined_text("ProtectedColumns_Cont")
        if not txt:
            return set()
        cols = [norm(c) for c in re.split(r"[;,|]", txt) if c.strip()]
        return set(cols)

    def _build_write_plan(self, lo, table_name: str) -> WritePlan:
        dest_headers = self.gw.read_table_headers(lo)
        dest_headers_norm = [norm(h) for h in dest_headers]
        protected = self._detect_protected_columns()

        # Detectar columnas con fórmula (si hay filas actuales)
        formula_cols: Set[int] = set()
        row_count = self.gw.get_table_row_count(lo)
        if row_count > 0:
            for idx in range(1, len(dest_headers_norm) + 1):
                if self.gw.read_column_has_formula(lo, idx):
                    formula_cols.add(idx)

        # Columna de ID centro (para modo REPLACE por centro)
        id_idx: Optional[int] = None
        for idx, hnorm in enumerate(dest_headers_norm, start=1):
            if hnorm in CENTER_ID_COL_CANDIDATES_NORM:
                id_idx = idx
                break

        return WritePlan(
            table_name=table_name,
            dest_headers=dest_headers,
            dest_headers_norm=dest_headers_norm,
            protected_cols_norm=protected,
            formula_cols_indices=formula_cols,
            id_center_col_idx=id_idx
        )

    def _delete_existing_centers(self, lo, id_col_idx: int, centers_to_replace: Set[str]) -> int:
        """Borra filas cuyo ID_CENTRO esté en centers_to_replace. Devuelve nº filas borradas."""
        if not centers_to_replace:
            return 0
        vals = self.gw.read_column_values(lo, id_col_idx)
        if not vals:
            return 0
        # Borrar de abajo arriba
        deleted = 0
        for i in range(len(vals), 0, -1):
            v = vals[i - 1]
            vtxt = str(v).strip().upper() if v is not None else ""
            if vtxt in centers_to_replace:
                lo.ListRows(i).Delete()
                deleted += 1
        if deleted:
            logger.info(f"      - Borradas {deleted} fila(s) previas por ID_CENTRO en '{lo.Name}'.")
        return deleted

    def write_matrix(self, lo, plan: WritePlan, matrix: List[List[Any]], centers_in_batch: Set[str]) -> None:
        # Gestionar totales
        totals_before = self.gw.is_total_shown(lo)
        if totals_before:
            self.gw.set_total_shown(lo, False)

        try:
            # Modo REPLACE por centro si hay columna identificadora
            if plan.id_center_col_idx is not None and centers_in_batch:
                self._delete_existing_centers(lo, plan.id_center_col_idx, centers_in_batch)

            # Redimensionar filas
            current = self.gw.get_table_row_count(lo)
            target = len(matrix)
            if target > current:
                self.gw.add_table_rows(lo, target - current)
            elif target < current:
                self.gw.delete_table_rows_from_bottom(lo, current - target)

            # Escribir por columnas (respetando protegidas y fórmulas)
            if target > 0:
                # Transponer por columnas
                cols_count = len(plan.dest_headers_norm)
                for j in range(cols_count):
                    if (plan.dest_headers_norm[j] in plan.protected_cols_norm) or ((j + 1) in plan.formula_cols_indices):
                        continue
                    col_vals = [matrix[i][j] if j < len(matrix[i]) else None for i in range(target)]
                    self.gw.write_column_values(lo, j + 1, col_vals)

        finally:
            # Restaurar totales
            self.gw.set_total_shown(lo, totals_before)


# --------------- Caso de uso: orquestación end-to-end ---------------
class LoadSonigeoIntoContUseCase:
    def __init__(self, gateway: ExcelGateway, reader: SonigeoReader, writer: TableWriter):
        self.gw = gateway
        self.reader = reader
        self.writer = writer

    def _load_mapping(self) -> Dict[str, str]:
        txt = self.gw.get_defined_text("SonigeoMap_JSON_Cont")
        if not txt:
            logger.info("• Usando mapeo hoja→tabla por defecto.")
            return DEFAULT_SONIGEO_MAP.copy()
        try:
            mp = json.loads(txt)
            logger.info("• Usando mapeo hoja→tabla desde 'SonigeoMap_JSON_Cont'.")
            return {k: v for k, v in mp.items()}
        except Exception as e:
            logger.warning(f"   ! Error leyendo SonigeoMap_JSON_Cont: {e}. Se usa mapeo por defecto.")
            return DEFAULT_SONIGEO_MAP.copy()

    def _autodetect_table_for_sheet(
        self,
        cont_sheet: xw.Sheet,
        tables_by_name: Dict[str, Any],
        source_block: SourceBlock
    ) -> Optional[str]:
        """Heurística: elige la tabla con mayor intersección de cabeceras."""
        best_name = None
        best_score = 0.0
        for tname_norm, lo in tables_by_name.items():
            dest_headers = self.gw.read_table_headers(lo)
            dest_norm = [norm(h) for h in dest_headers]
            score = ColumnMapper.jaccard_similarity(dest_norm, source_block.header_norm)
            if score > best_score:
                best_score = score
                best_name = tname_norm
        if best_name and best_score >= 0.45:  # umbral razonable
            logger.info(f"   · Autodetectada tabla '{best_name}' (score={best_score:.2f}).")
            return best_name
        logger.warning("   ! No se pudo autodetectar tabla destino de forma fiable (score bajo).")
        return None

    def run(self) -> None:
        # 1) Determinar ruta Sonigeo
        root_txt = self.gw.get_defined_text("Ruta_SonigeoRoot")
        if not root_txt:
            logger.info("• No hay valor en 'Ruta_SonigeoRoot'. Abriendo selector de carpeta...")
            root_txt = self.gw.choose_folder("Selecciona la carpeta raíz de Sonigeo")
        if not root_txt:
            raise RuntimeError("No se ha proporcionado la carpeta raíz de Sonigeo.")
        root = Path(root_txt)
        if not root.exists():
            raise RuntimeError(f"La carpeta especificada no existe: {root}")

        logger.info(f"• Carpeta raíz Sonigeo: {root}")

        # 2) Descubrir centros
        centers = self.reader.discover_centers(root)
        if not centers:
            raise RuntimeError("No se encontraron centros válidos en la carpeta de Sonigeo.")
        logger.info(f"• Centros detectados: {len(centers)}")

        # 3) Leer todos los exceles y acumular por hoja
        aggregated: Dict[str, SourceBlock] = {}
        centers_in_batch: Set[str] = set()

        # Mejorar rendimiento de Excel:
        app = self.gw.app
        calc_prev = app.api.Calculation
        scr_prev = app.screen_updating
        events_prev = app.enable_events
        app.api.Calculation = -4135  # xlCalculationManual
        app.screen_updating = False
        app.enable_events = False

        try:
            for center_id, xlsx_path in centers:
                logger.info(f"· Leyendo centro {center_id}: {xlsx_path.name}")
                centers_in_batch.add(center_id)
                book_src = None
                try:
                    book_src = self.reader.read_center_book(xlsx_path)
                    sb_map = self.reader.read_sheets(book_src, center_id)
                    for sheet_name, sb in sb_map.items():
                        if sheet_name not in aggregated:
                            aggregated[sheet_name] = SourceBlock(
                                header_original=sb.header_original,
                                header_norm=sb.header_norm,
                                rows=[]
                            )
                        aggregated[sheet_name].rows.extend(sb.rows)
                except Exception as e:
                    logger.error(f"   ! Error leyendo {xlsx_path}: {e}")
                finally:
                    if book_src:
                        try:
                            book_src.close()
                        except Exception:
                            pass
            if not aggregated:
                logger.warning("No se recopiló ningún dato desde Sonigeo.")
                return

            # 4) Preparar escritura en 'Cont'
            cont = self.gw.get_cont_sheet()
            tables = self.gw.list_tables_in_sheet(cont)
            sheet_map = self._load_mapping()  # hoja -> nombre tabla (idealmente)

            writer = self.writer

            # 5) Por cada hoja agregada, localizar tabla destino y escribir
            for sheet_name, sblock in aggregated.items():
                desired_table_name = sheet_map.get(sheet_name)
                lo = None
                lo_key = None

                if desired_table_name:
                    lo_key = norm(desired_table_name)
                    lo = tables.get(lo_key)
                    if not lo:
                        logger.warning(f"   ! Tabla '{desired_table_name}' no existe en 'Cont'. Intentando autodetección...")
                if lo is None:
                    # Autodetectar por intersección de cabeceras
                    auto_key = self._autodetect_table_for_sheet(cont, tables, sblock)
                    if auto_key:
                        lo = tables[auto_key]
                        lo_key = auto_key

                if lo is None:
                    logger.warning(f"   ! No se encontró tabla destino para hoja '{sheet_name}'. Se omite.")
                    continue

                # Plan de escritura
                plan = writer._build_write_plan(lo, table_name=lo.Name)

                # Construir matriz alineada a destino
                matrix = ColumnMapper.build_row_matrix_for_dest(
                    dest_headers=plan.dest_headers,
                    dest_headers_norm=plan.dest_headers_norm,
                    protected_cols_norm=plan.protected_cols_norm,
                    formula_cols_indices=plan.formula_cols_indices,
                    source_block=sblock
                )

                logger.info(f"· Escribiendo {len(matrix)} fila(s) en tabla '{plan.table_name}' (hoja '{sheet_name}')...")
                writer.write_matrix(lo, plan, matrix, centers_in_batch)

        finally:
            # Restaurar ajustes de Excel
            try:
                app.api.Calculation = calc_prev
            except Exception:
                pass
            try:
                app.screen_updating = scr_prev
            except Exception:
                pass
            try:
                app.enable_events = events_prev
            except Exception:
                pass

        logger.info("• Carga Sonigeo finalizada.")


# --------------- Punto de entrada ---------------
def main():
    try:
        # Ejecutado desde Excel
        book = xw.Book.caller()
    except Exception:
        # Fallback: si se ejecuta desde un intérprete externo, usar el libro activo.
        app = xw.App(visible=False)
        try:
            if not app.books:
                raise RuntimeError("No hay libro abierto para operar.")
            book = app.books.active
        except Exception as e:
            logger.error(f"No se puede obtener el libro activo: {e}")
            raise

    gw = ExcelGateway(book)
    reader = SonigeoReader(book.app)
    writer = TableWriter(gw)
    use_case = LoadSonigeoIntoContUseCase(gw, reader, writer)

    try:
        use_case.run()
    except Exception as e:
        logger.error(f"Error inesperado: {e}")
        raise
