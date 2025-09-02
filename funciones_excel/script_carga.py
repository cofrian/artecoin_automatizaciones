# -*- coding: utf-8 -*-
"""
script_carga.py

Carga datos en la hoja activa (CM, Sec, Via o ListCP) del libro Excel que llama
al script vía xlwings, leyendo ÚNICAMENTE la ruta del archivo origen que
corresponde a dicha hoja. No borra nada que esté por debajo de la tabla:
redimensiona la tabla con ListRows.Add/Delete, preservando el contenido inferior.

Protecciones:
- No sobrescribe columnas con fórmula (HasFormula).
- No sobrescribe columnas listadas como protegidas:
  * Definidas en el script (CONFIG_BY_SHEET.protected_dest_columns).
  * Definidas en Excel como nombres: ProtectedColumns_<Hoja> o ProtectedCols_<Hoja>.
- Las columnas protegidas de CONFIG se filtran:
  * Deben existir en la tabla destino (match relajado por nombre).
  * Si hay mapeo explícito para la hoja, además deben existir en el mapeo.

Fila de Totales:
- Se desactiva antes de redimensionar/escribir y se restaura después.

Requisitos:
- Windows + Office 365 + xlwings
- Ejecutar desde Excel con:
    RunPython("import script_carga; script_carga.main()")
  (o compatibilidad)
    RunPython("import script_carga; script_carga.cargar_datos()")
"""
from __future__ import annotations

import json
import os
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple, Set

import xlwings as xw


# --------------------------- Utilidades de nombres --------------------------- #

def _strip_accents_lower(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def _norm(s: str) -> str:
    """Normaliza para comparar cabeceras: minúsculas, sin acentos y solo alfanumérico."""
    if s is None:
        return ""
    s = _strip_accents_lower(str(s))
    return "".join(ch for ch in s if ch.isalnum())

_STOPWORDS = {"de", "del", "la", "el", "los", "las", "y", "en", "por", "para", "con", "sin", "al", "a"}

def _norm_relaxed(s: str) -> str:
    """
    Normalización relajada para emparejar cabeceras humanas vs técnicas:
    - minúsculas, sin acentos
    - elimina guiones/underscores convirtiéndolos en espacios
    - quita stopwords frecuentes
    - deja solo alfanumérico al final
    """
    if s is None:
        return ""
    s = _strip_accents_lower(str(s)).replace("_", " ").replace("-", " ")
    tokens = [t for t in s.split() if t and t not in _STOPWORDS]
    s2 = "".join(tokens)
    return "".join(ch for ch in s2 if ch.isalnum())


# ------------------------------- Configuración --------------------------------#

@dataclass(frozen=True)
class SheetConfig:
    """
    Config por hoja destino.
    - path_names: posibles nombres definidos para la ruta del libro origen.
    - source_sheet_candidates: posibles nombres de hoja en el libro origen.
    - dest_table_name: nombre de la tabla destino (opcional; si None, se usa la primera).
    - explicit_column_map: mapeo explícito DESTINO->ORIGEN (opcional).
    - protected_dest_columns: lista de cabeceras de destino a NO sobrescribir jamás (humanas).
    """
    path_names: Sequence[str]
    source_sheet_candidates: Sequence[str]
    dest_table_name: Optional[str] = None
    explicit_column_map: Optional[Dict[str, str]] = None
    protected_dest_columns: Optional[Sequence[str]] = None


# ----- Mapeos (ORIGEN -> DESTINO). Se invertirán para usar DESTINO->ORIGEN. ----- #

RAW_MAP_CM_SRC_TO_DEST: Dict[str, str] = {
    'Coord.X (m)': 'coord.X (m)',
    'Coord.Y (m)': 'coord.Y (m)',
    'id_centro_':   'id_centro_mando',
    'descripcio':   'descripción',
    'vial':         'id_vial',
    'Tipo_via':     'tipo_vía',
    'Nombre_via':   'nombre_vía',
    'numero':       'medido',
    'localizaci':   'localización',
    'modulo_med':   'módulo_medida',
    'estado':       'estado',
    'tension':      'tensión',
    'tipo_regul':   'tipo_regulación',
    'marca_regu':   'marca_regulación',
    'celula':       'célula',
    'tipo_reloj':   'tipo_reloj',
    'marca_relo':   'marca_reloj',
    'interrupto':   'interruptor_manual',
    'marca_inte':   'marca_interruptor_manual',
    'tipo_teleg':   'tipo_telegestión',
    'marca_tele':   'marca_telegestión',
    'observacio':   'observaciones',
    'marca_celu':   'marca_célula',
    'medido':       'medido',
}

RAW_MAP_SEC_SRC_TO_DEST: Dict[str, str] = {
  "id_seccion": "id_seccion",
  "nombre": "nombre_vía",
  "centro_man": "centro_mando",
  "zona": "sección",                           # ← revisar
  "clase_alum": "CLASIFICACIÓN ALUMBRADO",
  "acera_1": "acera_1",
  "aparcamien": "aparcamiento_1",
  "calzada_1": "calzada_1",
  "s1": "SUPERFICIE ESTUDIO",                  # ← revisar
  "mediana": "mediana",
  "s2": "SECCIÓN EQUIVALENTE",                 # ← revisar
  "calzada_2": "calzada_2",
  "aparcami_1": "aparcamiento_2",
  "acera_2": "acera_2",
  "altura": "altura",
  "interdista": "interdistancia",
  "disposicio": "disposición",
  "numero_pl": "número_pl",
  "luminaria_": "LUMINARIA FUTURA",            # ← revisar
  "tipo_lumin": "tipo_luminaria",
  "marca_lumi": "PATRÓN-ORDEN",                # ← revisar
  "modelo_lum": "MODELO LUMINARIA FUTURA",
  "numero_lum": "PL TOTALES",
  "ral": "UNIFORMIDAD",
  "tipo_sopor": "tipo_soporte",
  "marca_sopo": "SOPORTE",                     # ← revisar
  "modelo_sop": "TIPO VIAL",                   # ← revisar
  "longitud_b": "interdistancia_simplificada",
  "tipo_lampa": "tipo_lámpara",
  "marca_lamp": "ESTUDIOS",                    # ← revisar
  "modelo_lam": "CM",                          # ← revisar
  "potencia": "potencia",
  "equipo_aux": "equipo_luminaria",
  "tipo_via": "tipo_vía",
  "nombre_via": "nombre_vía",
  "numero": "PL ADICIONALES",                  # ← revisar
  "VERI SEC": "TIPO LEGALIZACIÓN"              # ← revisar
}

RAW_MAP_VIA_SRC_TO_DEST: Dict[str, str] = {
  "id_vial": "id_vial",
  "tipologia": "tipo_vía",
  "nombre": "nombre_vía",
  "tramo": "calle",
  "acera_1": "acera_1",
  "aparcamien": "aparcamiento_1",
  "calzada_1": "calzada_1",
  "mediana": "mediana",
  "calzada_2": "calzada_2",
  "aparcami_1": "aparcamiento_2",
  "acera_2": "acera_2",
  "num_carril": "num_carril_1",
  "num_carr_1": "num_carril_2",
  "interdista": "interdistancia",
  "disposicio": "disposición",
  "lux_numero": "número_registro",
  "ancho_tota": "ancho total"
}

def _invert(src_to_dest: Dict[str, str]) -> Dict[str, str]:
    """Convierte ORIGEN->DESTINO en DESTINO->ORIGEN (lo que usa el mapeador)."""
    inv: Dict[str, str] = {}
    for src, dest in src_to_dest.items():
        inv[str(dest)] = str(src)
    return inv


# ------------------------- Listas de columnas a blindar ------------------------#

PROTECTED_CM_RAW = [
    "CENTRO DE MANDO", "calle", "LOCALIZACIÓN FINAL", "MÓDULO DE MEDIDA", "S ACOME",
    "número_contador", "SISTEMA DE ENCENDIDO", "SISTEMA DE REDUCCIÓN DE POTENCIA",
    "ELEMENTOS DE MEDIDA", "Nº CONTADOR", "TARIFA ACTUAL", "PRECIO ELÉCTRICO",
    "CONSUMO FACTURACIÓN CORREGIDO", "COSTE FACTURACIÓN CORREGIDO", "POTENCIA ANALIZADOR",
    "Nº PL ACTUALES", "POTENCIA ACTUAL (kW)", "HORAS ACTUALES (h)",
    "CONSUMO ACTUAL (kWh/año)", "COSTE ANUAL (€)", "Nº PL FUTUROS",
    "POTENCIA FUTURA (kW)", "CONSUMO FUTURO (CLASIFICACIÓN)"
]

PROTECTED_SEC_RAW = [
    "CM", "CENTRO DE MANDO", "calle", "interdistancia_simplificada", "ANCHURA TOTAL",
    "SUPERFICIE ESTUDIO", "TIPO VIAL", "PL TOTALES", "POTENCIA CORREGIDA",
    "LUMINARIA FUTURA", "MODELO LUMINARIA FUTURA", "POTENCIA ESTUDIOS",
    "POTENCIA ESTUDIOS CORREGIDA", "CLASIFICACIÓN ALUMBRADO", "ILUMINANCIA MEDIA",
    "ILUMINANCIA MÍNIMA", "SECCIÓN EQUIVALENTE", "POTENCIA FUTURA (kW)",
    "HORAS FUTURAS", "CONSUMO FUTURO"
]

PROTECTED_VIA_RAW = [
    "id_vial", "calle", "interdistancia_simplificada"
]


# ----------------------- Configuración por hoja (final) -----------------------#

CONFIG_BY_SHEET: Dict[str, SheetConfig] = {
    "CM": SheetConfig(
        path_names=("RutaProducto", "Ruta_CM", "RutaProducto_CM"),
        source_sheet_candidates=("Centro_mando", "Centro mando", "CM", "Hoja1"),
        dest_table_name="Tabla3",
        explicit_column_map=_invert(RAW_MAP_CM_SRC_TO_DEST),
        protected_dest_columns=PROTECTED_CM_RAW
    ),
    "Sec": SheetConfig(
        path_names=("RutaSecciones", "Ruta_Sec"),
        source_sheet_candidates=("Seccion", "Secciones", "Sec", "Hoja1"),
        dest_table_name="Tabla2",
        explicit_column_map=_invert(RAW_MAP_SEC_SRC_TO_DEST),
        protected_dest_columns=PROTECTED_SEC_RAW
    ),
    "Via": SheetConfig(
        path_names=("RutaVial", "Ruta_Via"),
        source_sheet_candidates=("Vial", "Via", "Hoja1"),
        dest_table_name="Tabla5",
        explicit_column_map=_invert(RAW_MAP_VIA_SRC_TO_DEST),
        protected_dest_columns=PROTECTED_VIA_RAW
    ),
    "ListCP": SheetConfig(
        path_names=("RutaListCP", "RutaMunicipio", "Ruta_ListCP"),
        source_sheet_candidates=("ListCP", "Municipio", "lista", "Hoja1"),
        dest_table_name="Tabla7",
        explicit_column_map=None,               # heurística si no hay mapeo
        protected_dest_columns=None
    ),
}


# ----------------------------- Capa de Infra/Excel ----------------------------#

class ExcelGateway:
    """Encapsula operaciones xlwings/COM para libro que llama y libros origen."""

    def __init__(self, caller_wb: xw.Book):
        self.wb = caller_wb
        self.app = caller_wb.app

    # ---------- Lectura de nombres definidos ---------- #
    def try_get_defined_name_value(self, candidates: Sequence[str]) -> Optional[str]:
        """
        Devuelve el valor del primer nombre definido existente como TEXTO.
        Soporta:
        - Nombres que apuntan a un rango (refers_to_range.value)
        - Nombres que son un literal de texto: ="C:\ruta\archivo.xlsx"
          (se toma .refers_to, se quita '=' y comillas)
        """
        for nm in candidates:
            try:
                name = self.wb.names[nm]
            except Exception:
                continue

            # 1) Caso rango -> coger su valor
            try:
                rng = name.refers_to_range
                val = rng.value
                if isinstance(val, str) and val.strip():
                    return val.strip()
            except Exception:
                pass

            # 2) Caso literal de texto en .refers_to -> ="C:\...\archivo.xlsx"
            try:
                refers = name.refers_to  # p.ej. ="C:\...\archivo.xlsx"
                if isinstance(refers, str) and refers:
                    txt = refers.lstrip("=")
                    # comportamiento como tu leer_ruta: quitar TODAS las comillas
                    txt = txt.replace('"', "").strip()
                    if txt:
                        return txt
            except Exception:
                pass

        return None

    def try_get_json_mapping_from_name(self, sheet_name: str) -> Optional[Dict[str, str]]:
        candidates = (f"ColumnMap_JSON_{sheet_name}", f"JSONMap_{sheet_name}")
        txt = self.try_get_defined_name_value(candidates)
        if not txt:
            return None
        try:
            data = json.loads(txt)
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items()}
        except Exception:
            pass
        return None

    def try_get_protected_columns_from_names(self, sheet_name: str) -> Optional[List[str]]:
        candidates = (f"ProtectedColumns_{sheet_name}", f"ProtectedCols_{sheet_name}")
        for nm in candidates:
            try:
                name = self.wb.names[nm]
                rng = name.refers_to_range
                val = rng.value
                items: List[str] = []
                if isinstance(val, str):
                    parts = [p.strip() for p in val.replace(";", ",").split(",")]
                    items = [p for p in parts if p]
                elif isinstance(val, (list, tuple)):
                    def _flatten(v):
                        if isinstance(v, (list, tuple)):
                            for el in v:
                                yield from _flatten(el)
                        else:
                            yield v
                    items = [str(x).strip() for x in _flatten(val) if x and str(x).strip()]
                if items:
                    return items
            except Exception:
                continue
        return None

    # ---------- Apertura libro origen ---------- #
    def open_source(self, path: str) -> xw.Book:
        return self.app.books.open(path, read_only=True, update_links=False)

    # ---------- Utilidades tabla destino ---------- #
    @staticmethod
    def _get_first_listobject(sh: xw.Sheet):
        lo_count = sh.api.ListObjects.Count
        if lo_count == 0:
            raise RuntimeError(f"La hoja '{sh.name}' no tiene tablas (ListObjects).")
        return sh.api.ListObjects(1)

    def get_table(self, sh: xw.Sheet, table_name: Optional[str]):
        if table_name:
            try:
                return sh.api.ListObjects(table_name)
            except Exception:
                pass
        return self._get_first_listobject(sh)

    @staticmethod
    def get_table_headers(lo) -> List[str]:
        hdrs = list(lo.HeaderRowRange.Value[0])
        return [str(h or "").strip() for h in hdrs]

    @staticmethod
    def get_table_data_row_count(lo) -> int:
        try:
            return int(lo.DataBodyRange.Rows.Count)
        except Exception:
            return 0

    @staticmethod
    def resize_table_rows(lo, desired_rows: int) -> None:
        """Añade/elimina filas de tabla sin tocar nada fuera (Totales se maneja fuera)."""
        current_rows = ExcelGateway.get_table_data_row_count(lo)
        if desired_rows < 0:
            desired_rows = 0

        if desired_rows > current_rows:
            to_add = desired_rows - current_rows
            for _ in range(to_add):
                lo.ListRows.Add()
        elif desired_rows < current_rows:
            to_del = current_rows - desired_rows
            for _ in range(to_del):
                lo.ListRows(lo.ListRows.Count).Delete()

    @staticmethod
    def write_rows_to_table(
        sh: xw.Sheet,
        lo,
        values_2d: List[List],
        dest_headers: List[str],
        protected_headers: Optional[Sequence[str]] = None
    ):
        """
        Escribe la matriz 2D en la tabla, PERO sin tocar:
        - Columnas que tengan fórmula (HasFormula)
        - Columnas cuyo encabezado esté en protected_headers (match por normalización)
        Se escribe columna a columna.
        """
        rows = len(values_2d)
        if rows == 0:
            return

        first_row = int(lo.DataBodyRange.Row)
        first_col = int(lo.DataBodyRange.Column)
        n_dest_cols = int(lo.ListColumns.Count)
        n_src_cols = len(values_2d[0]) if rows else 0

        prot_norm: Set[str] = set(_norm(h) for h in (protected_headers or []) if h)
        header_norms = [_norm(h) for h in dest_headers]
        protected_mask: List[bool] = [hn in prot_norm for hn in header_norms]

        # Detectar columnas con fórmula (miramos la primera celda de datos de cada ListColumn)
        formula_mask: List[bool] = []
        for j in range(1, n_dest_cols + 1):
            has_formula = False
            try:
                dbr = lo.ListColumns(j).DataBodyRange
                if dbr is not None:
                    try:
                        has_formula = bool(dbr.Cells(1, 1).HasFormula)
                    except Exception:
                        has_formula = False
            except Exception:
                has_formula = False
            formula_mask.append(has_formula)

        # Escribir solo en columnas SIN fórmula y NO protegidas
        for j in range(min(n_dest_cols, n_src_cols)):
            if formula_mask[j] or protected_mask[j]:
                continue
            col_values = [[row[j]] for row in values_2d]
            target = sh.range((first_row, first_col + j)).resize(rows, 1)
            target.value = col_values

        # Logging informativo
        skipped_formula = [dest_headers[i] for i, f in enumerate(formula_mask) if i < len(dest_headers) and f]
        skipped_protected = [dest_headers[i] for i, p in enumerate(protected_mask) if i < len(dest_headers) and p]
        if skipped_formula:
            print(f"[INFO] Columnas no escritas por fórmula: {skipped_formula}")
        if skipped_protected:
            print(f"[INFO] Columnas no escritas por protección explícita: {skipped_protected}")


# ------------------------------ Lectura de origen -----------------------------#

class SourceReader:
    """Lee datos tabulares (cabeceras + filas) desde una hoja de un libro origen."""

    @staticmethod
    def read_table_from_sheet(src_sheet: xw.Sheet) -> Tuple[List[str], List[List]]:
        used = src_sheet.used_range
        data = used.value
        if not data:
            return [], []
        if not isinstance(data[0], (list, tuple)):
            data = [data]
        if len(data) == 1:
            return [str(c or "") for c in data[0]], []
        headers = [str(c or "").strip() for c in data[0]]
        rows = [list(r) for r in data[1:] if any(cell is not None and str(cell).strip() != "" for cell in r)]
        return headers, rows


# ------------------------------- Mapeo columnas -------------------------------#

class ColumnMapper:
    """Resuelve el mapeo origen->destino y transforma filas."""

    def __init__(self, dest_headers: List[str], source_headers: List[str], explicit_map: Optional[Dict[str, str]]):
        self.dest_headers = dest_headers
        self.source_headers = source_headers
        self.src_index_by_norm = {_norm(h): idx for idx, h in enumerate(source_headers)}
        self.explicit_map = explicit_map or {}

        self.src_idx_for_dest: List[Optional[int]] = []
        for d in dest_headers:
            idx = self._resolve_source_index_for_dest(d)
            self.src_idx_for_dest.append(idx)

    def _resolve_source_index_for_dest(self, dest_header: str) -> Optional[int]:
        if dest_header in self.explicit_map:
            src_name = self.explicit_map[dest_header]
            return self.src_index_by_norm.get(_norm(src_name))

        n = _norm(dest_header)
        if n in self.src_index_by_norm:
            return self.src_index_by_norm[n]

        for src_h, idx in self.src_index_by_norm.items():
            if n and (n in src_h or src_h in n):
                return idx

        return None

    def map_rows(self, src_rows: List[List]) -> List[List]:
        out: List[List] = []
        for r in src_rows:
            new_row: List = []
            for idx in self.src_idx_for_dest:
                val = r[idx] if (idx is not None and idx < len(r)) else None
                new_row.append(val)
            out.append(new_row)
        return out

    def report_unmapped(self) -> Tuple[List[str], List[str]]:
        used_idx = {i for i in self.src_idx_for_dest if i is not None}
        dest_without_src = [d for d, i in zip(self.dest_headers, self.src_idx_for_dest) if i is None]
        src_not_used = [h for i, h in enumerate(self.source_headers) if i not in used_idx]
        return dest_without_src, src_not_used


# ------------------------------ Guardián de Totales ---------------------------#

class TotalsRowGuard:
    """Context manager para desactivar/activar la fila de Totales alrededor del volcado."""
    def __init__(self, listobject):
        self.lo = listobject
        self.prev: Optional[bool] = None

    def __enter__(self):
        try:
            self.prev = bool(self.lo.ShowTotals)
            if self.prev:
                print("[INFO] Desactivando fila de Totales antes del volcado…")
            self.lo.ShowTotals = False
        except Exception:
            pass
        return self.lo

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.prev is not None:
                self.lo.ShowTotals = self.prev
                if self.prev:
                    print("[INFO] Fila de Totales restaurada tras el volcado.")
        except Exception:
            pass
        return False  # no suprime excepciones


# ------------------------------- Caso de uso ----------------------------------#

def _match_requested_to_dest_headers(requested: Sequence[str], dest_headers: List[str]) -> Dict[str, str]:
    """Devuelve un dict {requested_name -> dest_header_match} usando normalización relajada."""
    dest_relaxed_map = {_norm_relaxed(h): h for h in dest_headers}
    result: Dict[str, str] = {}
    for req in requested:
        r = _norm_relaxed(req)
        if not r:
            continue
        if r in dest_relaxed_map:
            result[req] = dest_relaxed_map[r]
            continue
        candidates = [h for k, h in dest_relaxed_map.items() if r in k or k in r]
        if candidates:
            result[req] = sorted(candidates, key=len)[0]
    return result


class LoadActiveSheetUseCase:
    """Orquestador de la carga."""

    def __init__(self, excel: ExcelGateway, config_by_sheet: Dict[str, SheetConfig]):
        self.excel = excel
        self.config_by_sheet = config_by_sheet

    def _resolve_config(self, sheet_name: str) -> SheetConfig:
        key = sheet_name.strip()
        if key not in self.config_by_sheet:
            raise RuntimeError(
                f"La hoja activa '{sheet_name}' no está configurada. "
                f"Hojas soportadas: {', '.join(self.config_by_sheet.keys())}"
            )
        return self.config_by_sheet[key]

    def _pick_source_sheet(self, src_wb: xw.Book, candidates: Sequence[str]) -> xw.Sheet:
        by_norm: Dict[str, xw.Sheet] = {_norm(sh.name): sh for sh in src_wb.sheets}
        for name in candidates:
            if name in src_wb.sheets:
                return src_wb.sheets[name]
            n = _norm(name)
            if n in by_norm:
                return by_norm[n]
        return src_wb.sheets[0]

    def _effective_protected_headers(
        self,
        sheet_name: str,
        dest_headers: List[str],
        explicit_map: Optional[Dict[str, str]],
        cfg: SheetConfig
    ) -> List[str]:
        """
        Calcula la lista FINAL de cabeceras protegidas:
        - Une: (a) definidas en Excel por nombre y (b) definidas en config.
        - Mapea nombres 'humanos' a cabeceras reales usando normalización relajada.
        - Si hay mapeo explícito, exige que la cabecera protegida esté en ese mapeo.
        """
        protected_from_excel = self.excel.try_get_protected_columns_from_names(sheet_name) or []
        protected_from_config = list(cfg.protected_dest_columns or [])
        requested = list({*protected_from_excel, *protected_from_config})

        matched = _match_requested_to_dest_headers(requested, dest_headers)  # req -> real_dest_header
        matched_dest_headers = list(sorted(set(matched.values()), key=lambda x: dest_headers.index(x)))

        if explicit_map:
            allowed = set(explicit_map.keys())
            matched_dest_headers = [h for h in matched_dest_headers if h in allowed]

        return matched_dest_headers

    def run(self) -> None:
        sh = self.excel.wb.sheets.active
        sheet_name = sh.name
        print(f"[INFO] Hoja activa destino: {sheet_name}")

        cfg = self._resolve_config(sheet_name)

        # 1) Ruta del libro origen (solo la ruta de ESTA hoja)
        src_path = self.excel.try_get_defined_name_value(cfg.path_names)
        if not src_path:
            raise RuntimeError(
                f"No se encontró la ruta del origen para la hoja '{sheet_name}'. "
                f"Define uno de estos nombres: {', '.join(cfg.path_names)}"
            )
        if not os.path.isfile(src_path):
            raise RuntimeError(f"La ruta leída no es un archivo válido o no existe: {src_path}")
        print(f"[INFO] Archivo origen: {src_path}")

        # 2) Tabla destino y cabeceras
        lo = self.excel.get_table(sh, cfg.dest_table_name)
        dest_headers = self.excel.get_table_headers(lo)
        print(f"[INFO] Tabla destino: {lo.Name} | Columnas destino: {len(dest_headers)}")

        # 3) Leer libro/hoja origen
        src_wb = self.excel.open_source(src_path)
        try:
            src_sh = self._pick_source_sheet(src_wb, cfg.source_sheet_candidates)
            print(f"[INFO] Hoja origen seleccionada: {src_sh.name}")
            src_headers, src_rows = SourceReader.read_table_from_sheet(src_sh)
            print(f"[INFO] Columnas origen: {len(src_headers)} | Filas origen: {len(src_rows)}")

            # 4) Resolver mapeo
            explicit_map = self.excel.try_get_json_mapping_from_name(sheet_name) or cfg.explicit_column_map
            mapper = ColumnMapper(dest_headers, src_headers, explicit_map)
            mapped_rows = mapper.map_rows(src_rows)

            missing_dest, unused_src = mapper.report_unmapped()
            if missing_dest:
                print(f"[WARN] Columnas destino SIN origen: {missing_dest}")
            if unused_src:
                print(f"[INFO] Columnas origen NO usadas (ok si sobran): {unused_src}")

            # 5) Determinar columnas protegidas efectivas (config + Excel), filtradas por mapeo si aplica
            protected_headers = self._effective_protected_headers(sheet_name, dest_headers, explicit_map, cfg)
            if protected_headers:
                print(f"[INFO] Columnas protegidas (efectivas): {protected_headers}")

            # 6) Desactivar Totales, redimensionar y volcar respetando protecciones/fórmulas
            with TotalsRowGuard(lo):
                ExcelGateway.resize_table_rows(lo, len(mapped_rows))
                ExcelGateway.write_rows_to_table(
                    sh=sh,
                    lo=lo,
                    values_2d=mapped_rows,
                    dest_headers=dest_headers,
                    protected_headers=protected_headers
                )

            print(f"[OK] Carga completada en '{sheet_name}' -> {lo.Name}. Filas escritas: {len(mapped_rows)}")
        finally:
            try:
                src_wb.close(save=False)
            except Exception:
                pass


# --------------------------------- Entrada -----------------------------------#

def main() -> None:
    """Punto de entrada principal para Excel."""
    try:
        wb = xw.Book.caller()
    except Exception:
        raise RuntimeError(
            "Este script debe ejecutarse desde Excel con xlwings (Book.caller()).\n"
            "Abre el xlsm y ejecuta la macro que llama a script_carga.main()."
        )

    excel = ExcelGateway(wb)
    use_case = LoadActiveSheetUseCase(excel, CONFIG_BY_SHEET)
    use_case.run()

# Alias compatibilidad
def cargar_datos() -> None:
    main()

if __name__ == "__main__":
    main()