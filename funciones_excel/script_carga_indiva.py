# -*- coding: utf-8 -*-
"""
script_carga_indiva.py

Carga datos desde un Excel INDIVA al libro de Auditoría (Artecoin),
escribiendo únicamente las columnas indicadas en DIFF_MAPS (INDIVA -> ARTECOIN).
No realiza fuzzy ni inferencias: SOLO usa esas columnas.

Ejecución típica desde Excel (xlwings):
    RunPython("import script_carga_indiva; script_carga_indiva.main()")

Nombres definidos opcionales en el libro Auditoría:
- Ruta_IndivaFile: ruta completa del .xlsx de INDIVA (si no, se solicitará).
- ProtectedColumns_Cont: lista separada por ; , | de columnas destino a NO escribir.
- ConsulTableStyle: nombre del estilo de tabla a reaplicar.

Requisitos:
    pip install xlwings
"""

from __future__ import annotations
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set

import xlwings as xw

# ----------------- TU DICCIONARIO INDIVA -> ARTECOIN -----------------
# Claves: cabeceras de INDIVA (origen)
# Valores: cabeceras de ARTECOIN (destino)
DIFF_MAPS: Dict[str, Dict[str, str]] = {
  "CENTRO": {
    "Nombre": "NOMBRE_CENTRO",
    "Dirección": "DIRECCIÓN",
    "Alquiler/propiedad": "ALQUILER/PROPIEDAD",
    "Reformas ambito": "REFORMAS_AMBITO",
    "Observaciones": "OBSERVACIONES",
    "TIPOLOGIA": "TIPO_CENTRO",
    # "HORARIO FINDE": "DESGLOSAR HORARIO",
    "ID CENTRO": "ID_CENTRO",
    "FOTO_CENTRO": "FOTO_CENTRO",
    "DÍAS APERTURA2": "DIAS_APERTURA_SEMANA",
    "HORA_APERTURA_MAÑANA": "HORA_APERTURA_MAÑANA",
    "HORA_CIERRE_MAÑANA": "HORA_CIERRE_MAÑANA",
    "HORA_APERTURA_TARDE": "HORA_APERTURA_TARDE",
    "HORA_CIERRE_TARDE": "HORA_CIERRE_TARDE",
    "HORA_APERTURA_NOCHE": "HORA_APERTURA_NOCHE",
    "HORA_CIERRE_NOCHE": "HORA_CIERRE_NOCHE",
    "NUMERO_EDIFICIOS": "NUMERO_EDIFICIOS"
  },
  "EDIFICIO": {
    "Nombre": "NOMBRE_EDIFICIO",
    "NUMERO _PLANTAS": "NUMERO _PLANTAS",
    "OBSERVACIONES": "OBSERVACIONES",
    "FOTO_EDIFICIO": "FOTO_EDIFICIO",
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "ALTURA PLANTA": "ALTURA_PLANTA"
  },
  "DATOS_ELECTRICOS_EDIFICIOS": {
    "CT_POTENCIA": "CT_POTENCIA",
    "CT_OBSERVACIONES": "CT_OBSERVACIONES",
    "CT": "CT",
    "CDRO_ELECT_PPAL_CONTADOR": "CDRO_ELECT_PPAL_CONTADOR",
    "CDRO_ELECT_PPAL_CONTADOR_SECUND": "CDRO_ELECT_PPAL_CONTADOR_SECUND",
    "BAT_CONDENSADORES": "BAT_CONDENSADORES",
    "CDRO_ELECT_SECUND_OBSERVACIONES": "CDRO_ELECT_SECUND_OBSERVACIONES",
    "Observaciones": "Observaciones",
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID EDIFICACION2": "ID_EDIFICIO"
  },
  "DEPENDENCIA": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "Nº DEPENDENCIAS": "NUMERO_DEPENDENCIA",
    "SUPERFICIE POR DEPENDENCIA": "SUPERFICIE (m2)"
  },
  "EQUIPOS_HORIZONTALES": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "NOMBRE_EQHORIZ2": "NOMBRE_EQHORIZ",
    "NUMERO_EQHORIZ2": "NUMERO_EQHORIZ",
    "ID EQUIPOS HORIZONTALES3": "ID_EQHORIZ",
    "OBSERVACIONES2": "OBSERVACIONES",
    "POTENCIA": "POT_EQHORIZ_W"
  },
  "OTROS_EQUIPOS": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO22": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "MODELO2": "MODELO",
    "OBSERVACIONES3": "OBSERVACIONES",
    "NOMBRE_EQOTRO2": "NOMBRE_EQOTRO",
    "NUMERO_EQOTRO4": "NUMERO_EQOTRO",
    "ID OTROS EQUIPOS3": "ID_EQOTRO",
    "POTENCIA": "POT_ABS_W",
    "TECNICO": "",
    "FOTO_EQOTRO": "",
    "ID_EDIFICIO2": "ID EDIFICACION2",
    "ID_EQOTRO2": "ID EDIFICACION2"
  },
  "ILUMINACIÓN": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO22": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    # "Etiqueta2": "",
    "Tipo lampara3": "TIPO_LAMPARA",
    "Tipo luminaria2": "TIPO_LUMINARIA",
    "POT_LAMPARA_W3": "POT_LAMPARA_W",
    "OBSERVACIONES4": "OBSERVACIONES",
    "CLASE_ILUM2": "CLSASE_ILUM",
    "ID ILUMINACION3": "ID_ILUM",
    "NUM_LAMPARAS2": "NUM_LAMPARAS",
    "ALTURA_LUMINARIAS3": "ALTURA_LUMINARIAS",
    "REGULACION4": "REGULACION",
    "NUM_LUMINARIAS2": "NUM_LUMINARIAS"
  },
  "ASCENSORES": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "OBSERVACIONES2": "OBSERVACIONES",
    "ID ASCENSORES2": "ID_EQELEV",
    "NOMBRE_EQELEV3": "NOMBRE_EQELEV",
    "TECNOLOGIA4": "TECNOLOGIA",
    "NUM_ASCENSORES2": "NUM_ASCENSORES",
    "POTENCIA": "POT_ABS_W"
  },
  "UDD_INT_CLIMATIZACION": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "OBSERVACIONES2": "OBSERVACIONES",
    "ID UDD INT CLIMATIZACION2": "ID_EQCLIMAINT",
    "NOMBRE_EQCLIMAINT3": "NOMBRE_EQCLIMAINT",
    "NUMERO_EQCLIMAINT4": "NUMERO_EQCLIMAINT",
    "NUMERO_ELEMENTOS_RADIADOR5": "NUMERO_ELEMENTOS_RADIADOR",
    "FLUIDO6": "FLUIDO",
    "SERVICIO7": "SERVICIO",
    "MODELO8": "MODELO",
    "OBSERVACIONES9": "OBSERVACIONES",
    "REGULACION11": "REGULACION",
    "MARCA12": "MARCA",
    "POT_FRIGORIFICA_TERMICA_W": "POT_FRIGORIFICA_TERMICA_W",
    "POT_CALORIFICA_TERMICA_W": "POT_CALORIFICA_TERMICA_W",
    "POT_ABS_FRIO_W": "POT_ABS_FRIO_W",
    "POT_ABS_CALOR_W": "POT_ABS_CALOR_W",
    "ALIMENTACION13": "ALIMENTACION"
  },
  "UDD_EXT_CLIMATIZACION": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "ID UDD EXT CLIMATIZACION2": "ID_EQCLIMAEXT",
    "NOMBRE_EQCLIMAEXT3": "NOMBRE_EQCLIMAEXT",
    "NUMERO_EQCLIMAEXT4": "NUMERO_EQCLIMAEXT",
    "TIPO_UD_TERMINAL5": "TIPO_UD_TERMINAL",
    "SERVICIO6": "SERVICIO",
    "REGULACION7": "REGULACION",
    "MODELO8": "MODELO",
    "OBSERVACIONES9": "OBSERVACIONES",
    "FABRICANTE2": "FABRICANTE",
    "POT_FRIGORIFICA_TERMICA_W": "POT_FRIGORIFICA_TERMICA_W",
    "POT_CALORIFICA_TERMICA_W": "POT_CALORIFICA_TERMICA_W",
    "POT_ABS_FRIO_W": "POT_ABS_FRIO_W",
    "POT_ABS_CALOR_W": "POT_ABS_CALOR_W",
    "ERR": "EER",
    "COP": "COP",
    "FLUIDO3": "FLUIDO",
    "RFRIGERANTE4": "RFRIGERANTE",
    "RECUPERACION_CALOR5": "RECUPERACION_CALOR",
    "CONTROL_MARCA-MODELO6": "CONTROL_MARCA-MODELO",
    "CONTROL_CIRCUITOS_SECUNDARIOS7": "CONTROL_CIRCUITOS_SECUNDARIOS"
  },
  "GENERADOR_DE_CALOR": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "POTENCIA2": "POT_ABS_W",
    "ID GENERADOR DE CALOR3": "ID_EQGEN",
    "NOMBRE_EQGEN2": "NOMBRE_EQGEN",
    "MARCA3": "MARCA",
    "NUMERO_EQGEN4": "NUMERO_EQGEN",
    "OBSERVACIONES5": "OBSERVACIONES",
    "MODELO2": "MODELO"
  },
  "CERRAMIENTOS": {
    "ID CERRAMIENTO": "ID_CERRAMIENTO",
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO22": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "FACHADA_TIPO_OBSERVACIONES": "FACHADA_TIPO_OBSERVACIONES",
    "NOMBRE_CERRAMIENTO": "NOMBRE_CERRAMIENTO",
    "FACHADA_TIPO": "FACHADA_TIPO",
    "CARPINTERIA": "CARPINTERIA",
    "CARPINTERIA TIPO": "CARPINTERIA TIPO",
    "MATERIAL_CARPINTERIA": "MATERIAL_CARPINTERIA",
    "ACRISALAMIENTO": "VENTANAS_ACRISTALAMIENTO",
    "CUBIERTAS_TIPO_OBSERVACIONES": "CUBIERTAS_TIPO_OBSERVACIONES",
    "PROTECCION_SOLAR": "VENTANAS_PROTECCION_SOLAR",
    "VENTANAS_NUM_UNIDADES": "VENTANAS_NUM_UNIDADES",
    "CUBIERTAS_TIPO": "CUBIERTAS_TIPO",
    "CUBIERTAS_ACABADO": "CUBIERTAS_ACABADO",
    "CUBIERTAS_AISLAMIENTO": "CUBIERTAS_AISLAMIENTO",
    "CUBIERTAS_LUCERNARIO": "CUBIERTAS_LUCERNARIO",
    "CUBIERTAS_LUCERNARIO_NUMERO_UNIDADES": "CUBIERTAS_LUCERNARIO_NUMERO_UNIDADES"
  },
  "BOMBAS": {
    "ID CENTRO2": "ID_CENTRO",
    "NOMBRE_CENTRO2": "NOMBRE_CENTRO",
    "ID EDIFICACION2": "ID_EDIFICIO",
    "NOMBRE_EDIFICIO2": "NOMBRE_EDIFICIO",
    "ID DEPENDENCIA2": "ID_DEPENDENCIA",
    "DEPENDENCIA": "NOMBRE_DEPENDENCIA",
    "CLASE_GRUPBOM2": "CLASE_GRUPBOM",
    "NOMBRE_GRUPBOM2": "NOMBRE_GRUPBOM",
    "TIPO_GRUPBOM3": "TIPO_GRUPBOM",
    "TIPO_IMPULSION4": "TIPO_IMPULSION",
    "NUMERO_GRUPBOM5": "NUMERO_GRUPBOM",
    "ID BOMBAS6": "ID_GRUPBOM",
    "OBSERVACIONES7": "OBSERVACIONES",
    "POTENCIA": "POT_ABS_W",
    "TIPO BOMBA": "TIPO_BOMBA",
    "SISTEMA REGULACIÓN": "SISTEMA_REGULACION",
    "MARCA": "MARCA",
    "MODELO": "MODELO",
    "ESTADO": "ESTADO"
  }
}

# ----------------- Mapeo hoja INDIVA -> tabla Consul -----------------
SHEET_TO_TABLE: Dict[str, str] = {
    "DATOS_ELECTRICOS_EDIFICIOS": "Tabla32",
    "CENTRO": "Tabla34",
    "EDIFICIO": "Tabla35",
    "DEPENDENCIA": "Tabla36",
    "CERRAMIENTOS": "Tabla37",
    "GENERADOR_DE_CALOR": "Tabla38",
    "UDD_INT_CLIMATIZACION": "Tabla40",
    "ASCENSORES": "Tabla41",
    "EQUIPOS_HORIZONTALES": "Tabla42",
    "BOMBAS": "Tabla43",
    "OTROS_EQUIPOS": "Tabla44",
    "ILUMINACIÓN": "Tabla45",
    "UDD_EXT_CLIMATIZACION": "Tabla46",
}

# ----------------- Utilidades -----------------
def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = _strip_accents(s).lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"[^a-z0-9 ()]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def looks_like_total_row(values: Iterable[Any]) -> bool:
    for v in values:
        if v is None:
            continue
        txt = norm(v)
        if txt.startswith("total"):
            return True
    return False

# ----------------- Modelos -----------------
@dataclass
class SourceBlock:
    header_original: List[str]
    header_norm: List[str]
    rows: List[Dict[str, Any]] = field(default_factory=list)  # dict por header_norm

# ----------------- Excel Gateway -----------------
class ExcelGateway:
    def __init__(self, book: xw.Book):
        self.book = book
        self.app = book.app

    def get_defined_text(self, name: str) -> Optional[str]:
        n = None
        try:
            n = self.book.names[name]
        except Exception:
            for nm in self.book.names:
                nm_name = str(nm.name).split("!")[-1]
                if nm_name.lower() == name.lower():
                    n = nm
                    break
        if n is None: 
            return None
        try:
            s = (n.refers_to or "").lstrip("=")
            if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
                return s[1:-1]
        except Exception:
            pass
        try:
            val = n.refers_to_range.value
            return (str(val).strip() if val is not None else None)
        except Exception:
            return None

    def choose_file(self, title: str) -> Optional[str]:
        fd = self.app.api.FileDialog(3)  # FilePicker
        fd.Title = title
        fd.AllowMultiSelect = False
        fd.Filters.Clear()
        fd.Filters.Add("Excel", "*.xlsx;*.xlsm;*.xls")
        return fd.SelectedItems(1) if fd.Show() == -1 else None

    def consul_sheet(self) -> xw.Sheet:
        try:
            return self.book.sheets["Consul"]
        except KeyError:
            raise RuntimeError("No se encuentra la hoja 'Consul'.")

    def list_tables(self, sheet: xw.Sheet) -> Dict[str, Any]:
        return {norm(lo.Name): lo for lo in sheet.api.ListObjects}

    def read_table_headers(self, lo) -> List[str]:
        rng = lo.HeaderRowRange
        vals = list(rng.Value[0]) if rng.Rows.Count == 1 else [c.Value for c in rng.Columns]
        return [str(v or "").strip() for v in vals]

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

    def table_row_count(self, lo) -> int:
        try:
            dbr = lo.DataBodyRange
            return int(dbr.Rows.Count) if dbr is not None else 0
        except Exception: 
            return 0

    def add_rows(self, lo, n: int) -> None:
        for _ in range(n): 
            lo.ListRows.Add()

    def del_rows_from_bottom(self, lo, n: int) -> None:
        for _ in range(n):
            cnt = self.table_row_count(lo)
            if cnt <= 0: 
                break
            lo.ListRows(cnt).Delete()

    def write_col_values(self, lo, j1: int, vals: List[Any]) -> None:
        dbr = lo.DataBodyRange
        if dbr is None: 
            raise RuntimeError("Tabla sin DataBodyRange.")
        col_rng = dbr.Columns(j1)
        col_rng.Value = [[v] for v in vals] if vals else []

    def col_has_formula(self, lo, j1: int) -> bool:
        dbr = lo.DataBodyRange
        if dbr is None: 
            return False
        try: 
            return bool(dbr.Columns(j1).Cells(1,1).HasFormula)
        except Exception: 
            return False

    def formula_cols(self, lo, dest_norm: List[str]) -> Set[int]:
        if self.table_row_count(lo) == 0: 
            return set()
        return {j for j in range(1, len(dest_norm)+1) if self.col_has_formula(lo, j)}

    def read_col_values(self, lo, j1: int) -> List[Any]:
        dbr = lo.DataBodyRange
        if dbr is None: 
            return []
        vals = dbr.Columns(j1).Value
        if vals is None: 
            return []
        if isinstance(vals, list) and vals and isinstance(vals[0], list):
            return [r[0] for r in vals]
        return vals if isinstance(vals, list) else [vals]

    def clear_body_fill(self, lo) -> None:
        try:
            dbr = lo.DataBodyRange
            if dbr is None: 
                return
            dbr.Interior.Pattern = -4142
            dbr.Interior.TintAndShade = 0
        except Exception: 
            pass

    def apply_table_style(self, lo, preferred: Optional[str] = None) -> None:
        return

    def set_body_font_auto(self, lo) -> None: 
        try:
            dbr = lo.DataBodyRange
            if dbr is None: 
                return
            dbr.Font.ColorIndex = -4105
        except Exception: 
            pass

# ----------------- Lector INDIVA -----------------
class IndivaReader:
    def __init__(self, app: xw.App):
        self.app = app

    def read_book(self, path: Path) -> xw.Book:
        return self.app.books.open(str(path), update_links=False, read_only=True)

    @staticmethod
    def _used_values(sh: xw.Sheet) -> Optional[List[List[Any]]]:
        ur = sh.used_range
        vals = ur.value if ur else None
        if vals is None: 
            return None
        if not isinstance(vals, list): 
            return [[vals]]
        if vals and not isinstance(vals[0], list): 
            return [[v] for v in vals]
        return vals

    def read_sheets(self, book: xw.Book, interested_sheets: List[str]) -> Dict[str, SourceBlock]:
        out: Dict[str, SourceBlock] = {}
        for sname in interested_sheets:
            try:
                sh = book.sheets[sname]
            except Exception:
                continue
            mat = self._used_values(sh)
            if not mat or len(mat) < 2: 
                continue
            header = [str(h or "").strip() for h in mat[0]]
            header_norm = [norm(h) for h in header]
            rows: List[Dict[str, Any]] = []
            for row in mat[1:]:
                if not any(c is not None and str(c).strip() for c in row):
                    continue
                if looks_like_total_row(row):
                    continue
                d: Dict[str, Any] = {}
                for i, key in enumerate(header_norm):
                    d[key] = row[i] if i < len(row) else None
                rows.append(d)
            if rows:
                out[sname] = SourceBlock(header, header_norm, rows)
        return out

# ----------------- Mapper columnas (INDIVA->ARTECOIN) -----------------
class ColumnMapper:
    @staticmethod
    def build_allowed_map_for_sheet(sheet_name: str,
                                    mapping_indiva_to_dest: Dict[str, str]): # -> Tuple(Set[str], Dict[str, List[str]]):
        """
        Devuelve:
         - allowed_dest_cols_norm: set de columnas destino (normalizadas) que se permiten escribir
         - dest_to_srcs_norm: dict dest_norm -> lista de posibles source_norm (en orden de preferencia)
        Solo usa entradas con destino no vacío.
        """
        allowed: Set[str] = set()
        dest_to_srcs: Dict[str, List[str]] = {}
        for src_orig, dest_orig in mapping_indiva_to_dest.items():
            if dest_orig is None or str(dest_orig).strip() == "":
                continue  # ignorar mapeos vacíos
            src_n = norm(src_orig)
            dest_n = norm(dest_orig)
            allowed.add(dest_n)
            dest_to_srcs.setdefault(dest_n, []).append(src_n)
        return allowed, dest_to_srcs

    @staticmethod
    def choose_first_nonempty(row_dict: Dict[str, Any], src_candidates_norm: List[str]) -> Any:
        for src_n in src_candidates_norm:
            v = row_dict.get(src_n, None)
            # Acepta 0, False, etc. como valor; solo descarta '', None, '   '
            if v is None:
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            return v
        return None

    @staticmethod
    def build_row_matrix(dest_headers: List[str],
                         dest_headers_norm: List[str],
                         protected_cols_norm: Set[str],
                         formula_cols_indices: Set[int],
                         source_block: SourceBlock,
                         allowed_dest_cols_norm: Set[str],
                         dest_to_srcs_norm: Dict[str, List[str]]) -> List[List[Any]]:
        """
        Construye la matriz de escritura alineada a las columnas destino.
        - SOLO escribe columnas cuyo dest_norm esté en allowed_dest_cols_norm.
        - Para cada columna destino permitida, usa el primer source con valor no vacío.
        - Para el resto, devuelve None (y el escritor las saltará).
        """
        matrix: List[List[Any]] = []
        for r in source_block.rows:
            out: List[Any] = []
            for j, dcol in enumerate(dest_headers_norm, start=1):
                if (dcol not in allowed_dest_cols_norm) or (dcol in protected_cols_norm) or (j in formula_cols_indices):
                    out.append(None)
                    continue
                src_candidates = dest_to_srcs_norm.get(dcol, [])
                value = ColumnMapper.choose_first_nonempty(r, src_candidates) if src_candidates else None
                out.append(value)
            matrix.append(out)
        return matrix

# ----------------- Escritor -----------------
@dataclass
class WritePlan:
    table_name: str
    dest_headers: List[str]
    dest_headers_norm: List[str]
    protected_cols_norm: Set[str]
    formula_cols_indices: Set[int]
    allowed_dest_cols_norm: Set[str]

class TableWriter:
    def __init__(self, gw: ExcelGateway):
        self.gw = gw

    def _protected_cols(self) -> Set[str]:
        txt = self.gw.get_defined_text("ProtectedColumns_Cont")
        if not txt: 
            return set()
        return {norm(x) for x in re.split(r"[;,|]", txt) if x.strip()}

    def make_plan(self, lo, table_name: str, allowed_dest_cols_norm: Set[str]) -> WritePlan:
        dest = self.gw.read_table_headers(lo)
        destn = [norm(h) for h in dest]
        protected = self._protected_cols()
        fcols = self.gw.formula_cols(lo, destn)
        return WritePlan(table_name, dest, destn, protected, fcols, allowed_dest_cols_norm)

    def write_matrix(self, lo, plan: WritePlan, matrix: List[List[Any]]) -> None:
        totals = self.gw.is_total_shown(lo)
        if totals: 
            self.gw.set_total_shown(lo, False)
        try:
            # Redimensionar filas (modo REPLACE completo de la tabla por este lote)
            current = self.gw.table_row_count(lo)
            target = len(matrix)
            if target > current: 
                self.gw.add_rows(lo, target - current)
            elif target < current: 
                self.gw.del_rows_from_bottom(lo, current - target)

            # Recalcular columnas con fórmula (por si cambió tras redimensionar)
            fcols = self.gw.formula_cols(lo, plan.dest_headers_norm) | plan.formula_cols_indices

            # Limpieza visual
            self.gw.clear_body_fill(lo)
            self.gw.set_body_font_auto(lo)

            # Escribir solo columnas permitidas
            if target > 0:
                for j in range(len(plan.dest_headers_norm)):
                    dcol = plan.dest_headers_norm[j]
                    if (dcol not in plan.allowed_dest_cols_norm) or ((j+1) in fcols) or (dcol in plan.protected_cols_norm):
                        continue  # saltar: no permitido/protegida/fórmula
                    col_vals = [matrix[i][j] for i in range(target)]
                    self.gw.write_col_values(lo, j+1, col_vals)
        finally:
            self.gw.set_total_shown(lo, totals)
            self.gw.apply_table_style(lo)
            self.gw.set_body_font_auto(lo)

# ----------------- Caso de uso -----------------
class LoadIndivaUseCase:
    def __init__(self, gw: ExcelGateway, reader: IndivaReader, writer: TableWriter):
        self.gw = gw
        self.reader = reader
        self.writer = writer

    def run(self):
        # 1) Ruta del archivo INDIVA
        path_txt = self.gw.get_defined_text("Ruta_IndivaFile") or self.gw.choose_file("Selecciona el Excel de INDIVA")
        if not path_txt:
            raise RuntimeError("No se proporcionó el archivo de INDIVA.")
        path = Path(path_txt)
        if not path.exists():
            raise RuntimeError(f"No existe el archivo: {path}")

        # 2) Leer hojas necesarias (las del diccionario)
        interested_sheets = list(DIFF_MAPS.keys())
        app = self.gw.app
        calc_prev, scr_prev, ev_prev = app.api.Calculation, app.screen_updating, app.enable_events
        app.api.Calculation = -4135
        app.screen_updating = False
        app.enable_events = False
        try:
            src_book = self.reader.read_book(path)
            try:
                blocks = self.reader.read_sheets(src_book, interested_sheets)
            finally:
                try: 
                    src_book.close()
                except Exception: 
                    pass

            if not blocks:
                print("No se encontraron hojas válidas en el Excel de INDIVA con los nombres esperados.")
                return

            # 3) Preparar hoja Consul / tablas
            consul = self.gw.consul_sheet()
            tables = self.gw.list_tables(consul)
            writer = self.writer

            # 4) Para cada hoja con datos, construir matriz y escribir SOLO columnas permitidas
            for sheet_name, sblock in blocks.items():
                table_name = SHEET_TO_TABLE.get(sheet_name)
                if not table_name:
                    print(f"   ! No se definió tabla destino para hoja '{sheet_name}'. Omite.")
                    continue
                lo = tables.get(norm(table_name))
                if lo is None:
                    print(f"   ! Tabla '{table_name}' no existe en 'Consul'. Omite.")
                    continue

                # allowed + dest->sources (por hoja)
                allowed_dest_cols_norm, dest_to_srcs_norm = ColumnMapper.build_allowed_map_for_sheet(
                    sheet_name,
                    DIFF_MAPS.get(sheet_name, {})
                )

                # Plan y matriz
                plan = writer.make_plan(lo, lo.Name, allowed_dest_cols_norm)
                matrix = ColumnMapper.build_row_matrix(
                    dest_headers=plan.dest_headers,
                    dest_headers_norm=plan.dest_headers_norm,
                    protected_cols_norm=plan.protected_cols_norm,
                    formula_cols_indices=plan.formula_cols_indices,
                    source_block=sblock,
                    allowed_dest_cols_norm=allowed_dest_cols_norm,
                    dest_to_srcs_norm=dest_to_srcs_norm
                )

                print(f"· Escribiendo {len(matrix)} fila(s) en '{plan.table_name}' desde hoja '{sheet_name}'"
                      f" (solo columnas mapeadas: {len(allowed_dest_cols_norm)})...")
                writer.write_matrix(lo, plan, matrix)

        finally:
            try: 
                app.api.Calculation = calc_prev
            except Exception: 
                pass
            try: 
                app.screen_updating = scr_prev
            except Exception: 
                pass
            try:
                app.enable_events = ev_prev
            except Exception: 
                pass

# ----------------- Entry point -----------------
def main():
    try:
        book = xw.Book.caller()
    except Exception:
        app = xw.App(visible=False)
        if not app.books:
            raise RuntimeError("No hay libro activo.")
        book = app.books.active
    gw = ExcelGateway(book)
    reader = IndivaReader(book.app)
    writer = TableWriter(gw)
    LoadIndivaUseCase(gw, reader, writer).run()

if __name__ == "__main__":
    main()