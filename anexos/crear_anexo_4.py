from __future__ import annotations

from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime

# NUEVOS IMPORTS MOVIDOS DESDE FUNCIONES
import subprocess
import time
import unicodedata
import re

# Imports opcionales de pywin32
try:
    import win32com.client as win32_client
    from win32com.client import constants as win32_constants
    import pythoncom

    HAS_PYWIN32 = True
except ImportError:
    win32_client = None
    win32_constants = None
    pythoncom = None
    HAS_PYWIN32 = False

# --------------------------------------------------------------------
# 1) CONFIGURACIÓN ----------------------------------------------------
# --------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = (
    BASE_DIR.parent
    / "excel/proyecto/ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1_V20.xlsx"
)
TEMPLATE_DOC = BASE_DIR.parent / "word/anexos/Plantilla_Anexo_4.docx"
OUTPUT_PATH = BASE_DIR / "ANEXO_4.docx"

SHEET_MAP = {
    "Envol": "SISTEMAS CONSTRUCTIVOS",
}

HEADER_ROW = 0  # primera fila
SKIP_ROWS = None
GROUP_COLUMN = "CENTRO"  # Columna de agrupación para generar anexos

# --------------------------------------------------------------------
# 2) LIMPIEZA DE DATAFRAMES -------------------------------------------
# --------------------------------------------------------------------


def delete_rows_optimized(df, columna="ID CENTRO"):
    """Elimina las filas vacías al final del DataFrame."""
    if columna in df.columns:
        # Encontrar la última fila con datos válidos
        for i in range(len(df) - 1, -1, -1):
            val = df[columna].iloc[i]
            if pd.notna(val) and str(val).strip() != "" and str(val).strip() != "0":
                return df.iloc[: i + 2].copy()
    return df.copy()


def clean_last_row(df):
    """Limpia la última fila de totales."""
    if df.empty:
        return df
    df2 = df.copy()
    mask = df2.iloc[-1:] == "Total"
    df2.iloc[-1:] = df2.iloc[-1:].mask(mask, pd.NA)
    return df2


def load_and_clean_sheets(xls_path, sheet_map):
    """Carga y limpia todas las hojas especificadas."""
    with pd.ExcelFile(xls_path) as xls:
        # Las claves del sheet_map son los nombres reales de las hojas
        missing = [k for k in sheet_map.keys() if k not in xls.sheet_names]
        if missing:
            raise ValueError(f"Hojas faltantes en el Excel: {', '.join(missing)}")

        result = {}
        for key, sheet_name in sheet_map.items():
            print(f"-> Procesando hoja: {key}")  # key es el nombre real de la hoja
            df = pd.read_excel(
                xls,
                key,
                header=HEADER_ROW,
                skiprows=SKIP_ROWS,
                dtype=str,  # usar key, no sheet_name
            )
            df_cleaned = clean_last_row(delete_rows_optimized(df))

            # Redondear valores decimales a dos decimales
            for col in df_cleaned.columns:
                # Intentar convertir a float, si falla, dejar como está
                try:
                    df_cleaned[col] = pd.to_numeric(df_cleaned[col])
                    if pd.api.types.is_float_dtype(df_cleaned[col]):
                        df_cleaned[col] = df_cleaned[col].round(2)
                except Exception:
                    pass

            df_cleaned = df_cleaned.fillna("")

            result[key] = df_cleaned

        return result


def clean_filename(filename):
    """
    Limpia el nombre del archivo eliminando caracteres no válidos
    y tildes para Windows.
    """
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


def get_totales_centro(
    df_full: pd.DataFrame, df_grupo: pd.DataFrame, nombre_seccion: str
) -> dict[str, str]:
    """Calcula totales por grupo (antes 'edificio'), usando la última fila como referencia.

    df_full: DataFrame completo (incluye fila final con totales globales precalculados).
    df_grupo: Subconjunto filtrado para un valor concreto de GROUP_COLUMN.
    nombre_seccion: Reservado para posibles usos futuros.
    """
    if df_full.empty:
        return {}

    ultima_full = df_full.iloc[-1]
    cols_a_sumar = [
        col
        for col in df_full.columns[1:]
        if pd.notna(ultima_full[col]) and str(ultima_full[col]).strip() != ""
    ]

    key_label = "EDIFICIO" if "EDIFICIO" in df_full.columns else GROUP_COLUMN
    totales: dict[str, str] = {key_label: "Total general"}

    for col in cols_a_sumar:
        vals = pd.to_numeric(df_grupo[col], errors="coerce").dropna()
        if vals.empty or vals.sum() == 0:
            totales[col] = ""
        else:
            s = vals.sum()
            if abs(s - round(s)) < 1e-6:
                totales[col] = str(int(round(s)))
            else:
                totales[col] = f"{s:.2f}".rstrip("0").rstrip(".")
    return totales


def cerrar_word_procesos():
    """Cierra todos los procesos de Word para evitar conflictos de archivos."""
    try:
        print("-> Cerrando procesos de Word...")

        # Intentar cerrar Word con COM si está disponible
        if HAS_PYWIN32:
            pythoncom.CoInitialize()
            try:
                try:
                    word_app = win32_client.GetActiveObject("Word.Application")
                except Exception:
                    word_app = None

                if word_app is not None:
                    if word_app.Documents.Count > 0:
                        print(
                            f"   Cerrando {word_app.Documents.Count} documentos abiertos..."
                        )
                        for doc in list(word_app.Documents):
                            try:
                                doc.Close(SaveChanges=False)
                            except Exception:
                                pass
                    word_app.Quit()
                    print("   * Word cerrado correctamente")
            finally:
                pythoncom.CoUninitialize()

        # Forzar cierre de cualquier proceso restante
        result = subprocess.run(
            ["taskkill", "/F", "/IM", "winword.exe", "/T"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode == 0:
            print("   * Procesos de Word forzados a cerrar")
        else:
            print("   * No había procesos de Word ejecutándose")

        time.sleep(2)

    except subprocess.TimeoutExpired:
        print("   ! Timeout al intentar cerrar Word")
    except Exception as e:
        print(f"   ! Error al cerrar Word: {e}")

    print("   ✓ Sistema listo para generar documentos")


def get_user_input():
    """Solicita al usuario el mes y año para el documento."""
    print("\n" + "=" * 50)
    print("CONFIGURACIÓN DEL DOCUMENTO")
    print("=" * 50)

    # Obtener el año actual como valor por defecto
    current_year = datetime.now().year
    current_month = datetime.now().month

    # Diccionario de meses en español
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

    # Solicitar mes
    while True:
        try:
            print(f"\nMes actual: {meses_espanol[current_month]} ({current_month})")
            mes_input = input(
                f"Ingrese el mes (1-12) [Enter para usar {current_month}]: "
            ).strip()

            if mes_input == "":
                mes_num = current_month
            else:
                mes_num = int(mes_input)

            if 1 <= mes_num <= 12:
                mes_nombre = meses_espanol[mes_num]
                break
            else:
                print("Error: El mes debe estar entre 1 y 12")
        except ValueError:
            print("Error: Por favor ingrese un número válido")

    # Solicitar año
    while True:
        try:
            anio_input = input(
                f"Ingrese el año [Enter para usar {current_year}]: "
            ).strip()
            if anio_input == "":
                anio = current_year
            else:
                anio = int(anio_input)

            min_year, max_year = current_year - 5, current_year + 5
            if min_year <= anio <= max_year:
                break
            else:
                print(f"Error: El año debe estar entre {min_year} y {max_year}")
        except ValueError:
            print("Error: Por favor ingrese un año válido")

    print("\nConfiguración seleccionada:")
    print(f"   Mes: {mes_nombre} ({mes_num})")
    print(f"   Año: {anio}")
    print("=" * 50)

    return mes_nombre, anio


# Obtener datos del usuario
mes_nombre, anio = get_user_input()

# NUEVO: Cerrar todos los procesos de Word antes de empezar
cerrar_word_procesos()

# Cargar y limpiar datos
print("-> Cargando datos del Excel...")
all_dataframes = load_and_clean_sheets(EXCEL_PATH, SHEET_MAP)

# Asignar a variables individuales
df_envol = all_dataframes["Envol"]

print("* Datos cargados y limpiados")


# Crear documento Word
# def update_word_fields(doc_path):
#     """Actualiza los campos del documento Word."""
#     try:
#         if not Path(doc_path).exists():
#             print(f"   ! El archivo no existe: {doc_path}")
#             return

#         if not HAS_PYWIN32:
#             print("   ! Para actualizar automáticamente el índice, instala: pip install pywin32")
#             return

#         pythoncom.CoInitialize()
#         try:
#             word_app = win32_client.Dispatch("Word.Application")
#             word_app.Visible = False
#             word_app.ScreenUpdating = False
#             word_app.DisplayAlerts = False

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
#                 word_app.Quit()
#                 return

#             try:
#                 count_fields = doc.Fields.Count
#                 if count_fields > 0:
#                     for i in range(1, count_fields + 1):
#                         try:
#                             fld = doc.Fields(i)
#                             t = fld.Type
#                             if win32_constants and t in (
#                                 win32_constants.wdFieldTOC,
#                                 win32_constants.wdFieldIndex,
#                                 win32_constants.wdFieldTOA,
#                             ):
#                                 continue
#                             try:
#                                 fld.Update()
#                             except Exception:
#                                 pass
#                         except Exception:
#                             pass
#                     print("   * Campos actualizados correctamente")
#                 else:
#                     print("   * No hay campos para actualizar")
#             except Exception as e:
#                 print(f"   ! Error al actualizar campos: {e}")

#             try:
#                 doc.Save()
#             except Exception as e:
#                 print(f"   ! Error al guardar: {e}")
#             finally:
#                 try:
#                     doc.Close(SaveChanges=False)
#                 except Exception:
#                     pass

#             word_app.Quit()
#             print("   * Proceso completado")

#         finally:
#             pythoncom.CoUninitialize()

#     except Exception as e:
#         print(f"   ! Error general al actualizar índice: {e}")
#         print("   * El documento se generó correctamente, solo falló la actualización del índice")


# Crear contexto para la plantilla
print("-> Renderizando documentos...")

centros = []
if GROUP_COLUMN in df_envol.columns:
    df_envol_filtrado = df_envol.copy()
    # Remover posibles filas totales (asumimos última si GROUP_COLUMN vacío)
    df_envol_filtrado = df_envol_filtrado[
        df_envol_filtrado[GROUP_COLUMN].notna()
        & (df_envol_filtrado[GROUP_COLUMN].astype(str).str.strip() != "")
    ]
    centros = sorted(df_envol_filtrado[GROUP_COLUMN].unique())
else:
    print(f"   ! No se encontró la columna {GROUP_COLUMN} en la hoja Envol")

for centro in centros:
    df_envol_centro = df_envol[df_envol[GROUP_COLUMN] == centro].copy()
    totales_envol = get_totales_centro(df_envol, df_envol_centro, "Envolvente")

    # Ajustar etiqueta de la primera columna en totales si corresponde
    label_key = "EDIFICIO" if "EDIFICIO" in df_envol.columns else GROUP_COLUMN
    if label_key in totales_envol:
        totales_envol[label_key] = f"Total {centro}"

    context = {
        "mes": mes_nombre,
        "anio": anio,
        "centro": centro,
        "df_envol": df_envol_centro.to_dict("records"),
        "totales_envol": [totales_envol],
    }

    doc = DocxTemplate(TEMPLATE_DOC)
    doc.render(context)

    try:
        nombre_centro = clean_filename(centro)
        output_file = f"Anexo 4 {nombre_centro}.docx"
        output_path = BASE_DIR.parent / "word" / "anexos" / output_file
        doc.save(str(output_path))
        print(f"* Documento generado: {output_file}")
    except PermissionError as e:
        print(f"   ! Error de permisos con {output_file}: {e}")
        print("   ! Saltando este archivo...")
        continue
    except Exception as e:
        print(f"   ! Error inesperado con {output_file}: {e}")
        continue

print("\nTodos los documentos generados correctamente.")


# Mensaje final sobre archivos generados
print(f"\n{'=' * 60}")
print("PROCESO COMPLETADO")
print(f"{'=' * 60}")
print("Los documentos se encuentran en:")
print(f"  {BASE_DIR.parent / 'word' / 'anexos'}")
print("\nSi algún archivo no se generó debido a errores de permisos,")
print("cierra Word completamente y vuelve a ejecutar el script.")
print(f"{'=' * 60}")
