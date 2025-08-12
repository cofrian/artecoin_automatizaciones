from __future__ import annotations

from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
from io import BytesIO
import time
import subprocess
import re
import unicodedata

# Hacer opcional pywin32
try:
    import win32com.client as win32_client
    import pythoncom
    from win32com.client import constants
except Exception:
    win32_client = None
    pythoncom = None
    constants = None

# --------------------------------------------------------------------
# 1) CONFIGURACIÓN ----------------------------------------------------
# --------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = (
    BASE_DIR.parent
    / "excel/proyecto/ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1_V20.xlsx"
)
TEMPLATE_DOC = BASE_DIR.parent / "word/anexos/Plantilla_Anexo_3.docx"
# Cachear la plantilla en memoria para reuso (más rápido que leer de disco cada vez)
TEMPLATE_BYTES = Path(TEMPLATE_DOC).read_bytes()
OUTPUT_PATH = BASE_DIR / "ANEXO_3.docx"

SHEET_MAP = {
    "Clima": "SISTEMAS DE CLIMATIZACIÓN",
    "SistCC": "SISTEMAS DE CALEFACCIÓN",
    "Eleva": "EQUIPOS ELEVADORES",
    "EqHoriz": "EQUIPOS HORIZONTALES",
    "Ilum": "SISTEMAS DE ILUMINACIÓN",
    "OtrosEq": "OTROS EQUIPOS",
}

HEADER_ROW = 0  # primera fila
SKIP_ROWS = None
GROUP_COLUMN = "CENTRO"  # Columna utilizada para agrupar y generar anexos

# --------------------------------------------------------------------
# 2) LIMPIEZA DE DATAFRAMES -------------------------------------------
# --------------------------------------------------------------------


def delete_rows_optimized(df, columna="ID EDIFICIO"):
    """Elimina las filas vacías al final del DataFrame."""
    if columna in df.columns:
        # Encontrar la última fila con datos válidos
        for i in range(len(df) - 1, -1, -1):
            val = df[columna].iloc[i]
            if pd.notna(val) and str(val).strip() != "" and str(val).strip() != "0":
                return df.iloc[: i + 2].copy()
    return df.copy()


def clean_last_row(df):
    """
    Limpia la última fila, que se corresponde con la de totales
    """
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


def update_word_fields_bulk(doc_paths: list[str], toc_mode: str = "pn"):
    """
    Actualiza campos en **lote** usando una sola instancia de Word.
    - Campos normales: se actualizan uno a uno saltando TOC/Index/TOA.
    - TOC (Tabla de contenido): si toc_mode == "pn", solo UpdatePageNumbers(); si "full", Update().
    """
    if win32_client is None or pythoncom is None:
        print("   ! pywin32 no disponible. Omite actualización de campos.")
        return
    pythoncom.CoInitialize()
    try:
        word_app = win32_client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.ScreenUpdating = False
        word_app.DisplayAlerts = False

        for doc_path in doc_paths:
            doc_path = str(doc_path)
            if not Path(doc_path).exists():
                print(f"   ! No existe: {doc_path}")
                continue
            try:
                doc = word_app.Documents.Open(
                    doc_path,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                )
            except Exception as e:
                print(f"   ! No se pudo abrir: {doc_path} -> {e}")
                continue

            # 1) Actualizar campos no-TOC/Index/TOA
            try:
                fields_count = doc.Fields.Count
                if fields_count > 0 and constants is not None:
                    for i in range(1, fields_count + 1):
                        try:
                            fld = doc.Fields(i)
                            t = fld.Type
                            if t in (
                                constants.wdFieldTOC,
                                constants.wdFieldIndex,
                                constants.wdFieldTOA,
                            ):
                                continue
                            fld.Update()
                        except Exception:
                            pass
            except Exception as e:
                print(f"   ! Error al actualizar campos normales en {doc_path}: {e}")

            # 2) TOC: solo paginación (rápido) o actualización completa si se pide
            try:
                toc_count = doc.TablesOfContents.Count
                if toc_count > 0:
                    for i in range(1, toc_count + 1):
                        toc = doc.TablesOfContents(i)
                        try:
                            if toc_mode == "full":
                                toc.Update()
                            else:
                                toc.UpdatePageNumbers()
                        except Exception:
                            pass
            except Exception as e:
                print(f"   ! Error al actualizar TOC en {doc_path}: {e}")

            # Guardar y cerrar
            try:
                doc.Save()
                doc.Close(SaveChanges=False)
            except Exception as e:
                print(f"   ! Error al guardar/cerrar {doc_path}: {e}")
        # Cerrar Word
        word_app.Quit()
    finally:
        pythoncom.CoUninitialize()


def clean_filename(filename):
    """Limpia el nombre del archivo eliminando caracteres no válidos y tildes para Windows."""
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


def get_totales_centro(
    df_full: pd.DataFrame, df_grupo: pd.DataFrame, nombre_seccion: str
) -> dict[str, str]:
    """Calcula totales por grupo (antes edificio) usando la última fila como referencia.

    df_full: DataFrame completo (incluye fila final con totales globales precalculados).
    df_grupo: Subconjunto filtrado para un valor concreto de GROUP_COLUMN.
    nombre_seccion: No se usa actualmente pero se mantiene por compatibilidad/futuras mejoras.
    """
    if df_full.empty:
        return {}

    ultima_full = df_full.iloc[-1]
    cols_a_sumar = [
        col
        for col in df_full.columns[1:]
        if pd.notna(ultima_full[col]) and str(ultima_full[col]).strip() != ""
    ]

    # Clave descriptiva; si la plantilla aún espera 'EDIFICIO' la mantenemos.
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

        # Intentar cerrar Word de forma elegante si pywin32 está disponible
        if win32_client is not None and pythoncom is not None:
            pythoncom.CoInitialize()
            try:
                try:
                    word_app = win32_client.GetActiveObject("Word.Application")
                    if word_app.Documents.Count > 0:
                        print(
                            f"   Cerrando {word_app.Documents.Count} documentos abiertos..."
                        )
                        for doc in word_app.Documents:
                            doc.Close(SaveChanges=False)
                    word_app.Quit()
                    print("   * Word cerrado correctamente")
                except Exception:
                    # Si no hay instancia activa, continuar
                    pass
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

            # Usar el año actual como referencia para el rango permitido
            if current_year - 5 <= anio <= current_year + 5:  # Rango razonable de años
                break
            else:
                print(
                    f"Error: El año debe estar entre {current_year - 5} y {current_year + 5}"
                )
        except ValueError:
            print("Error: Por favor ingrese un año válido")

    print("\nConfiguración seleccionada:")
    print(f"   Mes: {mes_nombre} ({mes_num})")
    print(f"   Año: {anio}")
    print("=" * 50)

    return mes_nombre, anio


# Obtener datos del usuario
def create_anex(
    BASE_DIR,
    EXCEL_PATH,
    TEMPLATE_BYTES,
    SHEET_MAP,
    load_and_clean_sheets,
    update_word_fields_bulk,
    clean_filename,
    get_totales_centro,
    cerrar_word_procesos,
    get_user_input,
):
    mes_nombre, anio = get_user_input()

    # NUEVO: Cerrar todos los procesos de Word antes de empezar
    cerrar_word_procesos()

    # Cargar y limpiar datos
    print("-> Cargando datos del Excel...")
    all_dataframes = load_and_clean_sheets(EXCEL_PATH, SHEET_MAP)

    print("* Datos cargados y limpiados")

    # Crear contexto para la plantilla
    print("-> Renderizando documentos...")

    generated_docs = []

    # Verificar existencia de la columna de agrupación en al menos una hoja
    hojas_con_grupo = [
        k for k, df in all_dataframes.items() if GROUP_COLUMN in df.columns
    ]
    if not hojas_con_grupo:
        print(
            f"   ! No se encontró la columna '{GROUP_COLUMN}' en ninguna hoja. Se detiene el proceso."
        )
        return

    # Obtener el conjunto total de grupos presentes en cualquier hoja
    grupos_totales: set[str] = set()
    for df in all_dataframes.values():
        if GROUP_COLUMN in df.columns:
            grupos_totales.update(df[GROUP_COLUMN].dropna().unique().tolist())

    # Remover posibles valores vacíos
    grupos_totales = {c for c in grupos_totales if str(c).strip() not in ("", "nan")}

    if not grupos_totales:
        print(f"   ! No hay valores de {GROUP_COLUMN} válidos para procesar.")
        return

    print(
        f"-> Se generarán documentos para {len(grupos_totales)} {GROUP_COLUMN.lower()}s"
    )

    for centro in sorted(grupos_totales):
        full_clima = all_dataframes["Clima"]
        full_sist_cc = all_dataframes["SistCC"]
        full_eleva = all_dataframes["Eleva"]
        full_eqhoriz = all_dataframes["EqHoriz"]
        full_ilum = all_dataframes["Ilum"]
        full_otros = all_dataframes["OtrosEq"]

        # Filtrar por centro (las filas totales globales no suelen tener CENTRO, por lo que no se incluyen)
        df_clima_centro = full_clima[full_clima.get(GROUP_COLUMN) == centro].copy()
        df_sist_cc_centro = full_sist_cc[
            full_sist_cc.get(GROUP_COLUMN) == centro
        ].copy()
        df_eleva_centro = full_eleva[full_eleva.get(GROUP_COLUMN) == centro].copy()
        df_eqhoriz_centro = full_eqhoriz[
            full_eqhoriz.get(GROUP_COLUMN) == centro
        ].copy()
        df_ilum_centro = full_ilum[full_ilum.get(GROUP_COLUMN) == centro].copy()
        df_otros_eq_centro = full_otros[full_otros.get(GROUP_COLUMN) == centro].copy()

        # Saltar si todas las tablas están vacías para este centro
        if all(
            len(df_) == 0
            for df_ in [
                df_clima_centro,
                df_sist_cc_centro,
                df_eleva_centro,
                df_eqhoriz_centro,
                df_ilum_centro,
                df_otros_eq_centro,
            ]
        ):
            continue

        # Calcular totales (reutilizamos función existente). Etiquetamos para indicar centro.
        totales_clima = get_totales_centro(full_clima, df_clima_centro, "Climatización")
        totales_sist_cc = get_totales_centro(
            full_sist_cc, df_sist_cc_centro, "Calefacción"
        )
        totales_eleva = get_totales_centro(full_eleva, df_eleva_centro, "Elevadores")
        totales_eqhoriz = get_totales_centro(
            full_eqhoriz, df_eqhoriz_centro, "H. Horizontales"
        )
        totales_ilum = get_totales_centro(full_ilum, df_ilum_centro, "Iluminación")
        totales_otros_eq = get_totales_centro(
            full_otros, df_otros_eq_centro, "Otros Equipos"
        )

        # Ajustar etiqueta de la primera columna para reflejar el centro
        # Etiquetar la clave descriptiva utilizada en totales (EDIFICIO o GROUP_COLUMN)
        label_key = "EDIFICIO" if "EDIFICIO" in full_clima.columns else GROUP_COLUMN
        for tot in [
            totales_clima,
            totales_sist_cc,
            totales_eleva,
            totales_eqhoriz,
            totales_ilum,
            totales_otros_eq,
        ]:
            if label_key in tot:
                tot[label_key] = f"Total {centro}"

        context = {
            "mes": mes_nombre,
            "anio": anio,
            "centro": centro,
            "df_clima": df_clima_centro.to_dict("records"),
            "df_sist_cc": df_sist_cc_centro.to_dict("records"),
            "df_eleva": df_eleva_centro.to_dict("records"),
            "df_eqhoriz": df_eqhoriz_centro.to_dict("records"),
            "df_ilum": df_ilum_centro.to_dict("records"),
            "df_otros_eq": df_otros_eq_centro.to_dict("records"),
            "totales_clima": [totales_clima],
            "totales_sist_cc": [totales_sist_cc],
            "totales_eleva": [totales_eleva],
            "totales_eqhoriz": [totales_eqhoriz],
            "totales_ilum": [totales_ilum],
            "totales_otros_eq": [totales_otros_eq],
        }

        doc = DocxTemplate(BytesIO(TEMPLATE_BYTES))
        doc.render(context)

        try:
            nombre_centro = clean_filename(str(centro))
            output_file = f"Anexo 3 {nombre_centro}.docx"
            output_path = BASE_DIR.parent / "word" / "anexos" / output_file
            doc.save(str(output_path))
            print(f"* Documento generado: {output_file}")
            generated_docs.append(str(output_path))
        except PermissionError as e:
            print(f"   ! Error de permisos con {output_file}: {e}")
            print("   ! Saltando este archivo...")
            continue
        except Exception as e:
            print(f"   ! Error inesperado con {output_file}: {e}")
            continue

    print("\nActualizando campos en lote (TOC: solo paginación)...")
    update_word_fields_bulk(generated_docs, toc_mode="pn")
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


create_anex(
    BASE_DIR,
    EXCEL_PATH,
    TEMPLATE_BYTES,
    SHEET_MAP,
    load_and_clean_sheets,
    update_word_fields_bulk,
    clean_filename,
    get_totales_centro,
    cerrar_word_procesos,
    get_user_input,
)


# Crear documento Word
def update_word_fields(doc_path):
    """Actualiza los campos del documento Word."""
    if win32_client is None or pythoncom is None:
        print(
            "   ! Para actualizar automáticamente el índice, instala: pip install pywin32"
        )
        return

    try:
        pythoncom.CoInitialize()
        try:
            word_app = win32_client.Dispatch("Word.Application")
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
                    print("   * Campos actualizados correctamente")
                else:
                    print("   * No hay campos para actualizar")
            except Exception as e:
                print(f"   ! Error al actualizar campos: {e}")

            try:
                doc.Save()
                doc.Close(SaveChanges=False)
            except Exception as e:
                print(f"   ! Error al guardar: {e}")
                doc.Close(SaveChanges=False)

            word_app.Quit()
            print("   * Proceso completado")
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        print(f"   ! Error general al actualizar índice: {e}")
        print(
            "   * El documento se generó correctamente, solo falló la actualización del índice"
        )
