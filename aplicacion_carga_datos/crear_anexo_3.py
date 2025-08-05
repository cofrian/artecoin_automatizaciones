from __future__ import annotations

from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime


# --------------------------------------------------------------------
# 1) CONFIGURACIÓN ----------------------------------------------------
# --------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = (
    BASE_DIR.parent
    / "excel/proyecto/ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1_V20.xlsx"
)
TEMPLATE_DOC = BASE_DIR.parent / "word/anejos/Plantilla_Anexo_3.docx"
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
            result[key] = df_cleaned

        return result


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
                print("❌ Error: El mes debe estar entre 1 y 12")
        except ValueError:
            print("❌ Error: Por favor ingrese un número válido")

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

            if 2020 <= anio <= 2030:  # Rango razonable de años
                break
            else:
                print("❌ Error: El año debe estar entre 2020 y 2030")
        except ValueError:
            print("❌ Error: Por favor ingrese un año válido")

    print("\n✅ Configuración seleccionada:")
    print(f"   Mes: {mes_nombre} ({mes_num})")
    print(f"   Año: {anio}")
    print("=" * 50)

    return mes_nombre, anio


# Obtener datos del usuario
mes_nombre, anio = get_user_input()

# Cargar y limpiar datos
print("\n-> Cargando datos del Excel...")
all_dataframes = load_and_clean_sheets(EXCEL_PATH, SHEET_MAP)

# Asignar a variables individuales
df_clima = all_dataframes["Clima"]
df_sist_cc = all_dataframes["SistCC"]
df_eleva = all_dataframes["Eleva"]
df_eqhoriz = all_dataframes["EqHoriz"]
df_ilum = all_dataframes["Ilum"]
df_otros_eq = all_dataframes["OtrosEq"]

print("* Datos cargados y limpiados")


# Crear documento Word
def update_word_fields(doc_path):
    """Actualiza los campos del documento Word de forma optimizada."""
    try:
        import win32com.client
        import pythoncom

        # Inicializar COM
        pythoncom.CoInitialize()

        try:
            # Intentar usar instancia existente de Word (más rápido)
            try:
                word_app = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                # Si no hay instancia, crear una nueva
                word_app = win32com.client.Dispatch("Word.Application")

            # Optimizaciones de rendimiento
            word_app.Visible = False
            word_app.ScreenUpdating = False
            word_app.DisplayAlerts = False  # Deshabilitar alertas

            # Abrir documento
            doc = word_app.Documents.Open(
                str(doc_path),
                ConfirmConversions=False,  # No confirmar conversiones
                ReadOnly=False,
                AddToRecentFiles=False,  # No añadir a archivos recientes
                Visible=False,
            )

            # Actualizar solo campos específicos (más rápido que UpdateFields general)
            doc.Range().Fields.Update()

            # Guardar sin crear backup
            doc.Save()
            doc.Close(SaveChanges=True)

            # Restaurar configuración
            word_app.ScreenUpdating = True
            word_app.DisplayAlerts = True

            # Solo cerrar Word si lo creamos nosotros
            if word_app.Documents.Count == 0:
                word_app.Quit()

            print("* Indice actualizado correctamente")

        finally:
            # Limpiar COM
            pythoncom.CoUninitialize()

    except ImportError:
        print(
            "! Para actualizar automaticamente el indice, instala: pip install pywin32"
        )
    except Exception as e:
        print(f"Error al actualizar indice: {e}")


# Crear contexto para la plantilla
print("-> Creando contexto...")
context = {
    "mes": mes_nombre,
    "anio": anio,
    "df_clima": df_clima.to_dict("records"),
    "df_sist_cc": df_sist_cc.to_dict("records"),
    "df_eleva": df_eleva.to_dict("records"),
    "df_eqhoriz": df_eqhoriz.to_dict("records"),
    "df_ilum": df_ilum.to_dict("records"),
    "df_otros_eq": df_otros_eq.to_dict("records"),
}

print("-> Renderizando documento...")
doc = DocxTemplate(TEMPLATE_DOC)
doc.render(context)

output_file = "Anexo 3.docx"
output_path = BASE_DIR.parent / "word" / "anejos" / output_file
doc.save(str(output_path))

print("-> Actualizando indice...")
update_word_fields(str(output_path))

print(f"* Documento generado: {output_path}")
