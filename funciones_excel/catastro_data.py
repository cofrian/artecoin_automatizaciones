import xlwings as xw
import pandas as pd

def completar_tabla_excel():
    # Abre el libro activo de Excel
    wb = xw.Book.caller()
    sht = wb.sheets["CEE"]  # O puedes usar el nombre: wb.sheets['Hoja1']

     # Obtén la tabla de Excel por su nombre
    tabla = sht.api.ListObjects("Tabla1")

    # Obtén el rango del cuerpo de la tabla (sin encabezados)
    data_range = tabla.DataBodyRange
    nrows = data_range.Rows.Count
    ncols = data_range.Columns.Count

    # Obtén el rango total (incluyendo encabezados)
    header_range = tabla.HeaderRowRange
    headers = [cell.Value for cell in header_range.Columns]

    # Lee los valores en una lista de listas
    values = data_range.Value
    if nrows == 1:
        values = [values]   # Por si la tabla solo tiene 1 fila

    # Construye el DataFrame
    df = pd.DataFrame(values, columns=headers)
    df.to_excel("tabla_catastro.xlsx", index=False)