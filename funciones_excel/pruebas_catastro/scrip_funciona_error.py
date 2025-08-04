import xlwings as xw
import os

def cargar_datos():
    wb = xw.Book.caller()
    ws_dest = wb.sheets['CM']
    main_name = 'Tabla3'
    comp_name = 'Tabla_comp_CM'

    def leer_ruta(nd):
        try:
            v = wb.names[nd].refers_to
            if v.startswith('='):
                v = v[1:]
            return v.replace('"','') if v else None
        except:
            return None

    # 1) Backup de CM
    copia = ws_dest.copy(before=ws_dest)
    copia.name = '_CM_backup'

    # 2) Borrar filas bajo Tabla3
    tbl = ws_dest.api.ListObjects[main_name]
    fila_ini = tbl.Range.Row
    nfilas   = tbl.Range.Rows.Count
    fila_fin = fila_ini + nfilas - 1
    ultima   = ws_dest.cells.last_cell.row
    if ultima > fila_fin:
        ws_dest.api.Rows(f"{fila_fin+1}:{ultima}").Delete()

    # 3) Desactivar alertas y abrir con xw.Book
    ruta = leer_ruta("RutaProducto")
    if not ruta or not os.path.isfile(ruta):
        print(f"❌ Archivo no válido: {ruta}")
        return

    app_com = wb.app.api
    app_com.DisplayAlerts    = False
    app_com.AskToUpdateLinks = False

    libro_origen = xw.Book(ruta)  # ahora es xlwings Book
    hoja_origen = libro_origen.sheets['Centro_mando']

    # 4) Mapeo de columnas
    columnas = {
        'Coord.X (m)': 'coord.X (m)',
        'Coord.Y (m)': 'coord.X (m)',
        # … resto de tu mapeo …
        'medido':       'medido'
    }

    # 5) Desactivar totales en Tabla3
    try:
        tot_prev = tbl.ShowTotals
        tbl.ShowTotals = False
    except:
        tot_prev = False

    encabezados = ws_dest.range("A1").expand('right').value
    max_filas   = 0

    # 6) Volcar datos
    for orig, dest in columnas.items():
        try:
            idx_o = hoja_origen.range("1:1").value.index(orig) + 1
            idx_d = encabezados.index(dest) + 1
            datos = hoja_origen.range((2, idx_o)).expand('down').value or []
            if not isinstance(datos, list):
                datos = [datos]
            filtrados = [
                v for i, v in enumerate(datos, start=2)
                if not hoja_origen.cells(i, idx_o).api.HasFormula
            ]
            if filtrados:
                ws_dest.range((2, idx_d)).options(transpose=True).value = filtrados
            max_filas = max(max_filas, len(filtrados))
        except Exception as e:
            print(f"⚠️ {orig}→{dest}: {e}")

    # 7) Redimensionar Tabla3 en altura
    hdr_rng = ws_dest.range("A1").expand('right')
    nueva_h = max_filas + 1
    tbl.Resize(hdr_rng.resize(nueva_h).api)

    # 8) Reactivar totales
    if tot_prev:
        tbl.ShowTotals = True

    # 9) Copiar Tabla_comp_CM desde backup
    try:
        bkp       = wb.sheets['_CM_backup']
        comp_lo   = bkp.api.ListObjects[comp_name]
        comp_rng  = comp_lo.Range
        vals      = comp_rng.Value

        main_lo   = ws_dest.api.ListObjects[main_name]
        start_row = main_lo.Range.Row + main_lo.Range.Rows.Count
        first_col = comp_rng.Columns(1).Column
        last_col  = comp_rng.Columns(comp_rng.Columns.Count).Column

        dest_rng = ws_dest.range((start_row+1, first_col),
                                 (start_row+len(vals), last_col))
        dest_rng.value = vals

        new_lo = ws_dest.api.ListObjects.Add(
            SourceType             = 0,
            Source                 = dest_rng.api,
            XlListObjectHasHeaders = 1
        )
        new_lo.Name = comp_name
    except Exception as e:
        print(f"⚠️ Error copiando '{comp_name}': {e}")

    # 10) Cerrar origen y restaurar alertas
    libro_origen.close()
    app_com.DisplayAlerts    = True
    app_com.AskToUpdateLinks = True

    print(f"✅ Carga completada: {max_filas} filas y '{comp_name}' copiada.")

