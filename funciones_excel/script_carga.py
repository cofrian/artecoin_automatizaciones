import xlwings as xw
import os

def cargar_datos():
    wb = xw.Book.caller()
    hoja_destino = wb.sheets['CM']
    nombre_tabla = 'Tabla3'

    def leer_ruta(nombre_definido):
        try:
            valor = wb.names[nombre_definido].refers_to
            if valor.startswith('='):
                valor = valor[1:]
            return valor.replace('"', '') if valor else None
        except:
            return None

    # 1) Copia de seguridad de la hoja CM
    copia = hoja_destino.copy(before=hoja_destino)
    copia.name = '_CM_backup'

    # 2) Borrar todo lo que haya debajo de la tabla antes de la carga
    tabla = hoja_destino.api.ListObjects(nombre_tabla)
    fila_inicio = tabla.Range.Row
    num_filas = tabla.Range.Rows.Count
    fila_fin = fila_inicio + num_filas - 1
    ultima_fila = hoja_destino.cells.last_cell.row
    if ultima_fila > fila_fin:
        hoja_destino.api.Rows(f"{fila_fin+1}:{ultima_fila}").Delete()

    # 3) Leer ruta y abrir libro origen
    ruta = leer_ruta("RutaProducto")
    if not ruta or not os.path.isfile(ruta):
        print(f"❌ Archivo no válido: {ruta}")
        return
    app_com = wb.app.api
    app_com.DisplayAlerts    = False
    app_com.AskToUpdateLinks = False

    libro_origen = xw.Book(ruta)
    hoja_origen = libro_origen.sheets['Centro_mando']

    # 4) Mapeo de columnas
    columnas = {
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
        'medido':       'medido'
    }

    # 5) Desactivar totales
    try:
        tot_prev = tabla.ShowTotals
        tabla.ShowTotals = False
    except:
        tot_prev = False

    encabezados = hoja_destino.range("A1").expand('right').value
    max_filas = 0

    # 6) Volcado de cada columna
    for orig, dest in columnas.items():
        try:
            idx_o = hoja_origen.range('1:1').value.index(orig) + 1
            idx_d = encabezados.index(dest) + 1
            datos = hoja_origen.range((2, idx_o)).expand('down').value or []
            if not isinstance(datos, list): datos = [datos]
            filt = [v for i, v in enumerate(datos, start=2)
                    if not hoja_origen.cells(i, idx_o).api.HasFormula]
            if filt:
                hoja_destino.range((2, idx_d)).options(transpose=True).value = filt
            max_filas = max(max_filas, len(filt))
        except Exception as e:
            print(f"⚠️ {orig}→{dest}: {e}")

    # 7) Redimensionar tabla solo en altura
    encabezado_rng = hoja_destino.range('A1').expand('right')
    nueva_alt = max_filas + 1
    nueva_rng = encabezado_rng.resize(nueva_alt)
    tabla.Resize(nueva_rng.api)

    # 8) Reactivar totales
    try:
        tabla.ShowTotals = tot_prev
    except:
        pass

    # 9) Copiar Tabla_comp_CM al final de Tabla3
    try:
        tabla_comp = hoja_destino.api.ListObjects('Tabla_comp_CM43')
        comp_rng = tabla_comp.DataBodyRange
        if comp_rng is not None:
            # Determinar punto de inserción: una fila abajo de fila_fin final
            fila_fin = tabla.Range.Row + nueva_alt - 1
            # Pegar valores de comp_rng en mismas columnas
            cols = comp_rng.Columns.Count
            for j in range(1, cols+1):
                col_idx = comp_rng.Columns(j).Column
                valores = hoja_destino.range(
                    comp_rng.DataBodyRange.Address).columns(j).value
                dest = hoja_destino.range(fila_fin+1, col_idx).resize(len(valores),1)
                dest.value = valores
    except Exception as e:
        print(f"⚠️ Error copiando Tabla_comp_CM: {e}")

    print(f"✅ Carga completada: {max_filas} filas y tabla_comp_CM copiada.")
    libro_origen.close()