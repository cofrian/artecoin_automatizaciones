# -*- coding: utf-8 -*-
# ------------------------------------------------------------
# Lógica de exportación CEE (pares izquierda / impares derecha)
# ------------------------------------------------------------
from pathlib import Path
import re
import xlwings as xw
try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfWriter, PdfReader
    except ImportError:
        print("Warning: pypdf o PyPDF2 no están instalados. No se podrán combinar PDFs.")
        PdfWriter = PdfReader = None

# ====== Parámetros por defecto ======
HOJA_DEF = "EE"
SLICER_LEGALIZACION = "SegmentaciónDeDatos_TIPO_LEGALIZACIÓN"
SLICER_PARES        = "SegmentaciónDeDatos_CM"    # izquierda
SLICER_IMPARES      = "SegmentaciónDeDatos_CM1"   # derecha
FULL_PRINT_FALLBACK = "L5:AE41"                   # área completa (2 etiquetas)
SINGLE_RIGHT_RANGE  = "AB5:AE41"                  # solo etiqueta derecha

# ====== Constantes Excel ======
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_LANDSCAPE = 2
XL_PAPERSIZE_A4 = 9

# ====== Helpers ======
def _sanitize(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip()

def _resolve_print_area(ws, fallback=FULL_PRINT_FALLBACK):
    """Devuelve el área de impresión por nombre 'Área_de_impresión' o fallback."""
    try:
        rng = ws.api.Range("Área_de_impresión")
        return rng.Address  # incluye $
    except Exception:
        return ws.range(fallback).address.replace("$", "")

def _apply_pagesetup(ws, print_area_address):
    ps = ws.api.PageSetup
    ps.Orientation = XL_LANDSCAPE
    ps.PaperSize = XL_PAPERSIZE_A4
    ps.CenterHorizontally = True
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.PrintArea = print_area_address

def _get_cache(wb, name_or_source):
    key = str(name_or_source).lower()
    for sc in wb.api.SlicerCaches:
        if key in str(sc.Name).lower() or key == str(sc.SourceName).lower():
            return sc
    raise KeyError(f"No encuentro SlicerCache '{name_or_source}'")

def _current_selected_item(sc):
    for si in sc.SlicerItems:
        if si.Selected:
            return si
    return None

def _select_single(sc, value_text):
    prev = _current_selected_item(sc)
    target = None
    for si in sc.SlicerItems:
        if str(si.Name).casefold() == str(value_text).casefold():
            target = si
            break
    if target is None:
        raise ValueError(f"'{value_text}' no existe en slicer '{sc.Name}'")
    if prev is not None and prev.Name == target.Name:
        return
    if prev is not None:
        prev.Selected = False
    target.Selected = True

def _list_items(sc, only_with_data=True):
    items = []
    for si in sc.SlicerItems:
        if (not only_with_data) or bool(getattr(si, "HasData", True)):
            items.append(str(si.Name))
    return items

def _init_single_first(sc):
    """Deja solo el primer ítem con datos en True, resto en False."""
    first = None
    for si in sc.SlicerItems:
        if getattr(si, "HasData", True):
            if first is None:
                si.Selected = True
                first = si
            else:
                si.Selected = False

def _combinar_pdfs_por_legalizacion(base_out: Path):
    """
    Combina todos los PDFs de cada carpeta de legalización en un único PDF
    y los guarda en la carpeta 'ARCHIVOS ETIQUETAS FINALES'.
    """
    if PdfWriter is None or PdfReader is None:
        print("No se puede combinar PDFs: pypdf o PyPDF2 no está instalado")
        return
    
    carpeta_final = base_out / "ARCHIVOS ETIQUETAS FINALES"
    carpeta_final.mkdir(parents=True, exist_ok=True)
    
    # Buscar todas las carpetas de legalización
    for carpeta_leg in base_out.iterdir():
        if carpeta_leg.is_dir() and carpeta_leg.name != "ARCHIVOS ETIQUETAS FINALES":
            # Buscar todos los PDFs en esta carpeta
            pdfs = sorted(list(carpeta_leg.glob("*.pdf")))
            
            if not pdfs:
                continue
                
            print(f"Combinando {len(pdfs)} PDFs de la legalización: {carpeta_leg.name}")
            
            # Crear el PDF combinado
            pdf_writer = PdfWriter()
            
            for pdf_path in pdfs:
                try:
                    with open(pdf_path, 'rb') as pdf_file:
                        pdf_reader = PdfReader(pdf_file)
                        for page_num in range(len(pdf_reader.pages)):
                            page = pdf_reader.pages[page_num]
                            pdf_writer.add_page(page)
                except Exception as e:
                    print(f"Error al procesar {pdf_path.name}: {e}")
                    continue
            
            # Guardar el PDF combinado
            nombre_final = f"{carpeta_leg.name}.pdf"
            ruta_final = carpeta_final / nombre_final
            
            try:
                with open(ruta_final, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)
                print(f"PDF combinado guardado: {nombre_final}")
            except Exception as e:
                print(f"Error al guardar PDF combinado {nombre_final}: {e}")
    
    print(f"Proceso de combinación completado. PDFs finales en: {carpeta_final}")

# ====== Función principal ======
def exportar_cee(xlsx_path: str, out_dir: str):
    """
    Exporta PDFs CEE controlando slicers:
    - Recorre todas las legalizaciones visibles.
    - CM pares (izquierda) vs CM impares (derecha) por página.
    - Al cambiar de legalización, resetea ambos CM (solo primer valor).
    - Si la cantidad de CM es impar, la última página imprime solo la etiqueta derecha.
    - Excel oculto (rápido) y PDFs guardados en subcarpetas por legalización.
    """
    xlsx_path = str(xlsx_path)
    base_out = Path(out_dir); base_out.mkdir(parents=True, exist_ok=True)

    # Excel oculto
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(xlsx_path)
        ws = wb.sheets[HOJA_DEF] if HOJA_DEF else wb.sheets.active

        full_area = _resolve_print_area(ws, fallback=FULL_PRINT_FALLBACK)
        _apply_pagesetup(ws, full_area)
        half_area = ws.range(SINGLE_RIGHT_RANGE).address.replace("$", "")

        sc_leg = _get_cache(wb, SLICER_LEGALIZACION)
        sc_pares = _get_cache(wb, SLICER_PARES)
        sc_impares = _get_cache(wb, SLICER_IMPARES)

        _init_single_first(sc_leg)
        _init_single_first(sc_pares)
        _init_single_first(sc_impares)
        app.api.Calculate()

        legalizaciones = _list_items(sc_leg)
        if not legalizaciones:
            raise RuntimeError("No hay elementos visibles en el slicer de legalización.")

        for leg in legalizaciones:
            _select_single(sc_leg, leg)
            app.api.Calculate()

            _init_single_first(sc_pares)
            _init_single_first(sc_impares)
            app.api.Calculate()

            base = _list_items(sc_pares, only_with_data=True)
            if not base:
                continue

            impares = base[0::2]  # 1,3,5,...
            pares   = base[1::2]  # 2,4,6,...

            out_leg = base_out / _sanitize(leg)
            out_leg.mkdir(parents=True, exist_ok=True)

            n = max(len(pares), len(impares))
            total_items = len(base)
            for k in range(n):
                left  = pares[k]   if k < len(pares)   else pares[-1]   if pares   else base[0]
                right = impares[k] if k < len(impares) else impares[-1] if impares else base[-1]

                if left == right and total_items > 1:
                    for alt in base:
                        if alt != left:
                            right = alt
                            break

                _select_single(sc_pares, left)
                _select_single(sc_impares, right)
                app.api.Calculate()

                is_last_iter = (k == n - 1)
                if (total_items % 2 == 1) and is_last_iter:
                    ws.api.PageSetup.PrintArea = half_area
                else:
                    ws.api.PageSetup.PrintArea = full_area

                pdf = out_leg / f"{_sanitize(leg)}__PAR-{_sanitize(left)}__IMPAR-{_sanitize(right)}.pdf"
                ws.api.ExportAsFixedFormat(
                    Type=XL_TYPE_PDF,
                    Filename=str(pdf),
                    Quality=XL_QUALITY_STANDARD,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )

            ws.api.PageSetup.PrintArea = full_area

        wb.close()
        
        # Combinar PDFs por legalización después de crear todos los individuales
        print("Iniciando combinación de PDFs por legalización...")
        _combinar_pdfs_por_legalizacion(base_out)
        
    finally:
        app.display_alerts = True
        app.screen_updating = True
        app.quit()
