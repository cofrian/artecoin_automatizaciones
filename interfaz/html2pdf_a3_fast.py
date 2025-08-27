# -*- coding: utf-8 -*-
"""
HTML → PDF A3 (Chromium/Playwright) + merges por sección **POR CENTRO**, con render concurrente.

Salida:
  <data>_pdf/<CENTRO>/<seccion>/*.pdf
  <data>_pdf/<CENTRO>/_<SECCION>__MERGED.pdf
  <data>_pdf/<CENTRO>/Anejo5_Reportaje_Fotografico.pdf
"""

import argparse
import asyncio
import sys
from pathlib import Path
from urllib.parse import quote
from collections import defaultdict

from pypdf import PdfWriter
from playwright.async_api import async_playwright

SECCIONES = {
    "centro": "CENTRO",
    "edificios": "EDIFICIOS",
    "envolventes": "ENVOLVENTES",
    "dependencias": "DEPENDENCIAS",
    "acometida": "ACOMETIDA",
    "cc": "CC",
    "clima": "CLIMA",
    "eqhoriz": "EQ_HORIZONTALES",
    "elevadores": "ELEVADORES",
    "iluminacion": "ILUMINACION",
    "otros_equipos": "OTROS_EQUIPOS",
}

def to_file_uri(p: Path) -> str:
    p = p.resolve()
    s = str(p).replace("\\", "/")
    return "file:///" + quote(s, safe="/:._-()")

def find_htmls(root: Path) -> list[Path]:
    return sorted([p for p in root.rglob("*.html") if p.is_file()])

def out_root_for(in_root: Path) -> Path:
    return in_root.with_name(in_root.name + "_pdf")

def section_from_path(html_path: Path, data_root: Path) -> str:
    """
    .../<CENTRO>/<seccion>/archivo.html  -> <seccion>
    """
    try:
        rel = html_path.relative_to(data_root)
        parts = [p.lower() for p in rel.parts]
        for p in parts:
            if p in SECCIONES:
                return p
    except Exception:
        pass
    parent = html_path.parent.name.lower()
    return parent if parent in SECCIONES else "otros"

def center_from_path(html_path: Path, data_root: Path) -> str:
    """
    Primer nivel bajo data_root es el centro (CC0001, CC0002, ...).
    Si no hay, usa 'ROOT'.
    """
    try:
        rel = html_path.relative_to(data_root)
        return rel.parts[0] if len(rel.parts) > 1 else "ROOT"
    except Exception:
        return "ROOT"

async def render_one(page, html_path: Path, pdf_path: Path, scale: float, prefer_css: bool, wait_ms: int):
    url = to_file_uri(html_path)
    await page.goto(url, wait_until="load")
    await page.wait_for_timeout(wait_ms)
    await page.emulate_media(media="print")
    pdf_opts = {
        "path": str(pdf_path),
        "format": "A3",
        "landscape": True,
        "print_background": True,
        "margin": {"top": "0", "right": "0", "bottom": "0", "left": "0"},
        "scale": scale,
        "prefer_css_page_size": prefer_css,
    }
    await page.pdf(**pdf_opts)

def merge_pdfs(pdf_paths: list[Path], out_path: Path):
    writer = PdfWriter()
    for p in pdf_paths:
        try:
            writer.append(str(p))
        except Exception:
            pass
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("wb") as f:
        writer.write(f)

async def main_async(args):
    data_root = Path(args.data).resolve()
    out_root  = Path(args.out).resolve() if args.out else out_root_for(data_root)
    out_root.mkdir(parents=True, exist_ok=True)

    htmls = find_htmls(data_root)
    if not htmls:
        print("[INFO] No se encontraron .html en", data_root)
        return set()

    rendered = await render_htmls_to_pdfs(htmls, data_root, out_root, args)

    if args.no_merge:
        print("[OK] Conversión terminada. Sin merges por sección.")
        return {c for c, _, _ in rendered}

    # --- MERGES POR (CENTRO, SECCION) ---
    buckets = defaultdict(list)
    centros = set()
    for centro, section, pdf in rendered:
        centros.add(centro)
        buckets[(centro, section)].append(pdf)

    for (centro, section), paths in buckets.items():
        # orden reproducible
        paths = sorted(paths)
        name = f"_{SECCIONES.get(section, section).upper()}__MERGED.pdf"
        out_merge = out_root / centro / name
        print(f"[MERGE] {centro} · {section} -> {out_merge.name} ({len(paths)} PDFs)")
        try:
            merge_pdfs(paths, out_merge)
        except Exception as e:
            print(f"[WARN] Merge {centro}/{section} falló: {e}")

    return centros

# --- RENDER CONCURRENTE ---
async def render_htmls_to_pdfs(htmls, data_root, out_root, args):
    rendered = []
    total = len(htmls)
    done_counter = 0
    lock = asyncio.Lock()
    sem = asyncio.Semaphore(max(1, args.concurrency))

    block_types = set(t.strip().lower() for t in args.block.split(",") if t.strip())

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=(args.chromium_arg or [])
        )
        context = await browser.new_context()

        if block_types:
            async def route_handler(route):
                r = route.request
                if r.resource_type in block_types:
                    await route.abort()
                else:
                    await route.continue_()
            await context.route("**/*", route_handler)

        async def worker(i, html):
            nonlocal done_counter
            rel = html.relative_to(data_root)
            section = section_from_path(html, data_root)
            centro  = center_from_path(html, data_root)

            # Guardamos el PDF junto a la estructura del centro/sección
            dst_dir = out_root / rel.parent
            dst_dir.mkdir(parents=True, exist_ok=True)
            pdf_name = html.with_suffix(".pdf").name
            pdf_path = dst_dir / pdf_name

            async with sem:
                page = await context.new_page()
                try:
                    print(f"[{i}/{total}] → {pdf_path}")
                    await render_one(page, html, pdf_path,
                                     scale=args.scale,
                                     prefer_css=not args.ignore_css_page,
                                     wait_ms=args.wait)
                    async with lock:
                        done_counter += 1
                        if (done_counter % args.log_every) == 0 or done_counter == total:
                            print(f"[PROGRESO] {done_counter}/{total} completados")
                    return (centro, section, pdf_path)
                except Exception as e:
                    async with lock:
                        print(f"[WARN] Falló {html}: {e}")
                    return None
                finally:
                    await page.close()

        tasks = [asyncio.create_task(worker(i, html)) for i, html in enumerate(htmls, 1)]
        for coro in asyncio.as_completed(tasks):
            res = await coro
            if res:
                rendered.append(res)

        await context.close()
        await browser.close()

    rendered.sort(key=lambda tup: str(tup[2]))
    return rendered

def parse_args():
    ap = argparse.ArgumentParser(description="HTML → PDF A3 (Chromium/Playwright) + merges por sección **por centro**")
    ap.add_argument("--data", required=True, help="Carpeta raíz con los .html generados")
    ap.add_argument("--out", default=None, help="Carpeta de salida (default: <data>_pdf)")
    ap.add_argument("--scale", type=float, default=1.0, help="Escala del render (1.0 por defecto)")
    ap.add_argument("--wait", type=int, default=300, help="Espera (ms) tras cargar cada HTML (por defecto 300)")
    ap.add_argument("--no-merge", action="store_true", help="No crear PDFs combinados por sección")
    ap.add_argument("--ignore-css-page", action="store_true",
                    help="Ignorar @page del CSS y forzar A3 landscape desde Chromium")

    ap.add_argument("--concurrency", type=int, default=4, help="Número de páginas en paralelo (default 4)")
    ap.add_argument("--log-every", type=int, default=10, help="Frecuencia de logs de progreso (default 10)")
    ap.add_argument("--block", default="", help="Tipos de recurso a bloquear separados por coma: image,font,media,stylesheet,script")
    ap.add_argument("--chromium-arg", action="append", default=[], help="Argumentos extra para Chromium (repetible)")

    ap.add_argument("--caratulas-dir", default=None, help="Ruta a las carátulas del Anejo (PDFs)")
    return ap.parse_args()

def crear_anejo_final_por_centro(args, out_root: Path, centros: set[str]):
    """
    Genera un Anejo5_Reportaje_Fotografico.pdf por cada centro dentro de <out_root>/<CENTRO>.
    Usa las mismas carátulas para todos (si existen).
    """
    from pypdf import PdfWriter

    orden_secciones = [
        ("PORTADA.pdf", None),
        ("CENTRO.pdf", "_CENTRO__MERGED.pdf"),
        ("EDIFICIO.pdf", "_EDIFICIOS__MERGED.pdf"),
        ("DEPENDENCIAS.pdf", "_DEPENDENCIAS__MERGED.pdf"),
        ("ACOM.pdf", "_ACOMETIDA__MERGED.pdf"),
        ("ENVOL.pdf", "_ENVOLVENTES__MERGED.pdf"),
        ("CALEFACCION.pdf", "_CC__MERGED.pdf"),
        ("CLIMA.pdf", "_CLIMA__MERGED.pdf"),
        ("EQHORIZ.pdf", "_EQ_HORIZONTALES__MERGED.pdf"),
        ("ELEVA.pdf", "_ELEVADORES__MERGED.pdf"),
        ("ILUM.pdf", "_ILUMINACION__MERGED.pdf"),
        ("OTROSEQ.pdf", "_OTROS_EQUIPOS__MERGED.pdf"),
    ]
    orden_merged_extra = [
        "_DEPENDENCIAS__MERGED.pdf",
        "_EDIFICIOS__MERGED.pdf",
        "_ELEVADORES__MERGED.pdf",
        "_ENVOLVENTES__MERGED.pdf",
        "_EQ_HORIZONTALES__MERGED.pdf",
    ]

    caratulas_dir = Path(args.caratulas_dir) if args.caratulas_dir else None

    for centro in centros:
        base = out_root / centro
        pdfs_final = []

        # Añadir PDFs según orden_secciones (carátula + merged)
        for caratula, merged in orden_secciones:
            p_caratula = caratulas_dir / caratula if caratulas_dir else None
            p_merged = base / merged if merged else None

            if merged and p_merged.exists():
                if p_caratula and p_caratula.exists():
                    pdfs_final.append(p_caratula)
                else:
                    if p_caratula:
                        print(f"[INFO] ({centro}) Carátula no encontrada: {p_caratula.name}, se omite.")
                pdfs_final.append(p_merged)
            elif not merged and p_caratula and p_caratula.exists():
                pdfs_final.append(p_caratula)
            else:
                if merged:
                    print(f"[INFO] ({centro}) Merged no encontrado: {p_merged.name}, se omite carátula y merged.")

        # Añadir PDFs extra
        for nombre in orden_merged_extra:
            p = base / nombre
            if p.exists():
                pdfs_final.append(p)
            else:
                print(f"[INFO] ({centro}) Merged no encontrado: {nombre}, se omite.")

        if not pdfs_final:
            print(f"[ERROR] ({centro}) No se encontraron PDFs merged válidos en: {base}")
            continue

        print(f"[INFO] ({centro}) PDFs a unir en el Anejo final:")
        for p in pdfs_final:
            print(" -", p.name)

        out_final = base / "Anejo5_Reportaje_Fotografico.pdf"
        writer = PdfWriter()
        for p in pdfs_final:
            try:
                writer.append(str(p))
            except Exception as e:
                print(f"[WARN] ({centro}) No se pudo añadir {p}: {e}")
        with out_final.open("wb") as f:
            writer.write(f)
        print(f"[OK] ({centro}) Anejo final generado: {out_final}")

def main():
    args = parse_args()
    try:
        centros = asyncio.run(main_async(args))
        # Genera SIEMPRE un Anejo final por cada centro
        out_root = Path(args.out).resolve() if args.out else out_root_for(Path(args.data).resolve())
        crear_anejo_final_por_centro(args, out_root, centros)
    except KeyboardInterrupt:
        print("\n[STOP] Cancelado por el usuario")
    except Exception as e:
        print("[ERROR]", e)
        sys.exit(1)

if __name__ == "__main__":
    main()
