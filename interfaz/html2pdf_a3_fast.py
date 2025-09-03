# -*- coding: utf-8 -*-
"""
HTML → PDF A3 (Chromium/Playwright) + merges por sección **POR CENTRO**, con render concurrente.

Salida:
  <data>_pdf/<CENTRO>/<seccion>/*.pdf
  <data>_pdf/<CENTRO>/_<SECCION>__MERGED.pdf
  <data>_pdf/<CENTRO>/05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf
"""

import argparse
import asyncio
import os
import re
import sys
import functools
import threading
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import time
from pathlib import Path
from urllib.parse import quote, unquote
from collections import defaultdict
import shutil
import contextlib

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

# ──────────────────────────────────────────────────────────────────────────────
# HTTP local (sirve --data por http://127.0.0.1:<PORT>/…)
def start_http_server(root: Path, port: int = 8800):
    handler = functools.partial(SimpleHTTPRequestHandler, directory=str(root))
    httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
    t = threading.Thread(target=httpd.serve_forever, daemon=True)
    t.start()
    return httpd
# ──────────────────────────────────────────────────────────────────────────────

def to_file_uri(p: Path) -> str:
    p = p.resolve()
    s = str(p).replace("\\", "/")
    return "file:///" + quote(s, safe="/:._-()")

def find_htmls(root: Path, fast_mode: bool = False) -> list[Path]:
    """Encuentra HTML y aplica validaciones básicas (opcional)."""
    html_files = sorted([p for p in root.rglob("*.html") if p.is_file()])

    if fast_mode:
        return [p for p in html_files if p.stat().st_size > 0]

    valid_htmls = []
    for html_file in html_files:
        try:
            if html_file.stat().st_size == 0:
                print(f"[SKIP] {html_file.name} es vacío, se omite")
                continue

            with open(html_file, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
                if not any(tag in content.lower() for tag in ['<html', '<body', '<!doctype']):
                    print(f"[SKIP] {html_file.name} no parece contener HTML válido, se omite")
                    continue

                broken_images = []
                src_patterns = re.findall(r'src=["\']([^"\']+)["\']', content, re.IGNORECASE)
                for src in src_patterns[:5]:
                    if src.startswith(('http://', 'https://', 'data:')):
                        continue
                    if src.startswith('file:///'):
                        img_path = Path(unquote(src.replace('file:///', '')))
                    elif not os.path.isabs(src):
                        img_path = html_file.parent / src
                    else:
                        img_path = Path(src)
                    if not img_path.exists():
                        broken_images.append(src[:50] + "..." if len(src) > 50 else src)
                if broken_images:
                    print(f"[WARNING] {html_file.name} tiene {len(broken_images)} imágenes no encontradas; ejemplo: {broken_images[:2]}")

            valid_htmls.append(html_file)
        except Exception as e:
            print(f"[SKIP] {html_file.name} no se puede leer ({e}), se omite")
            continue

    if len(valid_htmls) != len(html_files):
        print(f"[INFO] {len(html_files)-len(valid_htmls)} archivos HTML omitidos por ser inválidos o corruptos")

    return valid_htmls

def out_root_for(in_root: Path) -> Path:
    return in_root.with_name(in_root.name + "_pdf")

def section_from_path(html_path: Path, data_root: Path) -> str:
    """.../<CENTRO>/<seccion>/archivo.html  -> <seccion>"""
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
    """Primer nivel bajo data_root es el centro (CC0001, ...)."""
    try:
        rel = html_path.relative_to(data_root)
        return rel.parts[0] if len(rel.parts) > 1 else "ROOT"
    except Exception:
        return "ROOT"

# ──────────────────────────────────────────────────────────────────────────────
# Render rápido (sin reintentos)
async def render_one_fast(page, url: str, pdf_path: Path, scale: float, prefer_css: bool, wait_ms: int):
    pdf_opts = {
        "path": str(pdf_path),
        "format": "A3",
        "landscape": True,
        "print_background": True,
        "margin": {"top": "0", "right": "0", "bottom": "0", "left": "0"},
        "scale": scale,
        "prefer_css_page_size": prefer_css,
    }
    await asyncio.wait_for(page.goto(url, wait_until="domcontentloaded"), timeout=5.0)
    with contextlib.suppress(Exception):
        await page.evaluate("""() => { document.querySelectorAll('meta[http-equiv="refresh"]').forEach(m=>m.remove()); }""")
    with contextlib.suppress(Exception):
        await page.wait_for_load_state("networkidle", timeout=3000)
    await page.wait_for_timeout(wait_ms)
    await page.emulate_media(media="print")
    await page.pdf(**pdf_opts)

# Render robusto (con reintentos)
async def render_one(page, url: str, pdf_path: Path, scale: float, prefer_css: bool, wait_ms: int, retry_mode: bool = False):
    pdf_opts = {
        "path": str(pdf_path),
        "format": "A3",
        "landscape": True,
        "print_background": True,
        "margin": {"top": "0", "right": "0", "bottom": "0", "left": "0"},
        "scale": scale,
        "prefer_css_page_size": prefer_css,
    }
    name = pdf_path.name

    try:
        if page.is_closed():
            raise RuntimeError("Page is closed")

        # Intento 1 (moderado)
        await asyncio.wait_for(page.goto(url, wait_until="domcontentloaded"),
                               timeout=(10.0 if retry_mode else 15.0))
        with contextlib.suppress(Exception):
            await page.evaluate("""() => { document.querySelectorAll('meta[http-equiv="refresh"]').forEach(m=>m.remove()); }""")
        with contextlib.suppress(Exception):
            await page.wait_for_load_state("networkidle", timeout=3000)
        with contextlib.suppress(Exception):
            await page.evaluate("""async () => {
                if (document.fonts && document.fonts.status !== "loaded") {
                    try { await document.fonts.ready; } catch(e) {}
                }
            }""", timeout=3000)

        await page.wait_for_timeout(min(wait_ms, 800))
        await page.emulate_media(media="print")
        await page.pdf(**pdf_opts)
        return

    except (asyncio.TimeoutError, Exception) as e:
        msg = str(e)
        if "net::ERR_ABORTED" in msg:
            print(f"[RETRY] {name} ERR_ABORTED - posibles imágenes rotas o HTML corrupto (intento 1)")
        elif "frame was detached" in msg:
            print(f"[RETRY] {name} frame detached - problema de contexto del navegador (intento 1)")
        else:
            print(f"[RETRY] {name} falló (intento 1): {msg[:80]}...")

        # Intento 2 (robusto)
        try:
            if page.is_closed():
                raise RuntimeError("Page closed before retry")

            await asyncio.wait_for(page.goto(url, wait_until="domcontentloaded"), timeout=10.0)
            with contextlib.suppress(Exception):
                await page.evaluate("""() => { document.querySelectorAll('meta[http-equiv="refresh"]').forEach(m=>m.remove()); }""")
            with contextlib.suppress(Exception):
                await page.wait_for_load_state("networkidle", timeout=3000)

            await page.wait_for_timeout(max(wait_ms, 1500))
            await page.emulate_media(media="print")
            await page.pdf(**pdf_opts)
            print(f"[RECOVER] {name} recuperado en intento 2")
            return

        except (asyncio.TimeoutError, Exception) as e2:
            msg2 = str(e2)
            if "net::ERR_ABORTED" in msg2:
                print(f"[LAST_TRY] {name} HTML corrupted/malformed - ERR_ABORTED (intento 2)")
            elif "frame was detached" in msg2:
                print(f"[LAST_TRY] {name} frame detached - browser context issue (intento 2)")
            else:
                print(f"[LAST_TRY] {name} falló intento 2: {msg2[:80]}...")

            # Intento 3 (último recurso, muy corto)
            try:
                if page.is_closed():
                    raise RuntimeError("Page closed before final retry")
                await asyncio.wait_for(page.goto(url, wait_until="domcontentloaded"), timeout=6.0)
                await page.wait_for_timeout(50)
                await page.emulate_media(media="print")
                await page.pdf(**pdf_opts)
                print(f"[RECOVER] {name} recuperado en último intento")
                return
            except (asyncio.TimeoutError, Exception) as e3:
                print(f"[FAIL] {name} FALLÓ definitivamente: {str(e3)[:120]}")
                raise
# ──────────────────────────────────────────────────────────────────────────────

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

async def merge_pdfs_parallel_by_center(rendered_list, out_root: Path, args):
    """Agrupa PDFs por centro y hace merges en paralelo."""
    from concurrent.futures import ThreadPoolExecutor

    buckets_by_center = defaultdict(lambda: defaultdict(list))
    centros = set()

    for centro, section, pdf in rendered_list:
        centros.add(centro)
        buckets_by_center[centro][section].append(pdf)

    print(f"[INFO] Procesando {len(centros)} centros en paralelo...")

    def process_center(centro):
        start_time = time.time()
        try:
            center_buckets = buckets_by_center[centro]
            merged_count = 0
            for section, paths in center_buckets.items():
                paths = sorted(paths)
                name = f"_{SECCIONES.get(section, section).upper()}__MERGED.pdf"
                out_merge = out_root / centro / name
                try:
                    merge_pdfs(paths, out_merge)
                    merged_count += 1
                    print(f"[MERGE] {centro} · {section} -> {out_merge.name} ({len(paths)} PDFs)")
                except Exception as e:
                    print(f"[WARN] Merge {centro}/{section} falló: {e}")
            print(f"[CENTRO-OK] {centro}: {merged_count} merges en {time.time()-start_time:.1f}s")
            return centro
        except Exception as e:
            print(f"[CENTRO-FAIL] {centro}: ERROR: {e}")
            return None

    loop = asyncio.get_event_loop()
    with ThreadPoolExecutor(max_workers=args.merge_workers) as executor:
        futures = [loop.run_in_executor(executor, process_center, centro) for centro in centros]
        try:
            completed_centros = await asyncio.wait_for(asyncio.gather(*futures, return_exceptions=True), timeout=600.0)
        except asyncio.TimeoutError:
            print(f"[TIMEOUT] Algunos merges se colgaron (>10min), cancelando…")
            for f in futures:
                if not f.done():
                    f.cancel()
            completed_centros = []

    successful_centros = [c for c in completed_centros if isinstance(c, str) and c is not None]
    errors = [c for c in completed_centros if not isinstance(c, str) or c is None]
    if errors:
        print(f"[WARN] {len(errors)} centros tuvieron errores en merge")

    print(f"[MERGE-SUMMARY] {len(successful_centros)}/{len(centros)} centros OK")
    return set(successful_centros)

# ──────────────────────────────────────────────────────────────────────────────
# Normalización de imágenes: file:///C:/…  → assets/<archivo>
def fix_html_assets(html_path: Path):
    try:
        txt = html_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return
    changed = False
    assets_dir = html_path.parent / "assets"

    def repl_img(m):
        nonlocal changed
        url = m.group(1)
        if isinstance(url, str) and url.lower().startswith("file:///"):
            src_path = Path(unquote(url[8:]))
            if src_path.exists():
                assets_dir.mkdir(exist_ok=True)
                dst = assets_dir / src_path.name
                if not dst.exists():
                    shutil.copy2(src_path, dst)
                changed = True
                return f'src="assets/{dst.name}"'
        return m.group(0)

    txt2 = re.sub(r'src=["\']([^"\']+)["\']', repl_img, txt, flags=re.IGNORECASE)
    if txt2 != txt:
        html_path.write_text(txt2, encoding="utf-8")

def fix_all_htmls(root: Path):
    for p in root.rglob("*.html"):
        try:
            fix_html_assets(p)
        except Exception as e:
            print(f"[SKIP] {p}: {e}")
# ──────────────────────────────────────────────────────────────────────────────

async def main_async(args):
    # Ajustes de velocidad (sin bloquear imágenes)
    if args.ultra_fast:
        print("[SPEED] ULTRA-RÁPIDO: concurrencia alta, timeouts mínimos")
        args.concurrency = min(os.cpu_count() or 4, 24)
        args.merge_workers = min(os.cpu_count() or 4, 8)
        args.wait = 100
        args.block = "media"  # ¡No bloquear imágenes!
    elif args.fast:
        print("[SPEED] RÁPIDO: concurrencia alta, timeouts reducidos")
        args.concurrency = min(os.cpu_count() or 4, 16)
        args.merge_workers = min(os.cpu_count() or 4, 6)
        args.wait = 200
        args.block = "media"  # ¡No bloquear imágenes!

    data_root = Path(args.data).resolve()
    httpd = None
    base_url = None
    if not args.use_file_scheme:
        httpd = start_http_server(data_root, args.port)
        base_url = f"http://127.0.0.1:{args.port}"

    out_root = Path(args.out).resolve() if args.out else out_root_for(data_root)

    # Normaliza rutas file:/// de imágenes a assets/
    fix_all_htmls(data_root)
    out_root.mkdir(parents=True, exist_ok=True)

    print(f"[CONFIG] Concurrencia: {args.concurrency}, Workers merge: {args.merge_workers}, Espera: {args.wait}ms")
    if args.block:
        print(f"[CONFIG] Recursos bloqueados: {args.block}")

    fast_mode = hasattr(args, 'fast') and (args.fast or args.ultra_fast)
    htmls = find_htmls(data_root, fast_mode=fast_mode)
    if not htmls:
        print("[INFO] No se encontraron .html en", data_root)
        return set()

    rendered = await render_htmls_to_pdfs(htmls, data_root, out_root, args, base_url=base_url)

    if args.no_merge:
        print("[OK] Conversión terminada. Sin merges por sección.")
        return {c for c, _, _ in rendered}

    print(f"[INFO] Iniciando merges paralelos con {args.merge_workers} workers…")
    centros = await merge_pdfs_parallel_by_center(rendered, out_root, args)

    if httpd:
        httpd.shutdown()
    return centros

# ──────────────────────────────────────────────────────────────────────────────
# Render concurrente + reintentos diferidos
async def render_htmls_to_pdfs(htmls, data_root, out_root, args, base_url=None):
    rendered = []
    failed_htmls = []
    final_failures = []
    total = len(htmls)
    done_counter = 0
    lock = asyncio.Lock()
    sem = asyncio.Semaphore(args.concurrency)

    block_types = set(t.strip().lower() for t in args.block.split(",") if t.strip())

    async with async_playwright() as pw:
        chromium_args = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-accelerated-2d-canvas",
            "--no-first-run",
            "--no-zygote",
            "--disable-gpu",
            "--disable-features=CalculateNativeWinOcclusion",
            "--disable-background-networking",
            "--disable-background-timer-throttling",
            "--disable-backgrounding-occluded-windows",
            "--disable-breakpad",
            "--disable-client-side-phishing-detection",
            "--disable-component-update",
            "--disable-default-apps",
            "--disable-domain-reliability",
            "--disable-extensions",
            "--disable-features=TranslateUI",
            "--disable-hang-monitor",
            "--disable-ipc-flooding-protection",
            "--disable-popup-blocking",
            "--disable-prompt-on-repost",
            "--disable-renderer-backgrounding",
            "--disable-sync",
            "--disable-translate",
            "--disable-windows10-custom-titlebar",
            "--metrics-recording-only",
            "--no-default-browser-check",
            "--no-pings",
            "--password-store=basic",
            "--use-mock-keychain",
            "--js-flags=--max-old-space-size=4096",
        ]
        chromium_args.extend(args.chromium-args or []) if hasattr(args, 'chromium-arg') else chromium_args.extend(args.chromium_arg or [])

        browser = await pw.chromium.launch(headless=True, args=chromium_args)
        context = await browser.new_context()

        if block_types:
            async def route_handler(route):
                r = route.request
                urlp = r.url.lower()
                # No bloquear imágenes de secciones fotográficas
                if r.resource_type == "image" and any(seg in urlp for seg in ["/clima/", "/cc/", "/acom"]):
                    await route.continue_()
                    return
                if r.resource_type in block_types:
                    await route.abort()
                else:
                    await route.continue_()
            await context.route("**/*", route_handler)

        async def worker(i, html, retry_attempt=0):
            nonlocal done_counter
            rel = html.relative_to(data_root)
            section = section_from_path(html, data_root)
            centro = center_from_path(html, data_root)

            dst_dir = out_root / rel.parent
            dst_dir.mkdir(parents=True, exist_ok=True)
            pdf_name = html.with_suffix(".pdf").name
            pdf_path = dst_dir / pdf_name

            # URL preferente por HTTP
            rel_http = html.relative_to(data_root).as_posix()
            url = f"{base_url}/{rel_http}" if base_url else to_file_uri(html)

            async with sem:
                page = await context.new_page()
                try:
                    retry_info = f" [RETRY {retry_attempt}]" if retry_attempt > 0 else ""
                    print(f"[{i}/{total}]{retry_info} → {pdf_path}")

                    if hasattr(args, 'ultra_fast') and args.ultra_fast:
                        await asyncio.wait_for(
                            render_one_fast(page, url, pdf_path, scale=args.scale,
                                            prefer_css=not args.ignore_css_page, wait_ms=args.wait),
                            timeout=15.0
                        )
                    elif hasattr(args, 'fast') and args.fast:
                        await asyncio.wait_for(
                            render_one_fast(page, url, pdf_path, scale=args.scale,
                                            prefer_css=not args.ignore_css_page, wait_ms=args.wait),
                            timeout=30.0
                        )
                    else:
                        timeout_duration = 180.0 if retry_attempt > 0 else 120.0
                        await asyncio.wait_for(
                            render_one(page, url, pdf_path, scale=args.scale,
                                       prefer_css=not args.ignore_css_page, wait_ms=args.wait,
                                       retry_mode=(retry_attempt > 0)),
                            timeout=timeout_duration
                        )

                    async with lock:
                        done_counter += 1
                        if retry_attempt > 0:
                            print(f"[RECUPERADO] {html.name} exitoso en reintento {retry_attempt}")
                        if (done_counter % args.log_every) == 0 or done_counter == total:
                            print(f"[PROGRESO] {done_counter}/{total} completados")
                    return (centro, section, pdf_path)

                except asyncio.TimeoutError:
                    async with lock:
                        timeout_duration = 180.0 if retry_attempt > 0 else 120.0
                        print(f"[TIMEOUT] {html.name} se colgó (>{timeout_duration}s)")
                        failed_htmls.append((i, html, {
                            'file': str(html),
                            'error_type': 'TimeoutError',
                            'error_msg': f'Timeout after {timeout_duration}s',
                            'is_corrupted': False,
                            'is_detached': False,
                            'is_timeout': True
                        }))
                    return None

                except Exception as e:
                    async with lock:
                        print(f"[WARN] Falló {html}: {e}, se reintentará al final" if retry_attempt == 0 else f"[WARN] {html} CANCELADO definitivamente: {e}")
                        detail = {
                            'file': str(html),
                            'error_type': type(e).__name__,
                            'error_msg': str(e)[:200],
                            'is_corrupted': "net::ERR_ABORTED" in str(e),
                            'is_detached': "frame was detached" in str(e)
                        }
                        if retry_attempt == 0:
                            failed_htmls.append((i, html, detail))
                        else:
                            final_failures.append(detail)
                    return None

                finally:
                    with contextlib.suppress(Exception):
                        await page.close()

        print(f"[INFO] Iniciando conversión de {total} archivos HTML…")
        tasks = [asyncio.create_task(worker(i, html)) for i, html in enumerate(htmls, 1)]

        completed_tasks = 0
        for coro in asyncio.as_completed(tasks):
            try:
                res = await coro
                if res:
                    rendered.append(res)
                completed_tasks += 1
            except Exception as e:
                print(f"[ERROR] Tarea no controlada falló: {str(e)[:100]}...")
                completed_tasks += 1

        print(f"[INFO] Primera ronda completada: {completed_tasks}/{total} tareas")

        if failed_htmls:
            wait_time = min(10.0, max(3.0, len(failed_htmls) * 0.1))
            print(f"\n[RETRY] {len(failed_htmls)} conversiones fallaron, esperando {wait_time:.1f}s…")
            await asyncio.sleep(wait_time)

            print(f"[RETRY] Reintentando {len(failed_htmls)} conversiones fallidas…")
            retry_tasks = [asyncio.create_task(worker(i, html, retry_attempt=1)) for i, html, _ in failed_htmls]

            retry_completed = 0
            for coro in asyncio.as_completed(retry_tasks):
                try:
                    res = await coro
                    if res:
                        rendered.append(res)
                    retry_completed += 1
                except Exception as e:
                    print(f"[ERROR] Reintento no controlado falló: {str(e)[:100]}...")
                    retry_completed += 1

            print(f"[INFO] Reintentos completados: {retry_completed}/{len(failed_htmls)}")

            final_failed = total - len(rendered)
            if final_failed > 0:
                print(f"[FINAL] {final_failed}/{total} conversiones fallaron definitivamente")
                print(f"[FINAL] Tasa de éxito: {((total - final_failed) / total * 100):.1f}%")
                error_types = {}
                corrupted_files = 0
                timeout_files = 0
                for failure in final_failures:
                    et = failure['error_type']
                    error_types[et] = error_types.get(et, 0) + 1
                    if failure.get('is_corrupted', False):
                        corrupted_files += 1
                    if failure.get('is_timeout', False):
                        timeout_files += 1
                print("[DIAGNOSTICO] Tipos de errores:")
                for et, count in error_types.items():
                    print(f"  - {et}: {count} archivos")
                if corrupted_files > 0:
                    print(f"[DIAGNOSTICO] {corrupted_files} archivos posiblemente corruptos (ERR_ABORTED)")
                if timeout_files > 0:
                    print(f"[DIAGNOSTICO] {timeout_files} archivos con timeout")
            else:
                print(f"[SUCCESS] Todas las conversiones completadas tras reintentos")
        else:
            print(f"[SUCCESS] Todas las conversiones completadas en primera ronda")

        await asyncio.sleep(1.0)
        await context.close()
        await browser.close()

    rendered.sort(key=lambda tup: str(tup[2]))
    return rendered

# ──────────────────────────────────────────────────────────────────────────────

def parse_args():
    default_concurrency = min(os.cpu_count() or 4, 8)

    ap = argparse.ArgumentParser(description="HTML → PDF A3 (Chromium/Playwright) + merges por sección **por centro**")
    ap.add_argument("--data", required=True, help="Carpeta raíz con los .html generados")
    ap.add_argument("--out", default=None, help="Carpeta de salida (default: <data>_pdf)")
    ap.add_argument("--scale", type=float, default=1.0, help="Escala del render (1.0 por defecto)")
    ap.add_argument("--wait", type=int, default=500, help="Espera (ms) tras cargar cada HTML (500 por defecto)")
    ap.add_argument("--no-merge", action="store_true", help="No crear PDFs combinados por sección")
    ap.add_argument("--ignore-css-page", action="store_true", help="Ignorar @page del CSS y forzar A3 landscape")
    ap.add_argument("--concurrency", type=int, default=default_concurrency, help=f"Número de páginas en paralelo (default {default_concurrency})")
    ap.add_argument("--merge-workers", type=int, default=min(os.cpu_count() or 2, 4), help="Workers paralelos para merges por centro")
    ap.add_argument("--fast", action="store_true", help="Modo rápido: alta concurrencia, timeouts reducidos (sin reintentos)")
    ap.add_argument("--ultra-fast", action="store_true", help="Modo ultra-rápido extremo")
    ap.add_argument("--log-every", type=int, default=10, help="Frecuencia de logs de progreso")
    ap.add_argument("--block", default="", help="Tipos de recurso a bloquear: image,font,media,stylesheet,script")
    ap.add_argument("--chromium-arg", action="append", default=[], help="Argumentos extra para Chromium (repetible)")
    ap.add_argument("--caratulas-dir", default=None, help="Ruta a las carátulas del Anejo (PDFs)")
    ap.add_argument("--port", type=int, default=8800, help="Puerto HTTP local para servir --data")
    ap.add_argument("--use-file-scheme", action="store_true", help="Forzar file:// en lugar de HTTP (menos estable)")
    return ap.parse_args()

def crear_anejo_final_por_centro(args, out_root: Path, centros: set[str]):
    """
    Genera un 05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf por cada centro dentro de <out_root>/<CENTRO>.
    Usa las mismas carátulas para todos (si existen).
    """
    from concurrent.futures import ThreadPoolExecutor

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

    def process_centro_final(centro):
        base = out_root / centro
        pdfs_final = []

        for caratula, merged in orden_secciones:
            p_caratula = caratulas_dir / caratula if caratulas_dir else None
            p_merged = base / merged if merged else None

            if merged and p_merged.exists():
                if p_caratula and p_caratula.exists():
                    pdfs_final.append(p_caratula)
                elif p_caratula:
                    print(f"[INFO] ({centro}) Carátula no encontrada: {p_caratula.name}, se omite")
                pdfs_final.append(p_merged)
            elif not merged and p_caratula and p_caratula.exists():
                pdfs_final.append(p_caratula)
            else:
                if merged:
                    print(f"[INFO] ({centro}) Merged no encontrado: {p_merged.name}, se omite")

        for nombre in orden_merged_extra:
            p = base / nombre
            if p.exists():
                pdfs_final.append(p)
            else:
                print(f"[INFO] ({centro}) Merged no encontrado: {nombre}, se omite")

        if not pdfs_final:
            print(f"[ERROR] ({centro}) No se encontraron PDFs merged válidos en: {base}")
            return False

        out_final = base / "05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf"
        writer = PdfWriter()
        for p in pdfs_final:
            with contextlib.suppress(Exception):
                writer.append(str(p))
        try:
            with out_final.open("wb") as f:
                writer.write(f)
            print(f"[OK] ({centro}) Anejo final generado: {out_final}")
            return True
        except Exception as e:
            print(f"[ERROR] ({centro}) Error generando anejo final: {e}")
            return False

    merge_workers = getattr(args, 'merge_workers', min(os.cpu_count() or 4, 8))
    print(f"[FINAL] Generando anejos finales en paralelo con {merge_workers} workers…")
    with ThreadPoolExecutor(max_workers=merge_workers) as executor:
        results = list(executor.map(process_centro_final, centros))
    print(f"[FINAL-SUMMARY] {sum(results)}/{len(centros)} anejos finales generados exitosamente")

def main():
    args = parse_args()
    try:
        centros = asyncio.run(main_async(args))
        out_root = Path(args.out).resolve() if args.out else out_root_for(Path(args.data).resolve())
        crear_anejo_final_por_centro(args, out_root, centros)
    except KeyboardInterrupt:
        print("\n[STOP] Cancelado por el usuario")
    except Exception as e:
        print("[ERROR]", e)
        sys.exit(1)

if __name__ == "__main__":
    main()

