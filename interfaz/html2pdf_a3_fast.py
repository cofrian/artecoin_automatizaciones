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
import signal
import sys
import threading
import time
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

def find_htmls(root: Path, fast_mode: bool = False) -> list[Path]:
    """Find all HTML files and filter out potentially corrupted ones"""
    html_files = sorted([p for p in root.rglob("*.html") if p.is_file()])
    
    if fast_mode:
        # Modo rápido: validación mínima
        valid_htmls = []
        for html_file in html_files:
            try:
                if html_file.stat().st_size > 0:  # Solo verificar que no esté vacío
                    valid_htmls.append(html_file)
            except Exception:
                continue
        return valid_htmls
    
    # Modo normal: validación completa
    # Filter out empty or potentially corrupted files
    valid_htmls = []
    for html_file in html_files:
        try:
            if html_file.stat().st_size == 0:
                print(f"[SKIP] {html_file.name} es vacío, se omite")
                continue
                
            # Basic validation: check if file contains basic HTML structure
            with open(html_file, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()  # Read full content for broken image detection
                if not any(tag in content.lower() for tag in ['<html', '<body', '<!doctype']):
                    print(f"[SKIP] {html_file.name} no parece contener HTML válido, se omite")
                    continue
                
                # Check for broken image references (main cause of ERR_ABORTED)
                broken_images = []
                src_patterns = re.findall(r'src=["\']([^"\']+)["\']', content, re.IGNORECASE)
                for src in src_patterns[:5]:  # Check first 5 images only for performance
                    if src.startswith(('http://', 'https://', 'data:')):
                        continue  # Skip web URLs and data URIs
                        
                    # Check if image file exists
                    if src.startswith('file:///'):
                        # Extract path from file URI
                        from urllib.parse import unquote
                        img_path = Path(unquote(src.replace('file:///', '')))
                    elif not os.path.isabs(src):
                        img_path = html_file.parent / src
                    else:
                        img_path = Path(src)
                        
                    if not img_path.exists():
                        broken_images.append(src[:50] + "..." if len(src) > 50 else src)
                
                if broken_images:
                    print(f"[WARNING] {html_file.name} tiene {len(broken_images)} imágenes no encontradas")
                    print(f"[WARNING] Ejemplos: {broken_images[:2]}")
                    # Don't skip, but warn - let the improved error handling deal with it
                    
            valid_htmls.append(html_file)
        except Exception as e:
            print(f"[SKIP] {html_file.name} no se puede leer ({e}), se omite")
            continue
    
    if len(valid_htmls) != len(html_files):
        skipped = len(html_files) - len(valid_htmls)
        print(f"[INFO] {skipped} archivos HTML omitidos por ser inválidos o corruptos")
    
    return valid_htmls

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

async def render_one_fast(page, html_path: Path, pdf_path: Path, scale: float, prefer_css: bool, wait_ms: int):
    """Versión ultra rápida sin reintentos para modo rápido"""
    url = to_file_uri(html_path)
    
    pdf_opts = {
        "path": str(pdf_path),
        "format": "A3",
        "landscape": True,
        "print_background": True,
        "margin": {"top": "0", "right": "0", "bottom": "0", "left": "0"},
        "scale": scale,
        "prefer_css_page_size": prefer_css,
    }
    
    # Verificación básica
    if not html_path.exists() or html_path.stat().st_size == 0:
        raise FileNotFoundError(f"HTML file missing or empty: {html_path.name}")
    
    # Un solo intento con timeout muy corto
    await asyncio.wait_for(
        page.goto(url, wait_until="domcontentloaded"),
        timeout=5.0  # Timeout ultra corto
    )
    await page.wait_for_timeout(wait_ms)  # Espera configurada
    await page.emulate_media(media="print")
    await page.pdf(**pdf_opts)

async def render_one(page, html_path: Path, pdf_path: Path, scale: float, prefer_css: bool, wait_ms: int, retry_mode: bool = False):
    url = to_file_uri(html_path)
    
    pdf_opts = {
        "path": str(pdf_path),
        "format": "A3",
        "landscape": True,
        "print_background": True,
        "margin": {"top": "0", "right": "0", "bottom": "0", "left": "0"},
        "scale": scale,
        "prefer_css_page_size": prefer_css,
    }
    
    # Verificar que el archivo HTML existe y es válido
    if not html_path.exists() or html_path.stat().st_size == 0:
        raise FileNotFoundError(f"HTML file missing or empty: {html_path.name}")
    
    # En modo reintento, usar estrategia más conservadora
    if retry_mode:
        try:
            # Verificar que la página sigue activa
            if page.is_closed():
                raise RuntimeError("Page is closed")
                
            # Modo reintento: estrategia ultra-conservadora (8s timeout)
            await asyncio.wait_for(
                page.goto(url, wait_until="load"),
                timeout=8.0
            )
            await page.wait_for_timeout(max(wait_ms, 1500))  # Mínimo 1.5 segundos de espera
            await page.emulate_media(media="print")
            await page.pdf(**pdf_opts)
            return  # Éxito en modo reintento
            
        except (asyncio.TimeoutError, Exception) as e:
            error_type = type(e).__name__
            error_msg = str(e)
            if "net::ERR_ABORTED" in error_msg:
                print(f"[RETRY_FAIL] {html_path.name} HTML corrupted/malformed - ERR_ABORTED")
            elif "frame was detached" in error_msg:
                print(f"[RETRY_FAIL] {html_path.name} frame detached - browser context issue")
            else:
                print(f"[RETRY_FAIL] {html_path.name} falló en modo reintento ({error_type}): {str(e)[:60]}...")
            raise e
    
    try:
        # Verificar que la página sigue activa
        if page.is_closed():
            raise RuntimeError("Page is closed")
            
        # Intento 1: Estrategia moderada (15s timeout)
        await asyncio.wait_for(
            page.goto(url, wait_until="domcontentloaded"),
            timeout=15.0
        )
        await page.wait_for_timeout(min(wait_ms, 800))  # Máximo 800ms
        await page.emulate_media(media="print")
        await page.pdf(**pdf_opts)
        return  # Éxito en primer intento
        
    except (asyncio.TimeoutError, Exception) as e:
        error_type = type(e).__name__
        error_msg = str(e)
        
        if "net::ERR_ABORTED" in error_msg:
            print(f"[RETRY] {html_path.name} ERR_ABORTED - posibles imágenes rotas o HTML corrupto (intento 1)")
        elif "frame was detached" in error_msg:
            print(f"[RETRY] {html_path.name} frame detached - problema de contexto del navegador (intento 1)")
        else:
            print(f"[RETRY] {html_path.name} falló (intento 1): {str(e)[:80]}...")
        
        try:
            # Verificar que la página sigue activa antes del segundo intento
            if page.is_closed():
                raise RuntimeError("Page closed before retry")
                
            # Intento 2: Estrategia robusta (10s timeout)
            await asyncio.wait_for(
                page.goto(url, wait_until="load"),
                timeout=10.0
            )
            await page.wait_for_timeout(wait_ms)
            await page.emulate_media(media="print")
            await page.pdf(**pdf_opts)
            print(f"[RECOVER] {html_path.name} recuperado en intento 2")
            return
            
        except (asyncio.TimeoutError, Exception) as e2:
            error_type2 = type(e2).__name__
            error_msg2 = str(e2)
            
            if "net::ERR_ABORTED" in error_msg2:
                print(f"[LAST_TRY] {html_path.name} HTML corrupted/malformed - ERR_ABORTED (intento 2)")
            elif "frame was detached" in error_msg2:
                print(f"[LAST_TRY] {html_path.name} frame detached - browser context issue (intento 2)")
            else:
                print(f"[LAST_TRY] {html_path.name} falló intento 2: {str(e2)[:80]}...")
            
            try:
                # Verificar que la página sigue activa antes del último intento
                if page.is_closed():
                    raise RuntimeError("Page closed before final retry")
                    
                # Intento 3: Último recurso (6s timeout)
                await asyncio.wait_for(
                    page.goto(url, wait_until="domcontentloaded"),
                    timeout=6.0
                )
                await page.wait_for_timeout(50)  # Solo 50ms
                await page.emulate_media(media="print")
                await page.pdf(**pdf_opts)
                print(f"[RECOVER] {html_path.name} recuperado en último intento")
                return
                
            except (asyncio.TimeoutError, Exception) as e3:
                error_type3 = type(e3).__name__
                error_msg3 = str(e3)
                
                if "net::ERR_ABORTED" in error_msg3:
                    print(f"[FAIL] {html_path.name} FALLÓ definitivamente: HTML corrupted/malformed - ERR_ABORTED")
                elif "frame was detached" in error_msg3:
                    print(f"[FAIL] {html_path.name} FALLÓ definitivamente: frame detached - browser context issue")
                else:
                    print(f"[FAIL] {html_path.name} FALLÓ definitivamente: {str(e3)[:80]}")
                raise e3

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
    """
    Agrupa los PDFs por centro y hace merges paralelos.
    Cada centro se procesa independientemente en paralelo.
    """
    from concurrent.futures import ThreadPoolExecutor
    
    # Agrupar por centro
    buckets_by_center = defaultdict(lambda: defaultdict(list))
    centros = set()
    
    for centro, section, pdf in rendered_list:
        centros.add(centro)
        buckets_by_center[centro][section].append(pdf)
    
    print(f"[INFO] Procesando {len(centros)} centros en paralelo...")
    
    def process_center(centro):
        """Procesa todos los merges de UN centro"""
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
            
            elapsed = time.time() - start_time
            print(f"[CENTRO-OK] {centro}: {merged_count} merges completados en {elapsed:.1f}s")
            return centro
            
        except Exception as e:
            elapsed = time.time() - start_time
            print(f"[CENTRO-FAIL] {centro}: ERROR después de {elapsed:.1f}s: {e}")
            return None
    
    # Ejecutar centros en paralelo usando ThreadPoolExecutor con timeout
    loop = asyncio.get_event_loop()
    with ThreadPoolExecutor(max_workers=args.merge_workers) as executor:
        futures = [loop.run_in_executor(executor, process_center, centro) for centro in centros]
        
        try:
            # Timeout global para todos los merges: 10 minutos máximo
            completed_centros = await asyncio.wait_for(
                asyncio.gather(*futures, return_exceptions=True),
                timeout=600.0  # 10 minutos
            )
        except asyncio.TimeoutError:
            print(f"[TIMEOUT] Algunos merges se colgaron (>10min), cancelando...")
            # Cancelar futures pendientes
            for future in futures:
                if not future.done():
                    future.cancel()
            completed_centros = []
    
    # Filtrar resultados (solo strings válidos son centros exitosos)
    successful_centros = [c for c in completed_centros if isinstance(c, str) and c is not None]
    errors = [c for c in completed_centros if not isinstance(c, str) or c is None]
    
    if errors:
        print(f"[WARN] {len(errors)} centros tuvieron errores en merge")
    
    print(f"[MERGE-SUMMARY] {len(successful_centros)}/{len(centros)} centros procesados exitosamente")
    return set(successful_centros)

async def main_async(args):
    # Aplicar optimizaciones de velocidad
    if args.ultra_fast:
        print("[SPEED] Modo ULTRA RÁPIDO activado: máxima concurrencia, timeouts mínimos, sin validaciones")
        args.concurrency = min(os.cpu_count() or 4, 24)  # Concurrencia extrema
        args.merge_workers = min(os.cpu_count() or 4, 8)  # Más workers de merge
        args.wait = 100  # Espera mínima
        args.block = "image,font,media"  # Bloquear recursos pesados
    elif args.fast:
        print("[SPEED] Modo RÁPIDO activado: alta concurrencia, timeouts reducidos")
        args.concurrency = min(os.cpu_count() or 4, 16)  # Alta concurrencia
        args.merge_workers = min(os.cpu_count() or 4, 6)
        args.wait = 200  # Espera reducida
        args.block = "image,media"  # Bloquear algunos recursos pesados
    
    data_root = Path(args.data).resolve()
    out_root  = Path(args.out).resolve() if args.out else out_root_for(data_root)
    out_root.mkdir(parents=True, exist_ok=True)

    print(f"[CONFIG] Concurrencia: {args.concurrency}, Workers merge: {args.merge_workers}, Espera: {args.wait}ms")
    if args.block:
        print(f"[CONFIG] Recursos bloqueados: {args.block}")

    # Usar validación rápida en modos rápidos
    fast_mode = hasattr(args, 'fast') and (args.fast or args.ultra_fast)
    htmls = find_htmls(data_root, fast_mode=fast_mode)
    if not htmls:
        print("[INFO] No se encontraron .html en", data_root)
        return set()

    rendered = await render_htmls_to_pdfs(htmls, data_root, out_root, args)

    if args.no_merge:
        print("[OK] Conversión terminada. Sin merges por sección.")
        return {c for c, _, _ in rendered}

    # --- MERGES PARALELOS POR CENTRO ---
    print(f"[INFO] Iniciando merges paralelos con {args.merge_workers} workers...")
    centros = await merge_pdfs_parallel_by_center(rendered, out_root, args)
    
    return centros

# --- RENDER CONCURRENTE CON SISTEMA DE REINTENTOS DIFERIDOS ---
async def render_htmls_to_pdfs(htmls, data_root, out_root, args):
    rendered = []
    failed_htmls = []  # Lista para almacenar conversiones fallidas
    final_failures = []  # Lista para rastrear fallos definitivos con detalles
    total = len(htmls)
    done_counter = 0
    lock = asyncio.Lock()
    # CORREGIDO: usar exactamente args.concurrency, no forzar mínimo 16
    sem = asyncio.Semaphore(args.concurrency)

    block_types = set(t.strip().lower() for t in args.block.split(",") if t.strip())

    async with async_playwright() as pw:
        # Argumentos más conservadores para Chromium, especialmente para datasets grandes
        chromium_args = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-accelerated-2d-canvas",
            "--no-first-run",
            "--no-zygote",
            "--single-process",
            "--disable-gpu",
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
            "--memory-pressure-off",  # Evitar limitación de memoria automática
            "--max-old-generation-size=2048"  # Límite memoria V8 en MB (válido para Chromium)
        ]
        # Añadir argumentos personalizados
        chromium_args.extend(args.chromium_arg or [])
        
        browser = await pw.chromium.launch(
            headless=True,
            args=chromium_args
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

        async def worker(i, html, retry_attempt=0):
            nonlocal done_counter
            rel = html.relative_to(data_root)
            section = section_from_path(html, data_root)
            centro  = center_from_path(html, data_root)

            # Guardamos el PDF junto a la estructura del centro/sección
            dst_dir = out_root / rel.parent
            dst_dir.mkdir(parents=True, exist_ok=True)
            pdf_name = html.with_suffix(".pdf").name
            pdf_path = dst_dir / pdf_name

            # Crear página con semáforo (solo para limitar páginas activas)
            async with sem:
                page = await context.new_page()
                
                try:
                    retry_info = f" [RETRY {retry_attempt}]" if retry_attempt > 0 else ""
                    print(f"[{i}/{total}]{retry_info} → {pdf_path}")
                    
                    # Seleccionar función de render según el modo
                    if hasattr(args, 'ultra_fast') and args.ultra_fast:
                        # Modo ultra rápido: timeout corto, función simple
                        await asyncio.wait_for(
                            render_one_fast(page, html, pdf_path,
                                         scale=args.scale,
                                         prefer_css=not args.ignore_css_page,
                                         wait_ms=args.wait),
                            timeout=15.0  # Timeout muy corto
                        )
                    elif hasattr(args, 'fast') and args.fast:
                        # Modo rápido: timeout medio, función simple
                        await asyncio.wait_for(
                            render_one_fast(page, html, pdf_path,
                                         scale=args.scale,
                                         prefer_css=not args.ignore_css_page,
                                         wait_ms=args.wait),
                            timeout=30.0  # Timeout medio
                        )
                    else:
                        # Modo normal: timeout largo, función con reintentos
                        timeout_duration = 180.0 if retry_attempt > 0 else 120.0
                        await asyncio.wait_for(
                            render_one(page, html, pdf_path,
                                     scale=args.scale,
                                     prefer_css=not args.ignore_css_page,
                                     wait_ms=args.wait,
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
                        error_msg = f"[TIMEOUT] {html.name} se colgó (>{timeout_duration}s)"
                        
                        # Registrar detalles del timeout para diagnóstico
                        timeout_detail = {
                            'file': str(html),
                            'error_type': 'TimeoutError',
                            'error_msg': f'Timeout after {timeout_duration}s',
                            'is_corrupted': False,
                            'is_detached': False,
                            'is_timeout': True
                        }
                        
                        if retry_attempt == 0:
                            print(f"{error_msg}, se reintentará al final")
                            failed_htmls.append((i, html, timeout_detail))
                        else:
                            print(f"{error_msg}, CANCELADO definitivamente")
                            final_failures.append(timeout_detail)
                    return None
                    
                except Exception as e:
                    async with lock:
                        error_msg = f"[WARN] Falló {html}: {e}"
                        error_type = type(e).__name__
                        
                        # Registrar detalles del error para diagnóstico
                        error_detail = {
                            'file': str(html),
                            'error_type': error_type,
                            'error_msg': str(e)[:200],
                            'is_corrupted': "net::ERR_ABORTED" in str(e),
                            'is_detached': "frame was detached" in str(e)
                        }
                        
                        if retry_attempt == 0:
                            print(f"{error_msg}, se reintentará al final")
                            failed_htmls.append((i, html, error_detail))
                        else:
                            print(f"{error_msg}, CANCELADO definitivamente")
                            # Guardar información detallada del fallo final
                            final_failures.append(error_detail)
                    return None
                    
                finally:
                    # SIEMPRE cerrar la página, incluso si se cuelga
                    try:
                        await page.close()
                    except Exception:
                        pass  # Ignorar errores de cleanup

        # --- PRIMERA RONDA: Procesamiento normal ---
        print(f"[INFO] Iniciando conversión de {total} archivos HTML...")
        tasks = [asyncio.create_task(worker(i, html)) for i, html in enumerate(htmls, 1)]
        
        # Procesar todas las tareas con mejor manejo de excepciones
        completed_tasks = 0
        for coro in asyncio.as_completed(tasks):
            try:
                res = await coro
                if res:
                    rendered.append(res)
                completed_tasks += 1
            except Exception as e:
                # Capturar cualquier excepción no manejada para evitar "Future exception was never retrieved"
                print(f"[ERROR] Tarea no controlada falló: {str(e)[:100]}...")
                completed_tasks += 1
                continue
        
        print(f"[INFO] Primera ronda completada: {completed_tasks}/{total} tareas procesadas")

        # --- ESPERA Y REINTENTOS DIFERIDOS ---
        if failed_htmls:
            # Espera más tiempo para datasets grandes
            wait_time = min(10.0, max(3.0, len(failed_htmls) * 0.1))  # Entre 3-10 segundos
            print(f"\n[RETRY] {len(failed_htmls)} conversiones fallaron, esperando {wait_time:.1f} segundos antes de reintentar...")
            await asyncio.sleep(wait_time)
            
            print(f"[RETRY] Reintentando {len(failed_htmls)} conversiones fallidas con configuración ultra-conservadora...")
            retry_tasks = [asyncio.create_task(worker(i, html, retry_attempt=1)) for i, html, _ in failed_htmls]
            
            # Procesar reintentos con mejor manejo de excepciones
            retry_completed = 0
            for coro in asyncio.as_completed(retry_tasks):
                try:
                    res = await coro
                    if res:
                        rendered.append(res)
                    retry_completed += 1
                except Exception as e:
                    # Capturar cualquier excepción no manejada en reintentos
                    print(f"[ERROR] Reintento no controlado falló: {str(e)[:100]}...")
                    retry_completed += 1
                    continue
            
            print(f"[INFO] Reintentos completados: {retry_completed}/{len(failed_htmls)} tareas procesadas")
            
            # Mostrar estadísticas finales
            final_failed = total - len(rendered)
            if final_failed > 0:
                print(f"[FINAL] {final_failed}/{total} conversiones fallaron definitivamente")
                print(f"[FINAL] Tasa de éxito: {((total - final_failed) / total * 100):.1f}%")
                
                # Diagnóstico de tipos de errores
                error_types = {}
                corrupted_files = 0
                timeout_files = 0
                for failure in final_failures:
                    error_type = failure['error_type']
                    error_types[error_type] = error_types.get(error_type, 0) + 1
                    if failure.get('is_corrupted', False):
                        corrupted_files += 1
                    if failure.get('is_timeout', False):
                        timeout_files += 1
                
                print(f"[DIAGNOSTICO] Tipos de errores finales:")
                for error_type, count in error_types.items():
                    print(f"  - {error_type}: {count} archivos")
                
                if corrupted_files > 0:
                    print(f"[DIAGNOSTICO] {corrupted_files} archivos posiblemente corruptos (ERR_ABORTED)")
                if timeout_files > 0:
                    print(f"[DIAGNOSTICO] {timeout_files} archivos con timeout (>120s)")
            else:
                print(f"[SUCCESS] Todas las conversiones completadas exitosamente después de reintentos")
        else:
            print(f"[SUCCESS] Todas las conversiones completadas exitosamente en primera ronda")

        # Esperar un momento antes de cerrar el contexto para asegurar limpieza
        await asyncio.sleep(1.0)
        await context.close()
        await browser.close()

    rendered.sort(key=lambda tup: str(tup[2]))
    return rendered

def parse_args():
    # Usar concurrencia más conservadora por defecto
    default_concurrency = min(os.cpu_count() or 4, 12)  # Máximo 12, más agresivo para velocidad
    
    ap = argparse.ArgumentParser(description="HTML → PDF A3 (Chromium/Playwright) + merges por sección **por centro**")
    ap.add_argument("--data", required=True, help="Carpeta raíz con los .html generados")
    ap.add_argument("--out", default=None, help="Carpeta de salida (default: <data>_pdf)")
    ap.add_argument("--scale", type=float, default=1.0, help="Escala del render (1.0 por defecto)")
    ap.add_argument("--wait", type=int, default=500, help="Espera (ms) tras cargar cada HTML (por defecto 500ms - más robusto)")
    ap.add_argument("--no-merge", action="store_true", help="No crear PDFs combinados por sección")
    ap.add_argument("--ignore-css-page", action="store_true",
                    help="Ignorar @page del CSS y forzar A3 landscape desde Chromium")

    ap.add_argument("--concurrency", type=int, default=default_concurrency, help=f"Número de páginas en paralelo (default {default_concurrency})")
    ap.add_argument("--merge-workers", type=int, default=min(os.cpu_count() or 2, 4), help="Workers paralelos para merges por centro")
    ap.add_argument("--fast", action="store_true", help="Modo ultra rápido: concurrencia alta, timeouts mínimos, sin reintentos")
    ap.add_argument("--ultra-fast", action="store_true", help="Modo ultra rápido extremo: máxima concurrencia, timeouts mínimos")
    ap.add_argument("--log-every", type=int, default=10, help="Frecuencia de logs de progreso (default 10)")
    ap.add_argument("--block", default="", help="Tipos de recurso a bloquear separados por coma: image,font,media,stylesheet,script")
    ap.add_argument("--chromium-arg", action="append", default=[], help="Argumentos extra para Chromium (repetible)")

    ap.add_argument("--caratulas-dir", default=None, help="Ruta a las carátulas del Anejo (PDFs)")
    return ap.parse_args()

def crear_anejo_final_por_centro(args, out_root: Path, centros: set[str]):
    """
    Genera un 05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf por cada centro dentro de <out_root>/<CENTRO>.
    Usa las mismas carátulas para todos (si existen).
    VERSIÓN PARALELA: cada centro se procesa independientemente.
    """
    from concurrent.futures import ThreadPoolExecutor
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
    
    def process_centro_final(centro):
        """Crear anejo final para UN centro"""
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
            return False

        print(f"[INFO] ({centro}) PDFs a unir en el Anejo final: {len(pdfs_final)} archivos")

        out_final = base / "05_ANEJO 5. REPORTAJE FOTOGRÁFICO.pdf"
        writer = PdfWriter()
        for p in pdfs_final:
            try:
                writer.append(str(p))
            except Exception as e:
                print(f"[WARN] ({centro}) No se pudo añadir {p}: {e}")
        
        try:
            with out_final.open("wb") as f:
                writer.write(f)
            print(f"[OK] ({centro}) Anejo final generado: {out_final}")
            return True
        except Exception as e:
            print(f"[ERROR] ({centro}) Error generando anejo final: {e}")
            return False
    
    # Procesar centros en paralelo
    merge_workers = getattr(args, 'merge_workers', min(os.cpu_count() or 4, 8))
    print(f"[FINAL] Generando anejos finales en paralelo con {merge_workers} workers...")
    
    with ThreadPoolExecutor(max_workers=merge_workers) as executor:
        results = list(executor.map(process_centro_final, centros))
    
    successful = sum(results)
    print(f"[FINAL-SUMMARY] {successful}/{len(centros)} anejos finales generados exitosamente")

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
