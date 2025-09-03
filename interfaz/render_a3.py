# -*- coding: utf-8 -*-
r"""
Render A3 (CENTRO, EDIFICIOS, ENVOLVENTE, DEPENDENCIAS, ACOMETIDA, CC, CLIMA,
EQ HORIZONTALES, ELEVADORES, ILUMINACIÓN, OTROS EQUIPOS).

Características:
- Rejilla de fotos "max-fill" con clases .ph-grid y .photos-*
- Soporta 'fotos' y fallback desde 'fotos_paths'
- URIs file:/// para abrir rutas locales Z:\... en navegador
- Si NO hay fotos: se COPIAN A LA CARPETA DE SALIDA los SVG de placeholder
  (A3_FOTOS_ICONO.svg / A3_FOTOS_AUDITORIA_SIN_ICONO.svg) y se referencian
  en RELATIVO para evitar bloqueos del navegador entre discos/letras.
  Se buscan por prioridad:
    1) --svg / --svg2 (CLI)
    2) C:\Users\indiva\Documents\html_anexos\*.svg
    3) Carpeta de plantillas
    4) BASE_DIR (junto a datos/script)
- Clasificación ENVOLVENTE priorizando 'denominacion' (luego campos, luego 'tipo_envolvente')
- Slots de fotos compatibles (usa el específico y, si no existe, intenta un genérico)
- Tolerante a JSON/plantillas ausentes (SKIP)
- Descubrimiento de múltiples centros, o un centro único con JSON sueltos
- CLI: --data (carpeta datos), --out (salida), --tpl (plantillas), --svg/--svg2 (placeholders)
- Logs y métricas por bloque + resumen por centro

Salida:
- <OUT>/C<centro_id>/<bloque>/<archivo_por_item>.html (+ SVGs junto al HTML)
"""

import os
import re
import json
import argparse
import unicodedata
from pathlib import Path
from urllib.parse import quote
from collections import defaultdict
from shutil import copy2

# ===================== CONFIG GLOBAL (rellena con CLI) =====================
BASE_DIR = Path(os.getcwd())
PLANTILLAS_DIR = BASE_DIR / "plantillas_a3_unificadas"
SALIDA_BASE = BASE_DIR / "salida"

# Directorios que nunca tratamos como "centro"
IGNORED_DIRS = {"salida", "plantillas_a3_unificadas", ".git", "__pycache__"}

# Control de filtrado de elementos sin fotos
INCLUDE_WITHOUT_PHOTOS = True  # Por defecto incluye elementos sin fotos

# Se rellenan dinámicamente tras parsear CLI
TPLS = {}
T_ENVOL = {}

# Candidatos a SVG de "sin foto" (se rellena en parse_cli_and_set_paths)
SVG_CANDIDATES = []

DEFAULT_SVG_MAIN = "A3_FOTOS_ICONO.svg"
DEFAULT_SVG_ALT  = "A3_FOTOS_AUDITORIA_SIN_ICONO.svg"

# ===================== HELPERS =====================
def _strip(s):
    return "" if s is None else str(s).strip()

def _normalize_text(s):
    s = _strip(s).lower()
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def _is_set(val):
    if val is None:
        return False
    s = str(val).strip()
    return s not in ("", "–")

def _to_int(val, default=0):
    try:
        return int(str(val).strip())
    except Exception:
        try:
            return int(float(str(val).replace(",", ".").strip()))
        except Exception:
            return default

def _to_float(val, default=0.0):
    try:
        return float(str(val).replace(",", ".").strip())
    except Exception:
        return default

def to_file_uri(path_like: str) -> str:
    """
    'Z:\\a\\b c\\f.jpg' -> 'file:///Z:/a/b%20c/f.jpg'
    Evita doble 'file://'.
    """
    if not path_like:
        return ""
    p = str(path_like).replace("\\", "/")
    if p.lower().startswith("file://"):
        return p
    return "file:///" + quote(p, safe="/:._-()")

def _gather_svg_search_paths() -> list[Path]:
    """Candidatos donde buscar los SVG origen, en orden de prioridad."""
    candidates: list[Path] = []
    # 1) CLI (ya poblado en SVG_CANDIDATES)
    candidates.extend(SVG_CANDIDATES)
    # 2) Carpeta interfaz (misma carpeta que este script)
    interfaz_dir = Path(__file__).parent
    candidates.extend([interfaz_dir / DEFAULT_SVG_MAIN, interfaz_dir / DEFAULT_SVG_ALT])
    # 3) Rutas por defecto del usuario
    base_user = Path(r"C:\Users\indiva\Documents\html_anexos")
    candidates.extend([base_user / DEFAULT_SVG_MAIN, base_user / DEFAULT_SVG_ALT])
    # 4) Plantillas
    candidates.extend([PLANTILLAS_DIR / DEFAULT_SVG_MAIN, PLANTILLAS_DIR / DEFAULT_SVG_ALT])
    # 5) BASE_DIR (junto al script/datos)
    candidates.extend([BASE_DIR / DEFAULT_SVG_MAIN, BASE_DIR / DEFAULT_SVG_ALT])

    # Deduplicar manteniendo orden
    uniq, seen = [], set()
    for c in candidates:
        sc = str(c).lower()
        if sc not in seen:
            seen.add(sc)
            uniq.append(c)
    return uniq

def ensure_placeholders_in_outdir(out_dir: Path) -> dict:
    """
    Copia 1 o 2 SVG disponibles junto al HTML si existen.
    Devuelve rutas RELATIVAS (para usar en <img src="...">): {"main": "...", "alt": "..."}.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    search = [p for p in _gather_svg_search_paths() if p.exists()]
    result = {"main": "", "alt": ""}

    # Copia el principal (main)
    main_svg = None
    for p in search:
        if Path(p).name.lower() == Path(DEFAULT_SVG_MAIN).name.lower():
            main_svg = p
            break
    if main_svg:
        target = out_dir / Path(DEFAULT_SVG_MAIN).name
        if not target.exists():
            try:
                copy2(main_svg, target)
                print(f"[ASSET] Copiado placeholder -> {target}")
            except Exception as e:
                print(f"[WARN] No se pudo copiar {main_svg} -> {target}: {e}")
        if target.exists():
            result["main"] = Path(DEFAULT_SVG_MAIN).name

    # Copia el alternativo (alt)
    alt_svg = None
    for p in search:
        if Path(p).name.lower() == Path(DEFAULT_SVG_ALT).name.lower():
            alt_svg = p
            break
    if alt_svg:
        target = out_dir / Path(DEFAULT_SVG_ALT).name
        if not target.exists():
            try:
                copy2(alt_svg, target)
                print(f"[ASSET] Copiado placeholder alt -> {target}")
            except Exception as e:
                print(f"[WARN] No se pudo copiar {alt_svg} -> {target}: {e}")
        if target.exists():
            result["alt"] = Path(DEFAULT_SVG_ALT).name

    return result

def build_photos_grid(photos: list, placeholder_href: str = "") -> str:
    """
    photos: lista de dicts con 'path' o 'file_uri' y opcional 'name'/'id'
    placeholder_href: nombre de archivo relativo al HTML (si no hay fotos)
    """
    n = len(photos or [])
    if n == 0:
        if placeholder_href:
            return f"""
<div class="ph-grid photos-1">
  <div class="ph-card">
    <div class="ph-imgwrap" style="display:flex;align-items:center;justify-content:center;min-height:8cm;padding:0.5cm;">
      <img class="ph-img" src="{placeholder_href}" alt="Sin foto" style="max-width:100%;max-height:16cm;object-fit:contain;display:block;margin:auto;">
    </div>
    <div class="ph-cap" style="text-align:center; color:#888;">Sin foto disponible</div>
  </div>
</div>
""".strip()
        # Fallback final: texto simple
        return (
            '<div class="ph-grid photos-1">'
            '  <div class="ph-card"><div class="ph-imgwrap">'
            '    <div class="ph-cap">Sin foto disponible</div>'
            '  </div></div>'
            '</div>'
        )

    if n == 1:
        grid_cls = "photos-1"
    elif n == 2:
        grid_cls = "photos-2"
    elif 3 <= n <= 6:
        grid_cls = f"photos-{n}"
    else:
        grid_cls = "photos-many"
    cards = []
    for ph in photos:
        uri = ph.get("file_uri") or to_file_uri(ph.get("path"))
        name = _strip(ph.get("name") or ph.get("id") or "")
        cards.append(
            f"""
            <div class="ph-card">
              <div class="ph-imgwrap">
                <img class="ph-img" src="{uri}" alt="{name}">
              </div>
              <div class="ph-cap">{name}</div>
            </div>
            """.strip()
        )
    return f'<div class="ph-grid {grid_cls}">\n' + "\n".join(cards) + "\n</div>"

def _replace_tokens_simple(html: str, mapping: dict) -> str:
    out = html
    for k, v in mapping.items():
        out = out.replace("{{" + k + "}}", _strip(v))
    return out

def render_template(html_text: str, slot_tokens: tuple[str, ...], fotos_html: str, context_maps: list, title_keys=()):
    """
    - Reemplaza el primer slot encontrado en 'slot_tokens' por fotos_html.
    - Reemplaza tokens con varios diccionarios (en orden).
    - Asegura también title_keys si estuvieran en el HTML (p.ej. <title>).
    """
    result = html_text

    # 1) Fotos: usa el primer slot que exista; si ninguno está, intenta genérico de envolvente.
    slot_applied = False
    for tok in slot_tokens:
        if tok in result:
            result = result.replace(tok, fotos_html)
            slot_applied = True
            break
    if not slot_applied and "[[FOTOS_ENVOL]]" in result:
        result = result.replace("[[FOTOS_ENVOL]]", fotos_html)

    # 2) Campos
    for mp in context_maps:
        result = _replace_tokens_simple(result, mp)

    # 3) title_keys (por si la plantilla los usa en <title>)
    for k in title_keys:
        for mp in context_maps:
            if k in mp:
                result = result.replace("{{" + k + "}}", _strip(mp[k]))
                break

    return result

def ensure_outdir_for_centro(centro_id: str, tipo_subdir: str) -> Path:
    # Si centro_id ya empieza por C, no añadir otra C (evitar CC0007)
    clean_id = _strip(centro_id)
    if clean_id.startswith('C'):
        out_dir = SALIDA_BASE / clean_id / tipo_subdir
    else:
        out_dir = SALIDA_BASE / f"C{clean_id}" / tipo_subdir
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir

def basename_noext(p: str) -> str:
    try:
        return Path(p).stem
    except Exception:
        return _strip(p)

def has_photo(any_obj: dict) -> bool:
    """
    Verifica si un objeto tiene al menos una foto válida.
    Retorna True si encuentra fotos en 'fotos' o 'fotos_paths', False en caso contrario.
    """
    # Verificar campo 'fotos' (lista de objetos con 'path')
    fotos = any_obj.get("fotos") or []
    if fotos:
        for foto in fotos:
            path = foto.get("path", "").strip()
            if path and Path(path).exists():
                return True
    
    # Verificar campo 'fotos_paths' (lista de rutas directas)
    fotos_paths = any_obj.get("fotos_paths") or []
    if fotos_paths:
        for path in fotos_paths:
            path_str = str(path).strip()
            if path_str and Path(path_str).exists():
                return True
    
    return False

def collect_fotos(any_obj: dict) -> list[dict]:
    """
    Normaliza una lista de fotos a [{'path':..., 'file_uri':..., 'name':...}, ...]
    Prioriza 'fotos' si existe; si no, usa 'fotos_paths'.
    """
    fotos = any_obj.get("fotos") or []
    if not fotos:
        fps = any_obj.get("fotos_paths") or []
        fotos = [{"path": p, "name": basename_noext(p)} for p in fps]
    # Asegura file_uri y name
    norm = []
    for f in fotos:
        path = f.get("path") or ""
        file_uri = f.get("file_uri") or to_file_uri(path)
        name = f.get("name") or f.get("id") or basename_noext(path)
        norm.append({"path": path, "file_uri": file_uri, "name": name, "id": f.get("id", name)})
    return norm

# ===================== CLASIFICACIÓN ENVOLVENTE =====================
RE_SUBTIPO = [
    (re.compile(r"\bfachad(a|as)?\b", re.IGNORECASE), "FACHADA"),
    (re.compile(r"\bpuert(a|as)?\b",  re.IGNORECASE), "PUERTAS"),
    (re.compile(r"\bventan(a|as)?\b", re.IGNORECASE), "VENTANAS"),
    (re.compile(r"\bcubiert(a|as)?\b|\blucernari\w*\b|\bazotea\b", re.IGNORECASE), "CUBIERTA"),
]
RE_ORIENT = re.compile(r"\b(NO|NE|SO|SE|N|S|E|O)\b", re.IGNORECASE)

def _from_denominacion(denominacion):
    d = _normalize_text(denominacion)
    for rx, tipo in RE_SUBTIPO:
        if rx.search(d):
            return tipo
    return None

def _orient_from_denominacion(denominacion):
    if not denominacion:
        return None
    m = RE_ORIENT.search(denominacion.upper())
    return m.group(0) if m else None

def _from_fields(item):
    nv = _to_int(item.get("num_ventanas"), 0)
    if nv > 0 or any(_is_set(item.get(k)) for k in [
        "ventanas_tipo","ventanas_carpinteria","ventanas_acristalamiento","ventanas_proteccion_solar"
    ]):
        return "VENTANAS"
    np = _to_int(item.get("num_puertas"), 0)
    if np > 0 or any(_is_set(item.get(k)) for k in ["puertas_tipo","puertas_material","puertas_dimensiones"]):
        return "PUERTAS"
    sup_c = _to_float(item.get("sup_cubierta"), 0.0)
    if sup_c > 0 or any(_is_set(item.get(k)) for k in [
        "cubiertas_tipo","cubiertas_acabado","cubiertas_aislamiento","lucernario","lucernario_dimensiones"
    ]):
        return "CUBIERTA"
    if any(_is_set(item.get(k)) for k in ["fachada_tipo","fachada_aislamiento","fachada_huecos"]):
        return "FACHADA"
    return None

def _from_tipo_envolvente(item):
    t = _normalize_text(item.get("tipo_envolvente"))
    if "fachad" in t:  return "FACHADA"
    if "puert" in t:   return "PUERTAS"
    if "ventan" in t:  return "VENTANAS"
    if "cubiert" in t or "lucernari" in t or "azotea" in t: return "CUBIERTA"
    return None

def clasificar_envolvente(item):
    deno_tipo  = _from_denominacion(item.get("denominacion"))
    field_tipo = _from_fields(item) if not deno_tipo else None
    tipo_tipo  = _from_tipo_envolvente(item) if not (deno_tipo or field_tipo) else None
    subtipo = deno_tipo or field_tipo or tipo_tipo or "FACHADA"

    if deno_tipo and tipo_tipo and deno_tipo != tipo_tipo:
        print(f"[WARN] {item.get('id','(sin id)')}: denominación sugiere {deno_tipo}, "
              f"pero tipo_envolvente='{item.get('tipo_envolvente')}' sugiere {tipo_tipo}. "
              f"Se usa DENOMINACIÓN.")

    if not _is_set(item.get("orientacion")):
        oden = _orient_from_denominacion(item.get("denominacion"))
        if oden:
            item["orientacion"] = oden

    info = T_ENVOL[subtipo]
    titulo = _strip(item.get("denominacion")) or (
        subtipo.capitalize() + (f" { _strip(item.get('orientacion')) }" if _is_set(item.get("orientacion")) else "")
    )
    return {"subtipo": subtipo, "template_path": info["path"], "slots": info["slots"], "titulo": titulo}

# ===================== MÉTRICAS / LOGS =====================
# metrics[centro_id][bloque] = {"in": int, "out": int, "photos": int}
metrics = defaultdict(lambda: defaultdict(lambda: {"in": 0, "out": 0, "photos": 0}))

def add_metrics(centro_id: str, bloque: str, in_inc=0, out_inc=0, photos_inc=0):
    m = metrics[centro_id][bloque]
    m["in"] += in_inc
    m["out"] += out_inc
    m["photos"] += photos_inc

def log_block_summary(centro_id: str, bloque: str):
    m = metrics[centro_id][bloque]
    print(f"[SUMMARY] {centro_id}::{bloque} -> entradas={m['in']}, render={m['out']}, fotos={m['photos']}")

def log_center_summary(centro_id: str):
    print(f"\n=== RESUMEN CENTRO C{centro_id} ===")
    total_in = total_out = total_ph = 0
    for bloque, m in metrics[centro_id].items():
        print(f"  - {bloque:15s}  in:{m['in']:4d}  out:{m['out']:4d}  fotos:{m['photos']:4d}")
        total_in += m['in']; total_out += m['out']; total_ph += m['photos']
    print(f"  TOTAL -> in:{total_in}  out:{total_out}  fotos:{total_ph}")
    print("====================================\n")

# ===================== RENDERERS =====================
def process_centro(centro_json: Path):
    if not centro_json or not centro_json.exists():
        return
    data = json.loads(centro_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    if not centro:
        return
    tpl = TPLS["centro"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    html = tpl_path.read_text(encoding="utf-8")
    fotos = collect_fotos(centro)

    centro_id = centro.get("id", "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)
    fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))

    # métricas
    add_metrics(centro_id, "centro", in_inc=1, out_inc=1, photos_inc=len(fotos))

    flat  = {k: centro.get(k, "") for k in centro.keys()}
    pref  = {f"centro.{k}": centro.get(k, "") for k in centro.keys()}
    header = {"bloque": "CENTRO", "id": centro.get("id", "")}

    out_html = render_template(
        html_text=html,
        slot_tokens=tpl["slots"],
        fotos_html=fotos_html,
        context_maps=[flat, pref, header],
        title_keys=tpl["title_keys"],
    )
    out_name = f"{centro_id}_centro.html"
    (out_dir / out_name).write_text(out_html, encoding="utf-8")
    print(f"[OK] CENTRO -> {out_dir / out_name}")
    log_block_summary(centro_id, "centro")

def process_edificios(edificios_json: Path):
    if not edificios_json or not edificios_json.exists():
        return
    data = json.loads(edificios_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    edificios = data.get("edificios") or []
    if not edificios:
        return

    tpl = TPLS["edificios"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")

    centro_id = centro.get("id") or (edificios[0].get("id_centro") if edificios else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "edificios", in_inc=len(edificios))

    for e in edificios:
        # Aplicar filtro de fotos si está habilitado
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(e):
            print(f"[SKIP] Edificio {e.get('id', 'SINID')} - sin fotos (filtrado habilitado)")
            continue
            
        fotos = collect_fotos(e)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: e.get(k, "") for k in e.keys()}
        pref_e = {f"e.{k}": e.get(k, "") for k in e.keys()}
        pref_c = {f"centro.{k}": centro.get(k, "") for k in centro.keys()}
        header = {"bloque": e.get("bloque", "EDIFICIOS"), "id": e.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, pref_e, pref_c, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{e.get('id','SINID')}_edificio.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "edificios", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] EDIFICIO -> {out_dir / out_name}")

    log_block_summary(centro_id, "edificios")

def process_envolventes(envolventes_json: Path):
    if not envolventes_json or not envolventes_json.exists():
        return
    data = json.loads(envolventes_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    envolventes = data.get("envolventes") or []
    if not envolventes:
        return

    centro_id = centro.get("id") or (envolventes[0].get("id_centro") if envolventes else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, "envolventes")
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "envolventes", in_inc=len(envolventes))

    for env in envolventes:
        clasif = clasificar_envolvente(env)
        tpl_path = clasif["template_path"]
        if not tpl_path.exists():
            print(f"[AVISO] Falta plantilla: {tpl_path}")
            continue
        
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(env):
            print(f"[SKIP] Envolvente {env.get('id', 'SINID')} omitida (sin fotos)")
            continue
            
        tpl_html = tpl_path.read_text(encoding="utf-8")
        fotos = collect_fotos(env)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))

        flat   = {k: env.get(k, "") for k in env.keys()}
        header = {"bloque": clasif["subtipo"], "id": env.get("id", ""), "titulo": clasif["titulo"]}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=clasif["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=("bloque", "id", "titulo"),
        )

        out_name = f"{env.get('id','SINID')}_{clasif['subtipo'].lower()}.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "envolventes", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] ENVOLVENTE -> {out_dir / out_name}")

    log_block_summary(centro_id, "envolventes")

def process_dependencias(dependencias_json: Path):
    if not dependencias_json or not dependencias_json.exists():
        return
    data = json.loads(dependencias_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    dependencias = data.get("dependencias") or []
    if not dependencias:
        return
    tpl = TPLS["dependencias"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (dependencias[0].get("id_centro") if dependencias else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "dependencias", in_inc=len(dependencias))

    for dep in dependencias:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(dep):
            print(f"[SKIP] Dependencia {dep.get('id', 'SINID')} omitida (sin fotos)")
            continue
            
        fotos = collect_fotos(dep)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: dep.get(k, "") for k in dep.keys()}
        header = {"bloque": dep.get("bloque", "DEPENDENCIAS"), "id": dep.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{dep.get('id','SINID')}_dependencia.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "dependencias", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] DEPENDENCIA -> {out_dir / out_name}")

    log_block_summary(centro_id, "dependencias")

def process_acometida(acom_json: Path):
    if not acom_json or not acom_json.exists():
        return
    data = json.loads(acom_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    acoms = data.get("acom") or []
    if not acoms:
        return
    tpl = TPLS["acom"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (acoms[0].get("id_centro") if acoms else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "acometida", in_inc=len(acoms))

    for acom in acoms:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(acom):
            print(f"[SKIP] Acometida {acom.get('id', 'SINID')} omitida (sin fotos)")
            continue
            
        fotos = collect_fotos(acom)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: acom.get(k, "") for k in acom.keys()}
        header = {"bloque": acom.get("bloque", "ACOMETIDA"), "id": acom.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{acom.get('id','SINID')}_acom.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "acometida", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] ACOMETIDA -> {out_dir / out_name}")

    log_block_summary(centro_id, "acometida")

def process_cc(cc_json: Path):
    if not cc_json or not cc_json.exists():
        return
    data = json.loads(cc_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    sistemas = data.get("sistemas_cc") or []
    if not sistemas:
        return
    tpl = TPLS["cc"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (sistemas[0].get("id_centro") if sistemas else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "cc", in_inc=len(sistemas))

    for sis in sistemas:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(sis):
            print(f"[SKIP] Sistema CC {sis.get('id', 'SINID')} omitido (sin fotos)")
            continue
            
        fotos = collect_fotos(sis)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: sis.get(k, "") for k in sis.keys()}
        header = {"bloque": sis.get("bloque", "CALEFACCIÓN"), "id": sis.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{sis.get('id','SINID')}_cc.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "cc", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] CC -> {out_dir / out_name}")

    log_block_summary(centro_id, "cc")

def process_clima(clima_json: Path):
    if not clima_json or not clima_json.exists():
        return
    data = json.loads(clima_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    equipos = data.get("equipos_clima") or []
    if not equipos:
        return
    tpl = TPLS["clima"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (equipos[0].get("id_centro") if equipos else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "clima", in_inc=len(equipos))

    for eq in equipos:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(eq):
            print(f"[SKIP] Equipo clima {eq.get('id', 'SINID')} omitido (sin fotos)")
            continue
            
        fotos = collect_fotos(eq)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: eq.get(k, "") for k in eq.keys()}
        header = {"bloque": eq.get("bloque", "CLIMATIZACIÓN"), "id": eq.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{eq.get('id','SINID')}_clima.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "clima", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] CLIMA -> {out_dir / out_name}")

    log_block_summary(centro_id, "clima")

def process_eqhoriz(eqhoriz_json: Path):
    if not eqhoriz_json or not eqhoriz_json.exists():
        return
    data = json.loads(eqhoriz_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    equipos = data.get("equipos_horiz") or []
    if not equipos:
        return
    tpl = TPLS["eqh"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (equipos[0].get("id_centro") if equipos else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "eqhoriz", in_inc=len(equipos))

    for eq in equipos:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(eq):
            print(f"[SKIP] Equipo horizontal {eq.get('id', 'SINID')} omitido (sin fotos)")
            continue
            
        fotos = collect_fotos(eq)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: eq.get(k, "") for k in eq.keys()}
        header = {"bloque": eq.get("bloque", "EQUIPOS HORIZONTALES"), "id": eq.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{eq.get('id','SINID')}_eqh.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "eqhoriz", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] EQH -> {out_dir / out_name}")

    log_block_summary(centro_id, "eqhoriz")

def process_elevadores(eleva_json: Path):
    if not eleva_json or not eleva_json.exists():
        return
    data = json.loads(eleva_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    elevadores = data.get("elevadores") or []
    if not elevadores:
        return
    tpl = TPLS["eleva"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (elevadores[0].get("id_centro") if elevadores else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "elevadores", in_inc=len(elevadores))

    for eq in elevadores:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(eq):
            print(f"[SKIP] Elevador {eq.get('id', 'SINID')} omitido (sin fotos)")
            continue
            
        fotos = collect_fotos(eq)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: eq.get(k, "") for k in eq.keys()}
        header = {"bloque": eq.get("bloque", "ELEVADORES"), "id": eq.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{eq.get('id','SINID')}_eleva.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "elevadores", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] ELEVA -> {out_dir / out_name}")

    log_block_summary(centro_id, "elevadores")

def process_iluminacion(iluminacion_json: Path):
    if not iluminacion_json or not iluminacion_json.exists():
        return
    data = json.loads(iluminacion_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    elementos = data.get("iluminacion") or []
    if not elementos:
        return
    tpl = TPLS["iluminacion"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (elementos[0].get("id_centro") if elementos else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "iluminacion", in_inc=len(elementos))

    for eq in elementos:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(eq):
            print(f"[SKIP] Iluminación {eq.get('id', 'SINID')} omitida (sin fotos)")
            continue
            
        fotos = collect_fotos(eq)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: eq.get(k, "") for k in eq.keys()}
        header = {"bloque": eq.get("bloque", "ILUMINACIÓN"), "id": eq.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{eq.get('id','SINID')}_iluminacion.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "iluminacion", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] ILUMINACIÓN -> {out_dir / out_name}")

    log_block_summary(centro_id, "iluminacion")

def process_otrosequipos(otrosequipos_json: Path):
    if not otrosequipos_json or not otrosequipos_json.exists():
        return
    data = json.loads(otrosequipos_json.read_text(encoding="utf-8"))
    centro = data.get("centro") or {}
    elementos = data.get("otros_equipos") or []
    if not elementos:
        return
    tpl = TPLS["otros"]
    tpl_path = tpl["path"]
    if not tpl_path.exists():
        print(f"[AVISO] Falta plantilla: {tpl_path}")
        return
    tpl_html = tpl_path.read_text(encoding="utf-8")
    centro_id = centro.get("id") or (elementos[0].get("id_centro") if elementos else "SINID")
    out_dir = ensure_outdir_for_centro(centro_id, tpl["out_subdir"])
    placeholders = ensure_placeholders_in_outdir(out_dir)

    add_metrics(centro_id, "otros_equipos", in_inc=len(elementos))

    for eq in elementos:
        # Skip elements without photos if filtering is enabled
        if not INCLUDE_WITHOUT_PHOTOS and not has_photo(eq):
            print(f"[SKIP] Otro equipo {eq.get('id', 'SINID')} omitido (sin fotos)")
            continue
            
        fotos = collect_fotos(eq)
        fotos_html = build_photos_grid(fotos, placeholders.get("main") or placeholders.get("alt"))
        flat   = {k: eq.get(k, "") for k in eq.keys()}
        header = {"bloque": eq.get("bloque", "OTROS EQUIPOS"), "id": eq.get("id", "")}

        html = render_template(
            html_text=tpl_html,
            slot_tokens=tpl["slots"],
            fotos_html=fotos_html,
            context_maps=[flat, header],
            title_keys=tpl["title_keys"],
        )
        out_name = f"{eq.get('id','SINID')}_otros.html"
        (out_dir / out_name).write_text(html, encoding="utf-8")
        add_metrics(centro_id, "otros_equipos", out_inc=1, photos_inc=len(fotos))
        print(f"[OK] OTROS EQUIPOS -> {out_dir / out_name}")

    log_block_summary(centro_id, "otros_equipos")

# ===================== DISCOVERY & MAIN =====================
def is_center_dir(d: Path) -> bool:
    if not d.is_dir() or d.name in IGNORED_DIRS:
        return False
    names = {
        "centro.json", "edificios.json", "envolventes.json", "envol.json", "dependencias.json",
        "acom.json", "cc.json", "clima.json", "eqhoriz.json", "eleva.json", "iluminacion.json",
        "ilum.json", "otroseq.json"
    }
    return any((d / n).exists() for n in names)

def discover_center_dirs(base: Path) -> list[Path]:
    return [p for p in base.iterdir() if is_center_dir(p)]

def first_existing(*candidates):
    for c in candidates:
        if c and Path(c).exists():
            return Path(c)
    return None

def run_for_dir(d: Path):
    cj    = first_existing(d / "centro.json",             BASE_DIR / "centro.json")
    ej    = first_existing(d / "edificios.json",          BASE_DIR / "edificios.json")
    envj  = first_existing(d / "envolventes.json", d / "envol.json",
                           BASE_DIR / "envolventes.json", BASE_DIR / "envol.json")
    depj  = first_existing(d / "dependencias.json",       BASE_DIR / "dependencias.json")
    acomj = first_existing(d / "acom.json",               BASE_DIR / "acom.json")
    ccj   = first_existing(d / "cc.json",                 BASE_DIR / "cc.json")
    climaj= first_existing(d / "clima.json",              BASE_DIR / "clima.json")
    eqhj  = first_existing(d / "eqhoriz.json",            BASE_DIR / "eqhoriz.json")
    elevj = first_existing(d / "eleva.json",              BASE_DIR / "eleva.json")
    ilumj = first_existing(d / "iluminacion.json", d / "ilum.json",
                           BASE_DIR / "iluminacion.json", BASE_DIR / "ilum.json")
    otroj = first_existing(d / "otroseq.json",            BASE_DIR / "otroseq.json")

    if cj:    process_centro(cj)
    else:     print(f"[SKIP] centro.json no encontrado en {d}")

    if ej:    process_edificios(ej)
    else:     print(f"[SKIP] edificios.json no encontrado en {d}")

    if envj:  process_envolventes(envj)
    else:     print(f"[SKIP] envolventes/envol(.json) no encontrado en {d}")

    if depj:  process_dependencias(depj)
    else:     print(f"[SKIP] dependencias.json no encontrado en {d}")

    if acomj: process_acometida(acomj)
    else:     print(f"[SKIP] acom.json no encontrado en {d}")

    if ccj:   process_cc(ccj)
    else:     print(f"[SKIP] cc.json no encontrado en {d}")

    if climaj:process_clima(climaj)
    else:     print(f"[SKIP] clima.json no encontrado en {d}")

    if eqhj:  process_eqhoriz(eqhj)
    else:     print(f"[SKIP] eqhoriz.json no encontrado en {d}")

    if elevj: process_elevadores(elevj)
    else:     print(f"[SKIP] eleva.json no encontrado en {d}")

    if ilumj: process_iluminacion(ilumj)
    else:     print(f"[SKIP] iluminacion/ilum(.json) no encontrado en {d}")

    if otroj: process_otrosequipos(otroj)
    else:     print(f"[SKIP] otroseq.json no encontrado en {d}")

    # Resumen del centro si lo conocemos
    centro_id = None
    for candidate in (cj, ej, envj, depj, acomj, ccj, climaj, eqhj, elevj, ilumj, otroj):
        if candidate and candidate.exists():
            try:
                data = json.loads(candidate.read_text(encoding="utf-8"))
                c = data.get("centro") or {}
                centro_id = c.get("id")
                if not centro_id:
                    for key in ("edificios","envolventes","dependencias","acom","sistemas_cc",
                                "equipos_clima","equipos_horiz","elevadores","iluminacion","otros_equipos"):
                        li = data.get(key) or []
                        if li and isinstance(li, list):
                            centro_id = li[0].get("id_centro")
                            if centro_id:
                                break
                if centro_id:
                    break
            except Exception:
                pass
    if centro_id:
        log_center_summary(centro_id)

def build_template_maps():
    """Reconstruye los mapeos de plantillas con la PLANTILLAS_DIR actual."""
    global TPLS, T_ENVOL
    TPLS = {
        "centro": {
            "path": PLANTILLAS_DIR / "centro.html",
            "slots": ("[[FOTOS_CENTRO]]",),
            "out_subdir": "centro",
            "title_keys": ("bloque", "id"),
        },
        "edificios": {
            "path": PLANTILLAS_DIR / "edificios.html",
            "slots": ("[[FOTOS_EDIFICIOS]]",),
            "out_subdir": "edificios",
            "title_keys": ("bloque", "id"),
        },
        "dependencias": {
            "path": PLANTILLAS_DIR / "dependencias.html",
            "slots": ("[[FOTOS_DEP]]",),
            "out_subdir": "dependencias",
            "title_keys": ("bloque", "id"),
        },
        "acom": {
            "path": PLANTILLAS_DIR / "acom.html",
            "slots": ("[[FOTOS_ACOM]]",),
            "out_subdir": "acometida",
            "title_keys": ("bloque", "id"),
        },
        "cc": {
            "path": PLANTILLAS_DIR / "cc.html",
            "slots": ("[[FOTOS_CC]]",),
            "out_subdir": "cc",
            "title_keys": ("bloque", "id"),
        },
        "clima": {
            "path": PLANTILLAS_DIR / "clima.html",
            "slots": ("[[FOTOS_CLIMA]]",),
            "out_subdir": "clima",
            "title_keys": ("bloque", "id"),
        },
        "eqh": {
            "path": PLANTILLAS_DIR / "eqh.html",
            "slots": ("[[FOTOS_EQH]]",),
            "out_subdir": "eqhoriz",
            "title_keys": ("bloque", "id"),
        },
        "eleva": {
            "path": PLANTILLAS_DIR / "eleva.html",
            "slots": ("[[FOTOS_ELEVA]]",),
            "out_subdir": "elevadores",
            "title_keys": ("bloque", "id"),
        },
        "iluminacion": {
            "path": PLANTILLAS_DIR / "iluminacion.html",
            "slots": ("[[FOTOS_ILUM]]",),
            "out_subdir": "iluminacion",
            "title_keys": ("bloque", "id"),
        },
        "otros": {
            "path": PLANTILLAS_DIR / "otros.html",
            "slots": ("[[FOTOS_OTROS]]",),
            "out_subdir": "otros_equipos",
            "title_keys": ("bloque", "id"),
        },
    }
    T_ENVOL = {
        "FACHADA": {
            "path": PLANTILLAS_DIR / "envol_fachada.html",
            "slots": ("[[FOTOS_ENVOL_FACHADA]]", "[[FOTOS_ENVOL]]"),
        },
        "PUERTAS": {
            "path": PLANTILLAS_DIR / "envol_puertas.html",
            "slots": ("[[FOTOS_ENVOL_PUERTAS]]", "[[FOTOS_ENVOL]]"),
        },
        "VENTANAS": {
            "path": PLANTILLAS_DIR / "envol_ventanas.html",
            "slots": ("[[FOTOS_ENVOL_VENTANAS]]", "[[FOTOS_ENVOL]]"),
        },
        "CUBIERTA": {
            "path": PLANTILLAS_DIR / "envol_cubierta.html",
            "slots": ("[[FOTOS_ENVOL_CUBIERTA]]", "[[FOTOS_ENVOL]]"),
        },
    }

def parse_cli_and_set_paths():
    """
    --data  -> carpeta raíz con centros (o con un solo centro/JSONs sueltos)
    --out   -> carpeta de salida (default: <data>/salida)
    --tpl   -> carpeta de plantillas (default: <data>/plantillas_a3_unificadas)
    --svg   -> ruta a SVG placeholder principal (opcional)
    --svg2  -> ruta a SVG placeholder alternativo (opcional)
    --include-without-photos -> incluir elementos sin fotos (default: True)
    """
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=os.getcwd(), help="Carpeta raíz de datos (centros o un centro).")
    ap.add_argument("--out",  default=None,        help="Carpeta de salida (default: <data>/salida).")
    ap.add_argument("--tpl",  default=None,        help="Carpeta de plantillas (default: <data>/plantillas_a3_unificadas).")
    ap.add_argument("--svg",  default=None,        help="Ruta a SVG placeholder principal (opcional).")
    ap.add_argument("--svg2", default=None,        help="Ruta a SVG placeholder alternativo (opcional).")
    ap.add_argument("--exclude-without-photos", action="store_true", default=False,
                    help="Excluir elementos sin fotos del Anejo 5.")
    args = ap.parse_args()

    data_dir = Path(args.data).resolve()
    out_dir  = Path(args.out).resolve() if args.out else (data_dir / "salida")
    tpl_dir  = Path(args.tpl).resolve() if args.tpl else (data_dir / "plantillas_a3_unificadas")

    global BASE_DIR, SALIDA_BASE, PLANTILLAS_DIR, SVG_CANDIDATES, INCLUDE_WITHOUT_PHOTOS
    BASE_DIR = data_dir
    SALIDA_BASE = out_dir
    PLANTILLAS_DIR = tpl_dir
    
    # Configurar filtrado de elementos sin fotos
    # Por defecto incluir elementos sin fotos (True), solo excluir si se pasa --exclude-without-photos
    INCLUDE_WITHOUT_PHOTOS = not args.exclude_without_photos

    # Placeholders señalados por CLI primero en prioridad
    SVG_CANDIDATES = []
    if args.svg:
        SVG_CANDIDATES.append(Path(args.svg))
    if args.svg2:
        SVG_CANDIDATES.append(Path(args.svg2))

    SALIDA_BASE.mkdir(parents=True, exist_ok=True)

    # reconstruye mapeos de plantillas con la PLANTILLAS_DIR actual
    build_template_maps()

    print(f"[SETUP] DATA: {BASE_DIR}")
    print(f"[SETUP] OUT : {SALIDA_BASE}")
    print(f"[SETUP] TPL : {PLANTILLAS_DIR}")
    if SVG_CANDIDATES:
        print("[SETUP] SVG placeholders:", ", ".join(str(p) for p in SVG_CANDIDATES))

def main():
    parse_cli_and_set_paths()
    candidates = discover_center_dirs(BASE_DIR)
    if not candidates:
        print("[INFO] Modo 'un solo centro' (JSON sueltos en --data).")
        run_for_dir(BASE_DIR)
        return
    print(f"[INFO] Detectados {len(candidates)} centros.")
    for d in candidates:
        print(f"\n>>> Procesando centro en: {d}")
        run_for_dir(d)

if __name__ == "__main__":
    main()
