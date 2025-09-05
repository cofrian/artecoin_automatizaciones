# -*- coding: utf-8 -*-
"""
Script para generar memoria final completa por centro.

Este script tiene dos funcionalidades:
1. Generar índices generales (001_INDICE_GENERAL.docx + PDF)
2. Montar memoria completa (MEMORIA_COMPLETA.pdf)

Uso:
    python render_memoria.py --input-dir "Y:/2025/.../06_REDACCION" --output-dir "C:/salida" --center "C0007" --action "indices"
    python render_memoria.py --input-dir "Y:/2025/.../06_REDACCION" --output-dir "C:/salida" --center "C0007" --action "memoria"
    python render_memoria.py --input-dir "Y:/2025/.../06_REDACCION" --output-dir "C:/salida" --center "C0007" --action "all"
"""

import argparse
import logging
import os
import re
import shutil
import sys
import unicodedata
from pathlib import Path
from typing import Optional, List, Dict

import pandas as pd
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader, PdfMerger, PdfWriter

try:
    import win32com.client as win32_client
    import pythoncom
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

# ================================================================
# CONFIGURACIÓN GLOBAL
# ================================================================
NAS_GROUPS = ["01_VARIOS EDIFICIOS", "02_UN EDIFICIO"]

# Ruta absoluta de la plantilla de índice
TEMPLATE_PATH = Path(__file__).parent.parent / "word" / "anexos" / "001_INDICE GENERAL_PLANTILLA.docx"

TITULOS_FIJOS = [
    "METODOLOGÍA",
    "FACTURACIÓN ENERGÉTICA",
    "INVENTARIO ENERGÉTICO", 
    "INVENTARIO SISTEMA CONSTRUCTIVO",
    "REPORTAJE FOTOGRÁFICO",
    "CERTIFICADOS ENERGÉTICOS",
    "PLANOS",
]

OUTPUT_NAME = "001_INDICE_GENERAL.docx"

OFFSET_MEMORIA = 3
ANCHO_LINEA = 86

# ================================================================
# FUNCIONES AUXILIARES COMUNES
# ================================================================
CODE_RE = re.compile(r"(C[-_ ]?\d{4})", re.IGNORECASE)

def normalize_code(raw: str) -> str:
    raw = raw.upper().replace(" ", "").replace("-", "").replace("_", "")
    m = re.search(r"C(\d{4})", raw)
    return f"C{m.group(1)}" if m else raw

def best_code_from_path(path: Path) -> Optional[str]:
    for part in reversed(path.parts):
        m = CODE_RE.search(part.upper())
        if m:
            return normalize_code(m.group(1))
    return None

def find_anejos_dir(start: Path) -> Optional[Path]:
    cand = start / "ANEJOS"
    if cand.exists() and cand.is_dir():
        return cand
    for sub in start.iterdir():
        if sub.is_dir():
            cand2 = sub / "ANEJOS"
            if cand2.exists() and cand2.is_dir():
                return cand2
    return None

def is_pdf(p: Path) -> bool:
    return p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith("~")

def contar_paginas_pdf(pdf_path: Optional[Path]) -> int:
    if not pdf_path or not pdf_path.exists():
        return 0
    try:
        reader = PdfReader(str(pdf_path))
        return len(reader.pages)
    except Exception:
        return 0

def pdf_ok(path: Optional[Path]) -> bool:
    if not path or not path.exists() or not path.is_file():
        return False
    try:
        PdfReader(str(path))  # valida apertura
        return True
    except Exception:
        return False

# ================================================================
# FUNCIONES PARA GENERAR ÍNDICES
# ================================================================

def detect_portada(building_dir: Path) -> Optional[Path]:
    """Detectar portada en carpeta del centro - versión mejorada."""
    pdfs = [p for p in building_dir.glob("*.pdf") if p.is_file()]
    # Preferir nombre que contenga PORTADA
    for p in pdfs:
        if "PORTADA" in p.stem.upper():
            return p
    # Si no hay, None
    return None

def detect_auditoria(building_dir: Path, portada: Optional[Path]) -> Optional[Path]:
    """Detectar auditoría en carpeta del centro - versión mejorada."""
    pdfs = [p for p in building_dir.glob("*.pdf") if p.is_file()]
    # Excluir portada
    pdfs = [p for p in pdfs if p != portada]
    # Evitar coger anejos por si estuvieran en el raíz
    pdfs = [p for p in pdfs if "ANEJO" not in p.stem.upper()]
    # Preferir nombres con AUDITORIA / AUDITORÍA / DOCUMENTO 3
    for p in pdfs:
        u = p.stem.upper()
        if "AUDITORIA" in u or "AUDITORÍA" in u or "DOCUMENTO 3" in u:
            return p
    # Fallback: el primer PDF que no sea portada
    return pdfs[0] if pdfs else None
    return pdfs[0] if pdfs else None

def find_existing_anejos(anejos_dir: Path) -> List[Dict]:
    """Busca anejos existentes usando la lógica del creador_indice que funciona."""
    anexos = []
    if not anejos_dir or not anejos_dir.exists():
        logger.info(f"    Directorio ANEJOS no existe: {anejos_dir}")
        return anexos
        
    pdfs_anejos = list(anejos_dir.glob("*.pdf"))
    logger.info(f"    Encontrados {len(pdfs_anejos)} PDFs en ANEJOS: {[p.name for p in pdfs_anejos]}")
    
    # Usar la lógica exacta del creador_indice que funciona
    for i, titulo_fijo in enumerate(TITULOS_FIJOS, 1):
        patron = f"{i:02d}_ANEJO {i}."
        encontrados = [f for f in pdfs_anejos if f.name.upper().startswith(patron.upper())]
        if encontrados:
            anejo_file = encontrados[0]
            paginas = contar_paginas_pdf(anejo_file)
            logger.info(f"    ✓ Anejo {i}: {anejo_file.name} ({paginas} páginas)")
            anexos.append({"numero": str(i), "titulo": titulo_fijo, "extension": paginas})
        else:
            logger.info(f"    ✗ Anejo {i} no encontrado con patrón: {patron}")
            
    return anexos

def _titulo_compuesto(anejo: Dict) -> str:
    numero_txt = str(anejo.get("numero", "")).strip()
    m = re.search(r"\d+", numero_txt)
    num = m.group(0) if m else numero_txt
    base = str(anejo.get("titulo", "")).replace('_', '').strip()
    resultado = f"ANEJO {num}: {base}".upper()
    logger.debug(f"        _titulo_compuesto: {anejo} -> {resultado}")
    return resultado

def _visual_len(s: str) -> int:
    return sum(1 for c in s if unicodedata.category(c)[0] != 'C')

def convert_docx_to_pdf(docx_path: Path) -> Optional[Path]:
    """Convierte un archivo DOCX a PDF usando Microsoft Word - mantiene mismo nombre base."""
    if not WORD_AVAILABLE:
        logger.warning("Microsoft Word no disponible. No se puede convertir DOCX a PDF")
        return None
    
    # Generar PDF con el mismo nombre base (001_INDICE_GENERAL.pdf)
    pdf_path = docx_path.with_suffix(".pdf")
    
    try:
        pythoncom.CoInitialize()
        word_app = win32_client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.ScreenUpdating = False
        
        try:
            word_app.DisplayAlerts = False
        except Exception:
            pass
        
        try:
            doc = word_app.Documents.Open(str(docx_path))
            doc.SaveAs2(str(pdf_path), FileFormat=17)  # 17 = PDF format
            doc.Close()
            logger.info(f"   -> Convertido a PDF: {pdf_path.name}")
            return pdf_path
        except Exception as e:
            logger.error(f"Error al convertir {docx_path.name} a PDF: {e}")
            return None
        finally:
            try:
                word_app.Quit()
            except Exception:
                pass
    except Exception as e:
        logger.error(f"Error al inicializar Word para conversión: {e}")
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def render_indice_general(template_path: Path, output_path: Path, auditoria_paginas: int, anexos: List[Dict]):
    """Genera el índice general usando la lógica exacta del creador_indice que funciona."""
    logger.info(f"    Renderizando con plantilla: {template_path}")
    logger.info(f"    Archivo de salida: {output_path}")
    logger.info(f"    ¿Existe la plantilla?: {template_path.exists()}")
    
    if not template_path.exists():
        logger.error(f"    ERROR: Plantilla no existe en: {template_path}")
        raise FileNotFoundError(f"Plantilla no encontrada: {template_path}")
    
    doc = DocxTemplate(str(template_path))
    anexos_calc = calcular_paginas_inicio(anexos, auditoria_paginas)
    contexto = {"e_aud": OFFSET_MEMORIA, "anexos": anexos_calc}
    
    logger.info(f"    CONTEXTO PARA PLANTILLA:")
    logger.info(f"    - e_aud: {contexto['e_aud']}")
    logger.info(f"    - anexos: {len(contexto['anexos'])} elementos")
    for i, anejo in enumerate(contexto['anexos']):
        logger.info(f"      {i+1}: {anejo.get('numero', '?')} - Página {anejo.get('pagina_inicio', '?')}")
    
    doc.render(contexto)
    doc.save(str(output_path))
    
    # Generar también versión PDF
    pdf_path = convert_docx_to_pdf(output_path)
    if pdf_path:
        logger.info(f"   -> Índice PDF generado: {pdf_path.name}")
    else:
        logger.warning(f"   -> No se pudo generar PDF para: {output_path.name}")

def render_indice_general_mejorado(template_path: Path, output_path: Path, auditoria_paginas: int, anexos: List[Dict],
                          mostrar_inicio_doc1=True, offset_memoria=OFFSET_MEMORIA, ancho_linea=ANCHO_LINEA):
    """Genera el índice general usando la lógica exacta del creador_indice que funciona."""
    # Usar la función base que funciona
    render_indice_general(template_path, output_path, auditoria_paginas, anexos)

def calcular_paginas_inicio(anexos: List[Dict], auditoria_paginas: int,
                            offset_memoria=OFFSET_MEMORIA, ancho_linea=ANCHO_LINEA) -> List[Dict]:
    """Cálculo exacto del creador_indice que funciona."""
    pagina = offset_memoria + int(auditoria_paginas)
    out = []
    for a in anexos:
        ext = int(a.get("extension", 0))
        item = {**a}
        titulo = _titulo_compuesto(a).rstrip()
        l = _visual_len(titulo)
        if l < ancho_linea:
            titulo = titulo + ("_" * (ancho_linea - l))
        else:
            count, res = 0, ""
            for c in titulo:
                if count >= ancho_linea:
                    break
                if unicodedata.category(c)[0] != 'C':
                    count += 1
def calcular_paginas_inicio(anexos: List[Dict], auditoria_paginas: int,
                            offset_memoria=OFFSET_MEMORIA, ancho_linea=ANCHO_LINEA) -> List[Dict]:
    """Cálculo exacto del creador_indice que funciona - CORREGIDO para usar anejo.e"""
    logger.info(f"    CALCULANDO PÁGINAS:")
    logger.info(f"    - Offset memoria: {offset_memoria}")
    logger.info(f"    - Páginas de auditoría: {auditoria_paginas}")
    
    pagina = offset_memoria + int(auditoria_paginas)
    logger.info(f"    - Página inicial (offset + auditoría): {pagina}")
    
    out = []
    for i, a in enumerate(anexos):
        ext = int(a.get("extension", 0))
        item = {**a}
        
        # Título compuesto con formato exacto
        titulo = _titulo_compuesto(a).rstrip()
        l = _visual_len(titulo)
        if l < ancho_linea:
            titulo = titulo + ("_" * (ancho_linea - l))
        else:
            count, res = 0, ""
            for c in titulo:
                if count >= ancho_linea:
                    break
                if unicodedata.category(c)[0] != 'C':
                    count += 1
                res += c
            titulo = res
        
        item["titulo"] = titulo
        item["pagina_inicio"] = pagina  # para compatibilidad futura
        item["e"] = pagina              # ⭐ CLAVE: la plantilla usa {{ anejo.e }}
        
        logger.info(f"    - Anejo {i+1}: {a.get('numero', '?')} - {a.get('titulo', '?')} -> Página {pagina} (extensión: {ext})")
        
        out.append(item)
        pagina += ext
    
    logger.info(f"    RESULTADO: {len(out)} anejos procesados con variable 'e' para la plantilla")
    return out

def calcular_paginas_inicio_mejorado(anexos: List[Dict], auditoria_paginas: int,
                            offset_memoria=OFFSET_MEMORIA, ancho_linea=ANCHO_LINEA) -> List[Dict]:
    """Versión mejorada del cálculo - usa la función base que funciona."""
    return calcular_paginas_inicio(anexos, auditoria_paginas, offset_memoria, ancho_linea)

# ================================================================
# FUNCIONES PARA MEMORIA COMPLETA
# ================================================================

def pick_indice(center: Path) -> Optional[Path]:
    # Preferir 001_INDICE_GENERAL*.pdf (incluye COMPLETADO)
    for p in center.glob("001_*INDICE*GENERAL*.pdf"):
        if is_pdf(p): return p
    # fallback: cualquier PDF con INDICE y GENERAL en el nombre
    for p in center.glob("*.pdf"):
        u = p.stem.upper()
        if is_pdf(p) and ("INDICE" in u or "ÍNDICE" in u) and "GENERAL" in u:
            return p
    return None

def pick_auditoria_completa(center: Path, portada: Optional[Path], indice: Optional[Path]) -> Optional[Path]:
    # Candidatos en raíz que no sean portada/índice/anejos
    for p in center.glob("*.pdf"):
        if not is_pdf(p): 
            continue
        if portada and p == portada: 
            continue
        if indice and p == indice: 
            continue
        # evitar coger un anejo que por error esté en raíz
        if "ANEJO" in p.stem.upper(): 
            continue
        u = p.stem.upper()
        # preferir DOCUMENTO 1 / AUDITORIA
        if "DOCUMENTO" in u or "AUDITOR" in u:
            return p
    # si no hay preferidos, el primer PDF "restante"
    for p in center.glob("*.pdf"):
        if not is_pdf(p): 
            continue
        if portada and p == portada: 
            continue
        if indice and p == indice: 
            continue
        if "ANEJO" in p.stem.upper(): 
            continue
        return p
    return None

def list_anejospdf(anejos_dir: Path) -> Dict[int, Path]:
    """Devuelve {n: path} con el primer PDF que cumple NN_ANEJO N.* para n=1..7."""
    out: Dict[int, Path] = {}
    if not anejos_dir or not anejos_dir.exists():
        return out
    pdfs = [p for p in anejos_dir.glob("*.pdf") if is_pdf(p)]
    for n in range(1, 20):  # por si hay más de 7 en el futuro
        patron = f"{n:02d}_ANEJO {n}."
        for p in pdfs:
            if p.name.upper().startswith(patron.upper()):
                out[n] = p
                break
    return out

# ================================================================
# FUNCIONES PRINCIPALES
# ================================================================

# ================================================================
# Generar "001_INDICE_GENERAL.docx" en TODOS los centros
# - Recorre 01_VARIOS EDIFICIOS y 02_UN EDIFICIO  
# - Usa tu plantilla y tu lógica de títulos/anchos/paginación
# - Solo incluye anejos que EXISTEN (patrón "NN_ANEJO N.")
# - Genera también PDF automáticamente con el mismo nombre
# - Usa anejo.e en lugar de anejo.pagina_inicio para la plantilla
# ================================================================

def generar_indices(nas_root: Path, center_filter: str = None, template_path: Path = None) -> int:
    """Generar índices generales en todos los centros."""
    logger.info("=== GENERANDO ÍNDICES GENERALES ===")
    logger.info(f"NAS Root: {nas_root}")
    logger.info(f"Template path recibido como parámetro: {template_path}")
    
    # VALIDACIÓN CRÍTICA: Asegurar que se use la plantilla correcta para índices
    if template_path:
        template_name = Path(template_path).name
        if "001_INDICE GENERAL_PLANTILLA.docx" not in template_name:
            logger.warning(f"RECHAZANDO plantilla incorrecta: {template_path}")
            logger.warning("No es la plantilla de índice correcta, usando plantilla por defecto")
            template_path = None
    
    # Usar template_path proporcionado (si es correcto) o el valor por defecto
    final_template_path = template_path or TEMPLATE_PATH
    
    logger.info(f"TEMPLATE_PATH por defecto: {TEMPLATE_PATH}")
    logger.info(f"Template final a usar: {final_template_path}")
    logger.info(f"¿Existe la plantilla?: {final_template_path.exists()}")
    
    if not final_template_path.exists():
        logger.error(f"Plantilla no encontrada: {final_template_path}")
        # Buscar plantillas alternativas
        possible_paths = [
            Path(__file__).parent.parent / "word" / "anexos" / "001_INDICE GENERAL_PLANTILLA.docx",
            Path(__file__).parent / "word" / "anexos" / "001_INDICE GENERAL_PLANTILLA.docx",
            Path("Y:/DOCUMENTACION TRABAJO/CARPETAS PERSONAL/SO/github_app/artecoin_automatizaciones/word/anexos/001_INDICE GENERAL_PLANTILLA.docx"),
        ]
        for alt_path in possible_paths:
            logger.info(f"Probando plantilla alternativa: {alt_path}")
            if alt_path.exists():
                logger.info(f"Plantilla alternativa encontrada: {alt_path}")
                final_template_path = alt_path
                break
        else:
            logger.error("No se encontró ninguna plantilla válida")
            return 1
    
    # ---------------------------
    # Recorrer centros y generar índices
    # ---------------------------
    centros = []
    for grp in NAS_GROUPS:
        base = nas_root / grp
        if not base.exists():
            continue
        for child in base.iterdir():
            if not child.is_dir():
                continue
            code = best_code_from_path(child)
            if center_filter and code != center_filter.upper():
                continue
            anejos = find_anejos_dir(child)
            if code:
                centros.append(dict(group=grp, code=code, dir=child, anejos_dir=anejos))

    logger.info(f"Centros detectados: {len(centros)}")

    LOG = []
    for row in centros:
        building_dir: Path = row["dir"]
        anejos_dir: Optional[Path] = row["anejos_dir"]

        # Detectar portada y auditoría en carpeta del centro
        portada = detect_portada(building_dir)
        auditoria = detect_auditoria(building_dir, portada)

        portada_paginas = contar_paginas_pdf(portada) if portada else 1
        auditoria_paginas = contar_paginas_pdf(auditoria) if auditoria else 0

        logger.info(f"    Portada: {portada.name if portada else 'No encontrada'} ({portada_paginas} páginas)")
        logger.info(f"    Auditoría: {auditoria.name if auditoria else 'No encontrada'} ({auditoria_paginas} páginas)")

        # Anejos existentes
        anexos = find_existing_anejos(anejos_dir) if anejos_dir else []
        
        logger.info(f"- {row['code']}: Detectados {len(anexos)} anejos en {anejos_dir}")

        # Si no hay ningún anejo, saltar (opcional)
        if not anexos:
            logger.info(f"- {row['code']}: SKIP (sin anejos válidos en ANEJOS)")
            LOG.append(dict(code=row["code"], status="SKIP", reason="Sin anejos válidos en ANEJOS"))
            continue

        # Renderizar índice
        output_name = OUTPUT_NAME  # "001_INDICE_GENERAL_COMPLETADO.docx"
        out_path = building_dir / output_name
        try:
            render_indice_general_mejorado(final_template_path, out_path, auditoria_paginas, anexos,
                                  offset_memoria=OFFSET_MEMORIA, ancho_linea=ANCHO_LINEA)
            
            # Generar también versión PDF
            pdf_path = convert_docx_to_pdf(out_path)
            if pdf_path:
                logger.info(f"- {row['code']}: ✓ Índice generado (DOCX + PDF)")
            else:
                logger.info(f"- {row['code']}: ✓ Índice generado (DOCX solamente)")
                
            LOG.append(dict(code=row["code"], status="OK", portada=str(portada) if portada else "", 
                           auditoria=str(auditoria) if auditoria else "", salida=str(out_path)))
        except Exception as e:
            logger.error(f"- {row['code']}: ERROR - {e}")
            LOG.append(dict(code=row["code"], status="ERROR", error=str(e)))

    # Resultado
    exitosos = len([r for r in LOG if r["status"] == "OK"])
    logger.info(f"Índices generados exitosamente: {exitosos}/{len(centros)}")
    return 0

def generar_memoria_completa(nas_root: Path, center_filter: str = None) -> int:
    """Generar memoria completa PDF por centro."""
    logger.info("=== GENERANDO MEMORIA COMPLETA ===")
    logger.info(f"NAS Root: {nas_root}")
    
    centros = []
    for grp in NAS_GROUPS:
        base = nas_root / grp
        if not base.exists(): 
            continue
        for child in base.iterdir():
            if not child.is_dir(): 
                continue
            code = best_code_from_path(child)
            if center_filter and code != center_filter.upper():
                continue

            anejos_dir = find_anejos_dir(child)
            portada = detect_portada(child)
            indice = pick_indice(child)
            auditoria = pick_auditoria_completa(child, portada, indice)
            anejos = list_anejospdf(anejos_dir) if anejos_dir else {}

            order_files: List[Path] = []
            missing = []

            if pdf_ok(portada): 
                order_files.append(portada)
            else: 
                missing.append("PORTADA")

            if pdf_ok(indice): 
                order_files.append(indice)
            else: 
                missing.append("INDICE")

            if pdf_ok(auditoria): 
                order_files.append(auditoria)
            else: 
                missing.append("AUDITORIA")

            # Anejos
            for n in sorted(anejos.keys()):
                if pdf_ok(anejos[n]):
                    order_files.append(anejos[n])

            # Generar nombre de archivo con formato: 01_DOCUMENTO 1. AUDITORIA ENERGETICA_{NOMBRE_CARPETA}.pdf
            folder_name = child.name.replace(" ", "_").replace("-", "_")
            output_name = f"01_DOCUMENTO 1. AUDITORIA ENERGETICA_{folder_name}.pdf"
            
            centros.append({
                "group": grp,
                "code": code,
                "center_dir": child,
                "anejos_dir": anejos_dir,
                "missing": missing,
                "anejos_detectados": list(sorted(anejos.keys())),
                "out": child / output_name,
                "total": len(order_files),
                "_files": order_files,
            })

    logger.info(f"Centros para procesar: {len(centros)}")
    
    # Mostrar plan
    for c in centros:
        missing_str = ", ".join(c["missing"]) if c["missing"] else "Ninguno"
        anejos_str = ", ".join(map(str, c["anejos_detectados"])) if c["anejos_detectados"] else "Ninguno"
        logger.info(f"- {c['code']}: {c['total']} archivos | Faltantes: {missing_str} | Anejos: {anejos_str}")

    # Ejecutar generación
    resultados = []
    for c in centros:
        files: List[Path] = c["_files"]
        out_path = Path(c["out"])
        if not files:
            logger.info(f"- {c['code']}: SKIP (sin PDFs válidos)")
            resultados.append({"code": c["code"], "status": "SKIP", "reason": "Sin PDFs válidos"})
            continue
        try:
            merger = PdfMerger(strict=False)
            for f in files:
                merger.append(str(f))
            with open(out_path, "wb") as fp:
                merger.write(fp)
            merger.close()
            logger.info(f"- {c['code']}: ✓ Memoria completa generada ({len(files)} archivos)")
            
            resultados.append({"code": c["code"], "status": "OK", "archivos": len(files), "salida": str(out_path)})
        except Exception as e:
            logger.error(f"- {c['code']}: ERROR - {e}")
            resultados.append({"code": c["code"], "status": "ERROR", "error": str(e)})

    exitosos = len([r for r in resultados if r["status"] == "OK"])
    logger.info(f"Memorias generadas exitosamente: {exitosos}/{len(centros)}")
    return 0

def main():
    """Función principal del script."""
    parser = argparse.ArgumentParser(description="Generar memoria final completa")
    parser.add_argument("--input-dir", required=True, help="Carpeta raíz del NAS (06_REDACCION)")
    parser.add_argument("--output-dir", help="Carpeta de salida (no usado, se genera in-situ)")
    parser.add_argument("--center", help="Centro específico para procesar (ej: C0007)")
    parser.add_argument("--action", choices=["indices", "memoria", "all"], default="all", 
                        help="Acción a realizar: indices, memoria o ambos")
    parser.add_argument("--template-path", help="Ruta a la plantilla del índice general")
    
    args = parser.parse_args()
    
    # Verificar dependencias opcionales
    missing_deps = []
    if not WORD_AVAILABLE:
        missing_deps.append("pywin32 (para conversión DOCX→PDF)")
        
    if missing_deps:
        logger.warning(f"Dependencias opcionales no disponibles: {', '.join(missing_deps)}")
        logger.warning("La funcionalidad estará limitada pero el script seguirá funcionando.")
    
    # Convertir a Path
    nas_root = Path(args.input_dir)
    template_path = Path(args.template_path) if args.template_path else None
    
    if not nas_root.exists():
        logger.error(f"La carpeta de entrada no existe: {nas_root}")
        return 1
    
    center_filter = args.center.upper() if args.center else None
    if center_filter:
        logger.info(f"Filtrando por centro: {center_filter}")
    
    try:
        if args.action in ["indices", "all"]:
            result = generar_indices(nas_root, center_filter, template_path)
            if result != 0:
                return result
        
        if args.action in ["memoria", "all"]:
            result = generar_memoria_completa(nas_root, center_filter)
            if result != 0:
                return result
                
        logger.info("=== PROCESO COMPLETADO ===")
        logger.info("Archivos generados:")
        logger.info("- 001_INDICE_GENERAL.docx + .pdf (si se ejecutó 'indices')")
        logger.info("- 01_DOCUMENTO 1. AUDITORIA ENERGETICA_{NOMBRE_CARPETA}.pdf")
        return 0
        
    except Exception as e:
        logger.error(f"Error durante la ejecución: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())

