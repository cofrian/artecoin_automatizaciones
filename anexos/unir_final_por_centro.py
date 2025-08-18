from __future__ import annotations
from pathlib import Path
import re, os, sys, subprocess, datetime
from typing import List, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox

# ------------------ asegurar pypdf ------------------
def ensure_pypdf():
    try:
        import pypdf  # noqa
    except Exception:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
ensure_pypdf()
from pypdf import PdfReader, PdfWriter
# ----------------------------------------------------

CARP_ANEJOS = "ANEJOS"
SUBS_ESPERADAS = ["01_VARIOS EDIFICIOS", "02_UN EDIFICIO"]
ORDEN_ANEJOS = [1,2,3,4,5,6,7]

# Patrón robusto: captura "ANEXO/ANEJO" + separador + número 1..7, aunque después venga "_" o texto
ANEJO_PAT = re.compile(
    r"(?:ANEXO|ANEJO)\s*[_\-\.\s]*0*([1-7])(?=\D|$)",
    re.IGNORECASE
)

# Heurísticas si no aparece número claro (palabras clave típicas)
HEURISTICAS = {
    1: [r"METODOLOGIA", r"METODOLOGÍA"],
    5: [r"REPORTAJE", r"FOTOGRAF"],  # "Reportaje Fotográfico"
}

def numero_anejo(pdf_name: str) -> int | None:
    """Intenta extraer el número 1..7 del nombre del archivo."""
    m = ANEJO_PAT.search(pdf_name)
    if m:
        try:
            n = int(m.group(1))
            if 1 <= n <= 7:
                return n
        except Exception:
            pass

    # Heurísticas si no hay número claro
    up = pdf_name.upper()
    for num, pats in HEURISTICAS.items():
        for pat in pats:
            if re.search(pat, up):
                return num

    # Último recurso: si aparece "ANEJO/ANEXO" y después un dígito suelto 1..7 (no pegado a otros dígitos)
    if re.search(r"(ANEJO|ANEXO)", up):
        for ch in "1234567":
            if re.search(rf"(?<!\d)0*{ch}(?!\d)", up):
                return int(ch)

    return None

def elegir_memoria_pdf(dir_centro: Path) -> Path | None:
    """PDF en la raíz del centro (no subcarpetas); prioriza '001_*' y, si no, el más grande. Ignora FINAL*.pdf."""
    candidatos = []
    for p in dir_centro.iterdir():
        if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.upper().startswith("FINAL"):
            candidatos.append(p)
    if not candidatos:
        return None
    preferidos = [p for p in candidatos if p.name.startswith("001_")]
    if preferidos:
        return sorted(preferidos)[0]
    return max(candidatos, key=lambda x: x.stat().st_size)

def recoger_anejos_pdf(dir_centro: Path) -> List[tuple[int, Path]]:
    """Todos los PDF dentro de ANEJOS/ (recursivo); devuelve (número, ruta)."""
    out: List[tuple[int, Path]] = []
    raiz = dir_centro / CARP_ANEJOS
    if not raiz.is_dir():
        return out
    for p in raiz.rglob("*.pdf"):
        n = numero_anejo(p.name)
        if n is not None:
            out.append((n, p))
        # DEBUG opcional:
        # else:
        #     print(f"[DEBUG] No detectado nº para: {p.name}")
    return out

def nombre_versionado(dst: Path) -> Path:
    """Si existe dst, devuelve dst__v2, dst__v3, ..."""
    if not dst.exists():
        return dst
    i = 2
    while True:
        alt = dst.with_name(f"{dst.stem}__v{i}{dst.suffix}")
        if not alt.exists():
            return alt
        i += 1

def merge_por_centro(dir_centro: Path) -> str:
    centro = dir_centro.name
    memoria = elegir_memoria_pdf(dir_centro)
    anejos = recoger_anejos_pdf(dir_centro)

    if not memoria and not anejos:
        return f"[{centro}] Sin PDFs: no hay memoria ni anejos."

    # Orden: memoria (si existe) + anejos 1..7 (si faltan, se saltan); si hay varios de un nº, se incluyen todos (por nombre)
    orden: List[Path] = []
    if memoria:
        orden.append(memoria)
    grupos = {n: sorted([p for (k, p) in anejos if k == n], key=lambda x: x.name.lower())
              for n in ORDEN_ANEJOS}
    for n in ORDEN_ANEJOS:
        orden.extend(grupos.get(n, []))

    # Nombre de salida: <memoria>_final.pdf, o FINAL.pdf si no hay memoria
    dst_name = f"{memoria.stem}_final.pdf" if memoria else "FINAL.pdf"
    dst = nombre_versionado(dir_centro / dst_name)

    writer = PdfWriter()
    incluidos, omitidos = [], []

    for pdf in orden:
        try:
            with open(pdf, "rb") as f:
                reader = PdfReader(f)
                if getattr(reader, "is_encrypted", False):
                    try:
                        reader.decrypt("")  # intenta abrir encriptados sin contraseña
                    except Exception:
                        omitidos.append((pdf, "encriptado"))
                        continue
                for page in reader.pages:
                    writer.add_page(page)
            incluidos.append(pdf)
        except Exception as e:
            omitidos.append((pdf, str(e)))

    if not incluidos:
        return f"[{centro}] No se pudo crear {dst.name}: todos fallaron."

    with open(dst, "wb") as f_out:
        writer.write(f_out)

    # Log legible
    lines = [f"[{centro}] -> {dst.name}"]
    lines.append(f"  MEMORIA: {memoria.name if memoria else '(NO ENCONTRADA)'}")
    for n in ORDEN_ANEJOS:
        ps = grupos.get(n, [])
        if ps:
            for p in ps:
                lines.append(f"  ANEJO {n}: {p.relative_to(dir_centro)}")
        else:
            lines.append(f"  ANEJO {n}: (no hay)")
    if omitidos:
        lines.append("  OMITIDOS por error:")
        for p, err in omitidos:
            lines.append(f"    - {p.name}  [{err}]")
    return "\n".join(lines)

def listar_centros(base: Path) -> List[Path]:
    """Si existen las subcarpetas esperadas, las usa; si no, toma TODAS las subcarpetas de base (excepto 'VA')."""
    centros: List[Path] = []
    subdirs = []
    encontrados = False
    for sub in SUBS_ESPERADAS:
        d = base / sub
        if d.is_dir():
            encontrados = True
            subdirs.append(d)
    if not encontrados:
        subdirs = [base]
    for cont in subdirs:
        for x in sorted(cont.iterdir()):
            if x.is_dir() and x.name.upper() != "VA":
                centros.append(x)
    return centros

def seleccionar_base() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    ruta = filedialog.askdirectory(title="Selecciona la carpeta base (ej.: ...\\06_REDACCION)")
    root.update()
    root.destroy()
    if not ruta:
        return None
    return Path(ruta)

def main():
    base = seleccionar_base()
    if base is None:
        messagebox.showinfo("Cancelado", "No seleccionaste carpeta. No se hizo nada.")
        return
    if not base.exists():
        messagebox.showerror("Error", f"La carpeta no existe:\n{base}")
        return

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path.cwd() / f"LOG_FINAL_{stamp}.txt"

    centros = listar_centros(base)
    if not centros:
        messagebox.showwarning("Sin centros", "No se encontraron carpetas de centros bajo la ruta seleccionada.")
        return

    lines = []
    for c in centros:
        try:
            res = merge_por_centro(c)
        except Exception as e:
            res = f"[{c.name}] ERROR: {e}"
        print(res)
        lines.append(res)
        lines.append("")

    log_path.write_text("\n".join(lines), encoding="utf-8")
    messagebox.showinfo(
        "Listo",
        f"Combinación terminada.\n\nCentros procesados: {len(centros)}\n"
        f"Log guardado en:\n{log_path}\n\n"
        "Cada centro tiene ahora <memoria>_final.pdf (o FINAL.pdf si no había memoria)."
    )

if __name__ == "__main__":
    main()
