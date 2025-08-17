from __future__ import annotations
from pathlib import Path
import shutil
import os
import re
import unicodedata
import difflib
import datetime
import sys
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk, filedialog

# ======================
# CONFIGURACIÓN INICIAL
# ======================
# Rutas de origen
RUTA_ANEJO5 = Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\JBA\Anejo5_Colmenar")
RUTA_OTROS_ANEJOS = Path(r"C:\Users\indiva\Downloads\anexos_colmenar")

# Ruta base de memorias (NAS), con subcarpetas: 01_VARIOS EDIFICIOS y 02_UN EDIFICIO
RUTA_MEMORIAS = Path(r"Z:\2025\1-PROYECTOS\10010_25_PR_AUDIT_ED_COLMENAR VIEJO\06_REDACCION")

# Fijos que se añadirán a TODOS los destinos ANEJOS (si existen)
FIJOS = [
    Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\ANEJO_AUDITORIA_FIJOS\01_ANEJO 1. METODOLOGIA_V1.docx"),
    Path(r"Z:\DOCUMENTACION TRABAJO\CARPETAS PERSONAL\SO\ANEJO_AUDITORIA_FIJOS\01_ANEJO 1. METODOLOGIA_V1.pdf"),
]

# Subcarpeta dentro de cada memoria donde se copiará todo
SUBCARPETA_DESTINO = "ANEJOS"

# Umbral de similitud para emparejar nombres de “OTROS ANEJOS”
UMBRAL_SIMILITUD = 0.74

# ======================
# FUNCIONES AUXILIARES
# ======================

# Acepta C0001 o CC0001 (con separadores opcionales)
CODIGO_RE = re.compile(r"C{1,2}\s*[-_ ]?\s*(\d{4})", re.IGNORECASE)

def normalizar_texto(s: str) -> str:
    if not s:
        return ""
    s = s.replace("_", " ")
    s = s.replace("“", '"').replace("”", '"').replace("’", "'").replace("´", "'")
    s = (unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii"))
    s = re.sub(r"[^A-Za-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.upper()

def extraer_codigo_y_nombre(dir_memoria: Path):
    """
    De '01_C0001_COLEGIO ...' o '01_CC0006_...' -> (codigo_norm, nombre_norm, codigo_bruto)
    - codigo_norm = 'C0001' (normaliza doble C a una C)
    - codigo_bruto = lo que se detectó literalmente (p.ej. 'CC0006')
    """
    nombre = dir_memoria.name
    m = CODIGO_RE.search(nombre)
    if not m:
        return None, normalizar_texto(nombre), None
    start, end = m.span()
    cuatro = m.group(1)
    codigo_bruto = nombre[start:end].upper()
    codigo_norm = f"C{cuatro}"
    resto = nombre[end:].lstrip(" _-")
    return codigo_norm, normalizar_texto(resto), codigo_bruto

def _posibles_carpetas_codigo(codigo_norm: str):
    num = codigo_norm[1:] if codigo_norm and codigo_norm[0].upper() == "C" else codigo_norm
    return [RUTA_ANEJO5 / f"C{num}", RUTA_ANEJO5 / f"CC{num}"]

def buscar_anejo5_por_codigo(codigo_norm: str):
    """Busca un archivo con 'ANEJO' en la carpeta del código (tolera C#### o CC####)."""
    for carpeta in _posibles_carpetas_codigo(codigo_norm):
        if carpeta.is_dir():
            candidatos = sorted(
                [p for p in carpeta.glob("*") if p.is_file() and "ANEJO" in p.name.upper()],
                key=lambda p: ("ANEJO5" not in p.name.upper(), len(p.name)),
            )
            if candidatos:
                return candidatos[0]
    return None

def recolectar_memorias(ruta_memorias: Path):
    if not ruta_memorias.is_dir():
        return []
    candidatos = []
    for sub in ["01_VARIOS EDIFICIOS", "02_UN EDIFICIO"]:
        d = ruta_memorias / sub
        if d.is_dir():
            for x in d.iterdir():
                if x.is_dir():
                    candidatos.append(x)
    return candidatos

def recolectar_carpetas_otros_anejos(ruta: Path):
    out = {}
    if ruta.is_dir():
        for p in ruta.iterdir():
            if p.is_dir():
                out[normalizar_texto(p.name)] = p
    return out

def mapear_otros_anejos_por_nombre(nombre_centro_norm: str, carpetas_fuente):
    if not carpetas_fuente:
        return None
    nombres = list(carpetas_fuente.keys())
    # 1) Coincidencia directa por substring
    for n in nombres:
        if n in nombre_centro_norm or nombre_centro_norm in n:
            return carpetas_fuente[n]
    # 2) Aproximación por similitud
    mejor = None
    mejor_score = 0.0
    for n in nombres:
        score = difflib.SequenceMatcher(None, n, nombre_centro_norm).ratio()
        if score > mejor_score:
            mejor = n
            mejor_score = score
    if mejor and mejor_score >= UMBRAL_SIMILITUD:
        return carpetas_fuente[mejor]
    return None

def copiar_archivo_origen_destino(src: Path, dst: Path):
    """Copia preservando metadata; crea nombre __v2, __v3 si ya existe."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    objetivo = dst
    i = 2
    while objetivo.exists():
        objetivo = dst.with_name(f"{dst.stem}__v{i}{dst.suffix}")
        i += 1
    shutil.copy2(src, objetivo)
    return objetivo

def copiar_carpeta(src_dir: Path, dst_dir: Path, log_lines):
    archivos = 0
    bytes_total = 0
    for root, _, files in os.walk(src_dir):
        root_path = Path(root)
        rel = root_path.relative_to(src_dir)
        for f in files:
            s = root_path / f
            d = dst_dir / rel / f
            try:
                final = copiar_archivo_origen_destino(s, d)
                archivos += 1
                try:
                    bytes_total += s.stat().st_size
                except Exception:
                    pass
                log_lines.append(f"COPIADO: {s} -> {final}")
            except Exception as e:
                log_lines.append(f"[ERROR] Al copiar {s} -> {d}: {e}")
    return archivos, bytes_total

def bytes_a_hum(n: int) -> str:
    for unit in ["B","KB","MB","GB","TB"]:
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"

# ======================
# PLAN Y EJECUCIÓN
# ======================

def construir_plan():
    tareas = []
    warnings = []
    carpetas_otros = recolectar_carpetas_otros_anejos(RUTA_OTROS_ANEJOS)
    memorias = recolectar_memorias(RUTA_MEMORIAS)

    for dir_mem in sorted(memorias, key=lambda p: p.name):
        codigo_norm, nombre_norm, codigo_bruto = extraer_codigo_y_nombre(dir_mem)
        if not codigo_norm:
            warnings.append(f"[SIN CÓDIGO] {dir_mem}")
            continue

        anejo5 = buscar_anejo5_por_codigo(codigo_norm)
        carpeta_otros = mapear_otros_anejos_por_nombre(nombre_norm, carpetas_otros)

        otros_conteo = 0
        if carpeta_otros:
            for _, _, files in os.walk(carpeta_otros):
                otros_conteo += len(files)

        estado = "OK"
        if anejo5 is None or carpeta_otros is None:
            estado = "FALTAN"

        tareas.append({
            "codigo_norm": codigo_norm,
            "codigo_bruto": codigo_bruto,
            "destino": dir_mem,
            "dest_sub": dir_mem / SUBCARPETA_DESTINO,
            "nombre_norm": nombre_norm,
            "anejo5": anejo5,
            "carpeta_otros": carpeta_otros,
            "otros_conteo": otros_conteo,
            "incluir": True,
            "estado": estado,
        })

        if anejo5 is None:
            warnings.append(f"[ANEJO5 NO ENCONTRADO] código {codigo_norm} → {dir_mem}")
        if carpeta_otros is None:
            warnings.append(f"[OTROS ANEJOS NO ENCONTRADOS] '{nombre_norm}' → {dir_mem}")

    # Validación de fijos
    for f in FIJOS:
        if not f.exists():
            warnings.append(f"[FIJO NO ENCONTRADO] {f}")

    return tareas, warnings

def ejecutar_plan(tareas):
    log_lines = []
    tot_archivos = 0
    tot_bytes = 0

    for t in tareas:
        if not t.get("incluir", True):
            log_lines.append(f"[OMITIDO] {t['codigo_norm']} → {t['destino']}")
            continue

        destino = Path(t["dest_sub"])

        # 1) Anejo5
        if t["anejo5"] and Path(t["anejo5"]).exists():
            try:
                final = copiar_archivo_origen_destino(Path(t["anejo5"]), destino / Path(t["anejo5"]).name)
                log_lines.append(f"COPIADO: {t['anejo5']} -> {final}")
                tot_archivos += 1
                try:
                    tot_bytes += Path(t["anejo5"]).stat().st_size
                except Exception:
                    pass
            except Exception as e:
                log_lines.append(f"[ERROR] Al copiar Anejo5 {t['anejo5']}: {e}")

        # 2) OTROS
        if t["carpeta_otros"] and Path(t["carpeta_otros"]).is_dir():
            copiados, bytes_copiados = copiar_carpeta(Path(t["carpeta_otros"]), destino, log_lines)
            tot_archivos += copiados
            tot_bytes += bytes_copiados

        # 3) FIJOS
        for fijo in FIJOS:
            try:
                if Path(fijo).exists():
                    final = copiar_archivo_origen_destino(Path(fijo), destino / Path(fijo).name)
                    log_lines.append(f"COPIADO FIJO: {fijo} -> {final}")
                    tot_archivos += 1
                    try:
                        tot_bytes += Path(fijo).stat().st_size
                    except Exception:
                        pass
                else:
                    log_lines.append(f"[OMITIDO FIJO – NO EXISTE] {fijo}")
            except Exception as e:
                log_lines.append(f"[ERROR FIJO] {fijo}: {e}")

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path.cwd() / f"LOG_MOVIMIENTO_ANEJOS_{stamp}.txt"
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    resumen = f"Archivos copiados: {tot_archivos} | Tamaño total: {bytes_a_hum(tot_bytes)}\nLog: {log_path}"
    return resumen

# ======================
# UI (REVISIÓN POR ELEMENTO)
# ======================

class App(tk.Tk):
    def __init__(self, tareas, warnings):
        super().__init__()
        self.title("Anejos → Memorias (revisión por elemento)")
        self.geometry("1400x820")
        self.tareas = tareas
        self.warnings = warnings
        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="both", expand=True)

        lbl = ttk.Label(top, text="Selecciona qué relaciones ejecutar. Puedes Editar rutas (archivo Anejo5 y/o carpeta de otros anejos), Omitir o Reincluir cada fila. ENTER = Ejecutar | ESC = Cancelar.", anchor="w")
        lbl.pack(fill="x", pady=(0,6))

        cols = ("codigo", "memoria", "carpeta_a5", "archivo_a5", "carpeta_otros", "archivos", "estado")
        self.tree = ttk.Treeview(top, columns=cols, show="headings", height=22)
        widths = (90, 380, 110, 320, 320, 80, 90)
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c.upper())
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_details())

        self._reload_rows()

        btns = ttk.Frame(top)
        btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="✏️ Editar selección…", command=self._editar_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="🚫 Omitir", command=self._omitir_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="✅ Reincluir", command=self._reincluir_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="📂 Abrir carpeta destino", command=self._abrir_destino_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="📂 Abrir carpeta A5", command=self._abrir_a5_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="📂 Abrir carpeta OTROS", command=self._abrir_otros_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="👁️ Ver archivos OTROS…", command=self._ver_archivos_otros).pack(side="left", padx=4)

        right = ttk.Frame(btns)
        right.pack(side="right")
        ttk.Button(right, text="✅ Ejecutar (ENTER)", command=self._ok).pack(side="right", padx=6)
        ttk.Button(right, text="❌ Cancelar (ESC)", command=self._cancel).pack(side="right", padx=6)

        if self.warnings:
            box = scrolledtext.ScrolledText(top, height=6, wrap="word")
            box.pack(fill="x")
            box.insert("1.0", "\n".join(["AVISOS:"] + self.warnings + ["", "Se copiarán además los FIJOS a todos los destinos si existen:"] + [str(f) for f in FIJOS]))
            box.configure(state="disabled")

        self.details = scrolledtext.ScrolledText(top, height=6, wrap="word")
        self.details.pack(fill="both", expand=False, pady=(6,0))
        self._update_details()

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())

    def _row_values(self, t):
        carpeta_a5 = Path(t["anejo5"]).parent.name if t["anejo5"] else "—"
        archivo_a5 = Path(t["anejo5"]).name if t["anejo5"] else "—"
        carpeta_otros = Path(t["carpeta_otros"]).name if t["carpeta_otros"] else "—"
        estado = "OMITIDO" if not t.get("incluir", True) else t.get("estado", "OK")
        cod_display = t["codigo_norm"] if not t["codigo_bruto"] or t["codigo_bruto"] == t["codigo_norm"] else f"{t['codigo_bruto']} → {t['codigo_norm']}"
        return (cod_display, t["destino"].name, carpeta_a5, archivo_a5, carpeta_otros, str(t["otros_conteo"]), estado)

    def _reload_rows(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, t in enumerate(self.tareas):
            self.tree.insert("", "end", iid=str(idx), values=self._row_values(t))

    def _selected_idx(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _update_details(self):
        idx = self._selected_idx()
        self.details.configure(state="normal")
        self.details.delete("1.0", "end")
        if idx is None:
            self.details.insert("1.0", "Selecciona una fila para ver detalles de rutas.")
        else:
            t = self.tareas[idx]
            destino = t["dest_sub"]
            a5 = t["anejo5"]
            otros = t["carpeta_otros"]
            self.details.insert("1.0",
                f"Destino (subcarpeta ANEJOS): {destino}\n"
                f"Anejo5 (archivo): {a5 if a5 else '—'}\n"
                f"Carpeta OTROS: {otros if otros else '—'}\n"
                f"Fijos que se añadirán: {', '.join([f.name for f in FIJOS])}\n"
            )
        self.details.configure(state="disabled")

    def _ver_archivos_otros(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        t = self.tareas[idx]
        carpeta = t["carpeta_otros"]
        if not carpeta or not Path(carpeta).is_dir():
            messagebox.showinfo("Sin carpeta", "No hay carpeta OTROS seleccionada para esta fila.")
            return
        win = tk.Toplevel(self)
        win.title(f"Archivos en OTROS – {Path(carpeta).name}")
        win.geometry("700x500")
        box = scrolledtext.ScrolledText(win, wrap="word")
        box.pack(fill="both", expand=True)
        lines = []
        for root, _, files in os.walk(carpeta):
            root_path = Path(root)
            rel = root_path.relative_to(carpeta)
            for f in sorted(files):
                lines.append(str(rel / f))
        box.insert("1.0", "\n".join(lines if lines else ["(vacío)"]))
        box.configure(state="disabled")

    def _editar_sel(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        t = self.tareas[idx]

        win = tk.Toplevel(self)
        win.title(f"Editar selección – {t['destino'].name}")
        win.geometry("900x280")
        frm = ttk.Frame(win, padding=10)
        frm.pack(fill="both", expand=True)

        # Anejo5 (archivo)
        ttk.Label(frm, text="Archivo Anejo5:").grid(row=0, column=0, sticky="w")
        var_a = tk.StringVar(value=str(t["anejo5"] or ""))
        ent_a = ttk.Entry(frm, textvariable=var_a, width=100)
        ent_a.grid(row=0, column=1, sticky="we", padx=6)
        def pick_a():
            start = None
            for c in _posibles_carpetas_codigo(t["codigo_norm"]):
                if c.exists():
                    start = c
                    break
            f = filedialog.askopenfilename(initialdir=str(start or RUTA_ANEJO5), title="Elige archivo Anejo5")
            if f:
                var_a.set(f)
        ttk.Button(frm, text="Elegir…", command=pick_a).grid(row=0, column=2, padx=4)

        # Otros anejos (carpeta)
        ttk.Label(frm, text="Carpeta otros anejos:").grid(row=1, column=0, sticky="w", pady=(8,0))
        var_o = tk.StringVar(value=str(t["carpeta_otros"] or ""))
        ent_o = ttk.Entry(frm, textvariable=var_o, width=100)
        ent_o.grid(row=1, column=1, sticky="we", padx=6, pady=(8,0))
        def pick_o():
            d = filedialog.askdirectory(initialdir=str(RUTA_OTROS_ANEJOS), title="Elige carpeta de otros anejos")
            if d:
                var_o.set(d)
        ttk.Button(frm, text="Elegir…", command=pick_o).grid(row=1, column=2, padx=4, pady=(8,0))

        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=10)
        def save():
            a = var_a.get().strip()
            o = var_o.get().strip()
            t["anejo5"] = Path(a) if a else None
            t["carpeta_otros"] = Path(o) if o else None
            t["estado"] = "OK" if (t["anejo5"] or t["carpeta_otros"]) else "FALTAN"
            t["incluir"] = True
            win.destroy()
            self._reload_rows()
            self._update_details()
        ttk.Button(btns, text="Guardar", command=save).pack(side="right", padx=6)
        ttk.Button(btns, text="Cancelar", command=win.destroy).pack(side="right")

    def _omitir_sel(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        self.tareas[idx]["incluir"] = False
        self.tareas[idx]["estado"] = "OMITIDO"
        self._reload_rows()
        self._update_details()

    def _reincluir_sel(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        self.tareas[idx]["incluir"] = True
        if self.tareas[idx]["anejo5"] is None and self.tareas[idx]["carpeta_otros"] is None:
            self.tareas[idx]["estado"] = "FALTAN"
        else:
            self.tareas[idx]["estado"] = "OK"
        self._reload_rows()
        self._update_details()

    def _abrir_destino_sel(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        dest = self.tareas[idx]["dest_sub"]
        try:
            os.startfile(dest)  # Windows
        except Exception:
            messagebox.showinfo("Ruta", str(dest))

    def _abrir_a5_sel(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        a5 = self.tareas[idx]["anejo5"]
        if not a5:
            messagebox.showinfo("Anejo5", "No hay archivo Anejo5 para esta fila.")
            return
        folder = Path(a5).parent
        try:
            os.startfile(folder)
        except Exception:
            messagebox.showinfo("Ruta", str(folder))

    def _abrir_otros_sel(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Selecciona una fila", "Marca una fila primero.")
            return
        otros = self.tareas[idx]["carpeta_otros"]
        if not otros:
            messagebox.showinfo("OTROS", "No hay carpeta de OTROS para esta fila.")
            return
        try:
            os.startfile(otros)
        except Exception:
            messagebox.showinfo("Ruta", str(otros))

    def _ok(self):
        self.result = True
        self.destroy()
    def _cancel(self):
        self.result = False
        self.destroy()

def mostrar_ui_y_confirmar(tareas, warnings):
    app = App(tareas, warnings)
    app.mainloop()
    return getattr(app, "result", False), tareas

def main():
    if not RUTA_MEMORIAS.exists():
        messagebox.showerror("Error", f"No existe la ruta de memorias:\n{RUTA_MEMORIAS}")
        sys.exit(1)
    if not RUTA_ANEJO5.exists():
        messagebox.showwarning("Aviso", f"No existe la ruta de Anejo5:\n{RUTA_ANEJO5}\nSe continuará sin Anejo5.")
    if not RUTA_OTROS_ANEJOS.exists():
        messagebox.showwarning("Aviso", f"No existe la ruta de Otros Anejos:\n{RUTA_OTROS_ANEJOS}\nSe continuará sin esos anexos.")

    tareas, warnings = construir_plan()
    if not tareas:
        messagebox.showerror("Sin tareas", "No se encontraron carpetas de memorias en la ruta indicada.")
        sys.exit(1)

    ok, tareas_editadas = mostrar_ui_y_confirmar(tareas, warnings)
    if ok:
        resumen = ejecutar_plan(tareas_editadas)
        messagebox.showinfo("Listo", f"Copias realizadas.\n\n{resumen}")
    else:
        messagebox.showinfo("Cancelado", "Operación cancelada. No se ha copiado nada.")

if __name__ == "__main__":
    main()
