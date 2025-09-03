# -*- coding: utf-8 -*-
"""
GUI (customtkinter) para ejecutar 'anexos_creator.py' o 'anexos_creator.exe' sin usar la terminal.

NOVEDADES (esta versión):
- Pestaña "Mover anexos al NAS": permite mover los DOCX/PDF locales organizados por centro (Cxxxx)
  a la carpeta de centros del NAS (las típicas 01_Cxxxx_* o 02_Cxxxx_*), bajo la subcarpeta "ANEJOS".
- Validación básica de rutas y opción "Simular (dry-run)".
- Filtro opcional de centros (rango/lista: C0001-C0010, C0003, C0011).

Se mantiene:
- Pestaña "Generación de Anejos" (creación de anexos 2..7).
- Logs en tiempo real, cancelación segura y persistencia de configuración.

Arquitectura (SRP / SOLID):
- RunOptions / MoveOptions (dataclasses) modelan las opciones.
- Validator valida entradas (sin efectos colaterales).
- ConfigStore persiste/restaura ajustes del usuario.
- ProcessRunner aísla la ejecución de 'anexos_creator' (stdout/err + stdin).
- AnexosApp orquesta la UI.
"""

from __future__ import annotations

import os
import sys
import json
import time
import queue
import threading
import subprocess
import ctypes
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List

import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar, IntVar

# ---------------------------
# Configuración / Constantes
# ---------------------------

APP_NAME = "Anexos GUI"
CONFIG_FILENAME = os.path.join(os.path.expanduser("~"), ".anexos_gui_config.json")

def set_uniform_scaling() -> None:
    """
    Fuerza un escalado uniforme para que la UI tenga el mismo tamaño en todas las pantallas.
    - En Windows: fija la app como 'System DPI Aware' (constante por sesión) y
      evita el reescalado por monitor.
    """
    try:
        if sys.platform.startswith("win"):
            try:
                # 1 = PROCESS_SYSTEM_DPI_AWARE (uniforme en todos los monitores)
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                # Fallback para versiones antiguas
                ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# ---------------------------
# Modelo de dominio
# ---------------------------

@dataclass(frozen=True)
class RunOptions:
    """Opciones de ejecución para 'anexos_creator' (acción: generate)."""
    # Rutas base
    excel_dir: str
    word_dir: Optional[str]
    output_dir: Optional[str]
    html_templates_dir: Optional[str]
    photos_dir: Optional[str]
    cee_dir: Optional[str]
    plans_dir: Optional[str]

    # Selección de anexos (rango/lista tipo "1-3, 6")
    anexos_expr: Optional[str]

    # Parámetros fecha
    month: Optional[int]
    year: Optional[int]

    # Expresión de centros (rango/lista)
    center: Optional[str] = None
    
    # Photo filtering for Anejo 5
    include_without_photos: bool = True



@dataclass(frozen=True)
class MoveOptions:
    """Opciones de movimiento de anexos al NAS (acción: move-to-nas)."""
    local_out_root: str          # Carpeta local con subcarpetas Cxxxx (o Cxxxx/anexos)
    nas_centers_dir: str         # Carpeta del NAS con 01_Cxxxx_* / 02_Cxxxx_*
    centers_expr: Optional[str]  # Filtro opcional de centros
    dry_run: bool                # True: solo simula (no mueve)
    word_dir: Optional[str]      # Carpeta de plantillas Word (para las plantillas del Anejo 1)


# ---------------------------
# Persistencia de configuración
# ---------------------------

class ConfigStore:
    """Lee/Escribe configuración simple en JSON (no crítica)."""

    @staticmethod
    def load() -> dict:
        try:
            if os.path.isfile(CONFIG_FILENAME):
                with open(CONFIG_FILENAME, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    @staticmethod
    def save(data: dict) -> None:
        try:
            with open(CONFIG_FILENAME, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            # no bloquear la UX si falla
            pass


# ---------------------------
# Validación de entradas
# ---------------------------

class Validator:
    """Valida coherencia de opciones antes de ejecutar (sin efectos colaterales)."""

    @staticmethod
    def _parse_anexos_expr(expr: Optional[str]) -> Optional[list[int]]:
        if not expr or not str(expr).strip():
            return None
        raw = str(expr).strip()
        out: set[int] = set()
        parts = [p.strip() for p in raw.split(",")] if "," in raw else [raw]
        for part in parts:
            if not part:
                continue
            if "-" in part:
                a, b = [x.strip() for x in part.split("-", 1)]
                if a.isdigit() and b.isdigit():
                    ai, bi = int(a), int(b)
                    if ai > bi:
                        ai, bi = bi, ai
                    for n in range(ai, bi + 1):
                        out.add(n)
            else:
                if part.isdigit():
                    out.add(int(part))
        # Limitar a rango válido 1..7
        out = {n for n in out if 1 <= n <= 7}
        return sorted(out) or None

    @staticmethod
    def validate_options(opts: RunOptions) -> Tuple[bool, str]:
        # Excel es obligatorio
        if not opts.excel_dir or not os.path.isdir(opts.excel_dir):
            return False, "Debes seleccionar una carpeta de Excel válida."

        excel_exts = (".xlsx", ".xlsm", ".xls")
        try:
            has_excel = any(f.lower().endswith(excel_exts) for f in os.listdir(opts.excel_dir))
        except Exception:
            has_excel = False
        if not has_excel:
            return False, "La carpeta de Excel no contiene archivos .xls/.xlsx/.xlsm."

        # Mes/año
        if opts.month is not None:
            try:
                m = int(opts.month)
            except Exception:
                return False, "Mes inválido."
            if m < 1 or m > 12:
                return False, "El mes debe estar entre 1 y 12."
        if opts.year is not None:
            try:
                y = int(opts.year)
            except Exception:
                return False, "Año inválido."
            if y < 2000 or y > datetime.now().year + 5:
                return False, "Año fuera de rango razonable."

        # Rutas opcionales si vienen informadas
        if opts.word_dir and not os.path.isdir(opts.word_dir):
            return False, "La carpeta de Word indicada no existe."
        if opts.html_templates_dir and not os.path.isdir(opts.html_templates_dir):
            return False, "La carpeta de plantillas HTML indicada no existe."
        if opts.photos_dir and not os.path.isdir(opts.photos_dir):
            return False, "La carpeta de fotos indicada no existe."
        if opts.cee_dir and not os.path.isdir(opts.cee_dir):
            return False, "La carpeta de CEE indicada no existe."
        if opts.plans_dir and not os.path.isdir(opts.plans_dir):
            return False, "La carpeta de planos indicada no existe."

        # Validar expresión de anexos si el usuario restringe
        if opts.anexos_expr:
            parsed = Validator._parse_anexos_expr(opts.anexos_expr)
            if not parsed:
                return False, "Expresión de anexos inválida. Ejemplos: 1-3, 6   o   2,3,4"
        return True, ""

    @staticmethod
    def validate_move_options(opts: MoveOptions) -> Tuple[bool, str]:
        if not opts.local_out_root or not os.path.isdir(opts.local_out_root):
            return False, "Selecciona una carpeta local (raíz de anexos por centro) válida."
        if not opts.nas_centers_dir or not os.path.isdir(opts.nas_centers_dir):
            return False, "Selecciona una carpeta del NAS válida (la que contiene 01_Cxxxx_* / 02_Cxxxx_*)."
        return True, ""


# ---------------------------
# Ejecutor del proceso (subprocess)
# ---------------------------

class ProcessRunner:
    """
    Encapsula la ejecución del binario/script de anexos, capturando logs
    y permitiendo cancelación y entrada por stdin.
    """

    def __init__(self, log_queue: "queue.Queue[str]"):
        self._proc: Optional[subprocess.Popen] = None
        self._log_queue = log_queue
        self._reader_thread: Optional[threading.Thread] = None

    def is_running(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    def _find_target(self) -> Tuple[List[str], bool]:
        """
        Busca el ejecutable/script de 'anexos_creator' en orden:
          1) junto al ejecutable (ruta congelada) -> anexos_creator.exe
          2) junto al ejecutable (ruta congelada) -> anexos_creator.py (usa python actual -u)
          3) en el CWD actual -> .exe
          4) en el CWD actual -> .py (usa python actual -u)

        Devuelve (cmd_base, es_exe).
        """
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.argv[0]))
        candidates = [
            (os.path.join(base_dir, "anexos_creator.exe"), True),
            (os.path.join(base_dir, "anexos_creator.py"), False),
            (os.path.abspath("anexos_creator.exe"), True),
            (os.path.abspath("anexos_creator.py"), False),
        ]
        for path, is_exe in candidates:
            if os.path.isfile(path):
                return ([path] if is_exe else [sys.executable, "-u", path], is_exe)
        raise FileNotFoundError(
            "No se encontró 'anexos_creator.exe' ni 'anexos_creator.py'. "
            "Colócalo en la misma carpeta que esta aplicación."
        )

    def _launch(self, cmd: List[str]) -> None:
        # Log del comando para depuración
        try:
            self._log_queue.put("CMD: " + " ".join(cmd))
        except Exception:
            pass

        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        # Forzar UTF-8
        env_utf8 = dict(os.environ)
        env_utf8.setdefault("PYTHONIOENCODING", "utf-8")

        self._proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=0,  # Sin buffering para output inmediato
            universal_newlines=True,
            creationflags=creationflags,
            env=env_utf8,
        )

        def _reader() -> None:
            assert self._proc is not None
            try:
                while True:
                    # Leer línea por línea para mostrar output inmediato
                    line = self._proc.stdout.readline()
                    if not line and self._proc.poll() is not None:
                        break
                    if line:
                        self._log_queue.put(line.rstrip("\r\n"))
                        # Forzar flush inmediato
                        try:
                            self._proc.stdout.flush()
                        except:
                            pass
            except Exception as e:
                self._log_queue.put(f"[ERROR] Lectura del proceso: {e}")
            finally:
                self._proc.wait()
                self._log_queue.put(f"--- Proceso finalizado. Código: {self._proc.returncode} ---")

        self._reader_thread = threading.Thread(target=_reader, daemon=True)
        self._reader_thread.start()

    # Acción: GENERATE
    def start(self, opts: RunOptions) -> None:
        if self.is_running():
            raise RuntimeError("Ya hay un proceso en ejecución.")
        base_cmd, _ = self._find_target()

        cmd: list[str] = list(base_cmd)
        cmd += ["--action", "generate"]
        cmd += ["--excel-dir", opts.excel_dir]

        if opts.word_dir:
            cmd += ["--word-dir", opts.word_dir]

        # pasar rango/lista de anexos si el usuario restringe
        if opts.anexos_expr:
            cmd += ["--anexos", opts.anexos_expr]

        if opts.html_templates_dir:
            cmd += ["--html-templates-dir", opts.html_templates_dir]
        if opts.photos_dir:
            cmd += ["--photos-dir", opts.photos_dir]
        if opts.output_dir:
            cmd += ["--output-dir", opts.output_dir]
        if opts.cee_dir:
            cmd += ["--cee-dir", opts.cee_dir]
        if opts.plans_dir:
            cmd += ["--plans-dir", opts.plans_dir]

        if opts.month is not None:
            cmd += ["--month", str(opts.month)]
        if opts.year is not None:
            cmd += ["--year", str(opts.year)]

        if getattr(opts, "center", None):
            cmd += ["--center", str(opts.center)]
            
        # Add photo filtering parameter for Anejo 5
        if not opts.include_without_photos:
            cmd += ["--exclude-without-photos"]

        self._launch(cmd)


    # Acción: MOVE-TO-NAS
    def start_move(self, opts: MoveOptions) -> None:
        if self.is_running():
            raise RuntimeError("Ya hay un proceso en ejecución.")
        base_cmd, _ = self._find_target()

        cmd: list[str] = list(base_cmd)
        cmd += ["--action", "move-to-nas"]
        cmd += ["--local-out-root", opts.local_out_root]
        cmd += ["--nas-centers-dir", opts.nas_centers_dir]
        if opts.word_dir:
            cmd += ["--word-dir", opts.word_dir]
        if opts.centers_expr:
            cmd += ["--centers", opts.centers_expr]
        if opts.dry_run:
            cmd += ["--dry-run"]

        self._launch(cmd)

    def send_input(self, text: str) -> bool:
        """Envía una línea al stdin del proceso (para manejar input())."""
        if not self.is_running() or self._proc is None or self._proc.stdin is None:
            return False
        try:
            self._proc.stdin.write(text + "\n")
            self._proc.stdin.flush()
            return True
        except Exception:
            return False

    def stop(self) -> None:
        if not self.is_running():
            return
        try:
            self._proc.terminate()  # type: ignore[union-attr]
        except Exception:
            pass
        for _ in range(30):
            if not self.is_running():
                break
            time.sleep(0.1)
        if self.is_running():
            try:
                self._proc.kill()  # type: ignore[union-attr]
            except Exception:
                pass
        self._log_queue.put("--- Proceso detenido por el usuario ---")

    def run_async(self, cmd: List[str], callback_line=None, callback_finished=None) -> None:
        """
        Ejecuta un comando arbitrario de forma asíncrona.
        
        Args:
            cmd: Lista con el comando y argumentos
            callback_line: Función que se llama por cada línea de salida
            callback_finished: Función que se llama cuando termina el proceso
        """
        if self.is_running():
            raise RuntimeError("Ya hay un proceso en ejecución.")
        
        # Log del comando para depuración
        try:
            self._log_queue.put("CMD: " + " ".join(cmd))
        except Exception:
            pass

        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        # Forzar UTF-8
        env_utf8 = dict(os.environ)
        env_utf8.setdefault("PYTHONIOENCODING", "utf-8")

        self._proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=0,  # Sin buffering para output inmediato
            universal_newlines=True,
            creationflags=creationflags,
            env=env_utf8,
        )

        def _reader() -> None:
            assert self._proc is not None
            try:
                while True:
                    # Leer línea por línea para mostrar output inmediato
                    line = self._proc.stdout.readline()
                    if not line and self._proc.poll() is not None:
                        break
                    if line:
                        line_clean = line.rstrip("\r\n")
                        if callback_line:
                            callback_line(line_clean)
                        else:
                            self._log_queue.put(line_clean)
                        # Forzar flush inmediato
                        try:
                            self._proc.stdout.flush()
                        except:
                            pass
            except Exception as e:
                error_msg = f"[ERROR] Lectura del proceso: {e}"
                if callback_line:
                    callback_line(error_msg)
                else:
                    self._log_queue.put(error_msg)
            finally:
                self._proc.wait()
                finish_msg = f"--- Proceso finalizado. Código: {self._proc.returncode} ---"
                if callback_line:
                    callback_line(finish_msg)
                else:
                    self._log_queue.put(finish_msg)
                
                # Llamar callback de finalización
                if callback_finished:
                    callback_finished(self._proc.returncode)

        self._reader_thread = threading.Thread(target=_reader, daemon=True)
        self._reader_thread.start()


# ---------------------------
# Vista / Controlador (customtkinter)
# ---------------------------

class AnexosApp(ctk.CTk):

    def __init__(self) -> None:
        super().__init__()
        # Apariencia
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        ctk.set_widget_scaling(1.15)
        ctk.set_window_scaling(1.05)

        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)
        try:
            # Escalado de fuentes Tk (evita variaciones por DPI del monitor)
            self.tk.call("tk", "scaling", 1.0)
        except Exception:
            pass
        
        self.title(APP_NAME)
        self.geometry("1800x1200")
        self.minsize(1600, 1000)

        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.runner = ProcessRunner(self.log_queue)

        # 1) IMPORTANTE: construir el estado ANTES de construir pestañas
        self._build_state()

        # 2) Tabs
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_generacion = self.tabs.add("Generación de Anejos")
        self.tab_move = self.tabs.add("Mover anexos al NAS")
        self.tab_memoria = self.tabs.add("Memoria Final")
        self.tab_terminal = self.tabs.add("Historial de ejecución")

        # 3) Vistas
        self._build_terminal_tab()
        self._build_generacion_tab()
        self._build_move_tab()
        self._build_memoria_tab()

        # 4) Config y ciclo de logs
        self._apply_saved_config()
        self._poll_logs()
        self.protocol("WM_DELETE_WINDOW", self._on_close)


    # ---------- pestaña terminal ----------
    def _build_terminal_tab(self):
        frame = ctk.CTkFrame(self.tab_terminal)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        label = ctk.CTkLabel(frame, text="Terminal de salida de procesos", font=ctk.CTkFont(size=20, weight="bold"))
        label.pack(anchor="w", pady=(0, 8))
        self.txt_terminal = ctk.CTkTextbox(frame, height=500, font=ctk.CTkFont(size=16))
        self.txt_terminal.pack(fill="both", expand=True, padx=4, pady=(0, 8))

    def _log_terminal(self, text: str) -> None:
        self.txt_terminal.configure(state="normal")
        self.txt_terminal.insert("end", text + "\n")
        self.txt_terminal.see("end")
        self.txt_terminal.configure(state="disabled")

    # ---------- pestaña generar ----------
    def _build_generacion_tab(self):
        # Fuentes (más grandes)
        pad = {"padx": 20, "pady": 18}
        font_title = ctk.CTkFont(size=28, weight="bold")
        font_label = ctk.CTkFont(size=18)
        font_btn = ctk.CTkFont(size=18)
        font_hint = ctk.CTkFont(size=18)
        font_entry = ctk.CTkFont(size=17)

        # === CONTENEDOR PRINCIPAL (ocupa TODO) ===
        main = ctk.CTkFrame(self.tab_generacion)
        main.pack(fill="both", expand=True, padx=12, pady=12)
        # Grid del contenedor principal
        main.grid_rowconfigure(0, weight=0)  # Título
        main.grid_rowconfigure(1, weight=1)  # Paths
        main.grid_rowconfigure(2, weight=1)  # Opciones
        main.grid_rowconfigure(3, weight=0)  # Botonera
        main.grid_columnconfigure(0, weight=1)

        # --- Título ---
        ctk.CTkLabel(main, text="Interfaz de Anejos", font=font_title)\
            .grid(row=0, column=0, sticky="w", **pad)

        # --- Frame selección carpetas (EXPANDIBLE) ---
        paths_frame = ctk.CTkFrame(main)
        paths_frame.grid(row=1, column=0, sticky="nsew", **pad)
        # Columnas: [label | entry expandible | botón]
        paths_frame.grid_columnconfigure(0, weight=0)
        paths_frame.grid_columnconfigure(1, weight=1)  # <- el entry crece horizontalmente
        paths_frame.grid_columnconfigure(2, weight=0)

        def path_row(r: int, label: str, var: StringVar, browse_cmd, required: bool = False):
            lab_txt = f"{label}{' (obligatoria)' if required else ''}:"
            ctk.CTkLabel(paths_frame, text=lab_txt, font=font_label)\
                .grid(row=r, column=0, sticky="w", padx=10, pady=10)
            entry = ctk.CTkEntry(paths_frame, textvariable=var, height=46, font=font_entry, state="disabled")
            entry.grid(row=r, column=1, sticky="we", padx=10, pady=10)
            ctk.CTkButton(paths_frame, text="Seleccionar...", command=browse_cmd, height=46, font=font_btn)\
                .grid(row=r, column=2, padx=10, pady=10, sticky="e")

        # Filas de rutas
        path_row(0, "Carpeta de Excel", self.excel_dir, self._choose_excel, required=True)
        path_row(1, "Carpeta de Word (plantillas)", self.word_dir, self._choose_word)
        path_row(2, "Carpeta de plantillas HTML", self.html_templates_dir, self._choose_html_templates)
        path_row(3, "Carpeta de carátulas (Anejo 5)", self.caratulas_dir, self._choose_caratulas)
        path_row(4, "Carpeta de fotos", self.photos_dir, self._choose_photos)
        path_row(5, "Carpeta de salida (anexos)", self.output_dir, self._choose_output)
        path_row(6, "Carpeta de CEE", self.cee_dir, self._choose_cee)
        path_row(7, "Carpeta de planos", self.plans_dir, self._choose_plans)

        # --- Frame opciones (EXPANDIBLE) ---
        opts_frame = ctk.CTkFrame(main)
        opts_frame.grid(row=2, column=0, sticky="nsew", **pad)
        # Columnas: [label | col1 | col2 | col3 expandible]
        opts_frame.grid_columnconfigure(0, weight=0)
        opts_frame.grid_columnconfigure(1, weight=0)
        opts_frame.grid_columnconfigure(2, weight=0)
        opts_frame.grid_columnconfigure(3, weight=1)

        # Anexos
        ctk.CTkLabel(opts_frame, text="Anexos:", font=font_label)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=12)

        self.radio_ax_all = ctk.CTkRadioButton(
            opts_frame, text="Todos", variable=self.anexos_mode, value="all",
            command=self._on_anexos_mode_change, font=font_label
        )
        self.radio_ax_one = ctk.CTkRadioButton(
            opts_frame, text="Solo anexo(s):", variable=self.anexos_mode, value="one",
            command=self._on_anexos_mode_change, font=font_label
        )
        self.radio_ax_all.grid(row=0, column=1, sticky="w", padx=10, pady=12)
        self.radio_ax_one.grid(row=0, column=2, sticky="w", padx=10, pady=12)

        # Contenedor ayuda + entry ANEXOS
        anx_box = ctk.CTkFrame(opts_frame, fg_color="transparent")
        anx_box.grid(row=0, column=3, sticky="w", padx=10, pady=12)
        anx_box.grid_columnconfigure(0, weight=0)  # <- no expandir

        ctk.CTkLabel(
            anx_box, text="1-4, 6", font=font_hint, text_color=("gray50", "gray60")
        ).grid(row=0, column=0, sticky="w")

        self.entry_anexos = ctk.CTkEntry(
            anx_box,
            textvariable=self.anexos_value,
            height=46,
            font=font_entry,
            width=300  # <- ancho fijo estilo "Año"
        )
        self.entry_anexos.grid(row=1, column=0, sticky="w", pady=(6, 0))  # <- no 'we'
        self.entry_anexos.configure(state="disabled")

        # Mes / Año
        ctk.CTkLabel(opts_frame, text="Mes:", font=font_label)\
            .grid(row=1, column=0, sticky="w", padx=10, pady=(0, 20))
        month_names = [
            ("01", "01 - Enero"), ("02", "02 - Febrero"), ("03", "03 - Marzo"),
            ("04", "04 - Abril"), ("05", "05 - Mayo"), ("06", "06 - Junio"),
            ("07", "07 - Julio"), ("08", "08 - Agosto"), ("09", "09 - Septiembre"),
            ("10", "10 - Octubre"), ("11", "11 - Noviembre"), ("12", "12 - Diciembre"),
        ]
        self.combo_month = ctk.CTkComboBox(
            opts_frame,
            values=[label for _val, label in month_names],
            height=46, font=font_label, width=200, 
            command=lambda v: self.month_var.set(v.split(" - ")[0])
        )
        try:
            idx = int(self.month_var.get()) - 1
        except Exception:
            idx = 0
        self.combo_month.set(month_names[idx][1])
        self.combo_month.grid(row=1, column=1, sticky="w", padx=10, pady=(0, 20))

        ctk.CTkLabel(opts_frame, text="Año:", font=font_label)\
            .grid(row=1, column=2, sticky="e", padx=10, pady=(0, 12))
        year_now = datetime.now().year
        years = [str(y) for y in range(year_now - 50, year_now + 50)]
        self.combo_year = ctk.CTkComboBox(
            opts_frame,
            values=years,
            height=46, font=font_label, width=150,  
            command=lambda v: self.year_var.set(int(v))
        )
        self.combo_year.set(str(self.year_var.get()))
        self.combo_year.grid(row=1, column=3, sticky="w", padx=10, pady=(0, 12))

        # Centros
        ctk.CTkLabel(opts_frame, text="Centros:", font=font_label)\
            .grid(row=2, column=0, sticky="w", padx=10, pady=(12, 12))

        self.radio_centers_all = ctk.CTkRadioButton(
            opts_frame, text="Todos", variable=self.center_mode, value="all",
            command=self._on_center_mode_change, font=font_label
        )
        self.radio_centers_one = ctk.CTkRadioButton(
            opts_frame, text="Solo centro(s):", variable=self.center_mode, value="one",
            command=self._on_center_mode_change, font=font_label
        )
        self.radio_centers_all.grid(row=2, column=1, sticky="w", padx=10, pady=(12, 12))
        self.radio_centers_one.grid(row=2, column=2, sticky="w", padx=10, pady=(12, 12))

        # Contenedor ayuda + entry CENTROS
        ctr_box = ctk.CTkFrame(opts_frame, fg_color="transparent")
        ctr_box.grid(row=2, column=3, sticky="w", padx=10, pady=(12, 12))
        ctr_box.grid_columnconfigure(0, weight=0)  # <- no expandir

        ctk.CTkLabel(
            ctr_box, text="C0002-C10, C0023, C035", font=font_hint, text_color=("gray50", "gray60")
        ).grid(row=0, column=0, sticky="w")

        self.entry_center = ctk.CTkEntry(
            ctr_box,
            textvariable=self.center_value,
            height=46,
            font=font_entry,
            width=300 
        )
        self.entry_center.grid(row=1, column=0, sticky="w", pady=(6, 0))  # <- no 'we'
        self.entry_center.configure(state="disabled")
        
        # Photo filtering option for Anejo 5
        ctk.CTkLabel(opts_frame, text="Anejo 5:", font=font_label)\
            .grid(row=3, column=0, sticky="w", padx=10, pady=(12, 12))
        
        self.checkbox_include_without_photos = ctk.CTkCheckBox(
            opts_frame, 
            text="Incluir elementos sin fotos", 
            variable=self.include_without_photos,
            font=font_label
        )
        self.checkbox_include_without_photos.grid(row=3, column=1, columnspan=3, sticky="w", padx=10, pady=(12, 12))

        # --- Botonera (abajo) ---
        btn_frame = ctk.CTkFrame(main)
        btn_frame.grid(row=3, column=0, sticky="ew", **pad)
        btn_frame.grid_columnconfigure(0, weight=0)
        btn_frame.grid_columnconfigure(1, weight=0)
        btn_frame.grid_columnconfigure(2, weight=1)  # empuja el botón "Salir" a la derecha

        self.btn_run = ctk.CTkButton(btn_frame, text="Ejecutar", command=self._on_run, height=48, font=font_btn)
        self.btn_stop = ctk.CTkButton(btn_frame, text="Detener ejecución", command=self._on_stop,
                                    height=48, font=font_btn, state="disabled")
        self.btn_exit = ctk.CTkButton(btn_frame, text="Salir", command=self._on_close, height=48, font=font_btn)

        self.btn_run.grid(row=0, column=0, padx=10, pady=12, sticky="w")
        self.btn_stop.grid(row=0, column=1, padx=10, pady=12, sticky="w")
        self.btn_exit.grid(row=0, column=2, padx=10, pady=12, sticky="e")



    # ---------- pestaña mover ----------
    def _build_move_tab(self):
        # Defensa por si el estado no estaba creado aún
        # for attr, default in (
        #     ("mv_local_root", StringVar(value="")),
        #     ("mv_nas_root", StringVar(value="")),
        #     ("mv_centers_expr", StringVar(value="")),
        #     ("mv_dry_run", ctk.BooleanVar(value=False)),
        # ):
        #     if not hasattr(self, attr):
        #         setattr(self, attr, default)

        pad = {"padx": 14, "pady": 10}
        font_title = ctk.CTkFont(size=22, weight="bold")
        font_label = ctk.CTkFont(size=18)
        font_btn = ctk.CTkFont(size=18)

        ctk.CTkLabel(self.tab_move, text="Mover anexos al NAS", font=font_title)\
            .pack(**pad, anchor="w")

        frame = ctk.CTkFrame(self.tab_move)
        frame.pack(fill="x", **pad)

        def row(r: int, label: str, var: StringVar, browse_cmd):
            ctk.CTkLabel(frame, text=label, font=font_label).grid(row=r, column=0, sticky="w", padx=10, pady=10)
            ctk.CTkEntry(frame, textvariable=var, width=720, height=38, state="disabled")\
                .grid(row=r, column=1, sticky="we", padx=10, pady=10)
            ctk.CTkButton(frame, text="Seleccionar...", command=browse_cmd, height=38, font=font_btn)\
                .grid(row=r, column=2, padx=10, pady=10)

        row(0, "Carpeta local (raíz por centro):", self.mv_local_root, self._choose_mv_local_root)
        row(1, "Carpeta NAS (carpetas de centros):", self.mv_nas_root, self._choose_mv_nas_root)
        row(2, "Carpeta de plantillas Word:", self.word_dir, self._choose_word_synchronized)

        frame.grid_columnconfigure(1, weight=1)

        # Filtro de centros
        filter_frame = ctk.CTkFrame(self.tab_move)
        filter_frame.pack(fill="x", **pad)

        ctk.CTkLabel(filter_frame, text="Filtrar centros (opcional):", font=font_label)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.mv_centers_entry = ctk.CTkEntry(filter_frame, textvariable=self.mv_centers_expr,
                                             width=360, height=38,
                                             placeholder_text="C0001-C0010, C0012")
        self.mv_centers_entry.grid(row=0, column=1, sticky="w", padx=10, pady=10)

        # Dry-run
        self.mv_dry_run_chk = ctk.CTkCheckBox(filter_frame, text="Simular (no mueve archivos)", variable=self.mv_dry_run)
        self.mv_dry_run_chk.grid(row=0, column=2, sticky="w", padx=10, pady=10)

        # Botones
        btn_frame = ctk.CTkFrame(self.tab_move)
        btn_frame.pack(fill="x", **pad)
        self.btn_mv_run = ctk.CTkButton(btn_frame, text="Mover al NAS", command=self._on_mv_run, width=160, height=42, font=font_btn)
        self.btn_mv_stop = ctk.CTkButton(btn_frame, text="Detener", command=self._on_stop, width=120, height=42, font=font_btn, state="disabled")
        self.btn_mv_run.pack(side="left", padx=8, pady=12)
        self.btn_mv_stop.pack(side="left", padx=8, pady=12)

        # Logs de movimiento
        ctk.CTkLabel(self.tab_move, text="Salida:", font=ctk.CTkFont(size=20, weight="bold")).pack(**pad, anchor="w")
        self.txt_mv_logs = ctk.CTkTextbox(self.tab_move, height=360, font=ctk.CTkFont(size=16))
        self.txt_mv_logs.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self._log_mv("Configura las carpetas y pulsa 'Mover al NAS'.")

    def _build_memoria_tab(self):
        """Construye la pestaña para generar memoria final."""
        pad = {"padx": 14, "pady": 8}
        
        # Título
        ctk.CTkLabel(self.tab_memoria, text="Generar Memoria Final", 
                    font=ctk.CTkFont(size=24, weight="bold")).pack(**pad, anchor="w")
        
        # Frame principal sin scroll
        main_frame = ctk.CTkFrame(self.tab_memoria)
        main_frame.pack(fill="both", expand=True, **pad)
        
        # Sección de configuración
        config_frame = ctk.CTkFrame(main_frame)
        config_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(config_frame, text="Configuración", 
                    font=ctk.CTkFont(size=18, weight="bold")).pack(**pad, anchor="w")
        
        # Carpeta NAS (06_REDACCION)
        ctk.CTkLabel(config_frame, text="Carpeta NAS (06_REDACCION):").pack(**pad, anchor="w")
        input_frame = ctk.CTkFrame(config_frame)
        input_frame.pack(fill="x", **pad)
        self.entry_memoria_input = ctk.CTkEntry(input_frame, textvariable=self.memoria_input_dir,
                                               placeholder_text="Y:/2025/.../06_REDACCION")
        self.entry_memoria_input.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(input_frame, text="...", width=40, 
                     command=self._browse_memoria_input).pack(side="right")
        
        # Centro específico (opcional)
        ctk.CTkLabel(config_frame, text="Centro específico (opcional):").pack(**pad, anchor="w")
        self.entry_memoria_center = ctk.CTkEntry(config_frame, textvariable=self.memoria_center, 
                                               placeholder_text="C0007 (vacío = todos los centros)")
        self.entry_memoria_center.pack(fill="x", **pad)
        
        # Plantilla de índices
        ctk.CTkLabel(config_frame, text="Plantilla de índices (001_INDICE GENERAL_PLANTILLA.docx):").pack(**pad, anchor="w")
        template_frame = ctk.CTkFrame(config_frame)
        template_frame.pack(fill="x", **pad)
        self.entry_memoria_template = ctk.CTkEntry(template_frame, textvariable=self.memoria_template_path,
                                                  placeholder_text="Automático si se deja vacío...")
        self.entry_memoria_template.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(template_frame, text="...", width=40, 
                     command=self._browse_memoria_template).pack(side="right", padx=(0, 4))
        ctk.CTkButton(template_frame, text="Reset", width=60, 
                     command=self._reset_memoria_template).pack(side="right")
        
        # Tipo de acción
        ctk.CTkLabel(config_frame, text="Acción a realizar:").pack(**pad, anchor="w")
        self.memoria_action = ctk.StringVar(value="all")
        action_frame = ctk.CTkFrame(config_frame)
        action_frame.pack(fill="x", **pad)
        
        ctk.CTkRadioButton(action_frame, text="Solo índices", variable=self.memoria_action, 
                          value="indices").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(action_frame, text="Solo memoria PDF", variable=self.memoria_action, 
                          value="memoria").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(action_frame, text="Ambos", variable=self.memoria_action, 
                          value="all").pack(side="left")
        
        # Información
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(fill="x", pady=(10, 0))
        
        info_text = """ℹ️ INFORMACIÓN:
• Índices: Genera 001_INDICE_GENERAL_COMPLETADO en Word y PDF en cada centro
• Memoria PDF: Combina PORTADA + ÍNDICE + AUDITORÍA + ANEJOS en MEMORIA_COMPLETA.pdf
• Los archivos se generan directamente en las carpetas de cada centro del NAS"""
        
        ctk.CTkLabel(info_frame, text=info_text, font=ctk.CTkFont(size=16, weight="bold"), 
                    justify="left").pack(**pad, anchor="w")
        
        # Botones de acción
        btn_frame = ctk.CTkFrame(self.tab_memoria)
        btn_frame.pack(fill="x", **pad)
        
        self.btn_memoria_generate = ctk.CTkButton(btn_frame, text="Generar", 
                                                 command=self._on_generate_memoria,
                                                 font=ctk.CTkFont(size=16, weight="bold"), 
                                                 height=40)
        self.btn_memoria_generate.pack(side="left", padx=8, pady=12)
        
        self.btn_memoria_stop = ctk.CTkButton(btn_frame, text="Cancelar", 
                                            command=self._on_memoria_stop,
                                            font=ctk.CTkFont(size=16, weight="bold"),
                                            height=40,
                                            fg_color="gray40", hover_color="gray30")
        self.btn_memoria_stop.pack(side="left", padx=8, pady=12)
        
        # Logs de memoria
        ctk.CTkLabel(self.tab_memoria, text="Salida:", font=ctk.CTkFont(size=20, weight="bold")).pack(**pad, anchor="w")
        self.txt_memoria_logs = ctk.CTkTextbox(self.tab_memoria, height=360, font=ctk.CTkFont(size=16))
        self.txt_memoria_logs.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self._log_memoria("Configura la carpeta NAS y pulsa 'Generar'.")

    # ----- estado (tk variables) -----

    def _build_state(self) -> None:
        # Generación
        self.excel_dir = StringVar(value="")
        self.word_dir = StringVar(value="")
        self.html_templates_dir = StringVar(value="")
        self.photos_dir = StringVar(value="")
        self.output_dir = StringVar(value="")
        self.cee_dir = StringVar(value="")
        self.plans_dir = StringVar(value="")
        self.caratulas_dir = StringVar(value="")  # carpeta de carátulas (Anejo 5)

        # Agregar callback para sincronizar word_dir
        self.word_dir.trace_add('write', self._on_word_dir_changed)

        now = datetime.now()
        self.month_var = StringVar(value=str(now.month).zfill(2))
        self.year_var = IntVar(value=now.year)

        # (rango/lista de anexos)
        self.anexos_mode = ctk.StringVar(value="all")
        self.anexos_value = ctk.StringVar(value="")  # "1-3, 6"

        # Centros
        self.center_mode = ctk.StringVar(value="all")
        self.center_value = ctk.StringVar(value="")
        
        # Photo filtering for Anejo 5
        self.include_without_photos = ctk.BooleanVar(value=True)  # Default: include elements without photos

        # Movimiento (DEBEN existir antes de _build_move_tab)
        self.mv_local_root = StringVar(value="")
        self.mv_nas_root = StringVar(value="")
        self.mv_centers_expr = StringVar(value="")
        self.mv_dry_run = ctk.BooleanVar(value=False)

        # Memoria Final
        self.memoria_input_dir = StringVar(value="")  # Carpeta NAS 06_REDACCION
        self.memoria_output_dir = StringVar(value="")  # No usado (se genera in-situ)
        self.memoria_center = StringVar(value="")  # Centro específico (opcional)
        self.memoria_action = ctk.StringVar(value="all")  # Acción: indices, memoria, all
        self.memoria_template_path = StringVar(value="")  # Plantilla de índices

        # Control del estado de botones para evitar parpadeo
        self._last_runner_state = False


    # ----- helpers UI comunes -----

    def _log(self, text: str) -> None:
        # Logs de la pestaña "Generación de Anejos" (no escribir en la terminal aquí)
        if hasattr(self, "txt_logs"):
            self.txt_logs.configure(state="normal")
            self.txt_logs.insert("end", text + "\n")
            self.txt_logs.see("end")
            self.txt_logs.configure(state="disabled")


    def _log_mv(self, text: str) -> None:
        # Logs de la pestaña "Mover anexos al NAS" (no escribir en la terminal aquí)
        if hasattr(self, "txt_mv_logs"):
            self.txt_mv_logs.configure(state="normal")
            self.txt_mv_logs.insert("end", text + "\n")
            self.txt_mv_logs.see("end")
            self.txt_mv_logs.configure(state="disabled")

    def _log_memoria(self, text: str) -> None:
        # Logs de la pestaña "Memoria Final"
        if hasattr(self, "txt_memoria_logs"):
            self.txt_memoria_logs.configure(state="normal")
            self.txt_memoria_logs.insert("end", text + "\n")
            self.txt_memoria_logs.see("end")
            self.txt_memoria_logs.configure(state="disabled")

    def _snapshot_config(self) -> dict:
        """Toma un snapshot del estado actual para persistir en ConfigStore."""
        try:
            month_val = int(str(self.month_var.get()).strip())
        except Exception:
            month_val = None
        try:
            year_val = int(str(self.year_var.get()).strip())
        except Exception:
            year_val = None

        return {
            # Generación
            "excel_dir": (self.excel_dir.get() or "").strip(),
            "word_dir": (self.word_dir.get() or "").strip(),
            "html_templates_dir": (self.html_templates_dir.get() or "").strip(),
            "photos_dir": (self.photos_dir.get() or "").strip(),
            "output_dir": (self.output_dir.get() or "").strip(),
            "cee_dir": (self.cee_dir.get() or "").strip(),
            "plans_dir": (self.plans_dir.get() or "").strip(),
            "month": month_val if month_val is not None else 1,
            "year": year_val if year_val is not None else 2000,

            # Anexos (si está en 'one' guardamos el texto, si no lo dejamos vacío)
            "anexos_mode": self.anexos_mode.get(),
            "anexos_expr": (self.anexos_value.get() or "").strip() if self.anexos_mode.get() == "one" else "",

            # Centros (idem)
            "centers_mode": self.center_mode.get(),
            "centers": (self.center_value.get() or "").strip() if self.center_mode.get() == "one" else "",
            
            # Photo filtering for Anejo 5
            "include_without_photos": bool(self.include_without_photos.get()),

            # Movimiento
            "mv_local_root": (self.mv_local_root.get() or "").strip(),
            "mv_nas_root": (self.mv_nas_root.get() or "").strip(),
            "mv_centers_expr": (self.mv_centers_expr.get() or "").strip(),
            "mv_dry_run": bool(self.mv_dry_run.get()),

            # Memoria Final
            "memoria_input_dir": (self.memoria_input_dir.get() or "").strip(),
            "memoria_output_dir": (self.memoria_output_dir.get() or "").strip(),
            "memoria_center": (self.memoria_center.get() or "").strip(),
            "memoria_action": self.memoria_action.get(),
            "memoria_template_path": (self.memoria_template_path.get() or "").strip(),
        }

        

    # ----- browsers -----

    def _choose_excel(self) -> None:
        # Sugerir la carpeta proyecto por defecto si existe
        default_path = os.path.join(os.getcwd(), "excel", "proyecto")
        initial_dir = default_path if os.path.isdir(default_path) else None
        
        path = filedialog.askdirectory(
            title="Selecciona carpeta de Excel",
            initialdir=initial_dir
        )
        if path:
            self.excel_dir.set(path)

    def _choose_word(self) -> None:
        """Versión para la pestaña de Generación - sincronizada con Mover al NAS."""
        path = filedialog.askdirectory(title="Selecciona carpeta de Word (plantillas)")
        if path:
            self.word_dir.set(path)

    def _choose_word_synchronized(self) -> None:
        """Versión sincronizada para ambas pestañas (Generación y Mover al NAS)."""
        path = filedialog.askdirectory(title="Selecciona carpeta de Word (plantillas)")
        if path:
            self.word_dir.set(path)

    def _choose_html_templates(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de plantillas HTML")
        if path:
            self.html_templates_dir.set(path)

    def _choose_photos(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de fotos")
        if path:
            self.photos_dir.set(path)

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de salida para anexos")
        if path:
            self.output_dir.set(path)

    def _choose_cee(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de CEE")
        if path:
            self.cee_dir.set(path)

    def _choose_plans(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de planos")
        if path:
            self.plans_dir.set(path)

    def _choose_caratulas(self) -> None:
        path = filedialog.askdirectory(title="Selecciona la carpeta de carátulas (PDF)")
        if path:
            self.caratulas_dir.set(path)

    # Movimiento
    def _choose_mv_local_root(self) -> None:
        path = filedialog.askdirectory(title="Selecciona la carpeta local raíz (Cxxxx o varias Cxxxx)")
        if path:
            self.mv_local_root.set(path)

    def _choose_mv_nas_root(self) -> None:
        path = filedialog.askdirectory(title="Selecciona la carpeta del NAS con las carpetas de los centros (p.ej. 02_UN EDIFICIO)")
        if path:
            self.mv_nas_root.set(path)

    def _browse_memoria_input(self) -> None:
        path = filedialog.askdirectory(title="Selecciona la carpeta NAS 06_REDACCION")
        if path:
            self.memoria_input_dir.set(path)

    def _browse_memoria_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecciona la plantilla de índices (001_INDICE GENERAL_PLANTILLA.docx)",
            filetypes=[("Archivos Word", "*.docx"), ("Todos los archivos", "*.*")],
            initialdir="Y:/DOCUMENTACION TRABAJO/CARPETAS PERSONAL/SO/github_app/artecoin_automatizaciones/word/anexos"
        )
        if path:
            # Validar que sea la plantilla correcta
            if "001_INDICE GENERAL_PLANTILLA.docx" not in Path(path).name:
                messagebox.showwarning(
                    APP_NAME, 
                    "ADVERTENCIA: Esta no es la plantilla correcta para índices.\n\n"
                    "Debe seleccionar: 001_INDICE GENERAL_PLANTILLA.docx\n"
                    "No un anejo como: 01_ANEJO 1. METODOLOGIA_V1.docx",
                    parent=self
                )
                return
            self.memoria_template_path.set(path)

    def _reset_memoria_template(self) -> None:
        """Resetea la plantilla a la por defecto del sistema."""
        default_template = str(Path(__file__).parent.parent / "word" / "anexos" / "001_INDICE GENERAL_PLANTILLA.docx")
        self.memoria_template_path.set(default_template)
        messagebox.showinfo(APP_NAME, "Plantilla reseteada a la por defecto del sistema.", parent=self)

    # ----- eventos -----

    def _on_word_dir_changed(self, *args) -> None:
        """Callback que se ejecuta cuando cambia word_dir para mantener sincronizadas ambas pestañas."""
        try:
            # Guardar configuración cuando cambie word_dir
            cfg = ConfigStore.load()
            cfg["word_dir"] = (self.word_dir.get() or "").strip()
            ConfigStore.save(cfg)
        except Exception:
            pass  # No bloquear la UI si falla el guardado

    def _on_center_mode_change(self) -> None:
        if self.center_mode.get() == "one":
            self.entry_center.configure(state="normal")
            try:
                self.entry_center.focus()
            except Exception:
                pass
        else:
            self.entry_center.configure(state="disabled")
            self.center_value.set("")
    
    def _on_anexos_mode_change(self) -> None:
        if self.anexos_mode.get() == "one":
            self.entry_anexos.configure(state="normal")
            try:
                self.entry_anexos.focus()
            except Exception:
                pass
        else:
            self.entry_anexos.configure(state="disabled")
            self.anexos_value.set("")


    # ----- persistencia -----

    def _apply_saved_config(self) -> None:
        cfg = ConfigStore.load()

        # Generación
        self.excel_dir.set(cfg.get("excel_dir", ""))
        self.word_dir.set(cfg.get("word_dir", ""))
        self.html_templates_dir.set(cfg.get("html_templates_dir", ""))
        self.photos_dir.set(cfg.get("photos_dir", ""))
        self.output_dir.set(cfg.get("output_dir", ""))
        self.cee_dir.set(cfg.get("cee_dir", ""))
        self.plans_dir.set(cfg.get("plans_dir", ""))

        # Mes/Año
        m = str(cfg.get("month", str(datetime.now().month).zfill(2)))
        y = int(cfg.get("year", datetime.now().year))
        self.month_var.set(m.zfill(2))
        if hasattr(self, 'combo_month'):
            month_names = [
                ("01", "01 - Enero"), ("02", "02 - Febrero"), ("03", "03 - Marzo"),
                ("04", "04 - Abril"), ("05", "05 - Mayo"), ("06", "06 - Junio"),
                ("07", "07 - Julio"), ("08", "08 - Agosto"), ("09", "09 - Septiembre"),
                ("10", "10 - Octubre"), ("11", "11 - Noviembre"), ("12", "12 - Diciembre"),
            ]
            try:
                idx = int(self.month_var.get()) - 1
            except Exception:
                idx = 0
            self.combo_month.set(month_names[idx][1])
        self.year_var.set(y)
        if hasattr(self, 'combo_year'):
            self.combo_year.set(str(y))

        # Anexos
        self.anexos_mode.set(cfg.get("anexos_mode", "all"))
        saved_anexos = (cfg.get("anexos_expr", "") or "").strip()
        self.anexos_value.set(saved_anexos if self.anexos_mode.get() == "one" else "")
        self._on_anexos_mode_change()

        # Centros
        self.center_mode.set(cfg.get("centers_mode", "all"))
        saved_centers = (cfg.get("centers", "") or "").strip()
        self.center_value.set(saved_centers if self.center_mode.get() == "one" else "")
        self._on_center_mode_change()
        
        # Photo filtering for Anejo 5
        self.include_without_photos.set(cfg.get("include_without_photos", True))

        # Movimiento
        self.mv_local_root.set(cfg.get("mv_local_root", ""))
        self.mv_nas_root.set(cfg.get("mv_nas_root", ""))
        self.mv_centers_expr.set(cfg.get("mv_centers_expr", ""))
        self.mv_dry_run.set(cfg.get("mv_dry_run", False))

        # Memoria Final
        self.memoria_input_dir.set(cfg.get("memoria_input_dir", ""))
        self.memoria_output_dir.set(cfg.get("memoria_output_dir", ""))
        self.memoria_center.set(cfg.get("memoria_center", ""))
        self.memoria_action.set(cfg.get("memoria_action", "all"))
        
        # Establecer plantilla por defecto si no hay una guardada
        default_template = str(Path(__file__).parent.parent / "word" / "anexos" / "001_INDICE GENERAL_PLANTILLA.docx")
        saved_template = cfg.get("memoria_template_path", default_template)
        
        # Verificar que la plantilla guardada sea correcta para índices
        if saved_template and "001_INDICE GENERAL_PLANTILLA.docx" in saved_template:
            self.memoria_template_path.set(saved_template)
        else:
            # Si hay una plantilla incorrecta guardada, usar la por defecto
            self.memoria_template_path.set(default_template)


    # ----- helpers generación -----

    def _collect_options(self) -> RunOptions:
        """Empaqueta los valores de la UI de generación en un RunOptions."""
        def _clean(s: object) -> str | None:
            if s is None:
                return None
            s = str(s).strip()
            return s or None

        def _to_int(v: object) -> int | None:
            try:
                s = str(v).strip()
                if not s:
                    return None
                return int(s)
            except Exception:
                return None

        anexos_expr = None
        if self.anexos_mode.get() == "one":
            anexos_expr = _clean(self.anexos_value.get())

        center_expr = None
        if self.center_mode.get() == "one":
            center_expr = _clean(self.center_value.get())

        return RunOptions(
            excel_dir=(self.excel_dir.get() or "").strip(),
            word_dir=_clean(self.word_dir.get()),
            output_dir=_clean(self.output_dir.get()),
            html_templates_dir=_clean(self.html_templates_dir.get()),
            photos_dir=_clean(self.photos_dir.get()),
            cee_dir=_clean(self.cee_dir.get()),
            plans_dir=_clean(self.plans_dir.get()),
            anexos_expr=anexos_expr,
            month=_to_int(self.month_var.get()),
            year=_to_int(self.year_var.get()),
            center=center_expr,
            include_without_photos=bool(self.include_without_photos.get()),
        )

    def _ensure_output_dir(self, output_dir: Optional[str]) -> bool:
        """Si se ha indicado carpeta de salida y no existe, propone crearla."""
        if not output_dir:
            return True
        if os.path.isdir(output_dir):
            return True
        if messagebox.askyesno(APP_NAME, f"La carpeta de salida no existe:\n\n{output_dir}\n\n¿Deseas crearla?", parent=self):
            try:
                os.makedirs(output_dir, exist_ok=True)
                return True
            except Exception as e:
                messagebox.showerror(APP_NAME, f"No se pudo crear la carpeta:\n{e}", parent=self)
                return False
        return False

    def _on_run(self) -> None:
        opts = self._collect_options()
        ok, msg = Validator.validate_options(opts)
        if not ok:
            messagebox.showerror(APP_NAME, msg, parent=self)
            return

        if not self._ensure_output_dir(opts.output_dir):
            return

        # Guardar selección
        ConfigStore.save(self._snapshot_config())

        try:
            self.btn_run.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            self._log("\n=== Iniciando proceso de anexos ===")
            self.runner.start(opts)
        except FileNotFoundError as e:
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            messagebox.showerror(APP_NAME, str(e), parent=self)
        except Exception as e:
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            messagebox.showerror(APP_NAME, f"Error al iniciar el proceso:\n{e}", parent=self)


    # ----- movimiento -----

    def _collect_move_options(self) -> MoveOptions:
        return MoveOptions(
            local_out_root=(self.mv_local_root.get() or "").strip(),
            nas_centers_dir=(self.mv_nas_root.get() or "").strip(),
            centers_expr=(self.mv_centers_expr.get() or "").strip() or None,
            dry_run=bool(self.mv_dry_run.get()),
            word_dir=(self.word_dir.get() or "").strip() or None,
        )

    def _on_mv_run(self) -> None:
        opts = self._collect_move_options()
        ok, msg = Validator.validate_move_options(opts)
        if not ok:
            messagebox.showerror(APP_NAME, msg, parent=self)
            return

        # Guardar selección
        cfg = ConfigStore.load()
        cfg.update({
            "mv_local_root": opts.local_out_root,
            "mv_nas_root": opts.nas_centers_dir,
            "mv_centers_expr": opts.centers_expr or "",
            "mv_dry_run": bool(opts.dry_run),
        })
        ConfigStore.save(cfg)

        try:
            self.btn_mv_run.configure(state="disabled")
            self.btn_mv_stop.configure(state="normal")
            self._log_mv("\n=== Iniciando movimiento de anexos al NAS ===")
            self.runner.start_move(opts)
        except FileNotFoundError as e:
            self.btn_mv_run.configure(state="normal")
            self.btn_mv_stop.configure(state="disabled")
            messagebox.showerror(APP_NAME, str(e), parent=self)
        except Exception as e:
            self.btn_mv_run.configure(state="normal")
            self.btn_mv_stop.configure(state="disabled")
            messagebox.showerror(APP_NAME, f"Error al iniciar el movimiento:\n{e}", parent=self)

    def _on_generate_memoria(self) -> None:
        """Ejecuta la generación de memoria final."""
        # Validación básica
        input_dir = self.memoria_input_dir.get().strip()
        center = self.memoria_center.get().strip()
        action = self.memoria_action.get()
        template_path = self.memoria_template_path.get().strip()
        
        if not input_dir:
            messagebox.showerror(APP_NAME, "Debes seleccionar la carpeta NAS (06_REDACCION).", parent=self)
            return
            
        # Validar template path si se especifica
        if template_path:
            template_path_obj = Path(template_path)
            # Verificar si es la plantilla correcta para índices
            if "001_INDICE GENERAL_PLANTILLA.docx" not in template_path_obj.name:
                self._log_memoria(f"ADVERTENCIA: Template incorrecto para índices: {template_path}")
                self._log_memoria("Debe ser '001_INDICE GENERAL_PLANTILLA.docx', no un anejo")
                template_path = ""  # Forzar uso de plantilla por defecto
            elif not template_path_obj.exists():
                self._log_memoria(f"ADVERTENCIA: La plantilla especificada no existe: {template_path}")
                self._log_memoria("Se usará la plantilla por defecto del sistema")
                template_path = ""  # Usar plantilla por defecto
            else:
                self._log_memoria(f"Plantilla validada correctamente: {template_path}")
        else:
            self._log_memoria("No se especificó template_path, usando plantilla por defecto del sistema")
            
        # Guardar configuración
        ConfigStore.save(self._snapshot_config())
        
        try:
            # Limpiar logs previos
            self.txt_memoria_logs.configure(state="normal")
            self.txt_memoria_logs.delete("1.0", "end")
            self.txt_memoria_logs.configure(state="disabled")
            
            self._log_memoria("=== Iniciando generación de memoria final ===")
            self._log_memoria(f"Carpeta NAS: {input_dir}")
            if center:
                self._log_memoria(f"Centro específico: {center}")
            else:
                self._log_memoria("Procesando todos los centros")
            self._log_memoria(f"Acción: {action}")
            if template_path:
                self._log_memoria(f"Plantilla: {template_path}")
            self._log_memoria("")
            
            # Construir comando
            script_path = Path(__file__).parent / "render_memoria.py"
            cmd = [
                sys.executable, "-u", str(script_path),
                "--input-dir", input_dir,
                "--action", action
            ]
            
            if center:
                cmd.extend(["--center", center])
            
            if template_path:
                self._log_memoria(f"Usando template_path desde GUI: {template_path}")
                cmd.extend(["--template-path", template_path])
            else:
                self._log_memoria("No se especificó template_path, usando por defecto")
            
            cmd_str = " ".join(f'"{arg}"' if " " in arg else arg for arg in cmd)
            self._log_memoria(f"CMD: {cmd_str}")
            
            # Configurar botones
            self.btn_memoria_generate.configure(state="disabled")
            self.btn_memoria_stop.configure(state="normal")
            
            # Ejecutar
            self.runner.run_async(
                cmd=cmd,
                callback_line=lambda line: self.log_queue.put(("memoria", line)),
                callback_finished=self._on_memoria_finished
            )
            
        except Exception as e:
            self.btn_memoria_generate.configure(state="normal")
            self.btn_memoria_stop.configure(state="disabled")
            messagebox.showerror(APP_NAME, f"Error al iniciar la generación de memoria:\n{e}", parent=self)

    def _on_memoria_stop(self) -> None:
        """Cancela la generación de memoria."""
        if self.runner.is_running():
            self.runner.stop()
        self.btn_memoria_generate.configure(state="normal")
        self.btn_memoria_stop.configure(state="disabled")

    def _on_memoria_output(self, line: str) -> None:
        """Callback para mostrar salida del proceso de memoria."""
        self._log_memoria(line.rstrip())

    def _on_memoria_finished(self, returncode: int) -> None:
        """Callback cuando termina el proceso de memoria."""
        self.btn_memoria_generate.configure(state="normal")
        self.btn_memoria_stop.configure(state="disabled")
        self._log_memoria(f"--- Proceso finalizado. Código: {returncode} ---")

    # ----- cierre / polling -----

    def _on_stop(self) -> None:
        if self.runner.is_running():
            self.runner.stop()
        # reactivar botones de todas las pestañas
        if hasattr(self, "btn_stop"):
            self.btn_stop.configure(state="disabled")
        if hasattr(self, "btn_run"):
            self.btn_run.configure(state="normal")
        if hasattr(self, "btn_mv_stop"):
            self.btn_mv_stop.configure(state="disabled")
        if hasattr(self, "btn_mv_run"):
            self.btn_mv_run.configure(state="normal")
        if hasattr(self, "btn_memoria_stop"):
            self.btn_memoria_stop.configure(state="disabled")
        if hasattr(self, "btn_memoria_generate"):
            self.btn_memoria_generate.configure(state="normal")

    def _on_close(self) -> None:
        # Guardar estado actual (así los cuadros vacíos quedan vacíos al reabrir)
        try:
            ConfigStore.save(self._snapshot_config())
        except Exception:
            pass

        if self.runner.is_running():
            if not messagebox.askyesno(APP_NAME, "Hay un proceso en ejecución. ¿Detener y salir?", parent=self):
                return
            self.runner.stop()
            time.sleep(0.3)
        self.destroy()


    def _poll_logs(self) -> None:
        try:
            while True:
                item = self.log_queue.get_nowait()
                
                # El item puede ser una línea simple (string) o una tupla (tipo, línea)
                if isinstance(item, tuple) and len(item) == 2:
                    log_type, line = item
                    
                    # 1) Escribir en la pestaña Terminal
                    self._log_terminal(line)
                    
                    # 2) Escribir en la pestaña específica
                    if log_type == "memoria":
                        self._log_memoria(line)
                    elif log_type == "mv":
                        self._log_mv(line)
                    else:
                        # Tipo desconocido, escribir en todas las pestañas
                        self._log(line)
                        self._log_mv(line)
                        self._log_memoria(line)
                else:
                    # Formato anterior (string simple) - escribir en todas las pestañas
                    line = str(item)
                    self._log_terminal(line)
                    self._log(line)
                    self._log_mv(line)
                    self._log_memoria(line)
                
                # 3) Forzar actualización inmediata de la UI
                self.update_idletasks()
                
        except queue.Empty:
            pass

        # Actualizar estado de botones SOLO si hay un cambio real
        current_runner_state = self.runner.is_running()
        if current_runner_state != self._last_runner_state:
            self._last_runner_state = current_runner_state
            
            if current_runner_state:
                # Proceso corriendo - deshabilitar botones de inicio
                if hasattr(self, "btn_stop"):
                    self.btn_stop.configure(state="normal")
                if hasattr(self, "btn_run"):
                    self.btn_run.configure(state="disabled")
                if hasattr(self, "btn_mv_stop"):
                    self.btn_mv_stop.configure(state="normal")
                if hasattr(self, "btn_mv_run"):
                    self.btn_mv_run.configure(state="disabled")
                if hasattr(self, "btn_memoria_stop"):
                    self.btn_memoria_stop.configure(state="normal")
                if hasattr(self, "btn_memoria_generate"):
                    self.btn_memoria_generate.configure(state="disabled")
            else:
                # Proceso parado - habilitar botones de inicio
                if hasattr(self, "btn_stop"):
                    self.btn_stop.configure(state="disabled")
                if hasattr(self, "btn_run"):
                    self.btn_run.configure(state="normal")
                if hasattr(self, "btn_mv_stop"):
                    self.btn_mv_stop.configure(state="disabled")
                if hasattr(self, "btn_mv_run"):
                    self.btn_mv_run.configure(state="normal")
                if hasattr(self, "btn_memoria_stop"):
                    self.btn_memoria_stop.configure(state="disabled")
                if hasattr(self, "btn_memoria_generate"):
                    self.btn_memoria_generate.configure(state="normal")

        # Reprogramar el polling con frecuencia optimizada
        self.after(100, self._poll_logs)  # Reducido de 50ms a 100ms



def main() -> None:
    set_uniform_scaling()
    app = AnexosApp()
    app.mainloop()


if __name__ == "__main__":
    main()
