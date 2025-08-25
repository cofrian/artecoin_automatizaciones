#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
GUI (customtkinter) para ejecutar 'anexos_creator.py' o 'anexos_creator.exe' sin usar la terminal.

Cambios respecto a la versión anterior:
- UTF-8 explícito.
- Radio buttons con una única variable (alternancia correcta entre "todos" y "uno").
- Soporte de entradas interactivas al proceso (stdin) para manejar llamadas a input() en anexos_creator.py.
- Widgets más grandes (escalado y alturas) para mejor usabilidad.
- Estructura con SRP: Validator, ProcessRunner, ConfigStore, AnexosApp.

Requisitos:
    pip install customtkinter
"""

from __future__ import annotations
import os
import sys
import json
import time
import queue
import threading
import subprocess
from dataclasses import dataclass
from typing import Optional, Tuple, List

import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar, IntVar


# ---------------------------
# Configuración / Constantes
# ---------------------------

APP_NAME = "Anexos GUI"
CONFIG_FILENAME = os.path.join(os.path.expanduser("~"), ".anexos_gui_config.json")
ANEXO_CHOICES = [2, 3, 4, 5, 6, 7]


# ---------------------------
# Modelo de dominio
# ---------------------------

@dataclass(frozen=True)
class RunOptions:
    """Opciones de ejecución para 'anexos_creator'."""
    excel_dir: str
    word_dir: Optional[str]
    mode: str                 # "all" | "single"
    anexo: Optional[int]      # 2..7 cuando mode == "single"


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
    """Valida coherencia de opciones antes de ejecutar."""

    @staticmethod
    def validate_options(opts: RunOptions) -> Tuple[bool, str]:
        if not opts.excel_dir or not os.path.isdir(opts.excel_dir):
            return False, "Debes seleccionar una carpeta de Excel válida."

        excel_exts = (".xlsx", ".xlsm", ".xls")
        try:
            has_excel = any(f.lower().endswith(excel_exts) for f in os.listdir(opts.excel_dir))
        except Exception:
            has_excel = False
        if not has_excel:
            return False, "La carpeta de Excel no contiene archivos .xls/.xlsx/.xlsm."

        if opts.word_dir:
            if not os.path.isdir(opts.word_dir):
                return False, "La carpeta de Word indicada no existe."

        if opts.mode == "single":
            if opts.anexo not in ANEXO_CHOICES:
                return False, "Selecciona un anexo válido (2 a 7)."

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

    def start(self, opts: RunOptions) -> None:
        if self.is_running():
            raise RuntimeError("Ya hay un proceso en ejecución.")

        base_cmd, _ = self._find_target()

        cmd = list(base_cmd)
        cmd += ["--excel-dir", opts.excel_dir]
        if opts.word_dir:
            cmd += ["--word-dir", opts.word_dir]

        if opts.mode == "all":
            cmd += ["--mode", "all"]
        else:
            cmd += ["--mode", "single", "--anexo", str(opts.anexo)]

        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        self._proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,          # <--- para permitir input() desde la GUI
            text=True,
            bufsize=1,
            universal_newlines=True,
            creationflags=creationflags,
        )

        def _reader() -> None:
            assert self._proc is not None
            for line in self._proc.stdout:  # type: ignore[union-attr]
                self._log_queue.put(line.rstrip("\n"))
            self._proc.wait()
            self._log_queue.put(f"\n--- Proceso finalizado. Código: {self._proc.returncode} ---")

        self._reader_thread = threading.Thread(target=_reader, daemon=True)
        self._reader_thread.start()

    def stop(self) -> None:
        if not self.is_running():
            return
        try:
            self._proc.terminate()  # type: ignore[union-attr]
        except Exception:
            pass
        # Espera breve; si no muere, kill.
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


# ---------------------------
# Vista / Controlador (customtkinter)
# ---------------------------

class AnexosApp(ctk.CTk):
    """Ventana principal de la aplicación."""

    def __init__(self) -> None:
        super().__init__()
        # Apariencia y escalado más grande
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        ctk.set_widget_scaling(1.15)   # <--- widgets más grandes
        ctk.set_window_scaling(1.05)

        self.title(APP_NAME)
        self.geometry("980x680")
        self.minsize(880, 560)

        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.runner = ProcessRunner(self.log_queue)

        self._build_state()
        self._build_widgets()
        self._apply_saved_config()
        self._poll_logs()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ----- estado (tk variables) -----

    def _build_state(self) -> None:
        self.excel_dir = StringVar(value="")
        self.word_dir = StringVar(value="")
        self.mode_var = StringVar(value="all")  # <--- una única variable para alternar
        self.anexo_value = IntVar(value=ANEXO_CHOICES[0])

    # ----- UI -----

    def _build_widgets(self) -> None:
        pad = {"padx": 14, "pady": 10}
        font_title = ctk.CTkFont(size=22, weight="bold")
        font_label = ctk.CTkFont(size=14)
        font_btn = ctk.CTkFont(size=14)
        font_log = ctk.CTkFont(size=13)

        # Título
        title = ctk.CTkLabel(self, text="Interfaz de Anexos", font=font_title)
        title.pack(**pad, anchor="w")

        # Frame selección carpetas
        paths_frame = ctk.CTkFrame(self)
        paths_frame.pack(fill="x", **pad)

        # Excel
        ctk.CTkLabel(paths_frame, text="Carpeta de Excel (obligatoria):", font=font_label)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.entry_excel = ctk.CTkEntry(paths_frame, textvariable=self.excel_dir, width=600, height=38, state="disabled")
        self.entry_excel.grid(row=0, column=1, sticky="we", padx=10, pady=10)
        ctk.CTkButton(paths_frame, text="Seleccionar...", command=self._choose_excel, height=38, font=font_btn)\
            .grid(row=0, column=2, padx=10, pady=10)

        # Word
        ctk.CTkLabel(paths_frame, text="Carpeta de Word (plantillas):", font=font_label)\
            .grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.entry_word = ctk.CTkEntry(paths_frame, textvariable=self.word_dir, width=600, height=38, state="disabled")
        self.entry_word.grid(row=1, column=1, sticky="we", padx=10, pady=10)
        ctk.CTkButton(paths_frame, text="Seleccionar...", command=self._choose_word, height=38, font=font_btn)\
            .grid(row=1, column=2, padx=10, pady=10)

        paths_frame.grid_columnconfigure(1, weight=1)

        # Frame opciones
        opts_frame = ctk.CTkFrame(self)
        opts_frame.pack(fill="x", **pad)

        ctk.CTkLabel(opts_frame, text="Modo de ejecución:", font=font_label)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.radio_all = ctk.CTkRadioButton(
            opts_frame, text="Crear todos", variable=self.mode_var, value="all",
            command=self._on_mode_change, font=font_label)
        self.radio_all.grid(row=0, column=1, sticky="w", padx=10, pady=10)

        self.radio_single = ctk.CTkRadioButton(
            opts_frame, text="Crear uno", variable=self.mode_var, value="single",
            command=self._on_mode_change, font=font_label)
        self.radio_single.grid(row=0, column=2, sticky="w", padx=10, pady=10)

        ctk.CTkLabel(opts_frame, text="Anexo:", font=font_label)\
            .grid(row=0, column=3, sticky="e", padx=10, pady=10)
        self.combo_anexo = ctk.CTkComboBox(
            opts_frame, values=[f"Anexo {n}" for n in ANEXO_CHOICES],
            command=self._on_anexo_select, width=160, height=38, font=font_label)
        self.combo_anexo.set(f"Anexo {ANEXO_CHOICES[0]}")
        self.combo_anexo.configure(state="disabled")
        self.combo_anexo.grid(row=0, column=4, sticky="w", padx=10, pady=10)

        opts_frame.grid_columnconfigure(2, weight=1)

        # Botonera
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", **pad)

        self.btn_run = ctk.CTkButton(btn_frame, text="Ejecutar", command=self._on_run, width=140, height=40, font=font_btn)
        self.btn_stop = ctk.CTkButton(btn_frame, text="Detener ejecución", command=self._on_stop,
                                      width=180, height=40, font=font_btn, state="disabled")
        self.btn_exit = ctk.CTkButton(btn_frame, text="Salir", command=self._on_close, width=120, height=40, font=font_btn)

        self.btn_run.pack(side="left", padx=8, pady=12)
        self.btn_stop.pack(side="left", padx=8, pady=12)
        self.btn_exit.pack(side="right", padx=8, pady=12)

        # Logs
        ctk.CTkLabel(self, text="Logs:", font=ctk.CTkFont(size=15, weight="bold")).pack(**pad, anchor="w")
        self.txt_logs = ctk.CTkTextbox(self, height=360, font=font_log)
        self.txt_logs.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self._log("Listo. Configura las carpetas y el modo, luego pulsa 'Ejecutar'.")

        # Entrada para el proceso (stdin)
        stdin_frame = ctk.CTkFrame(self)
        stdin_frame.pack(fill="x", padx=14, pady=(6, 14))
        ctk.CTkLabel(stdin_frame, text="Entrada para el proceso (cuando pida input):", font=font_label)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.entry_stdin = ctk.CTkEntry(stdin_frame, width=620, height=38, font=font_label)
        self.entry_stdin.grid(row=0, column=1, sticky="we", padx=10, pady=10)
        self.btn_send = ctk.CTkButton(stdin_frame, text="Enviar", width=120, height=38, font=font_btn,
                                      command=self._on_send_input, state="disabled")
        self.btn_send.grid(row=0, column=2, padx=10, pady=10)
        self.entry_stdin.bind("<Return>", lambda _e: self._on_send_input())

        stdin_frame.grid_columnconfigure(1, weight=1)

        # Ayuda
        hint = ctk.CTkLabel(
            self,
            text="Consejo: si tu proceso usa plantillas Word, selecciona la carpeta de Word. "
                 "Si el proceso pide datos (ej. número de mes), escribe aquí y pulsa Enter.",
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=12, slant="italic"),
            wraplength=900
        )
        hint.pack(padx=14, pady=(0, 12), anchor="w")

    # ----- helpers UI -----

    def _log(self, text: str) -> None:
        self.txt_logs.configure(state="normal")
        self.txt_logs.insert("end", text + "\n")
        self.txt_logs.see("end")
        self.txt_logs.configure(state="disabled")

    def _choose_excel(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de Excel")
        if path:
            self.excel_dir.set(path)

    def _choose_word(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de Word (plantillas)")
        if path:
            self.word_dir.set(path)

    def _on_mode_change(self) -> None:
        # Habilitar/deshabilitar combo según modo
        if self.mode_var.get() == "single":
            self.combo_anexo.configure(state="normal")
        else:
            self.combo_anexo.configure(state="disabled")

    def _on_anexo_select(self, value: str) -> None:
        try:
            self.anexo_value.set(int(value.split()[-1]))
        except Exception:
            self.anexo_value.set(ANEXO_CHOICES[0])

    # ----- persistencia -----

    def _apply_saved_config(self) -> None:
        cfg = ConfigStore.load()
        self.excel_dir.set(cfg.get("excel_dir", ""))
        self.word_dir.set(cfg.get("word_dir", ""))

        mode = cfg.get("mode", "all")
        if mode not in ("all", "single"):
            mode = "all"
        self.mode_var.set(mode)
        self._on_mode_change()

        try:
            a = int(str(cfg.get("anexo", ANEXO_CHOICES[0])))
        except Exception:
            a = ANEXO_CHOICES[0]
        self.anexo_value.set(a)
        self.combo_anexo.set(f"Anexo {a}")

    # ----- eventos -----

    def _collect_options(self) -> RunOptions:
        mode = self.mode_var.get()
        anexo = self.anexo_value.get() if mode == "single" else None
        return RunOptions(
            excel_dir=self.excel_dir.get().strip(),
            word_dir=(self.word_dir.get().strip() or None),
            mode=mode,
            anexo=anexo
        )

    def _on_run(self) -> None:
        opts = self._collect_options()
        ok, msg = Validator.validate_options(opts)
        if not ok:
            messagebox.showerror(APP_NAME, msg, parent=self)
            return

        # Guardar selección
        ConfigStore.save({
            "excel_dir": opts.excel_dir,
            "word_dir": opts.word_dir or "",
            "mode": opts.mode,
            "anexo": opts.anexo or "",
        })

        if not opts.word_dir:
            messagebox.showinfo(APP_NAME, "No se seleccionó carpeta de Word (plantillas).", parent=self)

        try:
            self.btn_run.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            self.btn_send.configure(state="normal")  # permitir enviar stdin cuando esté corriendo
            self._log("\n=== Iniciando proceso de anexos ===")
            self.runner.start(opts)
        except FileNotFoundError as e:
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.btn_send.configure(state="disabled")
            messagebox.showerror(APP_NAME, str(e), parent=self)
        except Exception as e:
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.btn_send.configure(state="disabled")
            messagebox.showerror(APP_NAME, f"Error al iniciar el proceso:\n{e}", parent=self)

    def _on_stop(self) -> None:
        if self.runner.is_running():
            self.runner.stop()
        self.btn_stop.configure(state="disabled")
        self.btn_run.configure(state="normal")
        self.btn_send.configure(state="disabled")

    def _on_send_input(self) -> None:
        text = self.entry_stdin.get().strip()
        if not text:
            return
        if self.runner.send_input(text):
            self._log(f">>> {text}")
            self.entry_stdin.delete(0, "end")
        else:
            messagebox.showwarning(APP_NAME, "No hay un proceso en ejecución o no acepta entrada.", parent=self)

    def _on_close(self) -> None:
        if self.runner.is_running():
            if not messagebox.askyesno(APP_NAME, "Hay un proceso en ejecución. ¿Detener y salir?", parent=self):
                return
            self.runner.stop()
            time.sleep(0.3)
        self.destroy()

    # ----- polling de logs -----

    def _poll_logs(self) -> None:
        try:
            while True:
                line = self.log_queue.get_nowait()
                self._log(line)
        except queue.Empty:
            pass
        # auto-habilitar/deshabilitar botones según estado del proceso
        if self.runner.is_running():
            self.btn_stop.configure(state="normal")
            self.btn_run.configure(state="disabled")
            self.btn_send.configure(state="normal")
        else:
            self.btn_stop.configure(state="disabled")
            self.btn_run.configure(state="normal")
            self.btn_send.configure(state="disabled")
        # reprogramar
        self.after(100, self._poll_logs)


def main() -> None:
    app = AnexosApp()
    app.mainloop()


if __name__ == "__main__":
    main()
