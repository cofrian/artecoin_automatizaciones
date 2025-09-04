# -*- coding: utf-8 -*-
# ------------------------------------------------------------
# App Tkinter (azules/blancos) para exportar CEE
# El usuario SOLO elige xlsx_path y out_dir
# ------------------------------------------------------------
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading

from exporter_cee import exportar_cee  # importa la lógica

class CEEApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Exportador CEE · Pares/Impares")
        self.geometry("850x320")
        self._init_style()
        self._build_ui()

    def _init_style(self):
        # Paleta azul/blanco
        self.configure(bg="#f4f7fb")
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Title.TLabel",
                        background="#0B4F9E", foreground="white",
                        font=("Segoe UI Semibold", 14), padding=(12, 10))
        style.configure("Card.TFrame",
                        background="#ffffff", relief="flat", borderwidth=0)
        style.configure("TLabel",
                        background="#ffffff", foreground="#0B2E57",
                        font=("Segoe UI", 10))
        style.configure("TEntry",
                        fieldbackground="#ffffff", foreground="#0B2E57",
                        padding=6)
        style.configure("Primary.TButton",
                        background="#0B4F9E", foreground="white",
                        font=("Segoe UI Semibold", 10), padding=(12, 6))
        style.map("Primary.TButton", background=[("active", "#1565C0")])
        style.configure("Ghost.TButton",
                        background="#e9f1ff", foreground="#0B4F9E",
                        font=("Segoe UI", 10), padding=(10, 5))
        style.map("Ghost.TButton", background=[("active", "#d6e8ff")])

    def _build_ui(self):
        # Header
        header = ttk.Label(self, text="Exportador Etiquetas Energéticas", style="Title.TLabel", anchor="w")
        header.pack(fill="x")

        card = ttk.Frame(self, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True, padx=16, pady=16)

        # Variables
        self.xlsx_var = tk.StringVar()
        self.out_var  = tk.StringVar()
        self.procesando = False

        # Campos
        ttk.Label(card, text="Archivo Excel (.xlsx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(card, textvariable=self.xlsx_var, width=56).grid(row=1, column=0, sticky="we", pady=(4, 12))
        ttk.Button(card, text="Examinar…", style="Ghost.TButton",
                   command=self._pick_xlsx).grid(row=1, column=1, padx=(8, 0))

        ttk.Label(card, text="Carpeta de salida:").grid(row=2, column=0, sticky="w")
        ttk.Entry(card, textvariable=self.out_var, width=56).grid(row=3, column=0, sticky="we", pady=(4, 12))
        ttk.Button(card, text="Seleccionar…", style="Ghost.TButton",
                   command=self._pick_outdir).grid(row=3, column=1, padx=(8, 0))

        # Barra de progreso y estado
        self.progress_frame = ttk.Frame(card, style="Card.TFrame")
        self.progress_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=(8, 12))
        self.progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.status_label = ttk.Label(self.progress_frame, text="", foreground="#0B4F9E", font=("Segoe UI", 9))
        
        # Inicialmente ocultos
        self.progress_frame.grid_remove()

        # Botonera
        btns = ttk.Frame(card, style="Card.TFrame")
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(8, 0))
        self.export_btn = ttk.Button(btns, text="Exportar", style="Primary.TButton",
                                    command=self._run)
        self.export_btn.pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Salir", style="Ghost.TButton",
                   command=self.destroy).pack(side="right")

        card.columnconfigure(0, weight=1)

    def _pick_xlsx(self):
        f = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xlsb *.xls")]
        )
        if f:
            self.xlsx_var.set(f)

    def _pick_outdir(self):
        d = filedialog.askdirectory(title="Selecciona la carpeta de salida")
        if d:
            self.out_var.set(d)

    def _run(self):
        if self.procesando:
            return
            
        xlsx = self.xlsx_var.get().strip()
        outd = self.out_var.get().strip()
        if not xlsx or not Path(xlsx).exists():
            messagebox.showerror("Error", "Selecciona un archivo Excel válido.")
            return
        if not outd:
            messagebox.showerror("Error", "Selecciona una carpeta de salida.")
            return

        # Mostrar indicador de progreso
        self._mostrar_progreso("Iniciando exportación...")
        
        # Ejecutar en thread separado para no bloquear la UI
        thread = threading.Thread(target=self._ejecutar_exportacion, args=(xlsx, outd))
        thread.daemon = True
        thread.start()

    def _mostrar_progreso(self, mensaje):
        """Mostrar barra de progreso y mensaje de estado."""
        self.procesando = True
        self.export_btn.config(state="disabled")
        
        self.progress_frame.grid()
        self.progress_bar.grid(row=0, column=0, sticky="we", pady=(0, 4))
        self.status_label.grid(row=1, column=0, sticky="w")
        
        self.progress_bar.start(10)
        self.status_label.config(text=mensaje)
        
        # Actualizar la ventana
        self.update_idletasks()

    def _ocultar_progreso(self):
        """Ocultar indicadores de progreso."""
        self.procesando = False
        self.export_btn.config(state="normal")
        self.progress_bar.stop()
        self.progress_frame.grid_remove()

    def _ejecutar_exportacion(self, xlsx, outd):
        """Ejecutar la exportación en thread separado."""
        try:
            Path(outd).mkdir(parents=True, exist_ok=True)
            
            # Actualizar estado
            self.after(0, lambda: self._actualizar_estado("Procesando Excel y creando PDFs individuales..."))
            
            exportar_cee(xlsx_path=xlsx, out_dir=outd)
            
            # Calcular resultados
            carpeta_finales = Path(outd) / "ARCHIVOS ETIQUETAS FINALES"
            num_pdfs_combinados = len(list(carpeta_finales.glob("*.pdf"))) if carpeta_finales.exists() else 0
            
            # Mostrar mensaje de éxito y cerrar app
            self.after(0, lambda: self._mostrar_exito(outd, num_pdfs_combinados))
            
        except Exception as e:
            self.after(0, lambda: self._mostrar_error(str(e)))

    def _actualizar_estado(self, mensaje):
        """Actualizar mensaje de estado."""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=mensaje)

    def _mostrar_exito(self, outd, num_pdfs_combinados):
        """Mostrar mensaje de éxito y cerrar aplicación."""
        self._ocultar_progreso()
        
        mensaje = f"""🎉 ¡Exportación CEE Completada Exitosamente!

📁 Archivos guardados en: {outd}

✅ Se han creado:
   • PDFs individuales organizados por legalización
   • {num_pdfs_combinados} PDF(s) combinado(s) en la carpeta:
     'ARCHIVOS ETIQUETAS FINALES'

💡 Los PDFs combinados tienen como nombre el tipo de legalización y contienen todas las etiquetas de esa categoría.

La aplicación se cerrará al hacer clic en Aceptar."""
        
        messagebox.showinfo("🎉 Proceso Completado", mensaje)
        self.destroy()  # Cerrar la aplicación

    def _mostrar_error(self, error_msg):
        """Mostrar mensaje de error."""
        self._ocultar_progreso()
        messagebox.showerror("❌ Error", f"Se produjo un error durante la exportación:\n\n{error_msg}")

if __name__ == "__main__":
    CEEApp().mainloop()
