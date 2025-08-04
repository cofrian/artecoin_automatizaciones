import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 Sistema de Procesamiento Excel y HTML")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.ruta_excel = tk.StringVar()
        self.progreso = tk.IntVar()
        
        self.crear_interfaz()
        
    def crear_interfaz(self):
        # Título principal
        titulo_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        titulo_frame.pack(fill='x', padx=10, pady=10)
        titulo_frame.pack_propagate(False)
        
        titulo_label = tk.Label(titulo_frame, 
                               text="🚀 SISTEMA DE PROCESAMIENTO EXCEL Y HTML",
                               bg='#2c3e50', fg='white', 
                               font=('Arial', 16, 'bold'))
        titulo_label.pack(expand=True)
        
        # Frame para selección de archivo
        archivo_frame = tk.LabelFrame(self.root, text="📁 Selección de Archivo Excel", 
                                     font=('Arial', 12, 'bold'), bg='#f0f0f0', padx=10, pady=10)
        archivo_frame.pack(fill='x', padx=20, pady=10)
        
        # Campo de ruta
        ruta_frame = tk.Frame(archivo_frame, bg='#f0f0f0')
        ruta_frame.pack(fill='x', pady=5)
        
        self.ruta_entry = tk.Entry(ruta_frame, textvariable=self.ruta_excel, 
                                  font=('Arial', 10), width=60)
        self.ruta_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        examinar_btn = tk.Button(ruta_frame, text="📂 Examinar", 
                               command=self.examinar_archivo,
                               bg='#3498db', fg='white', font=('Arial', 10, 'bold'))
        examinar_btn.pack(side='right')
        
        # Frame para opciones de procesamiento
        opciones_frame = tk.LabelFrame(self.root, text="⚙️ Opciones de Procesamiento", 
                                      font=('Arial', 12, 'bold'), bg='#f0f0f0', padx=10, pady=10)
        opciones_frame.pack(fill='x', padx=20, pady=10)
        
        # Botones de procesamiento
        botones_frame = tk.Frame(opciones_frame, bg='#f0f0f0')
        botones_frame.pack(fill='x', pady=10)
        
        self.btn_completo = tk.Button(botones_frame, 
                                     text="🚀 Proceso Completo\n(Excel + HTML)",
                                     command=self.ejecutar_proceso_completo,
                                     bg='#27ae60', fg='white', 
                                     font=('Arial', 12, 'bold'),
                                     width=20, height=3)
        self.btn_completo.pack(side='left', padx=10)
        
        self.btn_excel = tk.Button(botones_frame, 
                                  text="📊 Solo Excel\nFiltrado",
                                  command=self.ejecutar_solo_excel,
                                  bg='#e67e22', fg='white', 
                                  font=('Arial', 12, 'bold'),
                                  width=15, height=3)
        self.btn_excel.pack(side='left', padx=10)
        
        self.btn_html = tk.Button(botones_frame, 
                                 text="🌐 Solo HTML\n(desde carpeta)",
                                 command=self.ejecutar_solo_html,
                                 bg='#9b59b6', fg='white', 
                                 font=('Arial', 12, 'bold'),
                                 width=15, height=3)
        self.btn_html.pack(side='left', padx=10)
        
        # Frame para progreso
        progreso_frame = tk.LabelFrame(self.root, text="📊 Progreso", 
                                      font=('Arial', 12, 'bold'), bg='#f0f0f0', padx=10, pady=10)
        progreso_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(progreso_frame, mode='indeterminate')
        self.progress_bar.pack(fill='x', pady=5)
        
        # Área de texto para logs
        texto_frame = tk.Frame(progreso_frame, bg='#f0f0f0')
        texto_frame.pack(fill='both', expand=True, pady=5)
        
        self.texto_log = tk.Text(texto_frame, height=15, font=('Consolas', 9))
        scrollbar = tk.Scrollbar(texto_frame, orient="vertical", command=self.texto_log.yview)
        self.texto_log.configure(yscrollcommand=scrollbar.set)
        
        self.texto_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Frame inferior con botones de acción
        inferior_frame = tk.Frame(self.root, bg='#f0f0f0')
        inferior_frame.pack(fill='x', padx=20, pady=10)
        
        limpiar_btn = tk.Button(inferior_frame, text="🗑️ Limpiar Log", 
                               command=self.limpiar_log,
                               bg='#95a5a6', fg='white', font=('Arial', 10))
        limpiar_btn.pack(side='left')
        
        salir_btn = tk.Button(inferior_frame, text="❌ Salir", 
                             command=self.root.quit,
                             bg='#e74c3c', fg='white', font=('Arial', 10))
        salir_btn.pack(side='right')
        
        # Mensaje inicial
        self.log("🎯 Bienvenido al Sistema de Procesamiento Excel y HTML")
        self.log("📝 Selecciona un archivo Excel y elige una opción de procesamiento")
    
    def examinar_archivo(self):
        """Abre el diálogo para seleccionar archivo Excel"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*")
            ]
        )
        if archivo:
            self.ruta_excel.set(archivo)
            self.log(f"📁 Archivo seleccionado: {Path(archivo).name}")
    
    def log(self, mensaje):
        """Añade mensaje al área de log"""
        self.texto_log.insert(tk.END, f"{mensaje}\n")
        self.texto_log.see(tk.END)
        self.root.update_idletasks()
    
    def limpiar_log(self):
        """Limpia el área de log"""
        self.texto_log.delete(1.0, tk.END)
    
    def validar_archivo(self):
        """Valida que el archivo seleccionado sea válido"""
        ruta = self.ruta_excel.get().strip()
        
        if not ruta:
            messagebox.showerror("Error", "Por favor selecciona un archivo Excel")
            return False
            
        ruta_path = Path(ruta)
        
        if not ruta_path.exists():
            messagebox.showerror("Error", f"El archivo no existe:\n{ruta}")
            return False
            
        if not ruta_path.suffix.lower() in ['.xlsx', '.xls']:
            messagebox.showerror("Error", "El archivo debe ser un Excel (.xlsx o .xls)")
            return False
            
        return True
    
    def deshabilitar_botones(self):
        """Deshabilita los botones durante el procesamiento"""
        self.btn_completo.config(state='disabled')
        self.btn_excel.config(state='disabled')
        self.btn_html.config(state='disabled')
        self.progress_bar.start()
    
    def habilitar_botones(self):
        """Habilita los botones después del procesamiento"""
        self.btn_completo.config(state='normal')
        self.btn_excel.config(state='normal')
        self.btn_html.config(state='normal')
        self.progress_bar.stop()
    
    def ejecutar_proceso_completo(self):
        """Ejecuta el proceso completo en un hilo separado"""
        if not self.validar_archivo():
            return
            
        thread = threading.Thread(target=self._proceso_completo_thread)
        thread.daemon = True
        thread.start()
    
    def _proceso_completo_thread(self):
        """Hilo para ejecutar el proceso completo"""
        try:
            self.deshabilitar_botones()
            ruta_excel = self.ruta_excel.get()
            
            self.log("🚀 Iniciando proceso completo...")
            self.log(f"📄 Archivo: {Path(ruta_excel).name}")
            self.log("-" * 50)
            
            # Paso 1: Excel filtrado
            self.log("📊 PASO 1: Generando archivos Excel filtrados...")
            sys.path.append(str(Path(__file__).parent))
            
            import descarga_plots_excel
            descarga_plots_excel.aplicar_filtros_iterativos(ruta_excel)
            self.log("✅ PASO 1 COMPLETADO: Archivos Excel filtrados generados")
            
            # Paso 2: HTML
            self.log("🌐 PASO 2: Generando archivos HTML...")
            import exportar_a_html
            exportar_a_html.procesar_carpeta_completa(ruta_excel)
            self.log("✅ PASO 2 COMPLETADO: Archivos HTML generados")
            
            # Resumen
            nombre_excel = Path(ruta_excel).stem
            self.log("🏁 PROCESAMIENTO COMPLETO EXITOSO")
            self.log("📁 ARCHIVOS GENERADOS EN:")
            self.log(f"   📊 Excel: Z:\\...\\FRONT\\{nombre_excel}")
            self.log(f"   🌐 HTML:  Z:\\...\\FRONT\\imagenes_{nombre_excel}")
            
            messagebox.showinfo("Éxito", "¡Proceso completado exitosamente!\n\nRevisa las carpetas de salida.")
            
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento:\n{str(e)}")
        finally:
            self.habilitar_botones()
    
    def ejecutar_solo_excel(self):
        """Ejecuta solo la generación de Excel filtrados"""
        if not self.validar_archivo():
            return
            
        thread = threading.Thread(target=self._solo_excel_thread)
        thread.daemon = True
        thread.start()
    
    def _solo_excel_thread(self):
        """Hilo para ejecutar solo Excel"""
        try:
            self.deshabilitar_botones()
            ruta_excel = self.ruta_excel.get()
            
            self.log("📊 Generando archivos Excel filtrados...")
            sys.path.append(str(Path(__file__).parent))
            
            import descarga_plots_excel
            descarga_plots_excel.aplicar_filtros_iterativos(ruta_excel)
            self.log("✅ Archivos Excel filtrados generados correctamente")
            
            messagebox.showinfo("Éxito", "¡Archivos Excel generados exitosamente!")
            
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento:\n{str(e)}")
        finally:
            self.habilitar_botones()
    
    def ejecutar_solo_html(self):
        """Ejecuta solo la generación de HTML"""
        thread = threading.Thread(target=self._solo_html_thread)
        thread.daemon = True
        thread.start()
    
    def _solo_html_thread(self):
        """Hilo para ejecutar solo HTML"""
        try:
            self.deshabilitar_botones()
            
            # Para el modo solo HTML, necesitamos obtener la última carpeta procesada
            # o permitir al usuario seleccionar una carpeta específica
            ruta_excel = self.ruta_excel.get()
            if not ruta_excel:
                # Si no hay archivo seleccionado, usar el modo legacy
                self.log("🌐 Generando archivos HTML desde carpeta existente...")
                import exportar_a_html
                exportar_a_html.procesar_carpeta_completa()
            else:
                # Si hay archivo seleccionado, usar su nombre para encontrar la carpeta
                self.log(f"🌐 Generando archivos HTML desde carpeta de: {Path(ruta_excel).stem}")
                import exportar_a_html
                exportar_a_html.procesar_carpeta_completa(ruta_excel)
            
            sys.path.append(str(Path(__file__).parent))
            
            self.log("✅ Archivos HTML generados correctamente")
            
            messagebox.showinfo("Éxito", "¡Archivos HTML generados exitosamente!")
            
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento:\n{str(e)}")
        finally:
            self.habilitar_botones()

def main():
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
