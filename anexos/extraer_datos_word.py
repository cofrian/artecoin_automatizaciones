#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
extraer_datos_word.py

Qué hace:
  1) Lee el Excel maestro y construye un contexto jerárquico:
     centro → edificios → dependencias / acom / envol / cc / clima / eqh / eleva / otros / ilum
  2) Resuelve fotos declaradas en la hoja “Consul” (todas las columnas FOTO_*) + disco.
     - ACOM usa: FOTO_BATERIA, FOTO_CT, FOTO_CDRO_PPAL, FOTO_CDRO_SECUND
  3) **NUEVO**: Busca fotos secuenciales automáticamente (ej: FOTO_01 → FOTO_02, FOTO_03...)
  4) **NUEVO**: Busca fotos adicionales por ID exacto de la entidad
  5) Guarda un JSON por centro (y combinado).
  6) (Opcional) tester: escribe TEST_FOTOS_<CENTRO>.txt y fotos_faltantes_por_id.json

Funcionalidades de búsqueda de fotos:
  - Fotos declaradas en Excel (columnas FOTO_*)
  - Fotos secuenciales automáticas (FOTO_F001_01, FOTO_F001_02...)
  - Búsqueda por ID exacto de la entidad
  - Fallback por carpeta usando patrones del ID
  - **RESTRICCIÓN UNIVERSAL**: Para todas las entidades sin fotos declaradas en Excel,
    solo incluye fotos cuyos nombres contengan parte del ID de la entidad.
    Ejemplo: EDIFICIO "C0007E001" → solo foto "E001_FE0001"
    Ejemplo: DEPENDENCIA "C0007E001D0001" → solo foto "D0001_FD0001"
    Ejemplo: EQUIPO "C0007E001D0001QE001" → solo foto "QE001_FQE0001"
  - **RESTRICCIÓN ACOM ESPECIAL**: Para ACOM aplica la misma regla universal.

MODOS DE EJECUCIÓN:

1) MODO INTERACTIVO CON INTERFAZ GRÁFICA (por defecto):
  py -3.13 .\extraer_datos_word.py
  - Abre exploradores de archivos para seleccionar Excel y carpetas
  - Interfaz visual fácil de usar con validaciones automáticas
  - Fallback automático a modo texto si no hay GUI disponible

2) MODO LÍNEA DE COMANDOS (avanzado):
  py -3.13 .\extraer_datos_word.py --no-interactivo --xlsx RUTA --fotos-root RUTA
  - Para automatización, scripts o usuarios expertos
  - Todos los parámetros por línea de comandos

Ejemplo interactivo:
  py -3.13 .\extraer_datos_word.py

Ejemplo línea de comandos completo:
  $xlsx = "Z:\...\ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1.xlsx"
  $root = "Z:\...\1_CONSULTA 1"
  py -3.13 .\extraer_datos_word.py `
    --no-interactivo `
    --xlsx $xlsx `
    --fotos-root $root `
    --outdir .\out_context `
    --centro C0007 `
    --fuzzy-threshold 0.88 `
    --buscar-secuenciales `
    --max-secuenciales 15 `
    --tester
"""

from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import argparse, json, re, unicodedata, difflib, math, os, sys
from datetime import datetime

import pandas as pd
import openpyxl  # noqa: F401
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False
    print("⚠️  Interfaz gráfica no disponible. Usando modo texto.")

# ==========================
# Interfaz de usuario
# ==========================
def _init_gui():
    """Inicializa la interfaz gráfica de tkinter."""
    if not GUI_AVAILABLE:
        return None
    root = tk.Tk()
    root.withdraw()  # Ocultar ventana principal
    return root

def _solicitar_excel() -> Path:
    """Solicita al usuario el archivo Excel usando interfaz gráfica o texto."""
    print("\n" + "="*60)
    print("🔍 SELECCIÓN DE ARCHIVO EXCEL")
    print("="*60)
    
    if GUI_AVAILABLE:
        return _solicitar_excel_gui()
    else:
        return _solicitar_excel_texto()

def _solicitar_excel_gui() -> Path:
    """Solicita el archivo Excel usando interfaz gráfica."""
    print("📂 Se abrirá el explorador de archivos para seleccionar el Excel...")
    
    # Inicializar GUI
    root = _init_gui()
    
    try:
        # Abrir diálogo de selección de archivo
        file_path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel de auditoría energética",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Excel 2007+", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Todos los archivos", "*.*")
            ],
            initialdir=os.getcwd()
        )
        
        if not file_path:
            print("❌ Operación cancelada por el usuario.")
            root.destroy()
            sys.exit(0)
        
        excel_path = Path(file_path)
        
        # Validar archivo
        if not excel_path.exists():
            messagebox.showerror("Error", f"El archivo no existe:\n{excel_path}")
            root.destroy()
            return _solicitar_excel_gui()  # Reintentar
            
        if not excel_path.suffix.lower() in ['.xlsx', '.xls']:
            messagebox.showerror("Error", "El archivo debe ser un Excel (.xlsx o .xls)")
            root.destroy()
            return _solicitar_excel_gui()  # Reintentar
        
        print(f"✅ Excel seleccionado: {excel_path.name}")
        print(f"📁 Ubicación: {excel_path.parent}")
        
        root.destroy()
        return excel_path
        
    except Exception as e:
        messagebox.showerror("Error", f"Error al seleccionar archivo:\n{e}")
        root.destroy()
        sys.exit(1)

def _solicitar_excel_texto() -> Path:
    """Solicita el archivo Excel usando interfaz de texto."""
    while True:
        ruta = input("Introduce la ruta completa del archivo Excel (.xlsx): ").strip()
        
        if not ruta:
            print("❌ Error: Debes introducir una ruta.")
            continue
            
        # Limpiar comillas si las tiene
        ruta = ruta.strip('"\'')
        
        try:
            excel_path = Path(ruta)
            if not excel_path.exists():
                print(f"❌ Error: El archivo no existe: {excel_path}")
                continue
                
            if not excel_path.suffix.lower() in ['.xlsx', '.xls']:
                print(f"❌ Error: El archivo debe ser .xlsx o .xls")
                continue
                
            print(f"✅ Excel encontrado: {excel_path.name}")
            return excel_path
            
        except Exception as e:
            print(f"❌ Error al procesar la ruta: {e}")
            continue

def _solicitar_carpeta_fotos() -> Path:
    """Solicita al usuario la carpeta de fotos usando interfaz gráfica o texto."""
    print("\n" + "="*60)
    print("📁 SELECCIÓN DE CARPETA DE FOTOS")
    print("="*60)
    
    if GUI_AVAILABLE:
        return _solicitar_carpeta_fotos_gui()
    else:
        return _solicitar_carpeta_fotos_texto()

def _solicitar_carpeta_fotos_gui() -> Path:
    """Solicita la carpeta de fotos usando interfaz gráfica."""
    print("📂 Se abrirá el explorador de archivos para seleccionar la carpeta de fotos...")
    
    # Inicializar GUI
    root = _init_gui()
    
    try:
        # Abrir diálogo de selección de carpeta
        folder_path = filedialog.askdirectory(
            title="Selecciona la carpeta RAÍZ que contiene las fotografías",
            initialdir=os.getcwd()
        )
        
        if not folder_path:
            print("❌ Operación cancelada por el usuario.")
            root.destroy()
            sys.exit(0)
        
        fotos_path = Path(folder_path)
        
        # Validar carpeta
        if not fotos_path.exists():
            messagebox.showerror("Error", f"La carpeta no existe:\n{fotos_path}")
            root.destroy()
            return _solicitar_carpeta_fotos_gui()  # Reintentar
            
        if not fotos_path.is_dir():
            messagebox.showerror("Error", "Debe seleccionar una carpeta, no un archivo")
            root.destroy()
            return _solicitar_carpeta_fotos_gui()  # Reintentar
        
        # Verificar contenido
        subdirs = [p for p in fotos_path.iterdir() if p.is_dir()]
        archivos_foto = [p for p in fotos_path.rglob("*") 
                        if p.is_file() and p.suffix.lower() in ALLOWED_EXTS]
        
        print(f"✅ Carpeta de fotos seleccionada: {fotos_path.name}")
        print(f"📁 Ubicación: {fotos_path}")
        print(f"📂 Subdirectorios encontrados: {len(subdirs)}")
        print(f"📸 Archivos de imagen encontrados: {len(archivos_foto)}")
        
        if len(archivos_foto) == 0:
            respuesta = messagebox.askyesno(
                "Advertencia", 
                f"No se encontraron imágenes en la carpeta seleccionada.\n\n"
                f"Carpeta: {fotos_path}\n\n"
                f"¿Deseas continuar de todas formas?"
            )
            if not respuesta:
                root.destroy()
                return _solicitar_carpeta_fotos_gui()  # Reintentar
        
        root.destroy()
        return fotos_path
        
    except Exception as e:
        messagebox.showerror("Error", f"Error al seleccionar carpeta:\n{e}")
        root.destroy()
        sys.exit(1)

def _solicitar_carpeta_fotos_texto() -> Path:
    """Solicita la carpeta de fotos usando interfaz de texto."""
    while True:
        ruta = input("Introduce la ruta de la carpeta RAÍZ de fotos: ").strip()
        
        if not ruta:
            print("❌ Error: Debes introducir una ruta.")
            continue
            
        # Limpiar comillas si las tiene
        ruta = ruta.strip('"\'')
        
        try:
            fotos_path = Path(ruta)
            if not fotos_path.exists():
                print(f"❌ Error: La carpeta no existe: {fotos_path}")
                continue
                
            if not fotos_path.is_dir():
                print(f"❌ Error: La ruta debe ser una carpeta, no un archivo")
                continue
                
            # Verificar si tiene subdirectorios (típico de estructura de fotos)
            subdirs = [p for p in fotos_path.iterdir() if p.is_dir()]
            archivos_foto = [p for p in fotos_path.rglob("*") 
                            if p.is_file() and p.suffix.lower() in ALLOWED_EXTS]
            
            print(f"✅ Carpeta de fotos encontrada con {len(subdirs)} subdirectorios")
            print(f"📸 Archivos de imagen encontrados: {len(archivos_foto)}")
            
            if len(archivos_foto) == 0:
                resp = input("⚠️  No se encontraron imágenes. ¿Continuar de todas formas? (s/n): ").strip().lower()
                if resp not in ['s', 'si', 'y', 'yes']:
                    continue
                
            return fotos_path
            
        except Exception as e:
            print(f"❌ Error al procesar la ruta: {e}")
            continue

def _solicitar_carpeta_salida(default_path: str = "./out_context") -> Path:
    """Solicita al usuario la carpeta de salida usando interfaz gráfica o texto."""
    print("\n📁 Selección de carpeta de salida...")
    
    # Preguntar si quiere usar la carpeta por defecto
    usar_default = input(f"¿Usar carpeta por defecto '{default_path}'? (s/n) [s]: ").strip().lower()
    
    if usar_default in ['', 's', 'si', 'y', 'yes']:
        return Path(default_path)
    
    if GUI_AVAILABLE:
        return _solicitar_carpeta_salida_gui(default_path)
    else:
        return _solicitar_carpeta_salida_texto(default_path)

def _solicitar_carpeta_salida_gui(default_path: str) -> Path:
    """Solicita la carpeta de salida usando interfaz gráfica."""
    print("📂 Se abrirá el explorador para seleccionar carpeta de salida...")
    
    # Inicializar GUI
    root = _init_gui()
    
    try:
        # Abrir diálogo de selección de carpeta
        folder_path = filedialog.askdirectory(
            title="Selecciona la carpeta donde guardar los archivos JSON",
            initialdir=os.getcwd()
        )
        
        if not folder_path:
            print(f"ℹ️  Usando carpeta por defecto: {default_path}")
            root.destroy()
            return Path(default_path)
        
        salida_path = Path(folder_path)
        print(f"✅ Carpeta de salida: {salida_path}")
        
        root.destroy()
        return salida_path
        
    except Exception as e:
        print(f"⚠️  Error al seleccionar carpeta: {e}")
        print(f"ℹ️  Usando carpeta por defecto: {default_path}")
        root.destroy()
        return Path(default_path)

def _solicitar_carpeta_salida_texto(default_path: str) -> Path:
    """Solicita la carpeta de salida usando interfaz de texto."""
    while True:
        salida = input(f"Ruta de carpeta de salida [{default_path}]: ").strip()
        if not salida:
            return Path(default_path)
        try:
            return Path(salida.strip('"\''))
        except Exception as e:
            print(f"❌ Error: {e}")

def _solicitar_opciones_adicionales() -> dict:
    """Solicita opciones adicionales al usuario."""
    print("\n" + "="*60)
    print("⚙️  OPCIONES ADICIONALES")
    print("="*60)
    
    opciones = {}
    
    # Centro específico
    centro = input("¿Filtrar por un centro específico? (ej: C0007, o ENTER para todos): ").strip()
    opciones['centro'] = centro if centro else None
    
    # Buscar secuenciales
    while True:
        resp = input("¿Buscar fotos secuenciales automáticamente? (s/n) [s]: ").strip().lower()
        if resp in ['', 's', 'si', 'y', 'yes']:
            opciones['buscar_secuenciales'] = True
            # Máximo secuenciales
            while True:
                try:
                    max_seq = input(f"Máximo fotos secuenciales por entidad [{MAX_SEQUENTIAL_PHOTOS}]: ").strip()
                    opciones['max_secuenciales'] = int(max_seq) if max_seq else MAX_SEQUENTIAL_PHOTOS
                    break
                except ValueError:
                    print("❌ Error: Introduce un número válido")
            break
        elif resp in ['n', 'no']:
            opciones['buscar_secuenciales'] = False
            opciones['max_secuenciales'] = MAX_SEQUENTIAL_PHOTOS
            break
        else:
            print("❌ Responde 's' para sí o 'n' para no")
    
    # Fuzzy threshold
    while True:
        try:
            fuzzy = input(f"Similitud para matching de fotos (0.0-1.0) [{FUZZY_THRESHOLD_DEFAULT}]: ").strip()
            opciones['fuzzy_threshold'] = float(fuzzy) if fuzzy else FUZZY_THRESHOLD_DEFAULT
            if 0.0 <= opciones['fuzzy_threshold'] <= 1.0:
                break
            else:
                print("❌ Error: El valor debe estar entre 0.0 y 1.0")
        except ValueError:
            print("❌ Error: Introduce un número válido")
    
    # Tester
    while True:
        resp = input("¿Generar archivos de testing y logs? (s/n) [s]: ").strip().lower()
        if resp in ['', 's', 'si', 'y', 'yes']:
            opciones['tester'] = True
            break
        elif resp in ['n', 'no']:
            opciones['tester'] = False
            break
        else:
            print("❌ Responde 's' para sí o 'n' para no")
    
    # JSONs separados
    while True:
        resp = input("¿Generar JSONs separados por tipo? (carpeta con centro.json, acom.json, etc.) (s/n) [n]: ").strip().lower()
        if resp in ['s', 'si', 'y', 'yes']:
            opciones['jsons_separados'] = True
            break
        elif resp in ['', 'n', 'no']:
            opciones['jsons_separados'] = False
            break
        else:
            print("❌ Responde 's' para sí o 'n' para no")
    
    # Carpeta de salida (ahora con interfaz gráfica opcional)
    opciones['outdir'] = _solicitar_carpeta_salida()
    
    return opciones

def _mostrar_resumen(excel_path: Path, fotos_path: Path, opciones: dict):
    """Muestra un resumen de la configuración antes de ejecutar."""
    print("\n" + "="*60)
    print("📋 RESUMEN DE CONFIGURACIÓN")
    print("="*60)
    print(f"📊 Excel:           {excel_path}")
    print(f"📁 Fotos:           {fotos_path}")
    print(f"🎯 Centro:          {opciones.get('centro', 'Todos')}")
    print(f"📸 Secuenciales:    {'Sí' if opciones.get('buscar_secuenciales') else 'No'}")
    if opciones.get('buscar_secuenciales'):
        print(f"🔢 Máx. secuenc.:   {opciones.get('max_secuenciales')}")
    print(f"🎚️  Fuzzy thresh.:   {opciones.get('fuzzy_threshold')}")
    print(f"🧪 Testing:         {'Sí' if opciones.get('tester') else 'No'}")
    print(f"� JSONs separados: {'Sí' if opciones.get('jsons_separados') else 'No'}")
    print(f"�💾 Salida:          {opciones.get('outdir')}")
    print("="*60)
    
    while True:
        resp = input("¿Continuar con esta configuración? (s/n) [s]: ").strip().lower()
        if resp in ['', 's', 'si', 'y', 'yes']:
            return True
        elif resp in ['n', 'no']:
            return False
        else:
            print("❌ Responde 's' para sí o 'n' para no")

# ==========================
# Config base y utilidades
# ==========================
NORMALIZE_EMPTY_TO = "–"
ALLOWED_EXTS = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp", ".bmp", ".gif"}
FUZZY_THRESHOLD_DEFAULT = 0.88  # 1.0 = solo exacto; <1.0 permite aproximado
MAX_SEQUENTIAL_PHOTOS = 10  # Máximo número de fotos secuenciales por entidad

SHEETS = {
    "CENT":   ["CENT", "Centro", "CENTRO"],
    "EDIF":   ["EDIF", "Edif", "EDIFICIO", "Edificio"],
    "DEPEN":  ["DEPEN", "Dependencia", "DEPENDENCIA"],
    "ACOM":   ["DAT_ELEC_EDIF", "ACOM", "DAT ELEC EDIF"],
    "ENVOL":  ["Envol", "CERRAMIENTOS", "ENVOL"],
    "SISTCC": ["SistCC", "SISTCC", "EQ GEN", "EQGEN"],
    "CLIMA":  ["Clima", "CLIMA"],
    "EQHORIZ":["EqHoriz", "EQHORIZ", "EQ HORIZ", "Horizontales"],
    "ELEVA":  ["Eleva", "ELEVA", "Elevadores"],
    "OTROSEQ":["OtrosEq", "OTROSEQ", "Otros Equipos"],
    "ILUM":   ["Ilum", "ILUM", "Iluminacion", "ILUMINACIÓN", "ILUMINACION"],
    "CONSUL": ["Consul", "CONSUL", "Consulta", "CONSULTA"],
}

# ==== MAPEOS  ====
MAP_CENTRO = {
    "id":              "ID CENTRO",
    "tipo":            "TIPO CENTRO",
    "nombre":          "CENTRO",
    "dias_apertura":   "DÍAS / SEMANA",
    "direccion":       "DIRECCIÓN",
    "hora_ap_m":       "APERTURA MAÑANA",
    "hora_ci_m":       "CIERRE MAÑANA",
    "hora_ap_t":       "APERTURA TARDE",
    "hora_ci_t":       "CIERRE TARDE",
    "hora_ap_n":       "APERTURA NOCHE",
    "hora_ci_n":       "CIERRE NOCHE",
    "num_edificios":   "Nº EDIFICIOS",
    "reformas":        "REFORMAS_AMBITO",
    "observaciones":   "OBSERVACIONES",
    "contacto":        "",
    "tecnico":         "",
    "foto":            "",
}
MAP_EDIF = {
    "id_centro":         "ID CENTRO",
    "id":                "ID EDIFICIO",
    "bloque":            "BLOQUE",
    "centros_propuestos":"EDIFICIOS PROPUESTOS",
    "tipo_centro":       "TIPO EDIFICIO",
    "centro":            "CENTRO",
    "nombre":            "EDIFICIO",
    "coord_x":           "COORD_X",
    "coord_y":           "COORD_Y",
    "num_plantas":       "NUMERO_PLANTAS",
    "altura_planta":     "ALTURA_PLANTA",
    "observaciones":     "OBSERVACIONES",
    "n_edificios":       "Nº EDIFICIOS",
    "superficie":        "SUPERFICIE",
}
MAP_DEPEN = {
    "id_centro":        "ID CENTRO",
    "id_edificio":      "ID EDIFICIO",
    "id":               "ID DEPENDENCIA",
    "bloque":           "BLOQUE",
    "tipo_edificio":    "TIPO EDIFICIO",
    "centro":           "CENTRO",
    "edificio":         "EDIFICIO",
    "nombre":           "DEPENDENCIA",
    "n_dependencias":   "N_DEPENDENCIAS",
    "planta":           "PLANTA_DEPENDENCIA",
    "superficie":       "SUPERFICIE",
    "observaciones":    "OBSERVACIONES",
    "zona_construida":  "ZONA. CONSTRUIDA",
    "zona_uso_principal":"ZONA. DE USO PRINCIPAL",
    "zona_almacenes":   "ZONA. ALMACENES Y SALAS TÉCNICAS",
    "zona_calefactada": "ZONA. CALEFACTADA",
    "zona_refrigerada": "ZONA. REFRIGERADA",
    "zona_ventilada":   "ZONA. VENTILADA",
    "zona_iluminada":   "ZONA. ILUMINADA",
    "construida_m2":    "CONSTRUIDA (m2)",
    "uso_principal_m2": "DE USO PRINCIPAL (m2)",
    "almacenes_m2":     "ALMACENES Y SALAS TÉCNICAS (m2)",
    "calefactada_m2":   "CALEFACTADA (m2)",
    "refrigerada_m2":   "REFRIGERADA (m2)",
    "ventilada_m2":     "VENTILADA (m2)",
    "iluminada_m2":     "ILUMINADA (m2)",
}
MAP_ACOM = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "bloque": "BLOQUE",
    "tipo_edificio": "TIPO EDIFICIO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "ct": "CT",
    "ct_num_transformadores": "CT_NUM_TRANSFORMADORES",
    "ct_potencia": "CT_POTENCIA",
    "ct_tension_primario": "CT_TENSION PRIMARIO",
    "ct_tension_secundario": "CT_TENSION SECUNDARIO",
    "ct_aislante": "CT_AISLANTE",
    "ct_marca": "CT_MARCA",
    "ct_modelo": "CT_MODELO",
    "ct_estado": "CT_ESTADO",
    "ct_observaciones": "CT_OBSERVACIONES",
    "cdro_ppal_contador": "CDRO_ELECT_PPAL_CONTADOR",
    "cdro_ppal_contador_sec": "CDRO_ELECT_PPAL_CONTADOR_SECUND",
    "cdro_ppal_estado": "CDRO_ELECT_PPAL_ESTADO",
    "cdro_ppal_pinzas": "CDRO_ELECT_PPAL_MEDIDA_PINZAS",
    "bat_condensadores": "BAT_CONDENSADORES",
    "num_condensadores": "NUM_CONDENSADORES",
    "bat_condensadores_pot": "BAT_CONDENSADORES_POT",
    "bat_condensadores_marca": "BAT_CONDENSADORES_MARCA",
    "bat_condensadores_modelo": "BAT_CONDENSADORES_MODELO",
    "cdro_sec_estado": "CDRO_ELECT_SECUND_ESTADO",
    "cdro_sec_observaciones": "CDRO_ELECT_SECUND_OBSERVACIONES",
}
MAP_ENVOL = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id": "ID",
    "bloque": "BLOQUE",
    "tipo_edificio": "TIPO EDIFICIO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "denominacion": "DENOMINACIÓN",
    "tipo_envolvente": "TIPO ENVOLVENTE",
    "orientacion": "ORIENTACIÓN",
    "fachada_tipo": "FACHADA_TIPO",
    "fachada_tipo_obs": "FACHADA_TIPO_OBSERVACIONES",
    "fachada_aislamiento": "FACHADA_AISLAMIENTO",
    "fachada_aislamiento_obs": "FACHADAS_AISLAMIENTO_OBSERVACIONES",
    "fachada_camara_aire": "FACHADA_CAMARA_AIRE",
    "fachada_huecos": "FACHADA_HUECOS",
    "fachada_huecos_obs": "FACHADA_HUECOS_OBSERVACIONES",
    "puertas_tipo": "PUERTAS_TIPO",
    "puertas_material": "PUERTAS_MATERIAL",
    "puertas_obs": "PUERTAS_OBSERVACIONES",
    "num_puertas": "N PUER",
    "puertas_dimensiones": "PUERTAS_DIMENSIONES",
    "sup_puertas": "SUP PUER",
    "ventanas_tipo": "VENTANAS_TIPO",
    "ventanas_carpinteria": "VENTANAS_CARPINTERIA",
    "ventanas_acristalamiento": "VENTANAS_ACRISTALAMIENTO",
    "ventanas_proteccion_solar": "VENTANAS_PROTECCION_SOLAR",
    "num_ventanas": "N VENT",
    "ventanas_dimensiones": "VENTANAS_DIMENSIONES",
    "sup_vent_unitaria": "SUP VENT UNITARIA",
    "sup_vent": "SUP VENT",
    "cubiertas_tipo": "CUBIERTAS_TIPO",
    "cubiertas_tipo_obs": "CUBIERTAS_TIPO_OBSERVACIONES",
    "cubiertas_acabado": "CUBIERTAS_ACABADO",
    "cubiertas_aislamiento": "CUBIERTAS_AISLAMIENTO",
    "cubiertas_aislamiento_obs": "CUBIERTAS_AISLAMIENTO_OBSERVACIONES",
    "lucernario": "CUBIERTAS_LUCERNARIO",
    "lucernario_dimensiones": "CUBIERTAS_LUCERNARIO_DIMENSIONES",
    "lucernario_unidades": "CUBIERTAS_LUCERNARIO_NUMERO_UNIDADES",
    "sup_cubierta": "SUP CUBIER",
}
MAP_SISTCC = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id_dependencia": "ID DEPENDENCIA",
    "id": "ID",
    "bloque": "BLOQUE",
    "subbloque": "SUBBLOQUE",
    "centros_propuestos": "CENTROS PROPUESTOS",
    "tipo_edificio": "TIPO EDIFICIO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "situacion": "SITUACIÓN",
    "dependencias": "DEPENDENCIAS",
    "denominacion": "DENOMINACIÓN",
    "tipo_equipo": "TIPO DE EQUIPO",
    "tipo_combustible": "TIPO DE COMBUSTIBLE",
    "marca": "MARCA",
    "modelo": "MODELO",
    "periodo_funcionamiento": "PERIODO FUNCIONAMIENTO",
    "servicio_calefaccion": "Servicio a calefacción",
    "servicio_acs": "Servicio a producción ACS",
    "rendimiento": "Rendimiento (%)",
    "estado": "ESTADO",
    "observaciones": "OBSERVACIONES",
}
MAP_CLIMA = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id_dependencia": "ID DEPENDENCIA",
    "id": "ID",
    "bloque": "BLOQUE",
    "tipo_edificio": "TIPO EDIFICIO",
    "tipo_edificio_geo": "TIPO ED - GEO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "situacion": "SITUACIÓN",
    "dependencias": "DEPENDENCIAS",
    "denominacion": "DENOMINACIÓN",
    "tipo_climatizacion": "TIPO DE CLIMATIZACIÓN",
    "tipo_terminal": "TIPO DE TERMINAL",
    "sistema_control": "SISTEMA CONTROL REGULACIÓN",
    "fabricante": "FABRICANTE",
    "modelo": "MODELO",
    "num_elementos_radiador": "NUMERO_ELEMENTOS_RADIADOR",
    "alimentacion": "ALIMENTACION",
    "fluido_calorportador": "FLUIDO CALOPORTADOR",
    "pot_frio": "POT_FRIGORIFICA_TERMICA_W",
    "pot_calor": "POT_CALORIFICA_TERMICA_W",
    "pot_abs_frio": "POT_ABS_FRIO_W",
    "pot_abs_calor": "POT_ABS_CALOR_W",
    "observaciones": "OBSERVACIONES",
}
MAP_EQH = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id_dependencia": "ID DEPENDENCIA",
    "id": "ID",
    "bloque": "BLOQUE",
    "centros_propuestos": "CENTROS PROPUESTOS",
    "tipo_edificio": "TIPO EDIFICIO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "dependencias": "DEPENDENCIAS",
    "denominacion": "DENOMINACIÓN",
    "denominacion_corregida": "DENOMINACIÓN CORREGIDA",
    "tipo_equipo": "TIPO DE EQUIPO",
    "voltaje": "Voltaje (V)",
    "observaciones": "observaciones",
    "potencia_nominal": "POTENCIA NOMINAL (kW)",
    "num_equipos": "Nº EQUIPOS",
}
MAP_ELEVA = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id_dependencia": "ID DEPENDENCIA",
    "id": "ID",
    "bloque": "BLOQUE",
    "centros_propuestos": "CENTROS PROPUESTOS",
    "tipo_edificio": "TIPO EDIFICIO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "dependencias": "DEPENDENCIAS",
    "denominacion": "DENOMINACIÓN",
    "tipo_equipo": "TIPO DE EQUIPO",
    "marca": "MARCA",
    "modelo": "MODELO",
    "plazas": "PLAZAS",
    "carga": "CARGA (kg)",
    "estado": "ESTADO",
    "observaciones": "observaciones",
    "voltaje": "Voltaje (V)",
    "potencia_nominal": "POTENCIA NOMINAL (kW)",
    "num_equipos": "Nº EQUIPOS",
}
MAP_OTRO = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id_dependencia": "ID DEPENDENCIA",
    "id": "ID",
    "bloque": "BLOQUE",
    "centros_propuestos": "CENTROS PROPUESTOS",
    "tipo_edificio": "TIPO EDIFICIO",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "dependencias": "DEPENDENCIAS",
    "denominacion": "DENOMINACIÓN",
    "tipo_equipo": "TIPO DE EQUIPO",
    "suministro": "SUMINISTRO",
    "tipo_impulsion": "TIPO IMPULSION",
    "tipo_bomba": "TIPO BOMBA",
    "marca": "MARCA",
    "modelo": "MODELO",
    "sistema_regulacion": "SISTEMA REGULACIÓN",
    "estado": "estado",
    "observaciones": "observaciones",
    "voltaje": "Voltaje (V)",
    "potencia_nominal": "POTENCIA NOMINAL (kW)",
    "num_equipos": "Nº EQUIPOS",
}
MAP_ILUM = {
    "id_centro": "ID CENTRO",
    "id_edificio": "ID EDIFICIO",
    "id_dependencia": "ID DEPENDENCIA",
    "id": "ID",
    "bloque": "BLOQUE",
    "centros_propuestos": "CENTROS PROPUESTOS",
    "tipo_edificio": "TIPO EDIFICIO",
    "tipo_ed_geo": "TIPO ED - GEO - ILUM",
    "centro": "CENTRO",
    "edificio": "EDIFICIO",
    "dependencias": "DEPENDENCIAS",
    "denominacion": "DENOMINACIÓN",
    "situacion_acortada": "SITUACIÓN ACORTADA",
    "situacion": "SITUACIÓN",
    "altura": "ALTURA",
    "tipo_soporte": "TIPO SOPORTE",
    "situacion_soporte": "SITUACIÓN SOPORTE",
    "tipo_luminaria": "tipo luminaria",
    "tipo_luminaria_corregida": "TIPO LUMINARIA CORREGIDA",
    "tipo_lampara": "tipo lámpara",
    "tipo_lampara_corregida": "TIPO LÁMPARA CORREGIDA",
    "equipo_auxiliar": "EQUIPO AUXILIAR",
    "regulacion": "REGULACIÓN",
    "pantalla_reflectante": "PANTALLA_REFLECTANTE",
    "nivel_iluminacion_lux": "NIVEL_ILUMINACION_MEDIO_LUX",
    "observaciones": "observaciones",
    "potencia_nominal": "POTENCIA NOMINAL (W)",
    "num_luminarias": "n_luminarias",
    "num_lamparas": "n_lamparas",
}

PHOTO_ID_FIELD = {
    "CENTRO":  "id",
    "EDIFICIO":"id",
    "DEPENDENCIA":"id",
    "ACOM":    "id_edificio",  # especial
    "ENVOL":   "id",
    "SISTCC":  "id",
    "CLIMA":   "id",
    "EQHORIZ": "id",
    "ELEVA":   "id",
    "OTROSEQ": "id",
    "ILUM":    "id",
}

# ==== Consul: bloques y columnas de fotos por tipo ====
_CONSUL_SPEC = [
    ("CENTRO",        "ID_CENTRO",        "CENT"),
    ("EDIFICIO",      "ID_EDIFICIO2",     "EDIF"),
    ("DEPENDENCIA",   "ID_DEPENDENCIA2",  "DEPEN"),
    ("CERRAMIENTOS",  "ID_CERRAMIENTO2",  "ENVOL"),
    ("EQ GEN",        "ID_EQGEN2",        "SISTCC"),
    ("EQ CLIMA EXT",  "ID_EQCLIMAEXT2",   "CLIMA"),
    ("EQ CLIMA INT",  "ID_EQCLIMAINT2",   "CLIMA"),
    ("EQ HORIZ",      "ID_EQHORIZ2",      "EQHORIZ"),
    ("EQ ELEV",       "ID_EQELEV2",       "ELEVA"),
    ("OTROS EQUIPOS", "ID_EQOTRO2",       "OTROSEQ"),
    ("EQ GRUPBOM",    "ID_GRUPBOM2",      "OTROSEQ"),
    ("ILUM",          "ID_ILUM2",         "ILUM"),
    ("DAT_ELEC_EDIF", "ID_EDIFICIO2",     "ACOM"),
]
# prioridad para ACOM
ACOM_FOTOS = {"FOTO_BATERIA","FOTO_CT","FOTO_CDRO_PPAL","FOTO_CDRO_SECUND"}

# --------------------------
# Limpieza y lectura Excel
# --------------------------
def _clean(v):
    if v is None:
        return NORMALIZE_EMPTY_TO
    try:
        if pd.isna(v) or (isinstance(v, float) and math.isnan(v)):
            return NORMALIZE_EMPTY_TO
    except Exception:
        pass
    s = str(v).strip()
    return NORMALIZE_EMPTY_TO if s == "" or s.lower() == "nan" else s

def _read_sheet(xls: pd.ExcelFile, names: List[str]) -> Optional[pd.DataFrame]:
    for nm in xls.sheet_names:
        for cand in names:
            if nm.strip().lower() == cand.strip().lower():
                df = xls.parse(nm, dtype=str, engine="openpyxl").fillna("")
                df.columns = [str(c).strip() for c in df.columns]
                for c in df.columns:
                    df[c] = df[c].apply(_clean)
                return df
    return None

def _read_all_sheets(xlsx: Path) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(xlsx, engine="openpyxl")
    return {k: _read_sheet(xls, v) for k, v in SHEETS.items()}

def _rename_row(sr: pd.Series, mapping: Dict[str, str]) -> Dict[str, str]:
    return {k: _clean(sr.get(col)) for k, col in mapping.items()}

def _safe_name(s: str) -> str:
    return re.sub(r"[^\w\-_.]+", "_", str(s).strip()) if s else "SIN_NOMBRE"

# --------------------------
# Optimizaciones y cachés
# --------------------------
from functools import lru_cache
from collections import defaultdict
import time

# Cache para normalización de slugs
_slug_cache = {}
# Cache para índices de archivos por carpeta
_file_index_cache = {}

@lru_cache(maxsize=10000)
def _norm_slug_cached(s: str) -> str:
    """Versión optimizada con caché de _norm_slug."""
    if not s:
        return ""
    
    s = _strip_ext(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper().strip()
    s = re.sub(r"[\s_.\-]+", "", s)
    return s

@lru_cache(maxsize=10000)
def _normalize_filename(filename: str) -> str:
    """Normaliza un nombre de archivo para comparaciones optimizadas."""
    return _norm_slug_cached(filename)

def _build_optimized_photo_index(root: Path) -> tuple[dict, dict, dict]:
    """
    Construye índices optimizados para búsqueda rápida de fotos.
    Returns: (exact_index, normalized_index, path_to_info)
    """
    if not root or not root.exists():
        return {}, {}, {}
    
    # Usar caché si ya existe para esta carpeta
    cache_key = str(root.resolve())
    if cache_key in _file_index_cache:
        return _file_index_cache[cache_key]
    
    print(f"🔍 Indexando fotos en: {root.name}...")
    start_time = time.time()
    
    exact_index = {}  # {nombre_exacto: Path}
    normalized_index = defaultdict(list)  # {slug_normalizado: [Path, ...]}
    path_to_info = {}  # {Path: {"stem": str, "normalized": str}}
    
    # Una sola pasada por todos los archivos
    foto_count = 0
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in ALLOWED_EXTS:
            stem = p.stem
            normalized = _norm_slug_cached(stem)
            
            # Índice exacto
            exact_index[stem] = p
            exact_index[stem.upper()] = p
            
            # Índice normalizado
            if normalized:
                normalized_index[normalized].append(p)
            
            # Información del archivo
            path_to_info[p] = {
                "stem": stem,
                "normalized": normalized,
                "realpath": os.path.realpath(str(p))
            }
            
            foto_count += 1
    
    # Convertir defaultdict a dict normal
    normalized_index = dict(normalized_index)
    
    result = (exact_index, normalized_index, path_to_info)
    _file_index_cache[cache_key] = result
    
    elapsed = time.time() - start_time
    print(f"✅ Índice creado: {foto_count} fotos en {elapsed:.2f}s")
    
    return result
# --------------------------
# Funciones básicas y utilitarias
# --------------------------
def _strip_ext(s: str) -> str:
    """Quita la extensión de un archivo."""
    return re.sub(r"\.(jpg|jpeg|png|heic|heif|webp|bmp|gif)$", "", str(s), flags=re.I)

def _norm_slug(s: str) -> str:
    """Versión backward-compatible de normalización (usa caché internamente)."""
    return _norm_slug_cached(s)

def _leaf_id(ident: str) -> str:
    """Extrae el ID principal de un identificador."""
    s = str(ident or "").upper().strip()
    m = re.search(r'([A-Z]{1,3}0*\d{1,5})$', s)
    return m.group(1) if m else s

# Funciones de compatibilidad (wrappers para las optimizadas)
def _resolve_name_to_path(name_from_excel: str, fotos_index: dict, fuzzy_threshold: float) -> Optional[Path]:
    """Wrapper de compatibilidad para la versión optimizada."""
    # Si fotos_index es el formato antiguo (dict simple), convertir
    if fotos_index and isinstance(list(fotos_index.values())[0], Path):
        # Es formato antiguo, necesita conversión
        exact_index = fotos_index
        normalized_index = {}
        for stem, path in fotos_index.items():
            norm = _norm_slug_cached(stem)
            if norm not in normalized_index:
                normalized_index[norm] = []
            normalized_index[norm].append(path)
        return _resolve_name_to_path_optimized(name_from_excel, exact_index, normalized_index, fuzzy_threshold)
    
    # Es formato nuevo (tupla de índices)
    if isinstance(fotos_index, tuple) and len(fotos_index) == 3:
        exact_index, normalized_index, _ = fotos_index
        return _resolve_name_to_path_optimized(name_from_excel, exact_index, normalized_index, fuzzy_threshold)
    
    return None

def _list_files_index(root: Path) -> Dict[str, Path]:
    """
    Función de compatibilidad que devuelve el diccionario de fotos 
    para funciones que aún lo requieren (como _tester_init).
    """
    if not root or not root.exists():
        return {}
    
    file_index = {}
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif']:
            stem = path.stem
            file_index[stem] = path
    
    return file_index

def _list_files_index_legacy(root: Path) -> Dict[str, Path]:
    """Alias para retrocompatibilidad."""
    return _list_files_index(root)

# --------------------------
# Búsqueda optimizada de fotos
# --------------------------
def _resolve_name_to_path_optimized(name_from_excel: str, exact_index: dict, 
                                  normalized_index: dict, fuzzy_threshold: float) -> Optional[Path]:
    """Versión optimizada de resolución de nombres a rutas."""
    if not name_from_excel:
        return None
    
    stem = Path(name_from_excel).stem
    
    # 1. Búsqueda exacta (más rápida)
    if stem in exact_index:
        return exact_index[stem]
    if stem.upper() in exact_index:
        return exact_index[stem.upper()]
    
    # 2. Búsqueda normalizada
    normalized_target = _norm_slug_cached(stem)
    if normalized_target in normalized_index:
        return normalized_index[normalized_target][0]  # Tomar el primero
    
    # 2.5. Búsqueda inteligente por patrones (para casos como C001_FC0001)
    # Buscar variantes del stem sin prefijos como C001_
    stem_variants = [stem]
    
    # Remover prefijos como "C001_", "C0007_", etc.
    if '_' in stem:
        parts = stem.split('_', 1)
        if len(parts) == 2 and re.match(r'^[A-Z]\d+$', parts[0]):
            stem_variants.append(parts[1])
    
    # Buscar cada variante
    for variant in stem_variants:
        if variant != stem:  # Evitar duplicados
            normalized_variant = _norm_slug_cached(variant)
            if normalized_variant in normalized_index:
                return normalized_index[normalized_variant][0]
    
    # 3. Fuzzy matching solo si es necesario
    if fuzzy_threshold < 1.0 and normalized_index:
        best_score = 0.0
        best_path = None
        
        for norm_key, paths in normalized_index.items():
            score = difflib.SequenceMatcher(None, normalized_target, norm_key).ratio()
            if score > best_score and score >= fuzzy_threshold:
                best_score = score
                best_path = paths[0]
        
        return best_path
    
    return None

def _buscar_fotos_secuenciales_optimized(foto_base: str, normalized_index: dict, 
                                        max_photos: int = MAX_SEQUENTIAL_PHOTOS) -> List[Path]:
    """Versión optimizada de búsqueda secuencial."""
    rutas_extra = []
    
    # Detectar patrón secuencial - mejorado para capturar el formato completo
    match = re.search(r"(.*_F\w*?)(\d+)$", foto_base, re.IGNORECASE)
    if not match:
        return rutas_extra
    
    prefijo = match.group(1)
    num_str_original = match.group(2)
    num_inicial = int(num_str_original)
    prefijo_norm = _norm_slug_cached(prefijo)
    
    # Detectar el formato original de dígitos
    formato_original = len(num_str_original)
    
    # Buscar secuenciales de manera eficiente
    for num in range(num_inicial + 1, num_inicial + max_photos + 1):
        # Probar diferentes formatos, priorizando el formato original
        formatos = []
        if formato_original >= 4:
            formatos.append(f"{num:04d}")
        formatos.extend([f"{num:0{formato_original}d}", f"{num:02d}", f"{num:03d}", str(num)])
        
        for formato in formatos:
            candidato_norm = prefijo_norm + formato
            
            if candidato_norm in normalized_index:
                rutas_extra.extend(normalized_index[candidato_norm])
                break  # Solo uno por número
        
        # Si no encuentra este número, asumir que no hay más
        else:
            break
    
    return rutas_extra

def _buscar_fotos_por_id_optimized(ident: str, normalized_index: dict, path_to_info: dict) -> List[Path]:
    """Versión optimizada de búsqueda por ID."""
    fotos_encontradas = []
    ident_norm = _norm_slug_cached(ident)
    leaf_id_norm = _norm_slug_cached(_leaf_id(ident))
    
    seen_realpaths = set()
    
    # Búsqueda directa en índice normalizado
    for norm_key, paths in normalized_index.items():
        if ident_norm in norm_key or leaf_id_norm in norm_key:
            for p in paths:
                realpath = path_to_info[p]["realpath"]
                if realpath not in seen_realpaths:
                    seen_realpaths.add(realpath)
                    fotos_encontradas.append(p)
    
    return fotos_encontradas

def _fallback_candidates_optimized(ident: str, exact_index: dict, normalized_index: dict, 
                                 path_to_info: dict, max_photos: int = 6, tipo: str = "") -> List[Path]:
    """
    Versión optimizada del fallback con filtrado ESTRICTO por tipo de entidad.
    - CENTRO: solo fotos con C en el nombre
    - EDIFICIO: solo fotos con E en el nombre (sin D ni Q)
    - DEPENDENCIA: solo fotos con D y el ID específico
    - EQUIPOS: solo fotos del ID específico del equipo
    """
    
    # Lógica especial para centros: capturar fotos que empiecen con C
    if tipo == "CENTRO":
        fotos_centro = []
        seen_realpaths = set()
        
        # Buscar todas las fotos que empiecen con C (como C001_FC0001.jpg)
        for norm_key, paths in normalized_index.items():
            for p in paths:
                stem = path_to_info[p]["stem"]
                # Si el nombre del archivo empieza con C seguido de números
                if re.match(r'^C\d+', stem, re.IGNORECASE):
                    realpath = path_to_info[p]["realpath"]
                    if realpath not in seen_realpaths:
                        seen_realpaths.add(realpath)
                        fotos_centro.append(p)
                        if len(fotos_centro) >= max_photos:
                            break
            if len(fotos_centro) >= max_photos:
                break
        
        if fotos_centro:
            return fotos_centro
    
    # Lógica especial para ACOM: solo buscar fotos específicas de acometida
    if tipo == "ACOM":
        fotos_acom = []
        seen_realpaths = set()
        
        # Patrones MÁS ESPECÍFICOS para acometidas - solo estas fotos
        acom_patterns = [
            r'.*acom.*',           # Contiene "acom"
            r'.*acometida.*',      # Contiene "acometida"
            r'.*bateria.*',        # Contiene "bateria"  
            r'.*ct\b.*',           # Contiene "ct" (centro transformación) como palabra completa
            r'.*centro.*transf.*', # Centro de transformación
            r'.*cdro.*ppal.*',     # Cuadro principal
            r'.*cdro.*secund.*',   # Cuadro secundario
            r'.*cuadro.*principal.*', # Cuadro principal
            r'.*cuadro.*secundario.*', # Cuadro secundario
            r'.*cgbt.*',           # Centro general de baja tensión
            r'.*suministro.*',     # Suministro eléctrico
            r'.*contador.*',       # Contador eléctrico
            r'.*medida.*',         # Medida eléctrica
        ]
        
        # FILTRADO ESTRICTO: solo fotos que coincidan EXACTAMENTE con patrones de acometida
        for norm_key, paths in normalized_index.items():
            for p in paths:
                stem_lower = path_to_info[p]["stem"].lower()
                
                # Verificar que coincide con al menos un patrón específico de acometida
                coincide_acom = any(re.search(pattern, stem_lower, re.IGNORECASE) for pattern in acom_patterns)
                
                if coincide_acom:
                    # Verificación adicional: NO debe contener patrones de otros equipos
                    patrones_excluidos = [
                        r'.*clima.*', r'.*calef.*', r'.*ilum.*', r'.*lamp.*',
                        r'.*ascens.*', r'.*elevador.*', r'.*bomba.*(?!.*contador)', 
                        r'.*ventil.*', r'.*radiador.*', r'.*termo.*'
                    ]
                    
                    # Si NO coincide con patrones excluidos, es válida
                    if not any(re.search(excl_pattern, stem_lower, re.IGNORECASE) for excl_pattern in patrones_excluidos):
                        realpath = path_to_info[p]["realpath"]
                        if realpath not in seen_realpaths:
                            seen_realpaths.add(realpath)
                            fotos_acom.append(p)
                            if len(fotos_acom) >= max_photos:
                                break
            if len(fotos_acom) >= max_photos:
                break
        
        return fotos_acom  # Siempre devolver las fotos encontradas, aunque sean pocas
    
    # FILTRADO ESTRICTO SEGÚN TIPO DE ENTIDAD
    ident_upper = ident.upper()
    fotos_filtradas = []
    seen_realpaths = set()
    
    for norm_key, paths in normalized_index.items():
        if len(fotos_filtradas) >= max_photos:
            break
            
        for p in paths:
            if len(fotos_filtradas) >= max_photos:
                break
                
            stem = path_to_info[p]["stem"]
            stem_upper = stem.upper()
            
            # FILTRADO ESTRICTO POR TIPO
            es_valida = False
            
            if tipo == "EDIFICIO":
                # EDIFICIOS: solo fotos cuyo nombre esté contenido en el ID de la entidad
                if ('E' in stem_upper and 'D' not in stem_upper and 'Q' not in stem_upper):
                    # Extraer la parte del nombre de la foto (ej: E001_FE0001 -> E001)
                    foto_parts = stem_upper.split('_')
                    foto_id_part = foto_parts[0] if foto_parts else stem_upper
                    
                    # Verificar que la parte del nombre de foto está contenida en el ID de la entidad
                    # Ejemplo: foto "E001_FE0001" -> parte "E001" debe estar en ID "C0007E001"
                    if foto_id_part in ident_upper:
                        es_valida = True
            
            elif tipo == "DEPENDENCIA":
                # DEPENDENCIAS: solo fotos cuyo nombre esté contenido en el ID de la entidad
                if 'D' in stem_upper:
                    # Extraer la parte del nombre de la foto (ej: D0006_FD0001 -> D0006)
                    foto_parts = stem_upper.split('_')
                    foto_id_part = foto_parts[0] if foto_parts else stem_upper
                    
                    # Verificar que la parte del nombre de foto está contenida en el ID de la entidad
                    # Ejemplo: foto "D0001_FD0001" -> parte "D0001" debe estar en ID "C0007E001D0001"
                    if foto_id_part in ident_upper:
                        # Verificar que NO tiene Q (no es un equipo)
                        if 'Q' not in stem_upper:
                            es_valida = True
            
            elif tipo in ["CLIMA", "EQHORIZ", "ELEVA", "OTROSEQ", "ILUM", "ENVOL", "SISTCC"]:
                # EQUIPOS: solo fotos cuyo nombre esté contenido en el ID de la entidad
                if 'Q' in stem_upper:
                    # Extraer la parte del nombre de la foto (ej: QE001_FQE0001 -> QE001)
                    foto_parts = stem_upper.split('_')
                    foto_id_part = foto_parts[0] if foto_parts else stem_upper
                    
                    # Verificar que la parte del nombre de foto está contenida en el ID de la entidad
                    # Ejemplo: foto "QE001_FQE0001" -> parte "QE001" debe estar en ID "C0007E001D0001QE001"
                    if foto_id_part in ident_upper:
                        es_valida = True
            
            else:
                # OTROS: lógica por defecto
                if ident_upper in stem_upper:
                    es_valida = True
            
            # Si la foto es válida, agregarla
            if es_valida:
                realpath = path_to_info[p]["realpath"]
                if realpath not in seen_realpaths:
                    seen_realpaths.add(realpath)
                    fotos_filtradas.append(p)
    
    return fotos_filtradas[:max_photos]
    
    for p in todas_fotos:
        realpath = path_to_info[p]["realpath"]
        if realpath not in seen_realpaths:
            seen_realpaths.add(realpath)
            fotos_unicas.append(p)
            if len(fotos_unicas) >= max_photos:
                break
    
    # Si aún necesitamos más fotos, buscar por prefijo MÁS ESPECÍFICO
    if len(fotos_unicas) < max_photos:
        leaf_norm = _norm_slug_cached(_leaf_id(ident))
        
        for norm_key, paths in normalized_index.items():
            if len(fotos_unicas) >= max_photos:
                break
            
            # CAMBIO: Buscar más específicamente - no solo que empiece con el prefijo
            # sino que sea más específico para evitar matches incorrectos
            if norm_key.startswith(leaf_norm):
                # Verificación adicional MÁS ESTRICTA: el ID debe ser exacto o muy similar
                for p in paths:
                    if len(fotos_unicas) >= max_photos:
                        break
                    
                    stem = path_to_info[p]["stem"]
                    ident_upper = ident.upper()
                    stem_upper = stem.upper()
                    
                    # NUEVA LÓGICA MÁS ESTRICTA POR TIPO DE ENTIDAD:
                    # Buscar por la parte relevante del ID en lugar del ID completo
                    # EDIFICIO C0007E001 -> buscar E001
                    # DEPENDENCIA C0007E001D0001 -> buscar D0001  
                    # EQUIPO C0007E001D0001QE001 -> buscar QE001
                    
                    # Determinar tipo de entidad basado en el ID
                    is_edificio = 'E' in ident_upper and 'D' not in ident_upper and 'Q' not in ident_upper
                    is_dependencia = 'D' in ident_upper and 'Q' not in ident_upper  
                    is_equipo = 'Q' in ident_upper
                    
                    valid_match = False
                    search_id = ident_upper  # Por defecto usar el ID completo
                    
                    if is_edificio:
                        # EDIFICIOS: verificar que la parte del nombre de foto está en el ID
                        foto_parts = stem_upper.split('_')
                        foto_id_part = foto_parts[0] if foto_parts else stem_upper
                        if foto_id_part in ident_upper and 'D' not in stem_upper and 'Q' not in stem_upper:
                            valid_match = True
                            
                    elif is_dependencia:
                        # DEPENDENCIAS: verificar que la parte del nombre de foto está en el ID
                        foto_parts = stem_upper.split('_')
                        foto_id_part = foto_parts[0] if foto_parts else stem_upper
                        if foto_id_part in ident_upper and 'Q' not in stem_upper:
                            valid_match = True
                            
                    elif is_equipo:
                        # EQUIPOS: verificar que la parte del nombre de foto está en el ID
                        foto_parts = stem_upper.split('_')
                        foto_id_part = foto_parts[0] if foto_parts else stem_upper
                        if foto_id_part in ident_upper:
                            valid_match = True
                            
                    else:
                        # OTROS: Lógica por defecto (usar ID completo)
                        if ident_upper in stem_upper:
                            valid_match = True
                    
                    if valid_match:
                        realpath = path_to_info[p]["realpath"]
                        if realpath not in seen_realpaths:
                            seen_realpaths.add(realpath)
                            fotos_unicas.append(p)
    
    return fotos_unicas[:max_photos]

def _buscar_fotos_secuenciales(fotos_base: str, index: Dict[str, Path]) -> List[Path]:
    """Busca fotos adicionales con secuencia incremental a partir de una foto base."""
    rutas_extra = []
    # Patrón para detectar secuencias como FOTO_01, FOTO_02, etc.
    match = re.search(r"(.*_F\w*?)(\d+)$", fotos_base, re.IGNORECASE)
    if match:
        prefijo = match.group(1)
        num = int(match.group(2))
        while True:
            num += 1
            # Probar diferentes formatos de numeración
            for formato in [f"{num:02d}", f"{num:03d}", str(num)]:
                candidato = f"{prefijo}{formato}"
                # Buscar en el índice con diferentes variaciones
                for k, p in index.items():
                    if _norm_slug(k) == _norm_slug(candidato):
                        rutas_extra.append(p)
                        break
            # Si no encuentra más fotos secuenciales, salir
            if not any(_norm_slug(k).startswith(_norm_slug(f"{prefijo}{num:02d}")) or 
                      _norm_slug(k).startswith(_norm_slug(f"{prefijo}{num:03d}")) or
                      _norm_slug(k).startswith(_norm_slug(f"{prefijo}{num}"))
                      for k in index.keys()):
                break
            if len(rutas_extra) >= MAX_SEQUENTIAL_PHOTOS:  # Limitar fotos secuenciales
                break
    return rutas_extra

def _buscar_fotos_por_id_exacto(ident: str, index: Dict[str, Path]) -> List[Path]:
    """Busca todas las fotos cuyo nombre contenga exactamente el ID."""
    fotos_encontradas = []
    ident_norm = _norm_slug(ident)
    leaf_id = _norm_slug(_leaf_id(ident))
    
    seen = set()
    for k, p in index.items():
        k_norm = _norm_slug(k)
        rp = os.path.realpath(str(p))
        
        # Buscar coincidencia exacta del ID completo o del leaf ID
        if (ident_norm in k_norm or leaf_id in k_norm) and rp not in seen:
            fotos_encontradas.append(p)
            seen.add(rp)
    
    return fotos_encontradas

def _fallback_candidates_from_folder(ident: str, index: Dict[str, Path], max_photos: int = 6) -> List[Path]:
    """
    Versión mejorada que busca:
    1. Fotos que coincidan exactamente con el ID
    2. Fotos secuenciales a partir de las encontradas
    3. Fallback por prefijo como antes
    """
    # Paso 1: Buscar fotos exactas por ID
    fotos_exactas = _buscar_fotos_por_id_exacto(ident, index)
    
    # Paso 2: Para cada foto exacta encontrada, buscar secuenciales
    fotos_secuenciales = []
    for foto_path in fotos_exactas:
        foto_stem = Path(foto_path).stem
        secuenciales = _buscar_fotos_secuenciales(foto_stem, index)
        fotos_secuenciales.extend(secuenciales)
    
    # Combinar fotos exactas + secuenciales
    todas_fotos = fotos_exactas + fotos_secuenciales
    
    # Eliminar duplicados manteniendo orden
    seen, ordered = set(), []
    for p in todas_fotos:
        rp = os.path.realpath(str(p))
        if rp not in seen:
            seen.add(rp)
            ordered.append(p)
    
    # Si aún no tenemos suficientes, usar el método original como fallback
    if len(ordered) < max_photos:
        leafN = _norm_slug(_leaf_id(ident))
        for k, p in index.items():
            if len(ordered) >= max_photos:
                break
            rp = os.path.realpath(str(p))
            if rp not in seen:
                if _norm_slug(k).startswith(leafN):
                    seen.add(rp)
                    ordered.append(p)
        
        # Segundo fallback: buscar ID dentro del nombre
        for k, p in index.items():
            if len(ordered) >= max_photos:
                break
            n = _norm_slug(k)
            rp = os.path.realpath(str(p))
            if leafN in n and not n.startswith(leafN) and rp not in seen:
                seen.add(rp)
                ordered.append(p)
    
    return ordered[:max_photos]

# --------------------------
# Consul → índice fotos
# --------------------------
def _split_multi(val) -> List[str]:
    if val is None:
        return []
    s = str(val).strip()
    if not s or s.lower() == "nan" or s == "0" or s == NORMALIZE_EMPTY_TO:
        return []
    parts = re.split(r'[,;|\s]+', s)
    seen, out = set(), []
    for p in parts:
        stem = Path(p).stem.strip()
        if stem and stem not in seen:
            seen.add(stem); out.append(stem)
    return out

def _read_consul(xlsx: Path) -> Optional[pd.DataFrame]:
    xls = pd.ExcelFile(xlsx, engine="openpyxl")
    nm = next((n for n in xls.sheet_names if n.strip().lower() == "consul"), None)
    if not nm:
        return None
    df = xls.parse(nm, dtype=str).fillna("")
    df.columns = df.columns.str.strip()
    for c in df.columns:
        df[c] = df[c].apply(_clean)
    return df

def _consul_index_from_sections(df_consul: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    """
    Devuelve: {TIPO: {ID: [slug1, slug2, ...]}}
    Usa los límites de columnas por bloque definidos en _CONSUL_SPEC
    y extrae TODAS las columnas FOTO_* (ACOM prioriza sus 4).
    """
    if df_consul is None or df_consul.empty:
        return {}
    cols = list(df_consul.columns)
    secciones = [(sec, cols.index(sec)) for (sec, _, _) in _CONSUL_SPEC if sec in cols]
    secciones.sort(key=lambda x: x[1])

    out: Dict[str, Dict[str, List[str]]] = {}
    for i, (sec, ini) in enumerate(secciones):
        fin = secciones[i+1][1] if i+1 < len(secciones) else len(cols)
        sub = df_consul.iloc[:, ini:fin]
        _, id_col, tipo = next(z for z in _CONSUL_SPEC if z[0] == sec)
        if id_col not in df_consul.columns:
            continue

        foto_cols = [c for c in sub.columns if str(c).strip().upper().startswith("FOTO_")]
        if not foto_cols:
            continue

        # Para ACOM, prioriza columnas conocidas en el orden típico
        if tipo == "ACOM":
            prefer = [c for c in sub.columns if c.strip().upper() in ACOM_FOTOS]
            # others  = [c for c in foto_cols if c not in prefer]
            # foto_cols = prefer + others
            foto_cols = prefer  # <-- SOLO las 4 principales
        bucket = out.setdefault(tipo, {})
        for idx, row in sub.iterrows():
            ident = str(df_consul.at[idx, id_col]).strip()
            if not ident:
                continue
            fotos_row = []
            for fc in foto_cols:
                fotos_row.extend(_split_multi(row.get(fc)))
            if not fotos_row:
                continue
            # dedup manteniendo orden
            seen, uniq = set(), []
            for f in fotos_row:
                if f not in seen:
                    seen.add(f); uniq.append(f)
            bucket.setdefault(ident, []).extend(uniq)

    # dedup global por id
    for tipo, d in out.items():
        for k, v in d.items():
            seen, uniq = set(), []
            for f in v:
                if f not in seen:
                    seen.add(f); uniq.append(f)
            d[k] = uniq
    return out

# --------------------------
# Funciones para organización de fotos en filas (compatible con Word)
# --------------------------
def _as_uri(p: str) -> str:
    """Convierte una ruta a URI para compatibilidad con Word."""
    try:
        return Path(p).absolute().as_uri()
    except Exception:
        return p

def _filas_fotos(paths: List[Path], incluir_uris: bool = True) -> List[List[Dict[str, str]]]:
    """
    Organiza las fotos en filas para el documento Word.
    Compatible con el formato original del script.
    """
    if not paths:
        return []
    
    # Anchos según número de fotos por fila
    _ANCHO_FULL, _ANCHO_HALF, _ANCHO_THIRD = 16.0, 12.0, 7.5
    
    n = len(paths)
    if n == 1:
        filas, w = [paths], _ANCHO_FULL
    elif n == 2:
        filas, w = [[paths[0]], [paths[1]]], _ANCHO_HALF
    else:
        # Agrupar de 3 en 3
        filas = [paths[i:i+3] for i in range(0, n, 3)]
        w = _ANCHO_THIRD
    
    resultado = []
    for fila in filas:
        fila_dict = []
        for path in fila:
            foto_dict = {
                "path": str(path),
                "name": path.stem,
                "width_cm": w
            }
            if incluir_uris:
                foto_dict["file_uri"] = _as_uri(str(path))
            fila_dict.append(foto_dict)
        resultado.append(fila_dict)
    
    return resultado

# --------------------------
# Tester (logs)
# --------------------------
def _tester_init(center_id: str, fotos_index: Dict[str, Path]) -> dict:
    uniq_paths = {os.path.realpath(str(p)) for p in fotos_index.values()}
    return {
        "center_id": center_id,
        "used_paths": set(),
        "declared_missing": [],   # [{"tipo","id","slug"}]
        "fallback_used": [],      # [{"tipo","id","count"}]
        "zero_entities": [],      # [{"tipo","id"}]
        "errors": [],             # [{"where","msg"}]
        "all_paths": uniq_paths,  # paths disponibles en carpeta
    }

def _tester_mark_used(tester: Optional[dict], path_str: str):
    if tester is None: return
    tester["used_paths"].add(os.path.realpath(str(path_str)))

def _tester_log_missing(tester: Optional[dict], tipo: str, ident: str, slug: str):
    if tester is None: return
    tester["declared_missing"].append({"tipo": tipo, "id": ident, "slug": Path(slug).stem})

def _tester_log_fallback(tester: Optional[dict], tipo: str, ident: str, count: int):
    if tester is None or count <= 0: return
    tester["fallback_used"].append({"tipo": tipo, "id": ident, "count": int(count)})

def _tester_log_sequential(tester: Optional[dict], tipo: str, ident: str, count: int):
    if tester is None or count <= 0: return
    tester.setdefault("sequential_found", []).append({"tipo": tipo, "id": ident, "count": int(count)})

def _tester_log_zero(tester: Optional[dict], tipo: str, ident: str):
    if tester is None: return
    tester["zero_entities"].append({"tipo": tipo, "id": ident})

def _tester_log_error(tester: Optional[dict], where: str, exc: Exception):
    if tester is None: return
    tester["errors"].append({"where": where, "msg": f"{type(exc).__name__}: {exc}"})

def _tester_write_txt(tester: Optional[dict], outdir: Path):
    if tester is None: return
    centro = tester["center_id"]
    unused = sorted(tester["all_paths"] - tester["used_paths"])
    path = Path(outdir) / f"TEST_FOTOS_{centro}.txt"
    lines = []
    lines.append(f"TEST FOTOS – Centro {centro}")
    lines.append(f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("="*72)

    lines.append("\n[1] Declaradas en Excel PERO NO encontradas en disco:")
    if not tester["declared_missing"]:
        lines.append("  (ninguna)")
    else:
        for it in tester["declared_missing"]:
            lines.append(f"  - {it['tipo']}:{it['id']}  →  {it['slug']}")

    lines.append("\n[2] Fotos secuenciales encontradas automáticamente:")
    if not tester.get("sequential_found", []):
        lines.append("  (ninguna)")
    else:
        for it in tester["sequential_found"]:
            lines.append(f"  - {it['tipo']}:{it['id']}  →  +{it['count']} fotos secuenciales")

    lines.append("\n[3] Fallback por carpeta (Excel vacío para ese ID):")
    if not tester["fallback_used"]:
        lines.append("  (ninguno)")
    else:
        for it in tester["fallback_used"]:
            lines.append(f"  - {it['tipo']}:{it['id']}  →  {it['count']} fotos")

    lines.append("\n[4] Entidades con 0 fotos después de Excel + secuenciales + fallback:")
    if not tester["zero_entities"]:
        lines.append("  (ninguna)")
    else:
        for it in tester["zero_entities"]:
            lines.append(f"  - {it['tipo']}:{it['id']}")

    lines.append("\n[5] Ficheros en carpeta SIN usar:")
    if not unused:
        lines.append("  (ninguno)")
    else:
        for p in unused:
            lines.append(f"  - {p}")

    lines.append("\n[6] Errores durante la extracción:")
    if not tester["errors"]:
        lines.append("  (ninguno)")
    else:
        for it in tester["errors"]:
            lines.append(f"  - {it['where']}: {it['msg']}")

    # Agregar sección de elementos descubiertos automáticamente
    discovered = tester.get("discovered", [])
    lines.append("\n[7] Elementos descubiertos automáticamente desde fotos disponibles:")
    if not discovered:
        lines.append("  (ninguno)")
    else:
        for it in discovered:
            lines.append(f"  - {it['tipo']}:{it['id']} → {it['count']} fotos ({it['key']})")

    # Agregar sección de fotos filtradas por restricción universal
    fotos_filtradas = tester.get("fotos_filtradas_universal", [])
    lines.append("\n[8] Fotos filtradas por restricción universal (nombre no contenido en ID):")
    if not fotos_filtradas:
        lines.append("  (ninguna)")
    else:
        for it in fotos_filtradas:
            lines.append(f"  - {it['entidad_tipo']}:{it['entidad_id']} | Foto: {it['foto_name']} | {it['razon']}")

    path.write_text("\n".join(lines), encoding="utf-8")
    print(f"[tester] Log escrito: {path}")

# --------------------------
# Contexto de datos
# --------------------------
def build_context(dfs: Dict[str, pd.DataFrame]) -> List[Dict]:
    out: List[Dict] = []
    d_cent = dfs["CENT"]
    if d_cent is None or d_cent.empty:
        raise RuntimeError("No se encontró la hoja CENT/CENTRO.")

    def sub(df_key: str, center_id: str, extra: Dict[str, str] = None) -> pd.DataFrame:
        df = dfs[df_key]
        if df is None or df.empty:
            return pd.DataFrame()
        m = (df.get("ID CENTRO", pd.Series(dtype=str)) == center_id)
        if extra:
            for c, v in extra.items():
                m &= (df.get(c, pd.Series(dtype=str)) == v)
        return df[m].copy()

    for _, row_c in d_cent.iterrows():
        cid = str(row_c.get("ID CENTRO")).strip()
        centro = _rename_row(row_c, MAP_CENTRO)

        edificios = []
        d_edif = sub("EDIF", cid)
        for _, row_e in d_edif.iterrows():
            eid = str(row_e.get("ID EDIFICIO")).strip()
            edif = _rename_row(row_e, MAP_EDIF)

            dep   = [_rename_row(r, MAP_DEPEN)  for _, r in sub("DEPEN",  cid, {"ID EDIFICIO": eid}).iterrows()]
            acom  = [_rename_row(r, MAP_ACOM)   for _, r in sub("ACOM",   cid, {"ID EDIFICIO": eid}).iterrows()]
            envol = [_rename_row(r, MAP_ENVOL)  for _, r in sub("ENVOL",  cid, {"ID EDIFICIO": eid}).iterrows()]
            cc    = [_rename_row(r, MAP_SISTCC) for _, r in sub("SISTCC", cid, {"ID EDIFICIO": eid}).iterrows()]
            clima = [_rename_row(r, MAP_CLIMA)  for _, r in sub("CLIMA",  cid, {"ID EDIFICIO": eid}).iterrows()]
            eqh   = [_rename_row(r, MAP_EQH)    for _, r in sub("EQHORIZ",cid, {"ID EDIFICIO": eid}).iterrows()]
            eleva = [_rename_row(r, MAP_ELEVA)  for _, r in sub("ELEVA",  cid, {"ID EDIFICIO": eid}).iterrows()]
            otros = [_rename_row(r, MAP_OTRO)   for _, r in sub("OTROSEQ",cid, {"ID EDIFICIO": eid}).iterrows()]
            ilum  = [_rename_row(r, MAP_ILUM)   for _, r in sub("ILUM",   cid, {"ID EDIFICIO": eid}).iterrows()]

            # especial: ACOM necesita id_edificio y un id único
            for i, it in enumerate(acom):
                it["id_edificio"] = eid
                # Agregar ID único para ACOM (para compatibilidad con plantilla Jinja)
                it["id"] = f"{eid}_ACOM_{i+1:02d}" if len(acom) > 1 else f"{eid}_ACOM"

            edif.update({
                "dependencias": dep, "acom": acom, "envolventes": envol,
                "sistemas_cc": cc, "equipos_clima": clima, "equipos_horiz": eqh,
                "elevadores": eleva, "otros_equipos": otros, "iluminacion": ilum,
            })
            edificios.append(edif)

        out.append({"centro": centro, "edif": edificios})
    return out

# --------------------------
# Fotos (Consul + disco)
# --------------------------
def _alt_ident_candidates(tipo: str, ident: str, cid: str) -> list:
    """
    Genera variantes para casar Consul ↔ contexto.
    Incluye manejo mejorado para casos como C001 vs C0007.
    """
    ident = str(ident or "").strip()
    cid = str(cid or "").strip()
    cands = {ident}
    
    if ident and cid:
        # Variantes básicas
        cands |= {
            f"{cid}{ident}", f"{cid}_{ident}", f"{cid}-{ident}",
            ident.replace(cid, ""), ident.replace("-", ""), ident.replace("_", "")
        }
        
        # Casos especiales para IDs numéricos
        # Ejemplo: C0007 debería encontrar C001_FC0001
        m_cid = re.match(r"^([A-Z])0*(\d+)$", cid, flags=re.I)
        m_ident = re.match(r"^([A-Z])0*(\d+)$", ident, flags=re.I)
        
        if m_cid and m_ident:
            # Extraer partes numéricas
            cid_letter, cid_num = m_cid.groups()
            ident_letter, ident_num = m_ident.groups()
            
            # Generar variantes sin ceros a la izquierda
            cid_short = f"{cid_letter}{cid_num}"  # C0007 -> C7
            ident_short = f"{ident_letter}{ident_num}"  # E001 -> E1
            
            cands.add(cid_short)
            cands.add(ident_short)
            cands.add(f"{cid_short}{ident}")
            cands.add(f"{cid_short}_{ident}")
            cands.add(f"{cid_short}-{ident}")
    
    # Quitar ceros a la izquierda en general
    m = re.match(r"^([A-Z]+)0*(\d+)$", ident, flags=re.I)
    if m:
        cands.add(m.group(1).upper() + m.group(2))
    
    return [c for c in cands if c]

def _declared_from_consul(consul_map: dict, tipo: str, ident: str, cid: str) -> list:
    bucket = consul_map.get(tipo, {}) or {}
    for k in _alt_ident_candidates(tipo, ident, cid):
        if k in bucket:
            return bucket[k]
    return []

def add_photos_to_context(ctx_all: List[Dict], df_consul: Optional[pd.DataFrame],
                          fotos_root: Path, fuzzy_threshold: float,
                          tester_on: bool, outdir: Path, buscar_secuenciales: bool = True,
                          max_secuenciales: int = MAX_SEQUENTIAL_PHOTOS, 
                          incluir_uris: bool = True) -> Tuple[List[Dict], Dict[str, List[str]]]:

    consul_map = _consul_index_from_sections(df_consul) if df_consul is not None else {}
    faltantes: Dict[str, List[str]] = {}

    def inject(entity: dict, tipo: str, cid: str, photo_indices: tuple, tester: Optional[dict]):
        ident = str(entity.get(PHOTO_ID_FIELD[tipo], "")).strip()
        if not ident:
            entity.update({
                "fotos_nombres": [],
                "fotos_paths": [],
                "fotos_por_filas": [],
                "fotos": [],
                "fotos_count": 0
            })
            return

        exact_index, normalized_index, path_to_info = photo_indices
        declared = _declared_from_consul(consul_map, tipo, ident, cid)
        fotos_list = []

       
        if declared:
            for s in declared:
                try:
                    p = _resolve_name_to_path_optimized(s, exact_index, normalized_index, fuzzy_threshold)
                    if p:
                        stem = path_to_info[p]["stem"]
                        fotos_list.append({"path": str(p), "name": stem, "id": stem})
                        _tester_mark_used(tester, str(p))
                        
                        # Buscar fotos secuenciales (solo si está habilitado)
                        if buscar_secuenciales:
                            try:
                                secuenciales = _buscar_fotos_secuenciales_optimized(stem, normalized_index, max_secuenciales)
                                if secuenciales:
                                    _tester_log_sequential(tester, tipo, ident, len(secuenciales))
                                for p_seq in secuenciales:
                                    stem_seq = path_to_info[p_seq]["stem"]
                                    fotos_list.append({"path": str(p_seq), "name": stem_seq, "id": stem_seq})
                                    _tester_mark_used(tester, str(p_seq))
                            except Exception as e:
                                _tester_log_error(tester, f"{tipo}:{ident}:sequential", e)
                    else:
                        _tester_log_missing(tester, tipo, ident, s)
                        faltantes.setdefault(f"{tipo}:{ident}", []).append(s)
                except Exception as e:
                    _tester_log_error(tester, f"{tipo}:{ident}:{s}", e)
        else:
            # Fallback optimizado por carpeta
            try:
                cands = _fallback_candidates_optimized(ident, exact_index, normalized_index, path_to_info, max_photos=6, tipo=tipo)
                for p in cands:
                    stem = path_to_info[p]["stem"]
                    fotos_list.append({"path": str(p), "name": stem, "id": stem})
                    _tester_mark_used(tester, str(p))
                _tester_log_fallback(tester, tipo, ident, len(cands))
            except Exception as e:
                _tester_log_error(tester, f"{tipo}:{ident}:fallback", e)

        # NUEVA RESTRICCIÓN UNIVERSAL: Para TODAS las entidades, filtrar fotos por nombre contenido en ID
        if not declared:  # Solo aplicar si no hay fotos declaradas en Excel
            fotos_list_filtradas = []
            for foto in fotos_list:
                foto_name = foto["name"]
                # Extraer la parte relevante del nombre de la foto (sin extensión y prefijos comunes)
                foto_clean = foto_name.upper()
                
                # Remover prefijos comunes de fotos
                prefixes_to_remove = ["FOTO_", "IMG_", "IMAGE_"]
                for prefix in prefixes_to_remove:
                    if foto_clean.startswith(prefix):
                        foto_clean = foto_clean[len(prefix):]
                        break
                
                # El nombre de la foto debe estar contenido en el ID de la entidad
                # Ejemplo: foto "D0001_FD0001" -> parte "D0001" debe estar en ID "C0007E001D0001"
                # Ejemplo: foto "E001_FE0001" -> parte "E001" debe estar en ID "C0007E001"
                # Ejemplo: foto "QE001_FQE0001" -> parte "QE001" debe estar en ID "C0007E001D0001QE001"
                
                # Extraer la primera parte del nombre de foto (antes del primer _)
                foto_parts = foto_clean.split('_')
                foto_id_part = foto_parts[0] if foto_parts else foto_clean
                
                # Verificar si el ID de la foto está contenido en el ID de la entidad
                incluir_foto = False
                if foto_id_part in ident.upper():
                    incluir_foto = True
                
                # EXCEPCIÓN ESPECIAL: Fotos del estilo C001, C0001, C007, etc. (código de centro)
                # Estas siempre van al centro porque solo hay uno y evitamos problemas
                if tipo == "CENTRO" and re.match(r'^C0*\d+$', foto_id_part):
                    incluir_foto = True
                
                if incluir_foto:
                    fotos_list_filtradas.append(foto)
                else:
                    # Log para debug - foto filtrada por restricción universal
                    if tester:
                        tester.setdefault("fotos_filtradas_universal", []).append({
                            "entidad_tipo": tipo,
                            "entidad_id": ident,
                            "foto_name": foto_name,
                            "foto_id_part": foto_id_part,
                            "razon": f"ID parte '{foto_id_part}' no encontrado en entidad '{ident}'"
                        })
            
            fotos_list = fotos_list_filtradas
        
        # Deduplicación optimizada usando realpath del índice
        seen_realpaths = set()
        fotos_list_unique = []
        for foto in fotos_list:
            path_obj = Path(foto["path"])
            if path_obj in path_to_info:
                realpath = path_to_info[path_obj]["realpath"]
            else:
                realpath = os.path.realpath(foto["path"])
            
            if realpath not in seen_realpaths:
                seen_realpaths.add(realpath)
                fotos_list_unique.append(foto)
        
        # Convertir a formato compatible con Word
        paths_list = [Path(foto["path"]) for foto in fotos_list_unique]
        
        # Generar la estructura de fotos compatible con el script original
        fotos_nombres = [foto["name"] for foto in fotos_list_unique]
        fotos_paths = [foto["path"] for foto in fotos_list_unique]
        fotos_por_filas = _filas_fotos(paths_list, incluir_uris=incluir_uris)
        
        if not fotos_list_unique:
            _tester_log_zero(tester, tipo, ident)

        # Actualizar con formato compatible
        entity.update({
            "fotos_nombres": fotos_nombres,
            "fotos_paths": fotos_paths,
            "fotos_por_filas": fotos_por_filas,
            "fotos": fotos_list_unique,  # Mantener formato actual también
            "fotos_count": len(fotos_list_unique),
        })

    for c in ctx_all:
        cid = c["centro"]["id"]
        
        # Filtrar centros con IDs vacíos o inválidos
        if not cid or cid.strip() in ["", "–", "-", "nan", "NaN", "None"]:
            print(f"⚠️  Omitiendo centro con ID inválido: '{cid}'")
            continue
            
        cid = cid.strip()
        print(f"🏢 Procesando centro: {cid}")

        # localizar carpeta del centro con búsqueda inteligente
        center_dir = None
        if fotos_root and fotos_root.exists():
            # Método 1: Búsqueda exacta
            cand1 = fotos_root / cid
            if cand1.exists():
                center_dir = cand1
            else:
                # Método 2: Búsqueda por nombre que contenga el ID
                hits = [p for p in fotos_root.iterdir() 
                       if p.is_dir() and cid.upper() in p.name.upper()]
                if hits:
                    center_dir = hits[0]
                else:
                    # Método 3: Usar carpeta raíz como fallback
                    center_dir = fotos_root
                    print(f"⚠️  No se encontró carpeta específica para {cid}, usando carpeta raíz")

        ref_dir = None
        for sub in ["Referencias/Fotografías referenciadas", "Referencias",
                    "FOTOGRAFÍAS", "FOTOGRAFIAS", "Fotos", "IMÁGENES", "IMAGENES"]:
            p = center_dir / sub if center_dir else None
            if p and p.exists():
                ref_dir = p; break
        ref_dir = ref_dir or center_dir or fotos_root

        # Construir índices optimizados para búsqueda de fotos
        photo_indices = _build_optimized_photo_index(ref_dir)
        fotos_index = _list_files_index(ref_dir)  # Mantener para compatibilidad con tester
        tester = _tester_init(cid, fotos_index) if tester_on else None

        inject(c["centro"], "CENTRO", cid, photo_indices, tester)
        for e in c["edif"]:
            inject(e, "EDIFICIO", cid, photo_indices, tester)
            for d in e.get("dependencias", []): inject(d, "DEPENDENCIA", cid, photo_indices, tester)
            for it in e.get("acom", []):        inject(it, "ACOM",        cid, photo_indices, tester)
            for it in e.get("envolventes", []): inject(it, "ENVOL",       cid, photo_indices, tester)
            for it in e.get("sistemas_cc", []): inject(it, "SISTCC",      cid, photo_indices, tester)
            for it in e.get("equipos_clima", []):inject(it,"CLIMA",       cid, photo_indices, tester)
            for it in e.get("equipos_horiz", []):inject(it,"EQHORIZ",     cid, photo_indices, tester)
            for it in e.get("elevadores", []):  inject(it, "ELEVA",       cid, photo_indices, tester)
            for it in e.get("otros_equipos", []):inject(it,"OTROSEQ",     cid, photo_indices, tester)
            for it in e.get("iluminacion", []): inject(it, "ILUM",        cid, photo_indices, tester)

        if tester_on:
            _tester_write_txt(tester, outdir)
            # Después de procesar elementos definidos, buscar elementos adicionales en disco
            _discover_missing_elements(c, cid, photo_indices, tester, outdir)

    return ctx_all, faltantes

def _discover_missing_elements(center_data: dict, center_id: str, photo_indices: tuple, tester: dict, outdir: str):
    """Descubre elementos adicionales (QE, QI, QH, etc.) basándose en fotos disponibles pero no usadas."""
    if not tester:
        return
    
    exact_index, normalized_index, path_to_info = photo_indices
    used_paths = tester.get("used_paths", set())
    
    # Buscar patrones de elementos no utilizados
    element_patterns = {}
    
    for path_obj, info in path_to_info.items():
        realpath = os.path.realpath(str(path_obj))
        if realpath in used_paths:
            continue
            
        stem = info["stem"]
        
        # Buscar patrones QE, QI, QH, etc. en fotos no utilizadas
        patterns = [
            (r'(QE)(\d+)', 'sistemas_cc', 'SISTCC'),
            (r'(QI)(\d+)', 'iluminacion', 'ILUM'), 
            (r'(QH)(\d+)', 'equipos_horiz', 'EQHORIZ'),
            (r'(QG)(\d+)', 'sistemas_cc', 'SISTCC'),
            (r'(CR)(\d+)', 'envolventes', 'ENVOL'),
            (r'(I)(\d+)', 'iluminacion', 'ILUM'),
            (r'(E)(\d+)', 'otros_equipos', 'OTROSEQ'),
            (r'(B)(\d+)', 'otros_equipos', 'OTROSEQ')
        ]
        
        for pattern_regex, category, tipo in patterns:
            match = re.search(pattern_regex, stem, re.IGNORECASE)
            if match:
                element_type = match.group(1).upper()
                element_num = match.group(2)
                element_key = f"{element_type}{element_num}"
                
                if element_key not in element_patterns:
                    element_patterns[element_key] = {
                        'type': tipo,
                        'category': category,
                        'photos': [],
                        'id_base': f"{center_id}E001D0001{element_key}"  # ID tentativo
                    }
                
                element_patterns[element_key]['photos'].append(path_obj)
    
    # Crear elementos automáticos para patrones encontrados
    if element_patterns:
        tester.setdefault("discovered", [])
        
        for element_key, info in element_patterns.items():
            photos = info['photos']
            if len(photos) >= 1:  # Solo crear elementos con al menos 1 foto
                # Crear elemento automático
                auto_element = {
                    'id': info['id_base'],
                    'nombre': f'Elemento {element_key} (auto-descubierto)',
                    'fotos': [],
                    'fotos_nombres': [],
                    'fotos_paths': [],
                    'fotos_por_filas': []
                }
                
                # Agregar fotos al elemento
                for photo_path in photos[:6]:  # Máximo 6 fotos por elemento
                    stem = path_to_info[photo_path]["stem"]
                    auto_element['fotos'].append({
                        'path': str(photo_path),
                        'name': stem,
                        'id': stem
                    })
                    auto_element['fotos_nombres'].append(stem)
                    auto_element['fotos_paths'].append(str(photo_path))
                    _tester_mark_used(tester, str(photo_path))
                
                # Generar fotos_por_filas
                auto_element['fotos_por_filas'] = _filas_fotos(
                    [Path(f['path']) for f in auto_element['fotos']], 
                    incluir_uris=True
                )
                
                # Agregar elemento al centro correspondiente
                if info['category']:
                    # Agregar a la categoría apropiada del primer edificio
                    if center_data.get('edif'):
                        edificio = center_data['edif'][0]
                        if info['category'] not in edificio:
                            edificio[info['category']] = []
                        edificio[info['category']].append(auto_element)
                
                tester["discovered"].append({
                    'tipo': info['type'],
                    'id': info['id_base'],
                    'count': len(photos),
                    'key': element_key
                })
                
        print(f"✨ Auto-descubiertos {len(element_patterns)} elementos desde fotos disponibles en {center_id}")
    else:
        tester["discovered"] = []

def _generar_jsons_por_tipo(contexto_centro: dict, carpeta_centro: Path) -> None:
    """
    Genera JSONs separados por tipo de entidad en una carpeta específica del centro.
    
    Args:
        contexto_centro: Diccionario con toda la info del centro
        carpeta_centro: Path a la carpeta donde crear los JSONs separados
    """
    carpeta_centro.mkdir(parents=True, exist_ok=True)
    
    # 1. JSON completo (igual que antes)
    completo_path = carpeta_centro / "completo.json"
    completo_path.write_text(
        json.dumps(contexto_centro, ensure_ascii=False, indent=2), 
        encoding="utf-8"
    )
    
    centro_info = contexto_centro.get("centro", {})
    edificios = contexto_centro.get("edif", [])
    
    # 2. JSON solo del centro (info básica + fotos del centro)
    centro_json = {
        "centro": centro_info
    }
    centro_path = carpeta_centro / "centro.json"
    centro_path.write_text(
        json.dumps(centro_json, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    
    # 3. JSON de edificios (info básica de edificios sin equipos)
    edificios_basicos = []
    for edif in edificios:
        edificio_basico = {k: v for k, v in edif.items() 
                          if k not in ['dependencias', 'acom', 'envolventes', 'sistemas_cc', 
                                     'equipos_clima', 'equipos_horiz', 'elevadores', 
                                     'otros_equipos', 'iluminacion']}
        edificios_basicos.append(edificio_basico)
    
    edificios_json = {
        "centro": centro_info,
        "edificios": edificios_basicos
    }
    edificios_path = carpeta_centro / "edificios.json"
    edificios_path.write_text(
        json.dumps(edificios_json, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    
    # 4-12. JSONs por tipo de equipo/elemento
    tipos_equipos = {
        "dependencias": "dependencias.json",
        "acom": "acom.json", 
        "envolventes": "envol.json",
        "sistemas_cc": "cc.json",
        "equipos_clima": "clima.json",
        "equipos_horiz": "eqhoriz.json",
        "elevadores": "eleva.json",
        "iluminacion": "ilum.json",
        "otros_equipos": "otroseq.json"
    }
    
    for tipo_key, archivo in tipos_equipos.items():
        # Recopilar todos los elementos de este tipo de todos los edificios
        elementos_tipo = []
        
        for edif in edificios:
            elementos = edif.get(tipo_key, [])
            for elemento in elementos:
                # Agregar contexto de edificio a cada elemento
                elemento_con_contexto = elemento.copy()
                elemento_con_contexto.update({
                    "edificio_id": edif.get("id"),
                    "edificio_nombre": edif.get("denominacion", ""),
                    "centro_id": centro_info.get("id"),
                    "centro_nombre": centro_info.get("nombre", "")
                })
                elementos_tipo.append(elemento_con_contexto)
        
        # Solo crear el archivo si hay elementos de este tipo
        if elementos_tipo:
            tipo_json = {
                "centro": centro_info,
                tipo_key: elementos_tipo,
                "resumen": {
                    "total_elementos": len(elementos_tipo),
                    "tipo": tipo_key,
                    "total_fotos": sum(elem.get('fotos_count', 0) for elem in elementos_tipo)
                }
            }
            
            tipo_path = carpeta_centro / archivo
            tipo_path.write_text(
                json.dumps(tipo_json, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
    
    print(f"    📁 JSONs separados generados en: {carpeta_centro.name}/")

# --------------------------
# Main CLI
# --------------------------
def main():
    ap = argparse.ArgumentParser(prog="extraer_datos_word.py",
                                description="Extrae datos de Excel de auditorías energéticas y vincula fotografías")
    ap.add_argument("--xlsx", help="Ruta al Excel maestro (si no se especifica, se solicitará interactivamente)")
    ap.add_argument("--outdir", help="Carpeta de salida JSON")
    ap.add_argument("--fotos-root", dest="fotos_root", help="Carpeta RAÍZ donde están las fotos")
    ap.add_argument("--centro", help="Procesar solo este ID de centro (ej. C0007)")
    ap.add_argument("--no-combinado", action="store_true", help="No generar JSON combinado")
    ap.add_argument("--fuzzy-threshold", type=float, default=FUZZY_THRESHOLD_DEFAULT,
                    help="Similitud 0–1 para matching de nombre (1.0 = exacto).")
    ap.add_argument("--buscar-secuenciales", action="store_true", 
                    help="Buscar fotos secuenciales automáticamente (ej: FOTO_01, FOTO_02...)")
    ap.add_argument("--max-secuenciales", type=int, default=MAX_SEQUENTIAL_PHOTOS,
                    help=f"Máximo número de fotos secuenciales por entidad (default: {MAX_SEQUENTIAL_PHOTOS})")
    ap.add_argument("--jsons-separados", action="store_true", 
                    help="Genera JSONs separados por tipo en carpetas individuales por centro")
    ap.add_argument("--tester", action="store_true", help="Escribe TEST_FOTOS_<CENTRO>.txt y faltantes.")
    ap.add_argument("--uris", action="store_true", help="Incluir file_uri en las fotos (para Word)")
    ap.add_argument("--interactivo", action="store_true", 
                    help="Modo interactivo: solicita rutas y opciones al usuario")
    ap.add_argument("--no-interactivo", action="store_true", 
                    help="Fuerza modo no interactivo (requiere --xlsx y --fotos-root)")
    args = ap.parse_args()

    print("\n🏢 EXTRACTOR DE DATOS DE AUDITORÍAS ENERGÉTICAS")
    print("=" * 60)
    
    # Determinar si usar modo interactivo
    modo_interactivo = True
    if args.no_interactivo:
        modo_interactivo = False
    elif args.interactivo:
        modo_interactivo = True
    elif args.xlsx and args.fotos_root:
        modo_interactivo = False
    
    if modo_interactivo:
        print("🔄 Modo interactivo activado")
        
        # Solicitar información al usuario
        excel_path = _solicitar_excel()
        fotos_path = _solicitar_carpeta_fotos()
        opciones = _solicitar_opciones_adicionales()
        
        # Mostrar resumen y confirmar
        if not _mostrar_resumen(excel_path, fotos_path, opciones):
            print("❌ Operación cancelada por el usuario.")
            sys.exit(0)
        
        # Configurar variables
        xlsx = excel_path
        fotos_root = fotos_path
        outdir = opciones['outdir']
        centro_filter = opciones['centro']
        buscar_secuenciales = opciones['buscar_secuenciales']
        max_secuenciales = opciones['max_secuenciales']
        fuzzy_threshold = opciones['fuzzy_threshold']
        tester_on = opciones['tester']
        jsons_separados = opciones['jsons_separados']
        incluir_uris = True  # Por defecto en modo interactivo
        
    else:
        print("🔧 Modo línea de comandos activado")
        
        # Validar parámetros requeridos
        if not args.xlsx:
            print("❌ Error: Se requiere --xlsx en modo no interactivo")
            sys.exit(1)
        if not args.fotos_root:
            print("❌ Error: Se requiere --fotos-root en modo no interactivo")
            sys.exit(1)
            
        xlsx = Path(args.xlsx)
        fotos_root = Path(args.fotos_root)
        outdir = Path(args.outdir) if args.outdir else Path("./out_context")
        centro_filter = args.centro
        buscar_secuenciales = args.buscar_secuenciales
        max_secuenciales = args.max_secuenciales
        fuzzy_threshold = args.fuzzy_threshold
        tester_on = args.tester
        jsons_separados = args.jsons_separados
        incluir_uris = args.uris

    # Validaciones comunes
    if not xlsx.exists():
        print(f"❌ Error: No existe el Excel: {xlsx}")
        sys.exit(1)

    if not fotos_root.exists():
        print(f"❌ Error: No existe la carpeta de fotos raíz: {fotos_root}")
        sys.exit(1)

    outdir.mkdir(parents=True, exist_ok=True)
    
    print(f"\n🚀 Iniciando procesamiento...")
    print(f"📊 Excel: {xlsx.name}")
    print(f"📁 Fotos: {fotos_root.name}")

    # 1) Leer hojas
    print("\n📖 Leyendo hojas del Excel...")
    dfs = _read_all_sheets(xlsx)

    # 2) Contexto base
    print("🏗️  Construyendo contexto de datos...")
    ctx = build_context(dfs)

    # 3) Filtro por centro (opcional)
    if centro_filter:
        print(f"🎯 Filtrando por centro: {centro_filter}")
        ctx = [c for c in ctx if c.get("centro", {}).get("id") == centro_filter]
        if not ctx:
            print(f"❌ Error: No se encontró el centro {centro_filter}")
            sys.exit(1)

    # 4) Fotos
    print("📸 Procesando fotografías...")
    ctx, falt = add_photos_to_context(ctx, _read_consul(xlsx), fotos_root,
                                      fuzzy_threshold, tester_on, outdir,
                                      buscar_secuenciales, max_secuenciales, incluir_uris)

    # 5) Guardar
    print("\n💾 Guardando archivos JSON...")
    all_ctx = []
    
    for c in ctx:
        cent = c["centro"]
        cid  = cent.get("id") or "CENTRO"
        nom  = cent.get("nombre") or cid
        
        # Generar JSONs separados por tipo si se solicita
        if jsons_separados:
            carpeta_centro = outdir / f"{cid}_{_safe_name(nom)}"
            _generar_jsons_por_tipo(c, carpeta_centro)
            print(f"  ✅ {carpeta_centro.name}/ (JSONs separados)")
        else:
            # Método tradicional: un JSON por centro
            outpath = outdir / f"{cid}_{_safe_name(nom)}.json"
            outpath.write_text(json.dumps(c, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  ✅ {outpath.name}")
        
        all_ctx.append(c)

    if not args.no_combinado:
        combinado_path = outdir / "contexto_con_fotos__COMBINADO.json"
        combinado_path.write_text(
            json.dumps(all_ctx, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"  ✅ {combinado_path.name}")

    if falt:
        faltantes_path = outdir / "fotos_faltantes_por_id.json"
        faltantes_path.write_text(
            json.dumps(falt, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"  ⚠️  {faltantes_path.name} ({len(falt)} entradas)")

    print(f"\n🎉 ¡Procesamiento completado exitosamente!")
    print(f"📂 Archivos generados en: {outdir.resolve()}")
    print(f"📊 Centros procesados: {len(all_ctx)}")
    
    # Estadísticas de fotos
    total_fotos = sum(
        sum(entity.get('fotos_count', 0) 
            for edificio in centro.get('edif', [])
            for seccion in ['dependencias', 'acom', 'envolventes', 'sistemas_cc', 
                          'equipos_clima', 'equipos_horiz', 'elevadores', 'otros_equipos', 'iluminacion']
            for entity in edificio.get(seccion, []))
        + sum(edificio.get('fotos_count', 0) for edificio in centro.get('edif', []))
        + centro.get('centro', {}).get('fotos_count', 0)
        for centro in all_ctx
    )
    print(f"📸 Total fotos vinculadas: {total_fotos}")
    
    if tester_on:
        print(f"🧪 Archivos de testing generados en: {outdir}")
    
    print("\n" + "="*60)

if __name__ == "__main__":
    main()
