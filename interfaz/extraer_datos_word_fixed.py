from __future__ import annotations

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
extraer_datos_word.py

Qué hace:
  1) Lee el Excel maestro y construye un contexto jerárquico:
     centro → edificios → dependencias / acom / envol / cc / clima / eqh / eleva / otros / ilum
  2) Resuelve fotos declaradas en la hoja "Consul" (todas las columnas FOTO_*) + disco.
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
  py -3.13 .\\extraer_datos_word.py
  - Abre exploradores de archivos para seleccionar Excel y carpetas
  - Interfaz visual fácil de usar con validaciones automáticas
  - Fallback automático a modo texto si no hay GUI disponible

2) MODO LÍNEA DE COMANDOS (avanzado):
  py -3.13 .\\extraer_datos_word.py --no-interactivo --xlsx RUTA --fotos-root RUTA
  - Para automatización, scripts o usuarios expertos
  - Todos los parámetros por línea de comandos

Ejemplo interactivo:
  py -3.13 .\\extraer_datos_word.py

Ejemplo línea de comandos completo:
  $xlsx = "Z:\\...\\ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1.xlsx"
  $root = "Z:\\...\\1_CONSULTA 1"
  py -3.13 .\\extraer_datos_word.py `
    --no-interactivo `
    --xlsx $xlsx `
    --fotos-root $root `
    --outdir .\\out_context `
    --centro C0007 `
    --fuzzy-threshold 0.88 `
    --buscar-secuenciales `
    --max-secuenciales 15 `
    --tester
"""

# Placeholder para mantener la estructura básica
# El archivo completo está corrompido y necesita reconstrucción completa
