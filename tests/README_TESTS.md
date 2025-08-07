# Tests para Crear Anexo 3

Este documento describe la suite de tests para el módulo `crear_anexo_3.py` que genera documentos Word con análisis de edificios y totales.

## Índice

- [Configuración](#configuración)
- [Ejecución de Tests](#ejecución-de-tests)
- [Descripción de Tests](#descripción-de-tests)
- [Cobertura de Funcionalidades](#cobertura-de-funcionalidades)
- [Solución de Problemas](#solución-de-problemas)

## Configuración

### Requisitos Previos

Asegúrate de tener instaladas todas las dependencias:

```powershell
# Activar el entorno virtual (si usas uno)
.\artecoin_venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt
```

### Estructura de Archivos

```
tests/
├── test_crear_anexo_3.py      # Tests unitarios
├── run_tests.py              # Script ejecutor de tests
├── setup_tests.ps1           # Script de configuración PowerShell
└── README_TESTS.md           # Esta documentación
```

## Ejecución de Tests

### Ejecución Básica

```powershell
# Ejecutar todos los tests
python run_tests.py

# Ejecutar con output detallado
python run_tests.py -v

# Detener en el primer fallo
python run_tests.py -f
```

### Ejecución de Tests Específicos

```powershell
# Ejecutar un test específico
python run_tests.py -t TestCrearAnexo3.test_clean_filename_basic

# Ejecutar toda una clase de tests
python run_tests.py -t TestCrearAnexo3
```

### Usando unittest directamente

```powershell
# Ejecutar todos los tests
python -m unittest test_crear_anexo_3 -v

# Ejecutar un test específico
python -m unittest test_crear_anexo_3.TestCrearAnexo3.test_clean_filename_basic
```

## Descripción de Tests

### TestCrearAnexo3

Clase principal que prueba las funciones del módulo `crear_anexo_3.py`.

#### Tests para `clean_filename()`

| Test | Descripción | Casos Probados |
|------|-------------|----------------|
| `test_clean_filename_basic` | Caracteres problemáticos básicos | Comillas, símbolos especiales |
| `test_clean_filename_reserved_names` | Nombres reservados de Windows | CON, PRN, NUL, etc. |
| `test_clean_filename_long_names` | Nombres muy largos | Truncamiento a 255 caracteres |
| `test_clean_filename_empty_or_none` | Casos extremos | Cadenas vacías, None |
| `test_clean_filename_only_invalid_chars` | Solo caracteres inválidos | Reemplazo por nombre por defecto |

#### Tests para `get_totales_edificio()`

| Test | Descripción | Casos Probados |
|------|-------------|----------------|
| `test_get_totales_edificio_basic` | Funcionalidad básica | Cálculo correcto de totales |
| `test_get_totales_edificio_tipo_column` | Columna Tipo | Verificar 'TOTALES' en resultado |
| `test_get_totales_edificio_empty_dataframes` | DataFrames vacíos | Manejo de datos vacíos |
| `test_get_totales_edificio_missing_columns` | Columnas faltantes | Manejo de errores |
| `test_get_totales_edificio_numeric_columns_only` | Solo columnas numéricas | Suma selectiva de datos |
| `test_get_totales_edificio_type_hints` | Tipos de parámetros | Validación de tipos |

### TestFileIntegration

Clase para tests de integración de archivos.

| Test | Descripción | Casos Probados |
|------|-------------|----------------|
| `test_clean_filename_file_creation` | Creación real de archivos | Verificar que nombres limpiados funcionan |

## Cobertura de Funcionalidades

### Funcionalidades Cubiertas

- **Limpieza de nombres de archivo**
  - Eliminación de caracteres inválidos en Windows
  - Manejo de nombres reservados del sistema
  - Truncamiento de nombres largos
  - Casos extremos (vacío, solo inválidos)

- **Cálculo de totales por edificio**
  - Suma de columnas numéricas
  - Manejo de DataFrames con diferentes estructuras
  - Validación de tipos de datos
  - Configuración correcta de columnas resultado

- **Integración de archivos**
  - Verificación de que nombres limpiados permiten crear archivos reales

### Limitaciones Conocidas

- Los tests no cubren la integración completa con Word (requiere instalación de Word)
- No se prueban casos con archivos Excel reales muy grandes
- Los tests de integración usan archivos temporales simples

## Interpretación de Resultados

### Salida Exitosa

```
======================================================================
EJECUTANDO TESTS PARA CREAR_ANEXO_3
======================================================================
...
======================================================================
RESUMEN DE RESULTADOS
======================================================================
Tests ejecutados: 12
Exitosos: 12
Fallos: 0
Errores: 0
Omitidos: 0

✅ Todos los tests PASARON
```

### Salida con Errores

```
FAIL: test_clean_filename_basic (test_crear_anexo_3.TestCrearAnexo3)
----------------------------------------------------------------------
AssertionError: 'Archivo (con comillas).docx' != 'Archivo con comillas.docx'

======================================================================
RESUMEN DE RESULTADOS
======================================================================
Tests ejecutados: 12
Exitosos: 11
Fallos: 1
Errores: 0
Omitidos: 0

FALLOS:
  - test_clean_filename_basic (test_crear_anexo_3.TestCrearAnexo3): AssertionError: 'Archivo (con comillas).docx' != 'Archivo con comillas.docx'

❌ Tests FALLARON (1 problemas)
```

## Solución de Problemas

### Error: "ModuleNotFoundError: No module named 'crear_anexo_3'"

**Solución:**
```powershell
# Asegúrate de estar en el directorio correcto
cd aplicacion_carga_datos

# Ejecutar desde este directorio
python run_tests.py
```

### Error: "No module named 'pandas'"

**Solución:**
```powershell
# Instalar dependencias
pip install -r requirements.txt
```

### Tests muy lentos

**Causa:** Los tests de integración crean archivos temporales.

**Solución:** Usar el flag `-f` para detener en el primer fallo durante desarrollo.

### Error de permisos en Windows

**Causa:** Antivirus o permisos de directorio.

**Solución:**
```powershell
# Ejecutar PowerShell como administrador
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# O usar el script de configuración
.\setup_tests.ps1
```

## Métricas de Calidad

- **Cobertura de código:** ~90% de las funciones principales
- **Tiempo de ejecución:** < 2 segundos para toda la suite
- **Casos de prueba:** 12+ escenarios diferentes
- **Robustez:** Manejo de casos extremos y errores

## Mantenimiento

### Añadir Nuevos Tests

1. Editar `test_crear_anexo_3.py`
2. Seguir la convención de nombres: `test_nombre_descriptivo`
3. Usar `setUp()` para datos de prueba comunes
4. Documentar el propósito de cada test

### Actualizar Tests Existentes

1. Mantener compatibilidad con versiones anteriores
2. Actualizar documentación si cambia la funcionalidad
3. Ejecutar toda la suite después de cambios

---

**Última actualización:** Diciembre 2024  
**Mantenedor:** Equipo de desarrollo artecoin_automatizaciones