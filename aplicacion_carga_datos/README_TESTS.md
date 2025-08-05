# Guía para Ejecutar Tests - Sistema de Generación de Anexos

## Descripción General

Este proyecto incluye un sistema completo de tests para validar la funcionalidad del módulo `crear_anexo_3.py`, que se encarga de procesar datos de Excel y generar documentos Word automatizados.

## Requisitos Previos

### 1. Python
- **Versión requerida**: Python 3.10 o superior
- **Versión recomendada**: Python 3.13.5 (versión validada)

### 2. Sistema Operativo
- **Compatible**: Windows 10/11 (requerido para funcionalidades de Word)
- **Nota**: Las funcionalidades de actualización automática de índices de Word solo funcionan en Windows

### 3. Microsoft Word (Opcional pero recomendado)
- Microsoft Word instalado para funcionalidades completas de automatización
- Sin Word, el sistema funcionará pero no podrá actualizar índices automáticamente

## Configuración del Entorno

### Paso 1: Clonar/Descargar el Proyecto
```bash
# Si usas Git
git clone [URL_DEL_REPOSITORIO]
cd artecoin_automatizaciones

# O descarga y extrae el ZIP del proyecto
```

### Paso 2: Configurar el Entorno Virtual

#### Opción A: Usar el entorno virtual existente (recomendado)
```powershell
# Navegar al directorio del proyecto
cd artecoin_automatizaciones

# Activar el entorno virtual existente
.\artecoin_venv\Scripts\Activate.ps1

# Si aparece error de políticas de ejecución, ejecutar:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Opción B: Crear un nuevo entorno virtual
```powershell
# Crear nuevo entorno virtual
python -m venv artecoin_venv

# Activar el entorno virtual
.\artecoin_venv\Scripts\Activate.ps1

# Instalar dependencias desde el requirements.txt principal
pip install -r ..\requirements.txt
```

### Paso 3: Verificar las Dependencias
Las dependencias principales incluyen:
- `pandas>=2.3.1` - Para manipulación de datos
- `openpyxl>=3.1.5` - Para lectura de archivos Excel
- `docxtpl` - Para plantillas de Word
- `pywin32` - Para automatización de Word (Windows)

## Estructura de Archivos de Test

```
aplicacion_carga_datos/
├── crear_anexo_3.py           # Módulo principal
├── test_crear_anexo_3.py      # Suite de tests unitarios
├── run_tests.py               # Script ejecutor de tests
└── README_TESTS.md            # Esta guía
```

## Ejecutar los Tests

### Método 1: Usar el Script Ejecutor (Recomendado)

```powershell
# Navegar al directorio de la aplicación
cd aplicacion_carga_datos

# Activar entorno virtual
..\artecoin_venv\Scripts\Activate.ps1

# Ejecutar tests con información detallada
python run_tests.py -v

# O simplemente
python run_tests.py
```

### Método 2: Ejecutar Tests Directamente

```powershell
# Activar entorno virtual
..\artecoin_venv\Scripts\Activate.ps1

# Ejecutar tests directamente
python -m unittest test_crear_anexo_3.py -v

# O ejecutar el archivo de tests
python test_crear_anexo_3.py
```

### Opciones del Script Ejecutor

```powershell
# Verificar solo el entorno (sin ejecutar tests)
python run_tests.py --check

# Instalar dependencias faltantes automáticamente
python run_tests.py --install-deps

# Ejecutar con salida detallada
python run_tests.py -v

# Ayuda
python run_tests.py -h
```

## Interpretación de Resultados

### Ejecución Exitosa
```
🔍 Verificando entorno...
✅ Python: 3.13.5
✅ Pandas: 2.3.1
✅ Openpyxl: 3.1.5
🧪 EJECUTANDO TESTS PARA CREAR_ANEXO_3.PY
============================================================
...
----------------------------------------------------------------------
Ran 15 tests in 2.088s
OK
✅ TODOS LOS TESTS PASARON EXITOSAMENTE
```

### ❌ Errores Comunes y Soluciones

#### Error: "ModuleNotFoundError: No module named 'pandas'"
**Solución**: 
```powershell
pip install pandas openpyxl docxtpl pywin32

# O instalar todas las dependencias:
pip install -r ..\requirements.txt
```

#### Error: "UnicodeEncodeError"
**Solución**: Ya está resuelto en la versión actual. Los caracteres Unicode han sido reemplazados por caracteres ASCII compatibles.

#### Error: "FileNotFoundError" para archivo Excel
**Solución**: Asegurar que existe el archivo:
`excel/proyecto/ANALISIS AUD-ENER_COLMENAR VIEJO_CONSULTA 1_V20.xlsx`

#### Error de políticas de PowerShell
**Solución**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Qué Validan los Tests

### Tests Unitarios (12 tests)
1. **TestDeleteRowsOptimized** (4 tests)
   - Eliminación correcta de filas vacías
   - Manejo de DataFrames vacíos
   - Comportamiento con columnas inexistentes
   - Validación con datos completamente válidos

2. **TestCleanLastRow** (3 tests)
   - Limpieza de filas con "Total"
   - Comportamiento sin filas "Total"
   - Manejo de DataFrames vacíos

3. **TestLoadAndCleanSheets** (3 tests)
   - Carga exitosa de hojas Excel
   - Manejo de hojas faltantes
   - Comportamiento con archivos inexistentes

4. **TestUpdateWordFields** (3 tests)
   - Verificación de funcionalidad básica
   - Validación de tipos de parámetros
   - Existencia y capacidad de llamada de la función

### Tests de Integración (2 tests)
5. **TestIntegration** (2 tests)
   - Consistencia del mapeo de hojas (SHEET_MAP)
   - Flujo completo de procesamiento de datos

## Archivos Generados Durante Tests

Durante la ejecución de tests se pueden generar:
- Archivos Excel temporales (se eliminan automáticamente)
- `Anexo 3.docx` en `word/anejos/` (archivo de salida real)

## Personalización y Mantenimiento

### Añadir Nuevos Tests
1. Editar `test_crear_anexo_3.py`
2. Seguir la estructura de clases existente
3. Usar métodos `setUp()` y `tearDown()` para preparación/limpieza

### Modificar Configuración
- Editar `SHEET_MAP` en `crear_anexo_3.py` para cambiar mapeo de hojas
- Modificar rutas de archivos en las constantes del módulo principal

## Soporte

Si encuentras problemas:

1. **Verificar entorno**: `python run_tests.py --check`
2. **Instalar dependencias**: `python run_tests.py --install-deps`
3. **Revisar versiones**: Asegurar Python 3.10+ y dependencias actualizadas
4. **Verificar archivos**: Confirmar que existen todos los archivos Excel necesarios

## Cobertura de Tests

- **Funciones principales**: 100% cubiertas
- **Manejo de errores**: Validado
- **Casos extremos**: Incluidos (DataFrames vacíos, archivos inexistentes)
- **Integración**: Flujo completo validado

---

**Última actualización**: Agosto 2025  
**Versión Python validada**: 3.13.5  
**Dependencias principales**: pandas 2.3.1, openpyxl 3.1.5
