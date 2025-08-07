# Inicio Rápido - Crear Anexo 3

Esta guía te ayudará a usar rápidamente el sistema de generación de documentos Word con análisis por edificios.

## Configuración Inicial

### 1. Activar Entorno Virtual

```powershell
# Navegar al directorio del proyecto
cd "C:\Users\ferma\Documents\repos\artecoin_automatizaciones"

# Activar entorno virtual
.\artecoin_venv\Scripts\Activate.ps1
```

### 2. Verificar Dependencias

```powershell
# Instalar o actualizar dependencias
pip install -r requirements.txt
```

### 3. Preparar Archivos

Asegúrate de tener:
- **Archivo Excel** con los datos fuente
- **Plantilla Word** (.docx) con las variables de template
- **Directorio de salida** para los documentos generados

## Uso Rápido

### Ejecutar Generación Completa

```powershell
# Navegar al directorio de tests
cd tests

# Ejecutar el script principal desde la raíz del proyecto
python ../anexos/crear_anexo_3.py
```

### ¿Qué hace el script?

1. **Lee datos** del archivo Excel especificado
2. **Filtra por edificios** únicos encontrados en los datos
3. **Calcula totales** específicos por cada edificio
4. **Genera documentos Word** individuales usando la plantilla
5. **Limpia nombres** de archivos para compatibilidad con Windows
6. **Actualiza campos** automáticamente en Word

## Funciones Principales

### `clean_filename(filename)`

Limpia nombres de archivo para compatibilidad con Windows.

**Ejemplo:**
```python
from crear_anexo_3 import clean_filename

# Antes
nombre_problema = 'Archivo "con comillas" y <símbolos>.docx'

# Después
nombre_limpio = clean_filename(nombre_problema)
# Resultado: 'Archivo (con comillas) y símbolos.docx'
```

**Características:**
- Elimina caracteres problemáticos (`"`, `<`, `>`, `|`, `:`, `*`, `?`, `\`, `/`)
- Maneja nombres reservados de Windows (CON, PRN, NUL, etc.)
- Trunca nombres muy largos (máximo 255 caracteres)
- Proporciona nombre por defecto si el resultado está vacío

### `get_totales_edificio(df_full, df_edificio)`

Calcula totales específicos para un edificio.

**Ejemplo:**
```python
from crear_anexo_3 import get_totales_edificio
import pandas as pd

# DataFrames de ejemplo
df_completo = pd.read_excel('datos_completos.xlsx')
df_edificio_a = df_completo[df_completo['Edificio'] == 'EDIFICIO A']

# Calcular totales
totales = get_totales_edificio(df_completo, df_edificio_a)
print(totales)
```

**Características:**
- Suma automática de todas las columnas numéricas
- Añade fila con `Tipo = 'TOTALES'`
- Manejo robusto de diferentes tipos de datos
- Validación de tipos con type hints

## Configuración Personalizada

### Modificar Rutas de Archivos

Edita las variables en `crear_anexo_3.py`:

```python
# Rutas principales (líneas ~200-210)
ruta_excel = r"C:\ruta\a\tu\archivo.xlsx"
ruta_plantilla = r"C:\ruta\a\tu\plantilla.docx"
directorio_salida = r"C:\ruta\de\salida"
```

### Personalizar Filtros

Modifica los filtros de datos según tus necesidades:

```python
# Ejemplo: filtrar por tipo de instalación
df_filtrado = df[df['Tipo'].isin(['Iluminación', 'Climatización'])]

# Ejemplo: filtrar por potencia mínima
df_filtrado = df[df['Potencia (W)'] > 100]
```

## Validación y Tests

### Ejecutar Tests

```powershell
# Tests completos con output detallado
python run_tests.py -v

# Test específico
python run_tests.py -t TestCrearAnexo3.test_clean_filename_basic
```

### Validar Salida

Después de ejecutar el script, verifica:

1. **Archivos generados** en el directorio de salida
2. **Nombres de archivo** sin caracteres problemáticos
3. **Contenido de documentos** con datos correctos y totales calculados
4. **Campos actualizados** automáticamente en Word

## Estructura de Salida

```
directorio_salida/
├── Anexo 3 EDIFICIO A - PRINCIPAL.docx
├── Anexo 3 EDIFICIO B - PRINCIPAL.docx
├── Anexo 3 CASA DE LA JUVENTUD - PRINCIPAL.docx
└── ...
```

Cada archivo contiene:
- Datos específicos del edificio
- Totales calculados automáticamente
- Campos de Word actualizados
- Formato consistente basado en la plantilla

## Solución de Problemas Comunes

### Error: "No se puede crear archivo"

**Causa:** Caracteres inválidos en nombre
**Solución:** La función `clean_filename()` debería resolverlo automáticamente

### Error: "KeyError en columna"

**Causa:** Estructura de datos diferente a la esperada
**Solución:** Verificar nombres de columnas en el Excel

### Error: "Word no responde"

**Causa:** Proceso de Word colgado
**Solución:** 
```powershell
# Terminar procesos de Word
Get-Process -Name "WINWORD" | Stop-Process -Force
```

### Error: "ModuleNotFoundError"

**Causa:** Dependencias no instaladas
**Solución:**
```powershell
pip install -r requirements.txt
```

## Contacto y Soporte

- **Documentación completa:** Ver `README_TESTS.md`
- **Tests:** Ejecutar `python run_tests.py -v`
- **Logs:** Revisar output de la consola para errores detallados

---


Ejecuta `python crear_anexo_3.py` y revisa la salida en tu directorio especificado.