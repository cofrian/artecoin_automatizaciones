# Artecoin Automatizaciones

Sistema automatizado para la generación de documentos Word con análisis energético por edificios. Este repositorio contiene herramientas para procesar datos de Excel, generar documentos con plantillas Word y realizar cálculos específicos por edificio.
## Características Principales

- **Generación automática de documentos Word** por edificio
- **Cálculo de totales** específicos por edificio y sección
- **Limpieza automática de nombres** de archivo para Windows
- **Procesamiento de datos Excel** con múltiples hojas
- **Actualización automática de campos** en documentos Word
- **Suite de tests completa** para validación de funciones
- **Documentación detallada** y guías de inicio rápido

## Estructura del Proyecto

```
artecoin_automatizaciones/
├── README.md                          # Este archivo
├── requirements.txt                   # Dependencias Python
├── anexos/                            # Módulo principal
│   └── crear_anexo_3.py              # Script de generación de documentos
├── tests/                             # Suite de testing y documentación
│   ├── test_crear_anexo_3.py         # Tests unitarios
│   ├── run_tests.py                  # Ejecutor de tests
│   ├── setup_tests.ps1               # Configuración PowerShell
│   ├── README_TESTS.md               # Documentación de tests
│   └── INICIO_RAPIDO.md              # Guía de inicio rápido
├── aplicacion_carga_datos/            # Aplicaciones auxiliares
│   └── ...                          # Otros archivos de aplicación
├── excel/                            # Datos fuente
│   └── proyecto/
│       └── *.xlsx                    # Archivos Excel con datos
├── word/                             # Plantillas y salida
│   └── anexos/
│       └── Plantilla_Anexo_3.docx   # Plantilla Word
├── funciones_excel/                  # Funciones auxiliares Excel
└── artecoin_venv/                    # Entorno virtual Python
```

## Inicio Rápido

### 1. Configuración del Entorno

```powershell
# Clonar el repositorio
git clone <url-del-repositorio>
cd artecoin_automatizaciones

# Activar entorno virtual
.\artecoin_venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt
```

### 2. Ejecutar Generación de Documentos

```powershell
# Navegar a la carpeta anexos
cd anexos

# Ejecutar el script principal
python crear_anexo_3.py
```

### 3. Validar con Tests

```powershell
# Navegar a la carpeta de tests
cd tests

# Ejecutar todos los tests
python run_tests.py -v

# Configuración automática (Windows)
.\setup_tests.ps1 -Verbose
```

## Funcionalidades Detalladas

### Generación de Documentos

El sistema procesa archivos Excel con datos energéticos y genera documentos Word individuales por edificio:

- Entrada: Archivo Excel con múltiples hojas (Clima, SistCC, Eleva, etc.)
- Procesamiento: Filtrado por edificio, cálculo de totales, limpieza de datos
- Salida: Documentos Word individuales con nombres limpios y datos específicos

### Funciones Principales

#### `clean_filename(filename)`

Limpia nombres de archivo para compatibilidad con Windows:

```python
from anexos.crear_anexo_3 import clean_filename

# Ejemplo de uso
nombre_limpio = clean_filename('Archivo "con comillas" <problemático>.docx')
# Resultado: 'Archivo con comillas problemático.docx'
```

Características:
- Elimina caracteres inválidos: `< > : " | ? * \ /`
- Maneja espacios múltiples y guiones bajos
- Trunca nombres muy largos (>100 caracteres)
- Compatible con limitaciones de Windows

#### `get_totales_edificio(df_full, df_edificio, nombre_seccion)`

Calcula totales específicos por edificio:

```python
from anexos.crear_anexo_3 import get_totales_edificio
import pandas as pd

# Ejemplo de uso
df_completo = pd.read_excel('datos.xlsx')
df_edificio = df_completo[df_completo['EDIFICIO'] == 'EDIFICIO A']
totales = get_totales_edificio(df_completo, df_edificio, "Sección")

# Resultado: {'EDIFICIO': 'Total general', 'Potencia (W)': '300', 'Consumo': '130'}
```

Características:
- Suma automática de columnas numéricas
- Formato inteligente (enteros vs decimales)
- Manejo robusto de tipos de datos
- Retorna diccionario con totales formateados

## Testing

### Ejecutar Tests

```powershell
# Tests básicos
python run_tests.py

# Tests detallados
python run_tests.py -v

# Test específico
python run_tests.py -t TestCrearAnexo3.test_clean_filename_basic

# Configuración automática
.\setup_tests.ps1
```

### Cobertura de Tests

- **12+ tests unitarios** cubriendo funciones principales
- **Tests de integración** para creación real de archivos
- **Casos extremos** (archivos vacíos, caracteres inválidos)
- **Validación de tipos** y formatos de salida
- **Tests de rendimiento** y límites

## Dependencias Principales

| Dependencia | Versión | Propósito |
|-------------|---------|-----------|
| **pandas** | ≥2.3.1 | Manipulación de datos y Excel |
| **docxtpl** | ≥0.20.1 | Plantillas Word y generación |
| **pywin32** | ≥310 | Automatización Word en Windows |
| **openpyxl** | ≥3.1.5 | Lectura de archivos Excel |
| **setuptools** | 80.9.0 | Compatibilidad y builds |

Ver `requirements.txt` para la lista completa de dependencias.

## Solución de Problemas Comunes

### Error: "No se puede crear archivo"
Causa: Caracteres inválidos en nombres  
Solución: La función `clean_filename()` los maneja automáticamente

### Error: "KeyError en columna"
Causa: Estructura de Excel diferente  
Solución: Verificar nombres de columnas en el archivo Excel fuente

### Error: "Word no responde"
Causa: Proceso Word colgado  
Solución:
```powershell
Get-Process -Name "WINWORD" | Stop-Process -Force
```

### Tests fallan con import error
Causa: Entorno virtual no activado  
Solución:
```powershell
.\artecoin_venv\Scripts\Activate.ps1
cd tests
python run_tests.py
```

## Documentación Adicional

- **[README_TESTS.md](tests/README_TESTS.md)** - Documentación completa de tests
- **[INICIO_RAPIDO.md](tests/INICIO_RAPIDO.md)** - Guía de inicio rápido detallada
- **Código fuente** - Comentarios detallados en `crear_anexo_3.py`

## Flujo de Trabajo Típico

1. Preparar datos: Colocar archivo Excel en `excel/proyecto/`
2. Verificar plantilla: Asegurar que `word/anexos/Plantilla_Anexo_3.docx` existe
3. Ejecutar script: `python anexos/crear_anexo_3.py`
4. Validar salida: Revisar documentos generados en la carpeta `anexos/`
5. Ejecutar tests: `python tests/run_tests.py -v`

## Contribuir

1. Fork el repositorio
2. Crear branch para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Ejecutar tests para asegurar compatibilidad
4. Commit tus cambios (`git commit -am 'Añadir nueva funcionalidad'`)
5. Push al branch (`git push origin feature/nueva-funcionalidad`)
6. Crear Pull Request

### Estándares de Código

- Tests requeridos para nuevas funciones
- Documentación actualizada en README y docstrings
- Compatibilidad Windows verificada
- Nombres descriptivos para variables y funciones

## Soporte y Contacto

- Issues: Usar el sistema de issues de GitHub
- Tests: Ejecutar `python run_tests.py -v` para diagnóstico
- Logs: Revisar output de consola para errores detallados
- Configuración: Usar `setup_tests.ps1` para configuración automática

---

## Licencia

Este proyecto es de uso interno para automatización de procesos energéticos y documentales.

---

Ejecuta `python anexos/crear_anexo_3.py` para comenzar.