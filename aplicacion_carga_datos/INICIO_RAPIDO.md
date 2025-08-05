# 🚀 INICIO RÁPIDO - TESTS

## Para usuarios con prisa:

### Windows - Opción 1 (Script automático):
```cmd
cd aplicacion_carga_datos
setup_tests.bat
```

### Windows - Opción 2 (PowerShell):
```powershell
cd aplicacion_carga_datos  
.\setup_tests.ps1
```

### Manual (3 pasos):
```powershell
# 1. Activar entorno virtual
cd aplicacion_carga_datos
..\artecoin_venv\Scripts\Activate.ps1

# 2. Verificar/instalar dependencias
python run_tests.py --install-deps

# 3. Ejecutar tests
python run_tests.py -v
```

## Resultado esperado:
```
✅ TODOS LOS TESTS PASARON EXITOSAMENTE
```

📖 **Guía completa**: Ver `README_TESTS.md`
