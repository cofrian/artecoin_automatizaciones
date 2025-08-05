# Script de configuración rápida para ejecutar tests
# Autor: Sistema automatizado  
# Fecha: Agosto 2025

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   CONFIGURACION RAPIDA - TESTS" -ForegroundColor Cyan  
Write-Host "========================================" -ForegroundColor Cyan

# Verificar si estamos en el directorio correcto
if (-not (Test-Path "crear_anexo_3.py")) {
    Write-Host "ERROR: No se encuentra crear_anexo_3.py" -ForegroundColor Red
    Write-Host "Asegurate de ejecutar este script desde el directorio: aplicacion_carga_datos" -ForegroundColor Red
    Read-Host "Presiona Enter para continuar"
    exit 1
}

# Verificar si existe el entorno virtual
if (-not (Test-Path "..\artecoin_venv\Scripts\Activate.ps1")) {
    Write-Host "ERROR: No se encuentra el entorno virtual" -ForegroundColor Yellow
    Write-Host "El entorno virtual debe estar en: ..\artecoin_venv\" -ForegroundColor Yellow
    Write-Host ""
    $crear_env = Read-Host "¿Quieres crear un nuevo entorno virtual? (s/n)"
    if ($crear_env -eq "s") {
        Write-Host "Creando entorno virtual..." -ForegroundColor Green
        Set-Location ..
        python -m venv artecoin_venv
        Set-Location aplicacion_carga_datos
        Write-Host "Entorno virtual creado exitosamente" -ForegroundColor Green
    } else {
        Write-Host "Cancelado por el usuario" -ForegroundColor Red
        Read-Host "Presiona Enter para continuar"
        exit 1
    }
}

Write-Host "Activando entorno virtual..." -ForegroundColor Green
& "..\artecoin_venv\Scripts\Activate.ps1"

Write-Host "Verificando dependencias..." -ForegroundColor Green
python run_tests.py --check

Write-Host ""
$instalar = Read-Host "¿Instalar dependencias faltantes automáticamente? (s/n)"
if ($instalar -eq "s") {
    Write-Host "Instalando dependencias..." -ForegroundColor Green
    python run_tests.py --install-deps
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   EJECUTANDO TESTS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
python run_tests.py -v

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   CONFIGURACION COMPLETADA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Para ejecutar tests en el futuro:" -ForegroundColor Green
Write-Host "1. cd aplicacion_carga_datos" -ForegroundColor White
Write-Host "2. ..\artecoin_venv\Scripts\Activate.ps1" -ForegroundColor White  
Write-Host "3. python run_tests.py -v" -ForegroundColor White
Write-Host ""
Write-Host "NOTA: Las dependencias se instalan desde ..\requirements.txt" -ForegroundColor Yellow
Write-Host ""
Read-Host "Presiona Enter para continuar"
