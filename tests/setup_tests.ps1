# Script de configuración para tests - Crear Anexo 3
# Este script prepara el entorno de testing en Windows

param(
    [switch]$Verbose,
    [switch]$Force
)

Write-Host "🔧 CONFIGURACIÓN DE TESTS - CREAR ANEXO 3" -ForegroundColor Cyan
Write-Host "=" * 50

# Variables de configuración
$ProjectRoot = Split-Path $PSScriptRoot -Parent
$VenvPath = Join-Path $ProjectRoot "artecoin_venv"
$RequirementsPath = Join-Path $ProjectRoot "requirements.txt"
$CurrentDir = $PSScriptRoot

# Función para escribir mensajes
function Write-Step {
    param([string]$Message, [string]$Color = "Yellow")
    Write-Host "📋 $Message" -ForegroundColor $Color
}

function Write-Success {
    param([string]$Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function Write-Error {
    param([string]$Message)
    Write-Host "❌ $Message" -ForegroundColor Red
}

function Write-Warning {
    param([string]$Message)
    Write-Host "⚠️  $Message" -ForegroundColor Yellow
}

# 1. Verificar PowerShell ExecutionPolicy
Write-Step "Verificando ExecutionPolicy de PowerShell..."
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -eq "Restricted") {
    Write-Warning "ExecutionPolicy está restringida. Intentando cambiar..."
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Success "ExecutionPolicy actualizada a RemoteSigned"
    }
    catch {
        Write-Error "No se pudo cambiar ExecutionPolicy. Ejecuta como administrador."
        Write-Host "Comando manual: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"
        if (-not $Force) { exit 1 }
    }
} else {
    Write-Success "ExecutionPolicy OK: $currentPolicy"
}

# 2. Verificar entorno virtual
Write-Step "Verificando entorno virtual..."
if (Test-Path $VenvPath) {
    Write-Success "Entorno virtual encontrado: $VenvPath"
} else {
    Write-Error "Entorno virtual no encontrado en: $VenvPath"
    Write-Host "Crear entorno virtual manualmente:"
    Write-Host "  python -m venv artecoin_venv"
    if (-not $Force) { exit 1 }
}

# 3. Activar entorno virtual
Write-Step "Activando entorno virtual..."
$ActivateScript = Join-Path $VenvPath "Scripts\Activate.ps1"
if (Test-Path $ActivateScript) {
    try {
        & $ActivateScript
        Write-Success "Entorno virtual activado"
    }
    catch {
        Write-Warning "Error activando entorno virtual: $_"
        Write-Host "Actívalo manualmente: .\artecoin_venv\Scripts\Activate.ps1"
    }
} else {
    Write-Error "Script de activación no encontrado: $ActivateScript"
}

# 4. Verificar Python
Write-Step "Verificando Python..."
try {
    $PythonVersion = python --version 2>&1
    Write-Success "Python disponible: $PythonVersion"
} catch {
    Write-Error "Python no encontrado en PATH"
    Write-Host "Asegúrate de que Python esté instalado y en PATH"
    if (-not $Force) { exit 1 }
}

# 5. Verificar pip
Write-Step "Verificando pip..."
try {
    $PipVersion = pip --version 2>&1
    Write-Success "pip disponible: $($PipVersion.Split(' ')[0..2] -join ' ')"
} catch {
    Write-Error "pip no encontrado"
    if (-not $Force) { exit 1 }
}

# 6. Instalar/actualizar dependencias
Write-Step "Instalando/actualizando dependencias..."
if (Test-Path $RequirementsPath) {
    try {
        if ($Verbose) {
            pip install -r $RequirementsPath --verbose
        } else {
            pip install -r $RequirementsPath --quiet
        }
        Write-Success "Dependencias instaladas desde requirements.txt"
    } catch {
        Write-Error "Error instalando dependencias: $_"
        if (-not $Force) { exit 1 }
    }
} else {
    Write-Warning "requirements.txt no encontrado en: $RequirementsPath"
}

# 7. Verificar módulos clave
Write-Step "Verificando módulos clave..."
$KeyModules = @("pandas", "docxtpl", "pywin32", "openpyxl")
foreach ($Module in $KeyModules) {
    try {
        python -c "import $Module; print(f'$Module OK')" 2>$null
        Write-Success "$Module disponible"
    } catch {
        Write-Warning "$Module no disponible"
    }
}

# 8. Verificar archivos de test
Write-Step "Verificando archivos de test..."
$TestFiles = @(
    "test_crear_anexo_3.py",
    "run_tests.py",
    "crear_anexo_3.py"
)

foreach ($File in $TestFiles) {
    $FilePath = Join-Path $CurrentDir $File
    if (Test-Path $FilePath) {
        Write-Success "Archivo encontrado: $File"
    } else {
        Write-Error "Archivo faltante: $File"
    }
}

# 9. Ejecutar tests de verificación
Write-Step "Ejecutando tests de verificación..."
Set-Location $CurrentDir
try {
    python -m unittest test_crear_anexo_3 -v 2>&1 | Out-String -Width 200
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Tests ejecutados correctamente"
    } else {
        Write-Warning "Algunos tests fallaron (código: $LASTEXITCODE)"
    }
} catch {
    Write-Error "Error ejecutando tests: $_"
}

# 10. Resumen final
Write-Host ""
Write-Host "🎯 RESUMEN DE CONFIGURACIÓN" -ForegroundColor Cyan
Write-Host "=" * 30

Write-Host "Directorio de trabajo: $CurrentDir" -ForegroundColor White
Write-Host "Entorno virtual: $VenvPath" -ForegroundColor White
Write-Host ""

Write-Host "✅ PARA EJECUTAR TESTS:" -ForegroundColor Green
Write-Host "  cd `"$CurrentDir`""
Write-Host "  python run_tests.py -v"
Write-Host ""

Write-Host "✅ PARA USAR EL MÓDULO:" -ForegroundColor Green
Write-Host "  python ../anexos/crear_anexo_3.py"
Write-Host ""

Write-Host "📚 DOCUMENTACIÓN:" -ForegroundColor Blue
Write-Host "  README_TESTS.md    - Documentación de tests"
Write-Host "  INICIO_RAPIDO.md   - Guía de inicio rápido"
Write-Host ""

Write-Success "Configuracion completada"