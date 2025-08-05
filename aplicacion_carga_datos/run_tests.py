#!/usr/bin/env python3
"""
Script para ejecutar todos los tests del proyecto de forma fácil y rápida.

Uso:
    python run_tests.py              # Ejecutar todos los tests
    python run_tests.py -v          # Ejecutar con output verbose
    python run_tests.py --help      # Mostrar ayuda
"""

import argparse
import sys
import subprocess
import os
from pathlib import Path


def run_tests(verbose=False, specific_test=None):
    """Ejecutar los tests del proyecto"""

    # Asegurarse de que estamos en el directorio correcto
    script_dir = Path(__file__).parent
    os.chdir(script_dir)

    print("🧪 EJECUTANDO TESTS PARA CREAR_ANEXO_3.PY")
    print("=" * 60)

    # Comando base
    cmd = [sys.executable, "-m", "unittest"]

    if specific_test:
        cmd.append(specific_test)
    else:
        cmd.append("test_crear_anexo_3")

    if verbose:
        cmd.append("-v")

    try:
        # Ejecutar tests
        result = subprocess.run(cmd, capture_output=True, text=True)

        # Mostrar output
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print("ERRORES:", result.stderr)

        # Mostrar resultado final
        if result.returncode == 0:
            print("✅ TODOS LOS TESTS PASARON EXITOSAMENTE")
        else:
            print("❌ ALGUNOS TESTS FALLARON")

        return result.returncode

    except Exception as e:
        print(f"❌ Error ejecutando tests: {e}")
        return 1


def install_dependencies():
    """Instalar dependencias necesarias para los tests"""
    print("📦 Instalando dependencias desde requirements.txt...")

    requirements_path = Path("../requirements.txt")

    if not requirements_path.exists():
        # Fallback: instalar dependencias individuales
        print("⚠️  requirements.txt no encontrado, instalando dependencias básicas...")
        dependencies = [
            "pandas>=2.3.1",
            "openpyxl>=3.1.5",
            "docxtpl>=0.20.1",
            "pywin32>=310",
        ]

        for dep in dependencies:
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", dep],
                    check=True,
                    capture_output=True,
                )
                print(f"✅ {dep} instalado correctamente")
            except subprocess.CalledProcessError as e:
                print(f"❌ Error instalando {dep}: {e}")
                return False
    else:
        # Instalar desde requirements.txt
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r", str(requirements_path)],
                check=True,
                capture_output=True,
            )
            print("✅ Dependencias instaladas desde requirements.txt")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error instalando desde requirements.txt: {e}")
            return False

    return True


def check_environment():
    """Verificar que el entorno esté configurado correctamente"""
    print("🔍 Verificando entorno...")

    # Verificar que el archivo principal existe
    main_file = Path("crear_anexo_3.py")
    if not main_file.exists():
        print(f"❌ No se encontró {main_file}")
        return False

    # Verificar que el archivo de tests existe
    test_file = Path("test_crear_anexo_3.py")
    if not test_file.exists():
        print(f"❌ No se encontró {test_file}")
        return False

    # Verificar Python
    print(f"✅ Python: {sys.version}")

    # Verificar imports básicos
    try:
        import pandas

        print(f"✅ Pandas: {pandas.__version__}")
    except ImportError:
        print("❌ Pandas no disponible")
        return False

    try:
        import openpyxl

        print(f"✅ Openpyxl: {openpyxl.__version__}")
    except ImportError:
        print("⚠️  Openpyxl no disponible (se instalará automáticamente)")

    return True


def main():
    """Función principal"""
    parser = argparse.ArgumentParser(
        description="Ejecutar tests para crear_anexo_3.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
    python run_tests.py                    # Ejecutar todos los tests
    python run_tests.py -v                # Ejecutar con verbose
    python run_tests.py --install-deps    # Instalar dependencias
    python run_tests.py --check           # Solo verificar entorno
    python run_tests.py -t TestDeleteRowsOptimized  # Test específico
        """,
    )

    parser.add_argument("-v", "--verbose", action="store_true", help="Output verbose")

    parser.add_argument(
        "--install-deps", action="store_true", help="Instalar dependencias necesarias"
    )

    parser.add_argument(
        "--check", action="store_true", help="Solo verificar el entorno"
    )

    parser.add_argument("-t", "--test", type=str, help="Ejecutar un test específico")

    args = parser.parse_args()

    # Instalar dependencias si se solicita
    if args.install_deps:
        if not install_dependencies():
            return 1

    # Verificar entorno
    if not check_environment():
        print("\n❌ El entorno no está configurado correctamente")
        print("💡 Ejecuta: python run_tests.py --install-deps")
        return 1

    # Solo verificar si se solicita
    if args.check:
        print("✅ Entorno configurado correctamente")
        return 0

    # Ejecutar tests
    return run_tests(verbose=args.verbose, specific_test=args.test)


if __name__ == "__main__":
    sys.exit(main())
