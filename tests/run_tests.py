#!/usr/bin/env python3
"""
Script para ejecutar todos los tests del módulo crear_anexo_3.
Ejecuta los tests unitarios y muestra un resumen de los resultados.
"""

import unittest
import sys
import os
from pathlib import Path
import argparse


def main():
    """Función principal para ejecutar los tests."""
    parser = argparse.ArgumentParser(
        description="Ejecutar tests para el módulo crear_anexo_3"
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Ejecutar tests en modo verbose"
    )
    parser.add_argument(
        "-f", "--failfast", action="store_true", help="Detener en el primer fallo"
    )
    parser.add_argument(
        "-t",
        "--test",
        type=str,
        help="Ejecutar un test específico (ej: TestCrearAnexo3.test_clean_filename_basic)",
    )

    args = parser.parse_args()

    # Configurar el directorio de trabajo - ahora estamos en tests/
    script_dir = Path(__file__).parent
    project_root = script_dir.parent  # artecoin_automatizaciones/
    os.chdir(script_dir)

    # Añadir el directorio raíz del proyecto al path de Python
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))  # Configurar el test loader
    loader = unittest.TestLoader()

    # Cargar tests
    if args.test:
        # Ejecutar un test específico
        suite = loader.loadTestsFromName(
            args.test, module=__import__("test_crear_anexo_3")
        )
    else:
        # Ejecutar todos los tests
        test_modules = [
            "test_crear_anexo_3",
            # Aquí puedes añadir más módulos de test si los hay
        ]

        suite = unittest.TestSuite()
        for module_name in test_modules:
            try:
                module = __import__(module_name)
                module_suite = loader.loadTestsFromModule(module)
                suite.addTest(module_suite)
                print(f"✓ Cargados tests de {module_name}")
            except ImportError as e:
                print(f"⚠ No se pudo cargar {module_name}: {e}")
                continue

    # Configurar el runner
    verbosity = 2 if args.verbose else 1
    runner = unittest.TextTestRunner(
        verbosity=verbosity, failfast=args.failfast, stream=sys.stdout
    )

    # Ejecutar tests
    print("\n" + "=" * 60)
    print("EJECUTANDO TESTS PARA CREAR_ANEXO_3")
    print("=" * 60)

    result = runner.run(suite)

    # Mostrar resumen
    print("\n" + "=" * 60)
    print("RESUMEN DE RESULTADOS")
    print("=" * 60)

    total_tests = result.testsRun
    failures = len(result.failures)
    errors = len(result.errors)
    skipped = len(result.skipped) if hasattr(result, "skipped") else 0

    print(f"Tests ejecutados: {total_tests}")
    print(f"Exitosos: {total_tests - failures - errors}")
    print(f"Fallos: {failures}")
    print(f"Errores: {errors}")
    print(f"Omitidos: {skipped}")

    if result.failures:
        print("\nFALLOS:")
        for test, traceback in result.failures:
            print(f"  - {test}: {traceback.split(chr(10))[0]}")

    if result.errors:
        print("\nERRORES:")
        for test, traceback in result.errors:
            print(f"  - {test}: {traceback.split(chr(10))[0]}")

    # Código de salida
    if failures or errors:
        print(f"\n❌ Tests FALLARON ({failures + errors} problemas)")
        return 1
    else:
        print("\n✅ Todos los tests PASARON")
        return 0


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
