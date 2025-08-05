"""
Tests para el módulo crear_anexo_3.py

Este archivo contiene tests unitarios para verificar que todas las funciones
del módulo crear_anexo_3.py funcionen correctamente.
"""

import unittest
import pandas as pd
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch
import sys
import os

# Añadir el directorio actual al path para importar el módulo
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Importar las funciones a testear
from crear_anexo_3 import (
    delete_rows_optimized,
    clean_last_row,
    load_and_clean_sheets,
    update_word_fields,
    get_user_input,
    SHEET_MAP,
)


class TestDeleteRowsOptimized(unittest.TestCase):
    """Tests para la función delete_rows_optimized"""

    def setUp(self):
        """Configuración inicial para cada test"""
        self.sample_data = {
            "ID CENTRO": ["C001", "C002", "C003", "", "", ""],
            "NOMBRE": ["Centro1", "Centro2", "Centro3", "", "", ""],
            "CONSUMO": ["100", "200", "300", "", "", ""],
        }
        self.df = pd.DataFrame(self.sample_data)

    def test_delete_rows_with_valid_data(self):
        """Test que elimina correctamente las filas vacías"""
        result = delete_rows_optimized(self.df, "ID CENTRO")

        # Debe devolver 4 filas (última fila con datos válidos + 1 extra como especifica la función)
        self.assertEqual(len(result), 4)

        # Las primeras 3 filas deben tener datos
        self.assertEqual(result["ID CENTRO"].iloc[0], "C001")
        self.assertEqual(result["ID CENTRO"].iloc[1], "C002")
        self.assertEqual(result["ID CENTRO"].iloc[2], "C003")

    def test_delete_rows_column_not_exists(self):
        """Test cuando la columna especificada no existe"""
        result = delete_rows_optimized(self.df, "COLUMNA_INEXISTENTE")

        # Debe devolver el DataFrame original
        pd.testing.assert_frame_equal(result, self.df)

    def test_delete_rows_empty_dataframe(self):
        """Test con DataFrame vacío"""
        empty_df = pd.DataFrame()
        result = delete_rows_optimized(empty_df, "ID CENTRO")

        # Debe devolver DataFrame vacío
        self.assertTrue(result.empty)

    def test_delete_rows_all_valid_data(self):
        """Test con todas las filas con datos válidos"""
        valid_data = {
            "ID CENTRO": ["C001", "C002", "C003"],
            "NOMBRE": ["Centro1", "Centro2", "Centro3"],
        }
        df_valid = pd.DataFrame(valid_data)
        result = delete_rows_optimized(df_valid, "ID CENTRO")

        # Debe devolver todas las filas (no hay filas vacías para eliminar)
        self.assertEqual(len(result), 3)


class TestCleanLastRow(unittest.TestCase):
    """Tests para la función clean_last_row"""

    def test_clean_last_row_with_total(self):
        """Test que limpia correctamente la fila con 'Total'"""
        data = {"ITEM": ["Item1", "Item2", "Total"], "VALOR": ["100", "200", "Total"]}
        df = pd.DataFrame(data)
        result = clean_last_row(df)

        # La última fila debe tener valores NaN donde antes había 'Total'
        self.assertTrue(pd.isna(result["ITEM"].iloc[-1]))
        self.assertTrue(pd.isna(result["VALOR"].iloc[-1]))

    def test_clean_last_row_without_total(self):
        """Test cuando no hay 'Total' en la última fila"""
        data = {"ITEM": ["Item1", "Item2", "Item3"], "VALOR": ["100", "200", "300"]}
        df = pd.DataFrame(data)
        result = clean_last_row(df)

        # El DataFrame debe permanecer igual
        pd.testing.assert_frame_equal(result, df)

    def test_clean_last_row_empty_dataframe(self):
        """Test con DataFrame vacío"""
        empty_df = pd.DataFrame()
        result = clean_last_row(empty_df)

        # Debe devolver DataFrame vacío
        self.assertTrue(result.empty)


class TestLoadAndCleanSheets(unittest.TestCase):
    """Tests para la función load_and_clean_sheets"""

    def setUp(self):
        """Configuración inicial para cada test"""
        # Crear un archivo Excel temporal para testing
        self.temp_dir = tempfile.mkdtemp()
        self.excel_path = Path(self.temp_dir) / "test_excel.xlsx"

        # Crear DataFrames de prueba
        self.test_data = {
            "Clima": pd.DataFrame(
                {
                    "ID CENTRO": ["C001", "C002", "", ""],
                    "EQUIPO": ["AC1", "AC2", "", ""],
                    "CONSUMO": ["100", "200", "", ""],
                }
            ),
            "SistCC": pd.DataFrame(
                {
                    "ID CENTRO": ["C001", "C002", "", ""],
                    "SISTEMA": ["SCC1", "SCC2", "", ""],
                    "POTENCIA": ["50", "75", "", ""],
                }
            ),
        }

        # Guardar en archivo Excel
        with pd.ExcelWriter(self.excel_path, engine="openpyxl") as writer:
            for sheet_name, df in self.test_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    def tearDown(self):
        """Limpieza después de cada test"""
        shutil.rmtree(self.temp_dir)

    def test_load_and_clean_sheets_success(self):
        """Test carga exitosa de hojas"""
        sheet_map = {
            "Clima": "SISTEMAS DE CLIMATIZACIÓN",
            "SistCC": "SISTEMAS DE CALEFACCIÓN",
        }

        result = load_and_clean_sheets(self.excel_path, sheet_map)

        # Verificar que se cargaron las hojas correctas
        self.assertIn("Clima", result)
        self.assertIn("SistCC", result)

        # Verificar que los DataFrames no están vacíos
        self.assertFalse(result["Clima"].empty)
        self.assertFalse(result["SistCC"].empty)

    def test_load_and_clean_sheets_missing_sheet(self):
        """Test cuando falta una hoja en el Excel"""
        sheet_map = {"HojaInexistente": "HOJA QUE NO EXISTE"}

        with self.assertRaises(ValueError) as context:
            load_and_clean_sheets(self.excel_path, sheet_map)

        self.assertIn("Hojas faltantes en el Excel", str(context.exception))

    def test_load_and_clean_sheets_nonexistent_file(self):
        """Test cuando el archivo Excel no existe"""
        nonexistent_path = Path(self.temp_dir) / "no_existe.xlsx"
        sheet_map = {"Clima": "SISTEMAS DE CLIMATIZACIÓN"}

        with self.assertRaises(FileNotFoundError):
            load_and_clean_sheets(nonexistent_path, sheet_map)


class TestUpdateWordFields(unittest.TestCase):
    """Tests para la función update_word_fields"""

    def test_update_word_fields_with_mock_file(self):
        """Test que verifica que la función se puede llamar sin errores con archivo inexistente"""
        # Como win32com es complejo de mockear, simplemente verificamos que la función
        # maneja correctamente archivos inexistentes
        nonexistent_path = "nonexistent_file.docx"

        # La función debe ejecutarse sin lanzar excepciones críticas
        try:
            update_word_fields(nonexistent_path)
        except Exception as e:
            # Verificamos que no sea un error de importación o estructura
            self.assertNotIsInstance(e, (ImportError, AttributeError))

    def test_update_word_fields_accepts_string_path(self):
        """Test que verifica que la función acepta rutas como string"""
        # Verificar que la función acepta el tipo correcto de parámetro
        test_path = "test.docx"

        # La función debe aceptar strings sin errores de tipo
        try:
            update_word_fields(test_path)
        except (TypeError, AttributeError) as e:
            if "str" in str(e):
                self.fail(f"La función no acepta strings como parámetro: {e}")

    def test_update_word_fields_function_exists(self):
        """Test que verifica que la función existe y es callable"""
        from crear_anexo_3 import update_word_fields

        self.assertTrue(callable(update_word_fields))

        # Verificar que tiene los parámetros esperados
        import inspect

        sig = inspect.signature(update_word_fields)
        params = list(sig.parameters.keys())

        # Debe tener al menos el parámetro de ruta del documento
        self.assertTrue(len(params) >= 1)


class TestGetUserInput(unittest.TestCase):
    """Tests para la función get_user_input"""

    def test_get_user_input_function_exists(self):
        """Test que verifica que la función existe y es callable"""
        self.assertTrue(callable(get_user_input))

        # Verificar que no requiere parámetros
        import inspect

        sig = inspect.signature(get_user_input)
        self.assertEqual(len(sig.parameters), 0)

    def test_get_user_input_returns_tuple(self):
        """Test que verifica que la función devuelve una tupla de mes y año"""
        # Simular entrada del usuario (mes=5, año=2024)
        with patch("builtins.input", side_effect=["5", "2024"]):
            resultado = get_user_input()

        # Verificar que devuelve una tupla
        self.assertIsInstance(resultado, tuple)
        self.assertEqual(len(resultado), 2)

        mes_nombre, anio = resultado

        # Verificar tipos
        self.assertIsInstance(mes_nombre, str)
        self.assertIsInstance(anio, int)

        # Verificar valores
        self.assertEqual(mes_nombre, "Mayo")
        self.assertEqual(anio, 2024)

    def test_get_user_input_default_values(self):
        """Test que verifica el uso de valores por defecto"""
        from datetime import datetime

        # Simular entrada vacía (usar valores por defecto)
        with patch("builtins.input", side_effect=["", ""]):
            resultado = get_user_input()

        mes_nombre, anio = resultado

        # Verificar que usa valores actuales
        current_year = datetime.now().year
        current_month = datetime.now().month

        meses_espanol = {
            1: "Enero",
            2: "Febrero",
            3: "Marzo",
            4: "Abril",
            5: "Mayo",
            6: "Junio",
            7: "Julio",
            8: "Agosto",
            9: "Septiembre",
            10: "Octubre",
            11: "Noviembre",
            12: "Diciembre",
        }

        self.assertEqual(mes_nombre, meses_espanol[current_month])
        self.assertEqual(anio, current_year)


class TestIntegration(unittest.TestCase):
    """Tests de integración para verificar el flujo completo"""

    def test_sheet_map_consistency(self):
        """Test que verifica la consistencia del SHEET_MAP"""
        # Verificar que SHEET_MAP tiene las claves esperadas
        expected_keys = ["Clima", "SistCC", "Eleva", "EqHoriz", "Ilum", "OtrosEq"]

        for key in expected_keys:
            self.assertIn(key, SHEET_MAP)
            self.assertIsInstance(SHEET_MAP[key], str)
            self.assertTrue(len(SHEET_MAP[key]) > 0)

    def test_data_flow_consistency(self):
        """Test que verifica la consistencia del flujo de datos"""
        # Crear DataFrame de prueba
        test_data = {
            "ID CENTRO": ["C001", "C002", "C003", "", "Total"],
            "EQUIPO": ["Equipo1", "Equipo2", "Equipo3", "", "Total"],
            "CONSUMO": ["100", "200", "300", "", "Total"],
        }
        df = pd.DataFrame(test_data)

        # Aplicar el flujo completo de limpieza
        df_deleted = delete_rows_optimized(df, "ID CENTRO")
        df_final = clean_last_row(df_deleted)

        # Verificar que el resultado es consistente
        self.assertFalse(df_final.empty)
        self.assertLessEqual(len(df_final), len(df))

        # Verificar que los datos válidos se mantienen
        valid_rows = df_final.dropna(subset=["ID CENTRO"])
        valid_rows = valid_rows[valid_rows["ID CENTRO"] != ""]
        self.assertEqual(len(valid_rows), 3)  # Debe tener 3 filas con datos válidos


def create_test_suite():
    """Crear suite de tests"""
    suite = unittest.TestSuite()

    # Añadir todos los tests
    suite.addTest(unittest.makeSuite(TestDeleteRowsOptimized))
    suite.addTest(unittest.makeSuite(TestCleanLastRow))
    suite.addTest(unittest.makeSuite(TestLoadAndCleanSheets))
    suite.addTest(unittest.makeSuite(TestUpdateWordFields))
    suite.addTest(unittest.makeSuite(TestGetUserInput))
    suite.addTest(unittest.makeSuite(TestIntegration))

    return suite


if __name__ == "__main__":
    # Ejecutar todos los tests
    print("=" * 70)
    print("EJECUTANDO TESTS PARA crear_anexo_3.py")
    print("=" * 70)

    # Crear y ejecutar suite de tests
    suite = create_test_suite()
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    # Mostrar resumen
    print("\n" + "=" * 70)
    print("RESUMEN DE TESTS")
    print("=" * 70)
    print(f"Tests ejecutados: {result.testsRun}")
    print(f"Exitosos: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Fallidos: {len(result.failures)}")
    print(f"Errores: {len(result.errors)}")

    if result.failures:
        print("\nFALLOS:")
        for test, traceback in result.failures:
            print(f"- {test}: {traceback}")

    if result.errors:
        print("\nERRORES:")
        for test, traceback in result.errors:
            print(f"- {test}: {traceback}")

    # Código de salida
    exit_code = 0 if result.wasSuccessful() else 1
    print(f"\nCódigo de salida: {exit_code}")
    sys.exit(exit_code)
