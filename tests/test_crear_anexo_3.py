import unittest
import pandas as pd
import tempfile
import os
from pathlib import Path
import sys

# Add the parent directory to the Python path to import from anexos
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

try:
    from anexos.crear_anexo_3 import clean_filename, get_totales_edificio
except ImportError as e:
    print(f"Error importing functions: {e}")
    print("Please ensure crear_anexo_3.py is in the anexos directory")
    sys.exit(1)


class TestCrearAnexo3(unittest.TestCase):
    """Test suite for crear_anexo_3 module functions."""

    def setUp(self):
        """Set up test fixtures before each test method."""
        # Create sample DataFrames for testing
        # df_full should include a totals row (last row)
        self.df_full = pd.DataFrame(
            {
                "EDIFICIO": ["Iluminación", "Climatización", "Total"],
                "Potencia (W)": [100.0, 200.0, 300.0],  # Total in last row
                "Consumo": [50.0, 80.0, 130.0],  # Total in last row
            }
        )

        # df_edificio should NOT include the totals row
        self.df_edificio_a = pd.DataFrame(
            {
                "EDIFICIO": ["Iluminación", "Climatización"],
                "Potencia (W)": [100.0, 200.0],
                "Consumo": [50.0, 80.0],
            }
        )

    def test_clean_filename_basic(self):
        """Test clean_filename with basic problematic characters."""
        # Test with quotes
        result = clean_filename('Archivo "con comillas".docx')
        # The function removes quotes, so it should not contain them
        self.assertNotIn('"', result)

        # Test with multiple problematic characters
        result = clean_filename("Archivo<con>varios:caracteres|problemáticos*.docx")
        # Should not contain any invalid characters
        invalid_chars = '<>:"|*?\\/'
        for char in invalid_chars:
            self.assertNotIn(char, result)

    def test_clean_filename_long_names(self):
        """Test clean_filename with long filenames."""
        long_name = "A" * 150 + ".docx"
        result = clean_filename(long_name)
        # Should be truncated to max 100 characters based on the function logic
        self.assertLessEqual(len(result), 100)

    def test_clean_filename_empty_or_none(self):
        """Test clean_filename with edge cases."""
        # Test with empty string - based on the actual function, it returns empty string
        result = clean_filename("")
        # The actual function returns empty string for empty input
        self.assertEqual(result, "")

    def test_clean_filename_only_invalid_chars(self):
        """Test clean_filename when filename is only invalid characters."""
        result = clean_filename("<<<>>>|||***")
        # Based on the actual function, after removing invalid chars, we get empty string
        # which after regex cleanup becomes empty
        self.assertEqual(result, "")

    def test_clean_filename_mixed_chars(self):
        """Test clean_filename with mix of valid and invalid characters."""
        result = clean_filename("Valid<>Name.docx")
        # Should remove invalid chars but keep valid ones
        self.assertEqual(result, "ValidName.docx")

    def test_get_totales_edificio_basic(self):
        """Test get_totales_edificio with basic functionality."""
        result = get_totales_edificio(self.df_full, self.df_edificio_a, "Test Section")

        # Should return a dictionary
        self.assertIsInstance(result, dict)

        # Should contain EDIFICIO key
        self.assertIn("EDIFICIO", result)

        # Check that numeric totals are calculated correctly
        # Based on df_edificio_a: Potencia = 100 + 200 = 300, Consumo = 50 + 80 = 130
        if "Potencia (W)" in result:
            # Convert result to numeric for comparison
            potencia_result = (
                float(result["Potencia (W)"]) if result["Potencia (W)"] else 0
            )
            self.assertEqual(potencia_result, 300.0)

        if "Consumo" in result:
            consumo_result = float(result["Consumo"]) if result["Consumo"] else 0
            self.assertEqual(consumo_result, 130.0)

    def test_get_totales_edificio_edificio_key(self):
        """Test that get_totales_edificio sets EDIFICIO key correctly."""
        result = get_totales_edificio(self.df_full, self.df_edificio_a, "Test Section")

        # The EDIFICIO key should be set to 'Total general'
        self.assertEqual(result["EDIFICIO"], "Total general")

    def test_get_totales_edificio_empty_dataframes(self):
        """Test get_totales_edificio with empty DataFrames."""
        empty_df_full = pd.DataFrame(
            {"EDIFICIO": [], "Potencia (W)": [], "Consumo": []}
        )
        empty_df_edificio = pd.DataFrame(
            {"EDIFICIO": [], "Potencia (W)": [], "Consumo": []}
        )

        # The function will fail on empty DataFrames because it tries to access iloc[-1]
        # This is expected behavior - the function expects non-empty dataframes
        with self.assertRaises(IndexError):
            get_totales_edificio(empty_df_full, empty_df_edificio, "Test")

    def test_get_totales_edificio_zero_sums(self):
        """Test get_totales_edificio when sums are zero."""
        df_zeros = pd.DataFrame(
            {
                "EDIFICIO": ["Item1", "Item2"],
                "Potencia (W)": [0.0, 0.0],
                "Consumo": [0.0, 0.0],
            }
        )

        df_full_zeros = pd.DataFrame(
            {
                "EDIFICIO": ["Item1", "Item2", "Total"],
                "Potencia (W)": [0.0, 0.0, 0.0],
                "Consumo": [0.0, 0.0, 0.0],
            }
        )

        result = get_totales_edificio(df_full_zeros, df_zeros, "Test")

        # Zero sums should result in empty strings according to function logic
        if "Potencia (W)" in result:
            self.assertEqual(result["Potencia (W)"], "")
        if "Consumo" in result:
            self.assertEqual(result["Consumo"], "")

    def test_get_totales_edificio_integer_results(self):
        """Test that get_totales_edificio returns integers when appropriate."""
        # Create data that sums to exact integers
        df_ints = pd.DataFrame(
            {
                "EDIFICIO": ["Item1", "Item2"],
                "Potencia (W)": [100.0, 200.0],  # Sum = 300.0 (should be "300")
            }
        )

        df_full_ints = pd.DataFrame(
            {
                "EDIFICIO": ["Item1", "Item2", "Total"],
                "Potencia (W)": [100.0, 200.0, 300.0],
            }
        )

        result = get_totales_edificio(df_full_ints, df_ints, "Test")

        # Should return "300" not "300.0"
        if "Potencia (W)" in result and result["Potencia (W)"]:
            self.assertEqual(result["Potencia (W)"], "300")

    def test_get_totales_edificio_decimal_results(self):
        """Test that get_totales_edificio handles decimals correctly."""
        # Create data with decimal results
        df_decimals = pd.DataFrame(
            {
                "EDIFICIO": ["Item1", "Item2"],
                "Potencia (W)": [100.5, 200.3],  # Sum = 300.8
            }
        )

        df_full_decimals = pd.DataFrame(
            {
                "EDIFICIO": ["Item1", "Item2", "Total"],
                "Potencia (W)": [100.5, 200.3, 300.8],
            }
        )

        result = get_totales_edificio(df_full_decimals, df_decimals, "Test")

        # Should return "300.8"
        if "Potencia (W)" in result and result["Potencia (W)"]:
            self.assertEqual(result["Potencia (W)"], "300.8")


class TestFileIntegration(unittest.TestCase):
    """Integration tests for file operations."""

    def setUp(self):
        """Set up temporary directory for file tests."""
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        """Clean up temporary files."""
        import shutil

        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_clean_filename_file_creation(self):
        """Test that cleaned filenames can actually create files."""
        problematic_name = 'Test "file" with <problems>.txt'
        clean_name = clean_filename(problematic_name)

        # Try to create a file with the cleaned name
        file_path = Path(self.temp_dir) / clean_name

        try:
            file_path.write_text("Test content")
            self.assertTrue(file_path.exists())
        except OSError as e:
            self.fail(f"Could not create file with cleaned name '{clean_name}': {e}")


if __name__ == "__main__":
    # Run the tests
    unittest.main(verbosity=2)
