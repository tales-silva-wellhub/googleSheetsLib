import unittest
from src.utils import (
    column_to_number, number_to_column, 
    validate_xrange, xrange_to_grid_range, 
    validate_grid_range, grid_range_to_xrange
)

class TestUtils(unittest.TestCase):

    # --- Testes de Conversão de Coluna ---
    def test_column_to_number(self):
        self.assertEqual(column_to_number("A"), 0)
        self.assertEqual(column_to_number("Z"), 25)
        self.assertEqual(column_to_number("AA"), 26)
        self.assertEqual(column_to_number("AB"), 27)

    def test_number_to_column(self):
        self.assertEqual(number_to_column(0), "A")
        self.assertEqual(number_to_column(25), "Z")
        self.assertEqual(number_to_column(26), "AA")
        self.assertEqual(number_to_column(27), "AB")

    # --- Testes de Validação de XRANGE ---
    def test_validate_xrange_valid(self):
        self.assertTrue(validate_xrange("A1:B2"))
        self.assertTrue(validate_xrange("A1"))
        self.assertTrue(validate_xrange("A:Z"))
        self.assertTrue(validate_xrange("1:10"))
        self.assertTrue(validate_xrange("Planilha1!A1:C3")) # Teste com sheet name

    def test_validate_xrange_invalid(self):
        self.assertFalse(validate_xrange("B2:A1"))      # Ordem invertida
        self.assertFalse(validate_xrange("Z:A"))       # Coluna invertida
        self.assertFalse(validate_xrange("10:1"))      # Linha invertida
        self.assertFalse(validate_xrange("A0"))        # Linha zero não existe
        self.assertFalse(validate_xrange("Célula"))    # Lixo

    # --- Testes de GridRange ---
    def test_xrange_to_grid_range(self):
        # A1:B2 -> Rows 0-2 (exclusivo), Cols 0-2 (exclusivo)
        expected = {
            "startRowIndex": 0, "endRowIndex": 2,
            "startColumnIndex": 0, "endColumnIndex": 2
        }
        self.assertEqual(xrange_to_grid_range("A1:B2"), expected)

    def test_validate_grid_range(self):
        valid_grid = {
            "startRowIndex": 0, "endRowIndex": 10,
            "startColumnIndex": 0, "endColumnIndex": 5
        }
        self.assertTrue(validate_grid_range(valid_grid))
        
        # Erro de lógica (start > end)
        invalid_grid = valid_grid.copy()
        invalid_grid["startRowIndex"] = 20
        self.assertFalse(validate_grid_range(invalid_grid))

    def test_grid_range_to_xrange(self):
        valid_grid_range = {'startRowIndex': 0,
            'endRowIndex': 10,
            'startColumnIndex': 0,
            'endColumnIndex': 5}
        self.assertEqual(grid_range_to_xrange(valid_grid_range), 'A1:E10')

if __name__ == '__main__':
    unittest.main()