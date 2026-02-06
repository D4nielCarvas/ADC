# -*- coding: utf-8 -*-
"""
Unit tests for ADC bug fixes and core functionality.
Tests encoding handling, empty sheet names, and DataFrame size validation.
"""
import sys
import os
import pandas as pd
import unittest
import time

sys.path.append(os.path.join(os.getcwd(), 'src'))
from core.cleaner import ADCLogic

class TestBugFixes(unittest.TestCase):
    """Test suite for ADC bug fixes and edge cases."""
    
    FILE_NAME = "test_dummy.xlsx"
    WIDE_FILE = "test_wide.xlsx"

    @classmethod
    def setUpClass(cls):
        """Create test Excel files with multiple sheets."""
        df1 = pd.DataFrame({'A': [1, 2, 3]})
        df2 = pd.DataFrame({'B': [4, 5, 6]})
        
        # Use context manager to ensure proper file closing
        with pd.ExcelWriter(cls.FILE_NAME, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sheet1', index=False)
            df2.to_excel(writer, sheet_name='Sheet2', index=False)
        
        # Small delay to ensure file is fully written
        time.sleep(0.1)

    @classmethod
    def tearDownClass(cls):
        """Clean up test files."""
        # Wait a bit to ensure all file handles are released
        time.sleep(0.2)
        
        for filename in [cls.FILE_NAME, cls.WIDE_FILE]:
            if os.path.exists(filename):
                try:
                    os.remove(filename)
                except PermissionError:
                    # If file is still locked, wait and retry
                    time.sleep(0.5)
                    try:
                        os.remove(filename)
                    except:
                        print(f"Warning: Could not remove {filename}")

    def test_empty_aba_loads_first_sheet(self):
        """Test that empty sheet name ('') loads the first available sheet."""
        logic = ADCLogic()
        print("\n[TEST] Testing empty aba loading...")
        
        # aba="" should automatically load Sheet1
        df = logic.carregar_planilha(self.FILE_NAME, aba="")
        
        self.assertIn('A', df.columns, "Expected column 'A' in first sheet")
        print("[PASS] Empty aba loaded first sheet successfully.")

    def test_gerar_resumo_small_df(self):
        """Test that gerar_resumo handles small DataFrames gracefully."""
        logic = ADCLogic()
        print("\n[TEST] Testing gerar_resumo with small dataframe...")
        
        # Sheet1 has only 1 column, should trigger error in result
        res = logic.gerar_resumo(self.FILE_NAME, aba="Sheet1")
        
        self.assertIn("erro", res, "Expected error key in result for small DataFrame")
        print(f"[PASS] Properly caught small dataframe error: {res['erro']}")

    def test_gerar_resumo_large_df(self):
        """Test that gerar_resumo processes large DataFrames correctly."""
        logic = ADCLogic()
        print("\n[TEST] Testing gerar_resumo with large dataframe...")
        
        # Create a DataFrame with 30 columns (more than required 27)
        cols = {f"col_{i}": [10] * 5 for i in range(30)}
        
        # Use context manager to ensure proper file closing
        with pd.ExcelWriter(self.WIDE_FILE, engine='openpyxl') as writer:
            pd.DataFrame(cols).to_excel(writer, sheet_name='Sheet1', index=False)
        
        time.sleep(0.1)  # Ensure file is written
        
        try:
            res = logic.gerar_resumo(self.WIDE_FILE, aba="Sheet1")
            
            self.assertNotIn("erro", res, "Should not have error for large DataFrame")
            self.assertEqual(res["total_pedidos"], 1, "Expected 1 unique order (all values are 10)")
            print("[PASS] Large dataframe processed successfully.")
        finally:
            # Ensure file is closed before cleanup
            time.sleep(0.1)

if __name__ == '__main__':
    # Run tests with verbose output
    unittest.main(verbosity=2)
