# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 15:45:35 2024

@author: pizzacoin
"""

# %%

from editors import ExcelEditor
from scripts import script_mapping

# %%

def run_macro(file_path, macro_name):
    excel = ExcelEditor(file_path, debug=True)
    excel.run_macro(macro_name)
    excel.close_workbook()

def add_macro_and_run(file_path, macro_name):
    excel = ExcelEditor(file_path, debug=True)
    vba_code = script_mapping[macro_name]["code"]
    excel.add_vba_to_excel(macro_name, vba_code)
    excel.run_macro(macro_name)
    excel.close_workbook()
    
# %%

