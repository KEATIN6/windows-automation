# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 11:08:52 2024

@author: ghostshade
"""

# %%

class ExcelEditorDebugger:
    TITLE = "EXCEL EDITOR V1.0"
    AUTHOR = "ekeating"
    BORDER = "-" * 100
    
    def __init__(self, parent):
        self.parent = parent
        self.print_title()
        
    def print_title(self):
        print(self.BORDER)
        print(self.TITLE)
        print(f"@author: {self.AUTHOR}")
        print()
        print(self.BORDER)
        
    def print_status(self, status):
        status_text = f"""{status}\n{self.BORDER}"""        
        print(status_text)
        
    def load(self):
        return f"Loading workbook from <{self.parent.file_path}>."
        
    def macro_exists(self, message, macro_name):
        return f"Checking {message} the Macro '{macro_name}' Exists>."
        
    def extension(self):
        return "Checking that file extension is Macro-Enabled."
    
    def close(self):
        return f"Closing workbook from <{self.parent.file_path}>."
    
    def run_macro(self, macro_name):
        return f"Attempting to Run Macro: {macro_name}."
    
    def add_macro(self, macro_name):
        return f"Adding VBA Code '{macro_name}' to Excel File."
    
    def convert(self, extension, to_extension):
        return f"Converting from {extension} to .{to_extension}."
        
    def process(self, action, variables=None):
        status = None
        if action == "load":
            status = self.load()
        elif action == "macro_exists":
            message, macro_name = variables
            status = self.macro_exists(message, macro_name)
        elif action == "extension":
            status = self.extension()
        elif action == "close":
            status = self.close()
        elif action == "run_macro":
            macro_name = variables[0]
            status = self.run_macro(macro_name)
        elif action == "add_macro":
            macro_name = variables[0]
            status = self.add_macro(macro_name)
        elif action == "convert":
            extension, to_extension = variables
            status = self.convert(extension, to_extension)
        self.print_status(status)
    
# %%

