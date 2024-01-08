# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 11:09:34 2024

@author: ghostshade
"""

# %%

from datetime import datetime
from debuggers import ExcelEditorDebugger
from win32com.client import Dispatch
import os

# %%

DOWNLOAD_PATH = r"C:\Users\brenn\Documents\Projects\WindowsAutomation\Downloads"

class OutlookEditor:
    def __init__(self, visiblity=True, debug=False):
        self.visiblity = visiblity
        self.debug = debug
        self.start_application()
        
    def start_application(self) -> bool:
        self.outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        return True
    
    def download_attachment(self, email, extension=".xlsx"):
        for attachment in email.Attachments:
            if attachment.FileName.endswith(extension):
                file_path = DOWNLOAD_PATH + "/" + attachment.FileName
                attachment.SaveAsFile(file_path)
    
    def download_excel_from_email(self, sender=None, subject=None):
        inbox = self.outlook.GetDefaultFolder(6)
        today = datetime.now().date()
        query = "[ReceivedTime] >= '" + today.strftime('%m/%d/%Y') + "'"
        emails = inbox.Items.Restrict(query)
        if sender:
            emails = [
                email for email in emails 
                if email.SenderEmailAddress == sender
            ]
        if subject: 
            emails = [
                email for email in emails 
                if subject in email.Subject
            ]
        for email in emails:
            print("Subject:", email.Subject)
            print("Sender:", email.SenderEmailAddress)
            print("Received Time:", email.ReceivedTime)
            print()
            self.download_attachment(email)
        

        
outlook = OutlookEditor()
outlook.download_excel_from_email("evan.c.keating@gmail.com", "Billing Tracker")

# %%
  
class ExcelEditor:
    def __init__(self, file_path, visiblity=True, debug=False):
        self.file_path = file_path
        self.visiblity = visiblity
        self.debug = debug
        self.debugger = ExcelEditorDebugger(self) if debug else None
        self.start_application()
        self.load_workbook()
        
    def start_application(self) -> bool:
        self.excel = Dispatch("Excel.Application")
        self.excel.Visible = self.visiblity
        return True
        
    def handle_debug(self, status, variables=None):
        if self.debug:
            self.debugger.process(status, variables)
        
    def load_workbook(self) -> bool:
        self.handle_debug("load")
        self.workbook = self.excel.Workbooks.Open(self.file_path)
        return True
        
    def check_macro_exists(self, macro_name, message="If or Where") -> tuple:
        self.handle_debug("macro_exists", (message, macro_name))
        vb_project = self.workbook.VBProject
        for components in vb_project.VBComponents:
            component = components.Name
            if "Module" in component:
                module = vb_project.VBComponents(component).CodeModule
                if module.Find(macro_name)[0]:
                    return (True, component)
        return (False, None)
    
    def check_extension(self):
        self.handle_debug("extension")
        _, extension = os.path.splitext(self.file_path)
        if extension in ["xlsm", "xltm"]:
            return True
        return False
        
    def add_vba_to_excel(self, macro_name, vba_code) -> bool:
        self.handle_debug("add_macro", (macro_name,))
        if not self.check_extension():
            self.convert_workbook()
        vba_project = self.workbook.VBProject 
        exists, module = self.check_macro_exists(macro_name, "If")
        if not exists:
            new_module = vba_project.VBComponents.Add(1)
            new_module.CodeModule.AddFromString(vba_code)
            self.workbook.Save()
            return True
        return False
    
    def convert_workbook(self, to_extension="xlsm") -> bool:
        base_path, extension = os.path.splitext(self.file_path)
        if extension == to_extension:
            return False
        self.handle_debug("convert", (extension, to_extension))
        new_path = base_path + "." + to_extension
        self.workbook.SaveAs(new_path, FileFormat=52)
        self.close_workbook()
        self.file_path = new_path
        self.load_workbook()
        return True
        
    def run_macro(self, macro_name):
        exists, module = self.check_macro_exists(macro_name, "Where")
        self.handle_debug("run_macro", (macro_name,))
        if not exists:
            return False
        self.excel.Run(f"{module}.{macro_name}")
        return True
        
    def close_workbook(self):
        self.handle_debug("close")
        self.workbook.Close()    
    
# %%