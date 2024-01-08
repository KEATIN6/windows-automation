# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 11:11:20 2024

@author: ghostshade
"""

# %%

my_macro = """
    Sub MyMacro()
        MsgBox "Hello from VBA!"
    End Sub
"""

# %%

script_mapping = {
    "MyMacro": {
        "name": "MyMacro",
        "code": my_macro
    }
}

# %%

