from cx_Freeze import setup, Executable
import os

base = None

executables = [Executable("XLPLAN.py", base=base)]

includes = []
include_files = [r"C:\Users\Frede\AppData\Local\Programs\Python\Python36-32\DLLs\tcl86t.dll",
                 r"C:\Users\Frede\AppData\Local\Programs\Python\Python36-32\DLLs\tk86t.dll"]
os.environ['TCL_LIBRARY'] = r'C:\Users\Frede\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\Frede\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'

packages = ["idna", "openpyxl", "tkinter"]
options = {
    'build_exe': {
        'packages':packages,
        "includes": includes,
        "include_files": include_files
    },
}

setup(
    name = "XLPLAN",
    options = options,
    version = "1.0",
    description = 'Af Frederik Schrøder Koefoed',
    executables = executables
)
