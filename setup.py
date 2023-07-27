import sys

from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "includes": ["tkinter"], "includes": ["pandas"], "includes": ["logging"], "includes": ["fitz"]}

# GUI applications require a different base on Windows (the default is for
# a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Aplicativo de Renomear Arquivo",
    version="0.1",
    description="Aplicação para Renomear arquivos - By: Felipe Barcelos",
    options={"build_exe": build_exe_options},
    executables=[Executable("rename.py", base=base)]
)