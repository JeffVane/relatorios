from cx_Freeze import setup, Executable
import sys

# Caminho do ícone da aplicação (opcional)
icon_path = "crc.ico"

# Configuração do executável
executables = [
    Executable(
        script="main.py",
        base="Win32GUI" if sys.platform == "win32" else None,
        target_name="Fiscalização(Relatórios).exe",  # Nome final do executável
        icon=icon_path  # Caminho do ícone (opcional)
    )
]

# Lista de pacotes extras
build_exe_options = {
    "packages": ["pandas", "tkinter", "sqlite3", "reportlab", "os", "sys"],
    "include_files": ["crc.ico"]  # Inclua o ícone e outros arquivos necessários
}

setup(
    name="Fiscalização",
    version="1.0",
    description="Aplicativo de Fiscalização",
    options={"build_exe": build_exe_options},
    executables=executables
)
