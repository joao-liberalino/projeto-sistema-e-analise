from cx_Freeze import setup, Executable

setup(
    name="ProjetoDisbra",
    version="1.0",
    description="Sistema de cadastro Disbra",
    executables=[Executable("projetodisbra.py")]
)
