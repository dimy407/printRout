from cx_Freeze import setup, Executable

setup(
    name = "PrintRoute",
    version = "0.1",
    description = "PrintRoute",
    executables = [Executable("main.py")]
)