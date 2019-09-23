from cx_Freeze import setup, Executable
import sys

base = None

if sys.platform == 'win32':
    base = None


executables = [Executable("WATSON v0.4-beta.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "WATSON",
    options = options,
    version = "0.4",
    description = 'Hyperpersonalised Automation Tool',
    executables = executables
)
