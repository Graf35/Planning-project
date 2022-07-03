import cx_Freeze
import sys


base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("main.py", base=base, icon = "icon.ico")]
build_exe_options = {"includes":["Main_window"]
                     }

cx_Freeze.setup(
    name = "Planning-project",
    options = {"build_exe": build_exe_options},
    version = "0.1",
    description = "A demo version of the repair planning system",
    executables = executables
    )
