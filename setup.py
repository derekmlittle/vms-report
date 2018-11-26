import sys
from cx_Freeze import setup, Executable
import os.path

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')



options = {
    'build_exe': {
				'packages':["tkinter", "pandas", "numpy", "os", "xlsxwriter"],
        'include_files':[
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
         ],
    },
}



executables = [
    Executable('vms.py', base=base)
]
setup(options = options,
			console = ['vms.py'],
			name = "vms.py" ,
      version = "0.1" ,
      description = "" ,
      executables = executables)
