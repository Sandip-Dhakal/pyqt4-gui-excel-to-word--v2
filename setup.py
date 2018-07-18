from distutils.core import setup
import sys
import py2exe

sys.setrecursionlimit(3000)
setup(
    options = {
            "py2exe":{
             'bundle_files': 1, 
             "includes": ["pandas", "docx"],  
            "includes":["sip"],
            "dll_excludes": ["MSVCP90.dll", "HID.DLL", "w9xpopen.exe"],
            
        }
    },
    # console = [{'script': 'demo.py'}],
    windows=['demo.py'],
)