import sysconfig
import os
import sys
from cx_Freeze import setup, Executable


#This file is for setting up the application for CX_Freeze
#CX Freeze compiles all dependencies for the App so that it can then be turned into an .exe with Inno

build_exe_options = {"packages": ["os"]}

#Path to TCL Library
os.environ['TCL_LIBRARY'] = r'C:\Users\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
#Path to TK Library
os.environ['TK_LIBRARY'] = r'C:\Users\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'

buildOptions = dict(
    packages = [],
    excludes = [],
    #Include path to libraries and icon file here or they won't load properly when building executable
    include_files=[r'C:\Users\AppData\Local\Programs\Python\Python36-32\DLLs\tcl86t.dll',
                   r'C:\Users\AppData\Local\Programs\Python\Python36-32\DLLs\tk86t.dll',
                   r'C:\Users\My Documents\LiClipse Workspace\Taps_Controller\taps.ico'])

base = 'Win32GUI' if sys.platform == 'win32' else None

# Icon is Path to app icon file
executables = [ Executable('Taps_Controller.py', base=base, icon=r'C:\Users\My Documents\LiClipse Workspace\Taps_Controller\taps.ico')]

setup(name='Tap Controller',
      version = '1.0.0',
      description = 'Tap Shoes Lighting Control Software',
              options = dict(build_exe = buildOptions),
              executables = executables)
