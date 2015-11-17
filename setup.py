__author__ = 'Duncan Lowder'

# -----------------------------------------------------------------------------
# setup.py
#
#   Description:
#       This file contains the basic setup information neccesary for the
#       py2exe script to run and successful build an executable file for
#       the get_beacon_status.py script.
#
#------------------------------------------------------------------------------

import py2exe
from distutils.core import setup
import os

icon_files = []
for files in os.listdir(r'C:\Users\Duncan Lowder\PycharmProjects\Beacon Status\icons'):
    f1 = r'C:\Users\Duncan Lowder\PycharmProjects\Beacon Status\icons\{0}'.format(files)
    if os.path.isfile(f1):
        f2 = 'icons', [f1]
        icon_files.append(f2)

setup(data_files=icon_files,
      options={"py2exe": {"includes": "decimal", "dll_excludes": ["MSVCP90.dll", "HID.DLL", "w9xpopen.exe"]}},
      windows=[{'script': 'get_beacon_status.py',
                "icon_resources": [(1, r"icons\app_icon_radar.ico")]}])