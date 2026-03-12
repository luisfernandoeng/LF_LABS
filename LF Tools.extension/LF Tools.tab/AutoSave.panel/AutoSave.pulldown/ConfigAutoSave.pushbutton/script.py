# -*- coding: utf-8 -*-
import sys
import os

# Adds the lib folder path inside LF Tools.extension to the system path
extensions_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extensions_dir, "lib")
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from pyrevit import script
from SmartAutoSave.config_window import show_config

if __name__ == '__main__':
    show_config()
