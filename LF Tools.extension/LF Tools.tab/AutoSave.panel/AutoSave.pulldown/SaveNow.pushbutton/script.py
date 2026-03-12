# -*- coding: utf-8 -*-
from pyrevit import script, forms
import sys
import os

extensions_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extensions_dir, "lib")
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from SmartAutoSave.autosave_manager import AutoSaveManager

if __name__ == '__main__':
    manager = AutoSaveManager(__revit__)
    manager.trigger_save_now()
