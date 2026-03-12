# -*- coding: utf-8 -*-
from pyrevit import script, forms
import sys
import os

extensions_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extensions_dir, "lib")
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from SmartAutoSave.autosave_manager import AutoSaveManager
from SmartAutoSave.config_manager import config

if __name__ == '__main__':
    enabled = config.get("enabled")
    new_state = not enabled
    
    config.set("enabled", new_state)
    
    manager = AutoSaveManager(__revit__)
    if new_state:
        manager.start()
        forms.alert("Smart AutoSave ATIVADO.", warn_icon=False)
    else:
        manager.stop()
        forms.alert("Smart AutoSave DESATIVADO.", warn_icon=False)
