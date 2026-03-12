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
    manager = AutoSaveManager(__revit__)
    
    if not config.get("enabled"):
        forms.alert("O Smart AutoSave está desligado. Você precisa ativá-arlo primeiro.")
    elif manager.is_paused:
        manager.resume()
        forms.alert("Timer de AutoSave RETOMADO.", warn_icon=False)
    else:
        manager.pause()
        forms.alert("Timer de AutoSave PAUSADO.", warn_icon=False)
