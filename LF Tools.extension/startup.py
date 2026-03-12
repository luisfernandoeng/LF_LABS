# -*- coding: utf-8 -*-
"""Inicializa o timer persistente do AutoSave assim que o Revit abre."""

import os
import sys

# Injeta explicitamente o caminho da biblioteca lib na lista do python sys.path
# Assim é possivel puxar módulos pelo nome SmartAutoSave
lib_path = os.path.join(os.path.dirname(__file__), "LF Tools.tab", "AutoSave.panel", "Smart AutoSave.pushbutton", "lib")
if lib_path not in sys.path:
    sys.path.append(lib_path)

from pyrevit import HOST_APP
from SmartAutoSave.config_manager import config
from SmartAutoSave.autosave_manager import AutoSaveManager

# Incializa a infraestrutura de tempo persistente se habilitado
manager = AutoSaveManager(HOST_APP.uiapp)

if config.get("enabled", True):
    manager.start()
