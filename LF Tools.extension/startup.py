# -*- coding: utf-8 -*-
"""Inicializa recursos persistentes da extensao LF Tools."""

import os
import sys


ROOT_PATH = os.path.dirname(__file__)
LF_LIB_PATH = os.path.join(ROOT_PATH, "lib")
if LF_LIB_PATH not in sys.path:
    sys.path.append(LF_LIB_PATH)


# Smart AutoSave
autosave_lib_path = os.path.join(
    ROOT_PATH,
    "LF Tools.tab", "AutoSave.panel",
    "Smart AutoSave.pushbutton", "lib"
)
if autosave_lib_path not in sys.path:
    sys.path.append(autosave_lib_path)

try:
    from SmartAutoSave.config_manager import config

    if config.get("enabled", True):
        from pyrevit import HOST_APP
        from SmartAutoSave.autosave_manager import AutoSaveManager

        AutoSaveManager._instance = None
        AutoSaveManager(HOST_APP.uiapp)

except Exception:
    pass


# Ocultar abas do Revit
try:
    from pyrevit import HOST_APP
    from lf_ribbon_tabs import install_persistent_hider

    install_persistent_hider(HOST_APP.uiapp, exclude_ext="LF Tools")
except Exception:
    pass
