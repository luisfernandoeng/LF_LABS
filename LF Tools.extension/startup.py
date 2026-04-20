# -*- coding: utf-8 -*-
"""Inicializa o AutoSave e aplica as abas ocultas assim que o Revit abre."""

import os
import sys

# Injeta o caminho da lib do SmartAutoSave no sys.path
lib_path = os.path.join(
    os.path.dirname(__file__),
    "LF Tools.tab", "AutoSave.panel",
    "Smart AutoSave.pushbutton", "lib"
)
if lib_path not in sys.path:
    sys.path.append(lib_path)

try:
    from SmartAutoSave.config_manager import config

    # Importação pesada (WPF / DispatcherTimer) só acontece se o AutoSave
    # estiver habilitado — evita carregar PresentationFramework no boot à toa.
    if config.get("enabled", True):
        from pyrevit import HOST_APP
        from SmartAutoSave.autosave_manager import AutoSaveManager

        # Garante instância limpa a cada reload do pyRevit
        AutoSaveManager._instance = None

        # __init__ já chama _stop_appdomain_timer() e start() internamente.
        # Não é necessário chamar start() novamente aqui.
        AutoSaveManager(HOST_APP.uiapp)

except Exception:
    # Nunca deixa o startup quebrar o carregamento da extensão
    pass

# ── Ocultar abas do Revit ──────────────────────────────────────────────────
try:
    from pyrevit import script as _script
    from pyrevit.runtime import types as _types

    _cfg = _script.get_config('lf_hidden_tabs')
    _hidden = _cfg.get_option('hidden_tabs', [])
    if _hidden:
        _types.RibbonTabVisibilityUtils.StartHidingTabs(_hidden)
except Exception:
    pass
