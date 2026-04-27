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

# ── Ocultar abas do Revit (Persistente) ─────────────────────────────────────
try:
    from pyrevit import script as _script, HOST_APP
    
    # Nome da função do handler para persistência entre reloads
    HANDLER_NAME = 'lf_tabs_visibility_handler'
    
    def _apply_hidden_tabs():
        """Lê config e aplica visibilidade na Ribbon."""
        cfg = _script.get_config('lf_hidden_tabs')
        is_active = cfg.get_option('is_active', False)
        hidden_list = set(cfg.get_option('hidden_tabs', []))
        
        if not is_active or not hidden_list:
            return

        import clr
        clr.AddReference('AdWindows')
        import Autodesk.Windows as adWin
        
        for tab in adWin.ComponentManager.Ribbon.Tabs:
            title = tab.Title or u""
            tab_id = tab.Id or u""
            
            # Pula abas de sistema e da própria extensão
            if not title or "LF Tools" in title or "Modify" in tab_id:
                continue
            try:
                if tab.IsContextualTab:
                    continue
            except Exception:
                continue
            
            # Se a aba está no perfil de ocultação, força False
            if title in hidden_list:
                if tab.IsVisible:
                    tab.IsVisible = False

    def _on_view_activated(sender, args):
        """Handler disparado sempre que uma vista é ativada (ex: troca de projeto)."""
        try:
            _apply_hidden_tabs()
        except Exception:
            pass

    # Registro do Evento de forma segura (evita duplicatas em reloads do pyRevit)
    # Procuramos no dicionário da UIApplication se já existe nosso handler
    if not hasattr(HOST_APP.uiapp, HANDLER_NAME):
        setattr(HOST_APP.uiapp, HANDLER_NAME, _on_view_activated)
        HOST_APP.uiapp.ViewActivated += _on_view_activated
    else:
        # Se já existe (reload), remove o antigo e põe o novo
        old_handler = getattr(HOST_APP.uiapp, HANDLER_NAME)
        try:
            HOST_APP.uiapp.ViewActivated -= old_handler
        except Exception:
            pass
        setattr(HOST_APP.uiapp, HANDLER_NAME, _on_view_activated)
        HOST_APP.uiapp.ViewActivated += _on_view_activated

    # Aplicação imediata no boot/reload
    _apply_hidden_tabs()

except Exception:
    pass
