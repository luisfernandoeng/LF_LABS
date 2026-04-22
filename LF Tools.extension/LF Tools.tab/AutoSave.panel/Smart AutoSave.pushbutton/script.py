# -*- coding: utf-8 -*-
__title__   = "Smart\nAutoSave"
__author__  = "LF Tools"
__persistentengine__ = True

import os
import sys

# Garante a importação do diretório nativo lib
lib_path = os.path.join(os.path.dirname(__file__), "lib")
if lib_path not in sys.path:
    sys.path.append(lib_path)

from SmartAutoSave.config_window import show_config
from SmartAutoSave.config_manager import config
from SmartAutoSave.autosave_manager import AutoSaveManager

if __name__ == '__main__':
    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        # Modo Configuração
        show_config()
    else:
        # Modo Ação (Salvar Agora)
        if config.get("enabled", True):
            from pyrevit import HOST_APP, forms
            manager = AutoSaveManager(HOST_APP.uiapp)
            if manager.is_paused:
                forms.alert(u"Timer Pausado. Clique com o botão direito para gerenciar, mas vamos forçar o save:", warn_icon=False)
            with forms.ProgressBar(title=u"AutoSave: Salvando...", cancellable=False) as pb:
                pb.update_progress(0, 1)
                manager.trigger_save_now()
                pb.update_progress(1, 1)
        else:
            from pyrevit import forms
            forms.alert(u"O AutoSave está desativado.\nSegure Shift e clique no botão para abrir as configurações e ativá-lo.", warn_icon=False)
