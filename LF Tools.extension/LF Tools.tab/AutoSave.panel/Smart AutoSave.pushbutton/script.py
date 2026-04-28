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
        from pyrevit import HOST_APP, forms
        if not config.get("enabled", True):
            forms.alert(u"O AutoSave está desativado.\nSegure Shift e clique no botão para abrir as configurações e ativá-lo.", warn_icon=False)
        else:
            manager = AutoSaveManager(HOST_APP.uiapp)
            manager._cancel_countdown()

            uidoc = HOST_APP.uiapp.ActiveUIDocument
            if uidoc is None:
                forms.alert(u"Nenhum documento ativo.")
            else:
                doc = uidoc.Document
                if not doc or doc.IsFamilyDocument:
                    forms.alert(u"Tipo de documento não suportado.")
                elif not doc.PathName:
                    forms.alert(u"Projeto ainda não salvo em disco. Use Salvar Como primeiro.")
                elif not doc.IsModified:
                    forms.toast(u"Sem alterações para salvar.", title=u"AutoSave")
                else:
                    with forms.ProgressBar(title=u"AutoSave: Salvando...", cancellable=False) as pb:
                        pb.update_progress(0, 1)
                        doc.Save()
                        pb.update_progress(1, 1)
                    manager.start()
                    forms.toast(u"Projeto salvo!", title=u"AutoSave")
