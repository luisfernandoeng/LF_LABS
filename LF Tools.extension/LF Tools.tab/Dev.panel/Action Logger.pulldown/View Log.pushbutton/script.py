# -*- coding: utf-8 -*-
__title__ = "View Log"
__doc__ = "Abre o arquivo de texto com o log das ações monitoradas."

import os
import System
from pyrevit import forms

desktop = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
LOG_FILE = os.path.join(desktop, "RevitActionLog.txt")

if os.path.exists(LOG_FILE):
    os.startfile(LOG_FILE)
else:
    forms.alert("O arquivo de log ainda não foi criado. Inicie o Logger e realize alguma ação no Revit.", title="Log não encontrado")
