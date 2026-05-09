# -*- coding: utf-8 -*-
__title__ = "View Log"
__doc__ = "Abre o arquivo de texto com o log das acoes monitoradas."

import os
from pyrevit import forms
import RevitActionLogger

LOG_FILE = RevitActionLogger.get_log_file()

if os.path.exists(LOG_FILE):
    os.startfile(LOG_FILE)
    forms.alert(RevitActionLogger.get_status_text(), title="Action Logger")
else:
    forms.alert(
        "O arquivo de log ainda nao foi criado. Inicie o Logger e realize alguma acao no Revit.\n\n"
        + RevitActionLogger.get_status_text(),
        title="Log nao encontrado"
    )
