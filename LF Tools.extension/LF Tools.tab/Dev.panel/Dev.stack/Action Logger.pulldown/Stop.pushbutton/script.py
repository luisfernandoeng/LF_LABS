# -*- coding: utf-8 -*-
__title__ = "Stop Log"
__doc__ = "Para o Action Logger em background e desvincula os eventos da API."

from pyrevit import HOST_APP, forms
import RevitActionLogger

uiapp = HOST_APP.uiapp

if RevitActionLogger.is_running():
    success = RevitActionLogger.stop_logger(uiapp)
    if success:
        forms.alert(
            "Revit Action Logger parado e eventos desvinculados.\n\n" + RevitActionLogger.get_status_text(),
            title="Logger Parado"
        )
    else:
        forms.alert(
            "Nao foi possivel desvincular o Logger. Ele pode ja estar parado.\n\n" + RevitActionLogger.get_status_text(),
            title="Erro"
        )
else:
    forms.alert(
        "O Logger nao esta em execucao.\n\n" + RevitActionLogger.get_status_text(),
        title="Aviso"
    )
