# -*- coding: utf-8 -*-
__title__ = "Stop Log"
__doc__ = "Para o Action Logger PRO em background e desvincula os eventos da API."

from pyrevit import HOST_APP, forms
import RevitActionLogger

uiapp = HOST_APP.uiapp

if RevitActionLogger.is_running():
    success = RevitActionLogger.stop_logger(uiapp)
    if success:
        forms.alert("Revit Action Logger PRO parado e eventos desvinculados com segurança.", title="Logger Parado")
    else:
        forms.alert("Não foi possível desvincular o Logger. Ele pode já estar parado.", title="Erro")
else:
    forms.alert("O Logger não está em execução.", title="Aviso")
