# -*- coding: utf-8 -*-
__title__ = "Start Log"
__doc__ = "Inicia o Action Logger em background para rastrear comandos e propriedades eletricas."

from pyrevit import HOST_APP, forms
import RevitActionLogger

uiapp = HOST_APP.uiapp

if RevitActionLogger.is_running():
    forms.alert(
        "O Action Logger ja esta em execucao.\n\n" + RevitActionLogger.get_status_text(),
        title="Logger Ativo"
    )
else:
    success = RevitActionLogger.start_logger(uiapp)
    if success:
        forms.alert(
            "Revit Action Logger iniciado.\n\n" + RevitActionLogger.get_status_text(),
            title="Logger Iniciado"
        )
    else:
        forms.alert(
            "Falha ao iniciar o Logger.\n\n" + RevitActionLogger.get_status_text(),
            title="Erro"
        )
