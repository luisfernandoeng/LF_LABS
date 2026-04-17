# -*- coding: utf-8 -*-
__title__ = "Start Log"
__doc__ = "Inicia o Action Logger PRO em background para rastrear comandos e propriedades elétricas."

from pyrevit import HOST_APP, forms
import RevitActionLogger

uiapp = HOST_APP.uiapp

if RevitActionLogger.is_running():
    forms.alert("O Action Logger já está em execução.", title="Logger Ativo")
else:
    success = RevitActionLogger.start_logger(uiapp)
    if success:
        forms.alert("Revit Action Logger PRO iniciado com sucesso!", title="Logger Iniciado")
    else:
        forms.alert("Falha ao iniciar o Logger.", title="Erro")
