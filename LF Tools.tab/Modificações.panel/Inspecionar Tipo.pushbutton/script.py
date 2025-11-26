# -*- coding: utf-8 -*-
from pyrevit import revit, script, forms
from Autodesk.Revit.DB import *
from System.Collections.Generic import List
import os
import datetime

doc = revit.doc
uidoc = revit.uidoc

log_path = os.path.join(os.environ["USERPROFILE"], "Desktop", "relatorio_familias.txt")

def get_safe_param_val(param):
    try:
        if param.StorageType == StorageType.String:
            return param.AsString()
        elif param.StorageType == StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            if eid.IntegerValue == -1:
                return "None"
            try:
                el = doc.GetElement(eid)
                return el.Name
            except:
                return "ElementId: {}".format(eid.IntegerValue)
        else:
            return "<desconhecido>"
    except:
        return "<erro>"

def escrever_log(texto):
    with open(log_path, "a") as f:
        f.write(texto + "\n")

# Início do script
with open(log_path, "w") as f:
    f.write("==== RELATÓRIO DE FAMÍLIAS E TIPOS ====\n\n")

sel_ids = uidoc.Selection.GetElementIds()
if not sel_ids:
    forms.alert("Selecione um ou mais elementos antes de rodar este script.")
    script.exit()

for elid in sel_ids:
    el = doc.GetElement(elid)

    escrever_log("ID: {}".format(el.Id))
    escrever_log("Classe: {}".format(el.__class__.__name__))
    escrever_log("Categoria: {}".format(el.Category.Name if el.Category else "Sem categoria"))

    try:
        tipo = el.Symbol if hasattr(el, "Symbol") else el
        tipo_nome = tipo.LookupParameter("Type Name").AsString()
        escrever_log("Tipo: {}".format(tipo_nome))
    except:
        escrever_log("Tipo: <erro ao obter tipo>")

    try:
        familia_nome = doc.GetElement(tipo.Family.Id).Name
        escrever_log("Família: {}".format(familia_nome))
    except:
        escrever_log("Família: <erro ou não aplicável>")

    escrever_log("--- Parâmetros ---")
    for p in el.Parameters:
        try:
            nome = p.Definition.Name
            valor = get_safe_param_val(p)
            escrever_log("  {0}: {1}".format(nome, valor))
        except:
            escrever_log("  <erro ao ler parâmetro>")

    escrever_log("-" * 50 + "\n")

forms.alert("✅ Relatório salvo em:\n{}".format(log_path))
