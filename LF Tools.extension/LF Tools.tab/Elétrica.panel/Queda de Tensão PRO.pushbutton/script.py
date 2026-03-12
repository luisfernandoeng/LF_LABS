# -*- coding: utf-8 -*-
__title__   = "Queda de\nTensão PRO"
__author__  = "LF Tools"

import os
import sys

# Insere pasta lib no root do Python Path para habilitar a sub-pasta QuedaTensao
lib_path = os.path.join(os.path.dirname(__file__), "lib")
if lib_path not in sys.path:
    sys.path.append(lib_path)

from pyrevit import revit, DB, forms
from QuedaTensao.queda_tensao_ui import QuedaTensaoWindow

doc = revit.doc

def is_conduit(element):
    return element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Conduit)

def get_conduit_data(element):
    """
    Extrai as medidas essenciais do eletroduto convertido para UI e Sistema Métrico
    """
    # 1. Obter Length
    try:
        length_ft = element.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    except Exception:
        length_ft = 0.0
    length_m = length_ft * 0.3048

    # 2. Obter Diameter
    try:
        diam_ft = element.get_Parameter(DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsDouble()
    except Exception:
        diam_ft = 0.0
    # Revit internal diameter storage is feet. 1ft = 304.8mm
    diam_mm = diam_ft * 304.8

    return {"length_m": round(length_m, 2), "diam_mm": round(diam_mm, 2)}

if __name__ == '__main__':
    selection = revit.get_selection()

    if not selection:
        forms.alert('Sem Seleção: Selecione exatamente 1 Eletroduto antes de usar a ferramenta.', warn_icon=True)
        sys.exit(0)
    
    selected_ids = selection.element_ids
    if len(selected_ids) != 1:
        forms.alert('Seleção Múltipla: Selecione exatamente apenas 1 Eletroduto para testar a Queda de Tensão.', warn_icon=True)
        sys.exit(0)
    
    el = doc.GetElement(selected_ids[0])
    if not is_conduit(el):
        forms.alert('Tipo Incorreto: O elemento selecionado não é um Eletroduto.', warn_icon=True)
        sys.exit(0)

    # Extrai o valor pre-computado
    data = get_conduit_data(el)

    # Invoca o WPF Form Window injetando o Active Element Id para eventual gravação de parâmetros
    cur_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(cur_dir, 'lib', 'QuedaTensao', 'queda_tensao_window.xaml')
    
    win = QuedaTensaoWindow(xaml_path, doc, el.Id, data)
    win.ShowDialog()
