# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import XYZ, Plane, ElementTransformUtils, Transaction
import os
import datetime

# Obtém o documento e vista ativa
doc = revit.doc
active_view = doc.ActiveView

# Caminho para log de erros
log_file = r"C:\Temp\flip_errors.txt"
if not os.path.exists(r"C:\Temp"):
    os.makedirs(r"C:\Temp")

# Pede ao usuário para escolher o escopo: todas ou selecionadas
scope = forms.SelectFromList.show(
    ['Todas as anotações na vista ativa', 'Apenas as selecionadas'],
    message='Escolha o escopo para espelhar e inverter:'
)
if not scope:
    forms.alert("Operação cancelada.")
    import sys
    sys.exit()

# Coleta elementos com base no escopo
if scope == 'Todas as anotações na vista ativa':
    collector = DB.FilteredElementCollector(doc, active_view.Id)
    elements = collector.OfCategory(DB.BuiltInCategory.OST_GenericAnnotation)\
                       .WhereElementIsNotElementType().ToElements()
else:
    selected = revit.get_selection()
    if not selected:
        elements = revit.pick_elements_by_category(
            DB.BuiltInCategory.OST_GenericAnnotation,
            "Selecione as anotações genéricas para espelhar e inverter."
        )
    else:
        elements = [el for el in selected 
                   if el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_GenericAnnotation)]

if not elements:
    forms.alert("Nenhuma anotação genérica encontrada ou selecionada.")
    import sys
    sys.exit()

# Pede ao usuário para escolher a direção de espelhamento
direction = forms.SelectFromList.show(
    ['Horizontal', 'Vertical'],
    message='Escolha a direção de espelhamento:'
)
if not direction:
    forms.alert("Operação cancelada.")
    import sys
    sys.exit()

# Inicia transação para espelhar e deletar originais
with Transaction(doc, "Espelhar e Inverter Anotações Genéricas") as t:
    t.Start()
    
    error_log = []
    for elem in elements:
        if isinstance(elem, DB.FamilyInstance):
            try:
                # Obtém o ponto central da anotação
                location = elem.Location
                if isinstance(location, DB.LocationPoint):
                    center = location.Point
                    
                    # Define o normal do plano com base na direção
                    if direction == 'Horizontal':
                        normal = XYZ(0, 1, 0)  # Horizontal: flip left-right (normal ao Y)
                    else:
                        normal = XYZ(1, 0, 0)  # Vertical: flip up-down (normal ao X)
                    
                    mirror_plane = Plane.CreateByNormalAndOrigin(normal, center)
                    
                    # Captura IDs antes do espelhamento
                    original_id = elem.Id
                    # Espelha (cria cópia)
                    ElementTransformUtils.MirrorElement(doc, original_id, mirror_plane)
                    
                    # Encontra a cópia recém-criada (baseado em proximidade ao original)
                    new_collector = DB.FilteredElementCollector(doc, active_view.Id)\
                                     .OfCategory(DB.BuiltInCategory.OST_GenericAnnotation)\
                                     .WhereElementIsNotElementType()
                    new_elem = None
                    for candidate in new_collector:
                        if candidate.Id != original_id and isinstance(candidate.Location, DB.LocationPoint):
                            if candidate.Location.Point.DistanceTo(center) < 0.01:  # Tolerância para mesmo local
                                new_elem = candidate
                                break
                    
                    if new_elem:
                        # Inverte orientação da cópia, se possível
                        if new_elem.CanFlipFacing:
                            new_elem.FlipFacing()
                        # Deleta o original
                        doc.Delete(original_id)
                    else:
                        error_msg = "Falha ao encontrar cópia espelhada para elemento ID {} (Família: {})".format(
                            elem.Id, elem.Symbol.FamilyName)
                        error_log.append(error_msg)
                else:
                    error_msg = "Elemento ID {} (Família: {}) não tem LocationPoint".format(
                        elem.Id, elem.Symbol.FamilyName)
                    error_log.append(error_msg)
            except Exception as e:
                error_msg = "Erro ao processar elemento ID {} (Família: {}): {}".format(
                    elem.Id, elem.Symbol.FamilyName, str(e))
                error_log.append(error_msg)
    
    t.Commit()

# Grava erros no log, se houver
if error_log:
    with open(log_file, 'a') as f:
        f.write("\n--- Log de erros: {} ---\n".format(datetime.datetime.now()))
        for msg in error_log:
            f.write(msg + "\n")
    forms.alert("Alguns elementos não foram processados. Verifique o log em: {}".format(log_file))
else:
    forms.alert("Anotações genéricas espelhadas e invertidas com sucesso (sem cópias)!")