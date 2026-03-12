# -*- coding: utf-8 -*-
"""
Inverter Anotação - Espelha anotações genéricas e notas de texto
preservando a posição do leader (chamada de detalhe/seta).
"""

__title__ = 'Inverter\nAnotação'
__author__ = 'Luis Fernando'

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import (
    XYZ, Plane, ElementTransformUtils, Transaction,
    TextNote, BuiltInCategory, FilteredElementCollector
)
import os
import datetime
import math

doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

# Caminho para log de erros
log_file = r"C:\Temp\flip_errors.txt"
if not os.path.exists(r"C:\Temp"):
    os.makedirs(r"C:\Temp")


# =====================================================================
#  FUNCOES AUXILIARES - LEADERS
# =====================================================================
def get_leader_info(elem):
    """Salva posicoes de todos os leaders (chamadas de detalhe) antes do mirror."""
    leader_data = []
    try:
        if hasattr(elem, 'GetLeaders'):
            leaders = elem.GetLeaders()
            if leaders:
                for leader in leaders:
                    data = {}
                    try:
                        data['end'] = XYZ(leader.End.X, leader.End.Y, leader.End.Z)
                    except:
                        pass
                    try:
                        data['elbow'] = XYZ(leader.Elbow.X, leader.Elbow.Y, leader.Elbow.Z)
                    except:
                        pass
                    if data:
                        leader_data.append(data)
    except:
        pass
    
    if not leader_data:
        try:
            if hasattr(elem, 'HasLeader') and elem.HasLeader:
                data = {}
                if hasattr(elem, 'LeaderEnd'):
                    end = elem.LeaderEnd
                    data['end'] = XYZ(end.X, end.Y, end.Z)
                if data:
                    leader_data.append(data)
        except:
            pass
    
    return leader_data


def restore_leader_info(elem, leader_data):
    """Restaura posicoes dos leaders salvos anteriormente."""
    if not leader_data:
        return
    try:
        if hasattr(elem, 'GetLeaders'):
            leaders = list(elem.GetLeaders())
            for i, leader in enumerate(leaders):
                if i < len(leader_data):
                    saved = leader_data[i]
                    try:
                        if 'end' in saved:
                            leader.End = saved['end']
                    except:
                        pass
                    try:
                        if 'elbow' in saved:
                            leader.Elbow = saved['elbow']
                    except:
                        pass
            return
    except:
        pass
    
    try:
        if hasattr(elem, 'HasLeader') and elem.HasLeader and leader_data:
            saved = leader_data[0]
            if 'end' in saved and hasattr(elem, 'LeaderEnd'):
                elem.LeaderEnd = saved['end']
    except:
        pass


# =====================================================================
#  FUNCOES DE MIRROR
# =====================================================================
def mirror_annotation(doc, elem, direction, active_view, error_log):
    """Espelha uma anotação genérica (FamilyInstance) preservando leaders."""
    location = elem.Location
    if not isinstance(location, DB.LocationPoint):
        error_log.append("ID {} sem LocationPoint".format(elem.Id))
        return False
    
    center = location.Point
    saved_leaders = get_leader_info(elem)
    
    if direction == 'Horizontal':
        normal = active_view.UpDirection
    else:
        normal = active_view.RightDirection
    
    mirror_plane = Plane.CreateByNormalAndOrigin(normal, center)
    original_id = elem.Id
    
    ElementTransformUtils.MirrorElement(doc, original_id, mirror_plane)
    
    # Encontrar cópia espelhada
    new_collector = FilteredElementCollector(doc, active_view.Id)\
                     .OfCategory(BuiltInCategory.OST_GenericAnnotation)\
                     .WhereElementIsNotElementType()
    
    new_elem = None
    for candidate in new_collector:
        if candidate.Id != original_id and isinstance(candidate.Location, DB.LocationPoint):
            if candidate.Location.Point.DistanceTo(center) < 0.01:
                new_elem = candidate
                break
    
    if new_elem:
        if saved_leaders:
            restore_leader_info(new_elem, saved_leaders)
        if new_elem.CanFlipFacing:
            new_elem.FlipFacing()
        doc.Delete(original_id)
        return new_elem
    else:
        error_log.append("ID {}: cópia espelhada não encontrada".format(elem.Id))
        return None


def mirror_text_note(doc, elem, direction, active_view, error_log):
    """Espelha uma TextNote preservando leader."""
    try:
        coord = elem.Coord
        
        # Salvar info do leader de TextNote
        has_leader = False
        leader_end = None
        leader_elbow = None
        try:
            if hasattr(elem, 'HasLeader') and elem.HasLeader:
                has_leader = True
            # TextNote leaders - via GetLeaders() em Revit 2020+
            if hasattr(elem, 'GetLeaders'):
                leaders = elem.GetLeaders()
                if leaders:
                    has_leader = True
                    for ldr in leaders:
                        try:
                            leader_end = XYZ(ldr.End.X, ldr.End.Y, ldr.End.Z)
                        except:
                            pass
                        try:
                            leader_elbow = XYZ(ldr.Elbow.X, ldr.Elbow.Y, ldr.Elbow.Z)
                        except:
                            pass
                        break  # Pega só o primeiro
        except:
            pass
        
        if direction == 'Horizontal':
            normal = active_view.UpDirection
        else:
            normal = active_view.RightDirection
        
        mirror_plane = Plane.CreateByNormalAndOrigin(normal, coord)
        original_id = elem.Id
        
        ElementTransformUtils.MirrorElement(doc, original_id, mirror_plane)
        
        # Encontrar cópia espelhada (TextNote)
        new_collector = FilteredElementCollector(doc, active_view.Id)\
                         .OfClass(TextNote)
        
        new_elem = None
        for candidate in new_collector:
            if candidate.Id != original_id:
                try:
                    if candidate.Coord.DistanceTo(coord) < 0.01:
                        new_elem = candidate
                        break
                except:
                    pass
        
        if new_elem:
            # Restaurar leader
            if has_leader and leader_end:
                try:
                    if hasattr(new_elem, 'GetLeaders'):
                        new_leaders = list(new_elem.GetLeaders())
                        if new_leaders:
                            try:
                                new_leaders[0].End = leader_end
                            except:
                                pass
                            if leader_elbow:
                                try:
                                    new_leaders[0].Elbow = leader_elbow
                                except:
                                    pass
                except:
                    pass
            
            doc.Delete(original_id)
            return new_elem
        else:
            error_log.append("TextNote ID {}: cópia espelhada não encontrada".format(elem.Id))
            return None
    except Exception as e:
        error_log.append("TextNote ID {}: {}".format(elem.Id, str(e)))
        return None


# =====================================================================
#  COLETA DE ELEMENTOS - Smart Detection
# =====================================================================
# Categorias suportadas
SUPPORTED_CATS = [
    int(BuiltInCategory.OST_GenericAnnotation),
    int(BuiltInCategory.OST_TextNotes)
]

# 1. Verifica seleção prévia
pre_selected = revit.get_selection()
elements = []

if pre_selected:
    for el in pre_selected:
        try:
            cat_id = el.Category.Id.IntegerValue
            if cat_id in SUPPORTED_CATS:
                elements.append(el)
        except:
            pass

# 2. Se não tem seleção prévia, pergunta escopo
if not elements:
    scope = forms.CommandSwitchWindow.show(
        ['📋 Todas na vista', '🖱️ Selecionar agora'],
        message='Nenhum elemento pré-selecionado.\nEscolha o escopo:'
    )
    
    if not scope:
        script.exit()
    
    if scope == '📋 Todas na vista':
        # Coleta anotações genéricas
        ann_collector = FilteredElementCollector(doc, active_view.Id)\
                        .OfCategory(BuiltInCategory.OST_GenericAnnotation)\
                        .WhereElementIsNotElementType().ToElements()
        elements.extend(ann_collector)
        
        # Coleta TextNotes
        txt_collector = FilteredElementCollector(doc, active_view.Id)\
                        .OfClass(TextNote).ToElements()
        elements.extend(txt_collector)
    else:
        # Seleção manual pelo usuario
        from Autodesk.Revit.UI.Selection import ObjectType
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element,
                'Selecione anotações genéricas e/ou notas de texto.'
            )
            for ref in refs:
                el = doc.GetElement(ref)
                try:
                    cat_id = el.Category.Id.IntegerValue
                    if cat_id in SUPPORTED_CATS:
                        elements.append(el)
                except:
                    pass
        except:
            script.exit()

if not elements:
    forms.alert("Nenhuma anotação genérica ou nota de texto encontrada.")
    script.exit()

# Classificar elementos
annotations = [e for e in elements if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericAnnotation)]
text_notes = [e for e in elements if isinstance(e, TextNote)]

# Resumo do que será processado
summary_items = []
if annotations:
    summary_items.append("{} anotação(ões)".format(len(annotations)))
if text_notes:
    summary_items.append("{} nota(s) de texto".format(len(text_notes)))

# 3. Escolher direção - Dialog único e nativo
direction = forms.CommandSwitchWindow.show(
    ['↔ Horizontal', '↕ Vertical', '🔄 Inverter 180°'],
    message='{}\nEscolha a direção a espelhar:'.format(" + ".join(summary_items))
)

if not direction:
    script.exit()

# Normalizar ação
action = 'Vertical'
if 'Horizontal' in direction:
    action = 'Horizontal'
elif '180°' in direction:
    action = 'Rotate180'

# =====================================================================
#  EXECUTAR MIRROR E VERIFICACAO
# =====================================================================
def process_annotation(doc, elem, action, active_view, error_log):
    target_elem = elem
    if action in ['Horizontal', 'Vertical']:
        target_elem = mirror_annotation(doc, elem, action, active_view, error_log)
    elif action == 'Rotate180':
        # Generic Annotation tem problema de recalcular leader ao rotacionar
        # Solução: Duplo Mirror (Horizontal + Vertical = 180 graus perfeito)
        temp_elem = mirror_annotation(doc, elem, 'Horizontal', active_view, error_log)
        if temp_elem:
           target_elem = mirror_annotation(doc, temp_elem, 'Vertical', active_view, error_log)
        else:
           target_elem = None

    return target_elem is not None

def process_text_note(doc, elem, action, active_view, error_log):
    target_elem = elem
    if action in ['Horizontal', 'Vertical']:
        target_elem = mirror_text_note(doc, elem, action, active_view, error_log)
    elif action == 'Rotate180':
        # TextNote tem problema ao girar o eixo Z usando RotateElement. Os leaders saem do lugar.
        # A solução perfeita de 180° para Texts no Revit é Espelhar Vertical + Espelhar Horizontal.
        temp_elem = mirror_text_note(doc, elem, 'Horizontal', active_view, error_log)
        if temp_elem:
           target_elem = mirror_text_note(doc, temp_elem, 'Vertical', active_view, error_log)
        else:
           target_elem = None
        
    return target_elem is not None


with Transaction(doc, "Inverter Anotações") as t:
    t.Start()
    
    error_log = []
    success_ann = 0
    success_txt = 0
    
    # Processar anotações genéricas
    for elem in annotations:
        if isinstance(elem, DB.FamilyInstance):
            try:
                if process_annotation(doc, elem, action, active_view, error_log):
                    success_ann += 1
            except Exception as e:
                error_log.append("Anotação ID {}: {}".format(elem.Id, str(e)))
    
    # Processar TextNotes
    for elem in text_notes:
        try:
            if process_text_note(doc, elem, action, active_view, error_log):
                success_txt += 1
        except Exception as e:
            error_log.append("TextNote ID {}: {}".format(elem.Id, str(e)))
    
    t.Commit()


# =====================================================================
#  RESULTADO
# =====================================================================
if error_log:
    with open(log_file, 'a') as f:
        f.write("\n--- Log de erros: {} ---\n".format(datetime.datetime.now()))
        for msg in error_log:
            f.write(msg + "\n")

# Mensagem final
result_parts = []
if success_ann > 0:
    result_parts.append("✓ {} anotação(ões)".format(success_ann))
if success_txt > 0:
    result_parts.append("✓ {} nota(s) de texto".format(success_txt))

total_success = success_ann + success_txt
total_errors = len(error_log)

if total_success == 0 and total_errors > 0:
    forms.alert("Nenhum elemento processado com sucesso.\nVerifique o log: {}".format(log_file))
elif total_errors > 0:
    forms.alert("{}\n\n⚠ {} erro(s) - veja o log: {}".format(
        "\n".join(result_parts), total_errors, log_file))
else:
    forms.alert("Invertido com sucesso!\n\n{}".format("\n".join(result_parts)))