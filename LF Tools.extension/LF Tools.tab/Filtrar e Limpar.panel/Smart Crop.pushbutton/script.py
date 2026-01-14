# -*- coding: utf-8 -*-
"""
Script: Smart Crop - Recorte Simples
Descrição: Recorta a vista com base na seleção com margem de 2m
Compatível: IronPython 2.7 / pyRevit / Revit 2024
"""

from pyrevit import forms, revit, DB
from System.Collections.Generic import List

# Variáveis Globais
doc = revit.doc
uidoc = revit.uidoc
view = revit.active_view

def recortar_vista():
    """Recorta a vista com base nos elementos selecionados."""
    
    # Pegar seleção
    selection = revit.get_selection()
    
    if not selection:
        forms.alert("Selecione elementos para recortar a vista.", 
                   warn_icon=True, 
                   title="Nenhum elemento selecionado")
        return
    
    # Filtrar elementos válidos (remover vistas, fixados, caixas 3D)
    elementos_validos = []
    
    for el in selection:
        # Ignorar vistas
        if isinstance(el, DB.View):
            continue
        
        # Ignorar elementos fixados (pinned)
        if hasattr(el, 'Pinned') and el.Pinned:
            continue
        
        # Ignorar Section Boxes (caixas 3D)
        if el.GetType().Name == 'Element' and el.Category and el.Category.Name == '3D Section Box':
            continue
        
        elementos_validos.append(el)
    
    if not elementos_validos:
        forms.alert("Nenhum elemento válido selecionado.\n\n"
                   "Elementos ignorados:\n"
                   "- Vistas\n"
                   "- Elementos fixados\n"
                   "- Caixas 3D", 
                   warn_icon=True,
                   title="Seleção inválida")
        return
    
    # Calcular limites de todos os elementos válidos
    min_x, min_y = float('inf'), float('inf')
    max_x, max_y = float('-inf'), float('-inf')
    
    elementos_com_bbox = 0
    
    for el in elementos_validos:
        bbox = el.get_BoundingBox(view)
        if bbox:
            min_x = min(min_x, bbox.Min.X)
            min_y = min(min_y, bbox.Min.Y)
            max_x = max(max_x, bbox.Max.X)
            max_y = max(max_y, bbox.Max.Y)
            elementos_com_bbox += 1
    
    if elementos_com_bbox == 0:
        forms.alert("Os elementos selecionados não possuem limites visíveis nesta vista.", 
                   warn_icon=True,
                   title="Sem limites")
        return
    
    # Margem de 2 metros (convertido para pés - unidade interna do Revit)
    # 2 metros = 6.56168 pés
    margin = 3.28084
    
    # Aplicar recorte com TransactionGroup
    tg = DB.TransactionGroup(doc, "Recortar Vista")
    tg.Start()
    
    try:
        # Ativar crop se necessário
        if not view.CropBoxActive:
            t1 = DB.Transaction(doc, "Ativar Crop")
            t1.Start()
            view.CropBoxActive = True
            view.CropBoxVisible = True
            t1.Commit()
        
        # Pegar cropbox atual para manter a transformação
        current = view.CropBox
        
        # Ajustar cropbox
        t2 = DB.Transaction(doc, "Ajustar Crop")
        t2.Start()
        
        new_crop = DB.BoundingBoxXYZ()
        new_crop.Min = DB.XYZ(min_x - margin, min_y - margin, current.Min.Z)
        new_crop.Max = DB.XYZ(max_x + margin, max_y + margin, current.Max.Z)
        new_crop.Transform = current.Transform
        
        view.CropBox = new_crop
        
        t2.Commit()
        
        # Finalizar TransactionGroup
        tg.Assimilate()
        
        forms.alert("Vista recortada com sucesso!\n\n"
                   "Elementos processados: {}\n"
                   "Margem aplicada: 2 metros".format(elementos_com_bbox),
                   title="Recorte concluído")
        
    except Exception as e:
        tg.RollBack()
        forms.alert("Erro ao recortar vista:\n\n{}".format(str(e)), 
                   warn_icon=True,
                   title="Erro")

# ============================================================================
# EXECUTAR
# ============================================================================

if __name__ == '__main__':
    recortar_vista()