# -*- coding: utf-8 -*-
"""
Script: Smart Crop Turbo - OTIMIZADO
Descricao: Recorte ultra-rapido (2 cliques) e analise de elementos perdidos.
Otimizacoes: Snaps minimos, caching, colecoes rapidas
"""

from pyrevit import forms, revit, DB, UI
from System.Collections.Generic import List # Import necessario para criar listas .NET

# Variaveis Globais
doc = revit.doc
uidoc = revit.uidoc
view = revit.active_view

# ============================================================================
# 1. ANALISE DE LIXO (OTIMIZADA)
# ============================================================================

def analyze_garbage():
    """Inverte a selecao para achar elementos longe. VERSAO OTIMIZADA."""
    selection = revit.get_selection()
    
    if not selection:
        forms.alert("Selecione o que esta CORRETO primeiro.", warn_icon=True)
        return

    # OTIMIZACAO 1: Calculo de limites em uma passada
    bounds = {'min_x': float('inf'), 'min_y': float('inf'), 
              'max_x': float('-inf'), 'max_y': float('-inf')}
    
    selected_ids = set()
    
    for el in selection:
        selected_ids.add(el.Id.IntegerValue)
        bbox = el.get_BoundingBox(view)
        if bbox:
            bounds['min_x'] = min(bounds['min_x'], bbox.Min.X)
            bounds['min_y'] = min(bounds['min_y'], bbox.Min.Y)
            bounds['max_x'] = max(bounds['max_x'], bbox.Max.X)
            bounds['max_y'] = max(bounds['max_y'], bbox.Max.Y)
    
    if bounds['min_x'] == float('inf'):
        return

    tolerance = 1.0
    view_id = view.Id.IntegerValue
    
    # OTIMIZACAO 2: Filtro rapido - apenas visiveis
    collector = (DB.FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType())
    
    # OTIMIZACAO 3: Lista de IDs direto (evita .ToElements())
    culprit_ids = []
    
    for el_id in collector.ToElementIds():
        id_val = el_id.IntegerValue
        
        # Skip rapido
        if id_val == view_id or id_val in selected_ids:
            continue
        
        # OTIMIZACAO 4: GetElement apenas se necessario
        el = doc.GetElement(el_id)
        if not el or not el.Category:
            continue
        
        bbox = el.get_BoundingBox(view)
        if not bbox:
            continue
        
        # Check rapido: esta fora?
        if (bbox.Max.X < bounds['min_x'] - tolerance or 
            bbox.Min.X > bounds['max_x'] + tolerance or 
            bbox.Max.Y < bounds['min_y'] - tolerance or 
            bbox.Min.Y > bounds['max_y'] + tolerance):
            culprit_ids.append(el_id)

    if culprit_ids:
        # OTIMIZACAO 5: SetElementIds direto (mais rapido que set_to)
        # CORRECAO APLICADA AQUI:
        uidoc.Selection.SetElementIds(List[DB.ElementId](culprit_ids))
        
        forms.alert("Encontrados {} elementos fora da area.".format(len(culprit_ids)), 
                   title="Analise Completa")
    else:
        forms.alert("Vista Limpa! Nada fora da area.", title="Analise Completa")

# ============================================================================
# 2. AJUSTE AUTOMATICO
# ============================================================================

def auto_adjust_crop():
    """Ajusta o cropbox automaticamente para a selecao. RAPIDO."""
    selection = revit.get_selection()
    
    if not selection:
        forms.alert("Selecione elementos primeiro.", warn_icon=True)
        return
    
    # Calculo rapido de limites
    min_x, min_y = float('inf'), float('inf')
    max_x, max_y = float('-inf'), float('-inf')
    
    for el in selection:
        bbox = el.get_BoundingBox(view)
        if bbox:
            min_x = min(min_x, bbox.Min.X)
            min_y = min(min_y, bbox.Min.Y)
            max_x = max(max_x, bbox.Max.X)
            max_y = max(max_y, bbox.Max.Y)
    
    if min_x == float('inf'):
        return
    
    # Margem
    margin = 2.0
    
    # OTIMIZACAO: TransactionGroup para evitar regeneracoes multiplas
    tg = DB.TransactionGroup(doc, "Auto Ajuste Crop")
    tg.Start()
    
    try:
        # Sub-transacao 1: Ativar crop
        if not view.CropBoxActive:
            t1 = DB.Transaction(doc, "Ativar Crop")
            t1.Start()
            view.CropBoxActive = True
            view.CropBoxVisible = True
            t1.Commit()
        
        current = view.CropBox
        
        # Sub-transacao 2: Ajustar
        t2 = DB.Transaction(doc, "Ajustar Crop")
        t2.Start()
        
        new_crop = DB.BoundingBoxXYZ()
        new_crop.Min = DB.XYZ(min_x - margin, min_y - margin, current.Min.Z)
        new_crop.Max = DB.XYZ(max_x + margin, max_y + margin, current.Max.Z)
        new_crop.Transform = current.Transform
        
        view.CropBox = new_crop
        
        t2.Commit()
        
        # COMMIT FINAL: Apenas UMA regeneracao
        tg.Assimilate()
        
    except:
        tg.RollBack()
        raise
    
    forms.alert("Cropbox ajustado para a selecao!", title="Sucesso")

# ============================================================================
# MAIN - INTERFACE OTIMIZADA
# ============================================================================

# Opcoes organizadas
opcoes = {
    '1. Analisar Lixo (Inverte Selecao)': 'ANALYZE',
    '2. Auto Ajustar (da Selecao)': 'AUTO'
}

resposta = forms.CommandSwitchWindow.show(
    sorted(opcoes.keys()),
    message='Smart Crop Turbo - Otimizado',
)

if resposta:
    modo = opcoes[resposta]
    
    if modo == 'ANALYZE':
        analyze_garbage()
    elif modo == 'AUTO':
        auto_adjust_crop()