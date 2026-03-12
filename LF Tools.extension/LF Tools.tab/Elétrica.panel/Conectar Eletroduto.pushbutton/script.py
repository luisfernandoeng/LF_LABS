# -*- coding: utf-8 -*-
"""
Conectar Eletroduto - Conexão inteligente entre elementos MEP.

Shift-Click: Abre menu de configuração (Conector específico vs Caixa inteira)
Normal: Seleção automática (conector mais próximo)

Lógica de traçado:
- Mesma caixa → 45° (3 segmentos, mais suave para puxar fio)
- Caixas diferentes → 90° ortogonal (instalação embutida)
- Distâncias curtas → trecho direto
"""

__title__ = 'Conectar\nEletroduto'
__author__ = 'Luis Fernando'

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.Exceptions import OperationCanceledException 
import math
import System

from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# =====================================================================
#  FUNÇÕES AUXILIARES
# =====================================================================
def get_connectors(element):
    """Retorna todos os conectores de um elemento MEP."""
    connectors = []
    try:
        if hasattr(element, "MEPModel") and element.MEPModel and element.MEPModel.ConnectorManager:
            for c in element.MEPModel.ConnectorManager.Connectors:
                connectors.append(c)
        elif hasattr(element, "ConnectorManager") and element.ConnectorManager:
            for c in element.ConnectorManager.Connectors:
                connectors.append(c)
    except:
        pass
    return connectors


def get_connector_name_or_desc(connector):
    """Pega o nome ou descrição do conector para identificar simetria."""
    try:
        return str(connector.Description) if hasattr(connector, "Description") else ""
    except:
        return ""


def get_conduit_type(doc):
    """Pega tipo do ÚLTIMO eletroduto criado no projeto (mais recente por Id)."""
    try:
        all_conduits = FilteredElementCollector(doc).OfClass(Conduit).ToElements()
        if all_conduits:
            last = max(all_conduits, key=lambda c: c.Id.IntegerValue)
            return last.GetTypeId()
    except:
        pass
    
    # Fallback: tipo padrão da UI do Revit
    try:
        default_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ConduitType)
        if default_id != ElementId.InvalidElementId:
            return default_id
    except:
        pass
    
    # Fallback final
    collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(ConduitType))
    return collector.FirstElementId()


# =====================================================================
#  SELEÇÃO DE ELEMENTOS E CONECTORES
# =====================================================================
def find_best_connector_pair(el1, el2):
    """Lógica automática: Encontra o par de conectores mais próximos entre os dois elementos."""
    conns1 = get_connectors(el1)
    conns2 = get_connectors(el2)
    
    if not conns1 or not conns2:
        return None, None
        
    best_c1 = None
    best_c2 = None
    min_dist = float('inf')
    
    for c1 in conns1:
        for c2 in conns2:
            dist = c1.Origin.DistanceTo(c2.Origin)
            if dist < min_dist:
                min_dist = dist
                best_c1 = c1
                best_c2 = c2
                
    return best_c1, best_c2


def find_best_connector_pair_same_box(el):
    """Para mesma caixa: Encontra o par de conectores mais DISTANTES (faces opostas/adjacentes)."""
    conns = get_connectors(el)
    
    if not conns or len(conns) < 2:
        return None, None
    
    best_c1 = None
    best_c2 = None
    max_dist = 0
    
    conn_list = list(conns)
    for i in range(len(conn_list)):
        for j in range(i + 1, len(conn_list)):
            dist = conn_list[i].Origin.DistanceTo(conn_list[j].Origin)
            if dist > max_dist:
                max_dist = dist
                best_c1 = conn_list[i]
                best_c2 = conn_list[j]
    
    return best_c1, best_c2


def find_symmetric_connector_pair(el1, pt1, el2, pt2):
    """Encontra par de conectores baseado nos pontos clicados."""
    conns1 = get_connectors(el1)
    conns2 = get_connectors(el2)
    
    if not conns1 or not conns2:
        return None, None
        
    best_c1 = min(conns1, key=lambda c: c.Origin.DistanceTo(pt1))
    
    # Simetria por nome se mesma família
    same_type = False
    try:
        if el1.Symbol.Id == el2.Symbol.Id:
            same_type = True
    except:
        pass
    
    best_c2 = None
    if same_type:
        desc1 = get_connector_name_or_desc(best_c1)
        if desc1:
            for c2 in conns2:
                if get_connector_name_or_desc(c2) == desc1:
                    best_c2 = c2
                    break
    
    if not best_c2:
        best_c2 = min(conns2, key=lambda c: c.Origin.DistanceTo(pt2))
        
    return best_c1, best_c2


def pick_elements_automatic():
    """Seleção automática: dois elementos (ou mesmo elemento 2x)."""
    try:
        selected_ids = list(uidoc.Selection.GetElementIds())
        if len(selected_ids) == 2:
            return doc.GetElement(selected_ids[0]), doc.GetElement(selected_ids[1])
        
        uidoc.Selection.SetElementIds(System.Collections.Generic.List[ElementId]())
        ref1 = uidoc.Selection.PickObject(ObjectType.Element, "1. Selecione o primeiro elemento (ou caixa)")
        el1 = doc.GetElement(ref1.ElementId)
        
        ref2 = uidoc.Selection.PickObject(ObjectType.Element, "2. Selecione o segundo elemento (ou a MESMA caixa)")
        el2 = doc.GetElement(ref2.ElementId)
        
        return el1, el2
    except OperationCanceledException:
        return None, None


def pick_elements_with_points():
    """Seleção manual: clicar no ponto exato do conector."""
    try:
        ref1 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "1. Clique no CONECTOR do primeiro elemento")
        pt1 = ref1.GlobalPoint
        el1 = doc.GetElement(ref1.ElementId)
        
        ref2 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "2. Clique no CONECTOR do segundo elemento")
        pt2 = ref2.GlobalPoint
        el2 = doc.GetElement(ref2.ElementId)
        
        return el1, pt1, el2, pt2
    except OperationCanceledException:
        return None, None, None, None


# =====================================================================
#  LÓGICA DE TRAÇADO (45° e 90°)
# =====================================================================
def create_45_degree_path(p_stub1, p_stub2, dir1, dir2):
    """Cria traçado entre stubs com mínimo de segmentos (mesma caixa, 45°).
    
    Os stubs já são o reto do conector. Entre eles:
    - 1 diagonal se já estiver ~45° ou alinhado
    - 2 segmentos (reto + diagonal) se precisar mudar de eixo
    O que não der fitting, não deu. Segmentos longos = menos fittings.
    """
    dx = p_stub2.X - p_stub1.X
    dy = p_stub2.Y - p_stub1.Y
    
    # Caso 1: Alinhados ou ~45° → 1 trecho direto
    if abs(dx) < 0.05 or abs(dy) < 0.05 or abs(abs(dx) - abs(dy)) < 0.1:
        return [(p_stub1, p_stub2)]
    
    # Caso 2: 2 segmentos (reto longo + diagonal 45°)
    if abs(dx) > abs(dy):
        # Predominante em X
        diag = abs(dy)
        sign_x = 1 if dx > 0 else -1
        reto = abs(dx) - diag
        pt_mid = XYZ(p_stub1.X + sign_x * reto, p_stub1.Y, p_stub1.Z)
    else:
        # Predominante em Y
        diag = abs(dx)
        sign_y = 1 if dy > 0 else -1
        reto = abs(dy) - diag
        pt_mid = XYZ(p_stub1.X, p_stub1.Y + sign_y * reto, p_stub1.Z)
    
    return [(p_stub1, pt_mid), (pt_mid, p_stub2)]


def create_90_degree_path(pt1, pt2, dir1, dir2):
    """Cria traçado para caixas diferentes.
    
    SEMPRE vai direto entre stubs (1 segmento diagonal/inclinado).
    Os stubs já fornecem o pedacinho reto de saída do conector.
    O fitting faz a curva entre o stub e o trecho diagonal.
    Se o fitting não caber, draw_conduit_and_connect ignora e segue.
    """
    return [(pt1, pt2)]


# =====================================================================
#  CRIAÇÃO DE ELETRODUTOS
# =====================================================================
def draw_conduit_and_connect(doc, conduit_type_id, p_start, p_end, level_id, diameter, prev_cond=None):
    """Cria trecho de eletroduto e tenta conectar ao anterior.
    Se o fitting não encaixar, deixa os tubos sem curva (o usuário ajusta manual)."""
    if p_start.DistanceTo(p_end) < 0.01:
        return prev_cond
        
    c_new = Conduit.Create(doc, conduit_type_id, p_start, p_end, level_id)
    c_new.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
    
    if prev_cond:
        try:
            c1 = None
            for c in prev_cond.ConnectorManager.Connectors:
                if c.Origin.DistanceTo(p_start) < 0.05:
                    c1 = c; break
            c2 = None
            for c in c_new.ConnectorManager.Connectors:
                if c.Origin.DistanceTo(p_start) < 0.05:
                    c2 = c; break
            if c1 and c2:
                try:
                    c1.ConnectTo(c2)
                    doc.Create.NewElbowFitting(c1, c2)
                except:
                    # Fitting não coube → deixa sem curva, segue em frente
                    pass
        except:
            pass
    return c_new


# =====================================================================
#  EXECUÇÃO PRINCIPAL
# =====================================================================
def execute_connection():
    # Shift-Click: abre menu de configuração
    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False
    
    # Determinar modo de seleção
    use_connector_mode = False
    
    if is_shift:
        mode = forms.CommandSwitchWindow.show(
            ['🔌 Conector específico', '📦 Caixa inteira'],
            message='Escolha o modo de seleção:'
        )
        if not mode:
            return
        use_connector_mode = 'Conector' in mode
    
    # Seleção de elementos
    if use_connector_mode:
        el1, pt_click1, el2, pt_click2 = pick_elements_with_points()
        if not el1: return
        
        same_box = (el1.Id == el2.Id)
        
        if same_box:
            # Mesma caixa com pontos: pegar os conectores mais próximos dos cliques
            conns = get_connectors(el1)
            if len(conns) < 2:
                TaskDialog.Show("Erro", "Caixa precisa ter pelo menos 2 conectores.")
                return
            conn1 = min(conns, key=lambda c: c.Origin.DistanceTo(pt_click1))
            remaining = [c for c in conns if c != conn1]
            conn2 = min(remaining, key=lambda c: c.Origin.DistanceTo(pt_click2))
        else:
            conn1, conn2 = find_symmetric_connector_pair(el1, pt_click1, el2, pt_click2)
    else:
        el1, el2 = pick_elements_automatic()
        if not el1: return
        
        same_box = (el1.Id == el2.Id)
        
        if same_box:
            conn1, conn2 = find_best_connector_pair_same_box(el1)
        else:
            conn1, conn2 = find_best_connector_pair(el1, el2)
    
    if not conn1 or not conn2:
        TaskDialog.Show("Erro", "Não foi possível encontrar conectores.")
        return
    
    pt1 = conn1.Origin
    pt2 = conn2.Origin
    
    # Determinar propriedades
    level_id = el1.LevelId
    if str(level_id) == str(ElementId.InvalidElementId):
        level_id = doc.ActiveView.GenLevel.Id if doc.ActiveView.GenLevel else FilteredElementCollector(doc).OfClass(Level).FirstElementId()
    
    conduit_type_id = get_conduit_type(doc)
    diameter = conn1.Radius * 2 if conn1.Shape == ConnectorProfileType.Round else 0.082021
    
    t = Transaction(doc, "Conectar Eletrodutos Inteligente")
    t.Start()
    
    try:
        # Direções de saída (Stub)
        stub_len = 0.82  # ~25cm (espaço suficiente para fittings)
        
        # Distância direta entre conectores
        dist_direct = pt1.DistanceTo(pt2)
        dz = abs(pt1.Z - pt2.Z)
        is_flat = dz < 0.1
        
        # Ajuste dinâmico do stub (mínimo 0.4ft ~12cm)
        if dist_direct < 1.5:  # ~45cm
            stub_len = max(dist_direct * 0.25, 0.4)
        
        dir1 = conn1.CoordinateSystem.BasisZ
        dir2 = conn2.CoordinateSystem.BasisZ
        
        p_stub1 = pt1 + dir1 * stub_len
        p_stub2 = pt2 + dir2 * stub_len
        
        # REGRA: Distância muito curta ou alinhados → trecho direto
        dist_stubs = p_stub1.DistanceTo(p_stub2)
        dx_stubs = abs(p_stub1.X - p_stub2.X)
        dy_stubs = abs(p_stub1.Y - p_stub2.Y)
        is_aligned = (dx_stubs < 0.05 or dy_stubs < 0.05)
        
        use_direct = False
        if dist_direct < 0.66:  # < 0.2m → muito perto, direto
            use_direct = True
        elif dist_direct < 3.28 and is_flat and is_aligned:  # < 1m, plano, alinhado
            use_direct = True
        
        # 1. Criar Stub inicial e conectar na caixa 1
        cond1 = Conduit.Create(doc, conduit_type_id, pt1, p_stub1, level_id)
        cond1.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
        for c in cond1.ConnectorManager.Connectors:
            if c.Origin.DistanceTo(pt1) < 0.05:
                c.ConnectTo(conn1); break
        
        last_cond = cond1
        
        if use_direct:
            # Conexão direta entre stubs
            last_cond = draw_conduit_and_connect(doc, conduit_type_id, p_stub1, p_stub2, level_id, diameter, last_cond)
        elif same_box:
            # MESMA CAIXA → 45° (3 segmentos, mais suave para puxar fio)
            segments = create_45_degree_path(p_stub1, p_stub2, dir1, dir2)
            for (p_start, p_end) in segments:
                if p_start.DistanceTo(p_end) > 0.05:
                    last_cond = draw_conduit_and_connect(doc, conduit_type_id, p_start, p_end, level_id, diameter, last_cond)
        else:
            # CAIXAS DIFERENTES → 90° ortogonal (instalação embutida)
            segments = create_90_degree_path(p_stub1, p_stub2, dir1, dir2)
            for (p_start, p_end) in segments:
                if p_start.DistanceTo(p_end) > 0.05:
                    last_cond = draw_conduit_and_connect(doc, conduit_type_id, p_start, p_end, level_id, diameter, last_cond)
        
        # 3. Stub final e conectar na caixa 2
        cond2 = Conduit.Create(doc, conduit_type_id, pt2, p_stub2, level_id)
        cond2.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
        for c in cond2.ConnectorManager.Connectors:
            if c.Origin.DistanceTo(pt2) < 0.05:
                c.ConnectTo(conn2); break
                
        # Ligar o último trecho do meio com o stub final
        if last_cond and last_cond.Id != cond2.Id:
            try:
                c_mid = next(c for c in last_cond.ConnectorManager.Connectors if c.Origin.DistanceTo(p_stub2) < 0.05)
                c_last = next(c for c in cond2.ConnectorManager.Connectors if c.Origin.DistanceTo(p_stub2) < 0.05)
                if c_mid and c_last:
                    c_mid.ConnectTo(c_last)
                    doc.Create.NewElbowFitting(c_mid, c_last)
            except:
                pass
                
        t.Commit()
    except Exception as e:
        if t.HasStarted(): t.RollBack()
        TaskDialog.Show("Erro", "Erro ao conectar: " + str(e))

if __name__ == "__main__":
    execute_connection()
