# -*- coding: utf-8 -*-
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

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def get_connectors(element):
    connectors = []
    if hasattr(element, "MEPModel") and element.MEPModel and element.MEPModel.ConnectorManager:
        for c in element.MEPModel.ConnectorManager.Connectors:
            connectors.append(c)
    elif type(element) is FamilyInstance:
        # Tenta pegar conectores da family instance (pode ser MechanicalEquipment, etc)
        try:
            for c in element.MEPModel.ConnectorManager.Connectors:
                connectors.append(c)
        except:
            pass
    return connectors

def find_best_connector_pair(el1, el2):
    """Encontra o par de conectores (um de cada elemento) mais próximos entre si."""
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

def get_conduit_type(doc):
    """Pega um tipo de eletroduto padrão válido para desenhar, preferencialmente o ultimo usado"""
    try:
        # Tenta pegar o default selecionado na UI do Revit (Revit 2022+)
        default_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ConduitType)
        if default_id != ElementId.InvalidElementId:
            return default_id
    except:
        pass
        
    try:
        # Fallback 1: Pega o ultimo eletroduto modelado (Maior ID)
        conduits = doc.GetElement(FilteredElementCollector(doc).OfClass(clr.GetClrType(Electrical.Conduit)).ToElementIds())
        # Na verdade, ToElements e sort
        all_conduits = FilteredElementCollector(doc).OfClass(clr.GetClrType(Electrical.Conduit)).ToElements()
        if all_conduits:
            last_conduit = sorted(all_conduits, key=lambda c: c.Id.IntegerValue)[-1]
            return last_conduit.GetTypeId()
    except:
        pass
        
    # Fallback final: Primeiro tipo carregado
    collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(Electrical.ConduitType))
    return collector.FirstElementId()
    
def select_two_boxes():
    # 1. Tentar pegar elementos pré-selecionados
    selected_ids = list(uidoc.Selection.GetElementIds())
    if len(selected_ids) == 2:
        return selected_ids[0], selected_ids[1]
    
    # Se não tiver exatamente 2 pré-selecionados, força limpeza e pede clique manual
    uidoc.Selection.SetElementIds(System.Collections.Generic.List[ElementId]())
    try:
        ref1 = uidoc.Selection.PickObject(ObjectType.Element, "1. Selecione a PRIMEIRA caixa/quadro (Origem)")
        ref2 = uidoc.Selection.PickObject(ObjectType.Element, "2. Selecione a SEGUNDA caixa/quadro (Destino)")
        return ref1.ElementId, ref2.ElementId
    except OperationCanceledException:
        return None, None
def create_sloped_conduits():
    id1, id2 = select_two_boxes()
    if not id1 or not id2:
        return
        
    el1 = doc.GetElement(id1)
    el2 = doc.GetElement(id2)
    
    conn1, conn2 = find_best_connector_pair(el1, el2)
    
    if not conn1 or not conn2:
        TaskDialog.Show("Erro", "Não foi possível encontrar conectores livres nas caixas selecionadas.")
        return
        
    # Extrair informacoes
    pt1 = conn1.Origin
    pt2 = conn2.Origin
    
    # Determinar propriedades do eletroduto
    # Copiar diametro e tipo do proprio conector se possivel, ou padrao
    level_id = el1.LevelId
    if str(level_id) == str(ElementId.InvalidElementId):
        level_id = el2.LevelId
        
    if str(level_id) == str(ElementId.InvalidElementId):
        if doc.ActiveView.GenLevel:
            level_id = doc.ActiveView.GenLevel.Id
        else:
            from Autodesk.Revit.DB import FilteredElementCollector, Level
            level_id = FilteredElementCollector(doc).OfClass(Level).FirstElementId()
        
    conduit_type_id = get_conduit_type(doc)
    
    diameter = 0.082021 # default 25mm (em pes)
    if conn1.Shape == ConnectorProfileType.Round:
        diameter = conn1.Radius * 2
        
    # O conceito da rampa (escadinha de eletrodutos):
    # A proposta do usuario: criar varios trechos inclinados gradualmente ate o objetivo.
    # Abordagem:
    # 1. Toco saindo reto de 1 (garante 90 graus na saida)
    # 2. Toco chegando reto em 2 (garante 90 graus na chegada)
    # 3. Interligacao inclinada (Opcionalmente em segmentos se o aclive for muito bruto)
    
    # Comprimento do toco (ex: 20 cm = 0.65 feet) para ter espaço de sobra pro cotovelo
    stub_length = 0.65
    
    # Ponto apos sair reto do conn1
    dir1 = conn1.CoordinateSystem.BasisZ
    pt_stub1 = pt1 + dir1 * stub_length
    
    # Ponto apos sair reto do conn2
    dir2 = conn2.CoordinateSystem.BasisZ
    pt_stub2 = pt2 + dir2 * stub_length
    
    # Agora ligar stub1 com stub2
    total_dist = pt_stub1.DistanceTo(pt_stub2)
    
    # Verifica Diferença de Nível (Plano vs Aclive)
    dz = abs(pt_stub1.Z - pt_stub2.Z)
    is_flat = dz < 0.1  # Considera plano se diferença for menor que ~3cm
    
    # Se for muito curta e reta
    if total_dist < 0.2 and not is_flat:
        TaskDialog.Show("Aviso", "A distância entre as caixas está muito curta para gerar cotovelos.")
        return
        
    t = Transaction(doc, "Conectar Eletrodutos")
    t.Start()
    
    try:
        from Autodesk.Revit.DB.Electrical import Conduit
        
        # Conexao DIRETA se muito perto e plano
        if total_dist < 0.2 and is_flat:
            # Traçado de um unico tubo conectando conn1 e conn2 diretamente
            cond_direct = Conduit.Create(doc, conduit_type_id, pt1, pt2, level_id)
            cond_direct.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
            for c in cond_direct.ConnectorManager.Connectors:
                if c.Origin.DistanceTo(pt1) < 0.05:
                    c.ConnectTo(conn1)
                elif c.Origin.DistanceTo(pt2) < 0.05:
                    c.ConnectTo(conn2)
            
            t.Commit()
            return
        
        # Helper interno para desenhar duto e tentar cotovelo com o anterior
        def draw_conduit_and_connect(p_start, p_end, prev_cond=None):
            if p_start.DistanceTo(p_end) < 0.05:
                # Distancia mt curta para duto isolado
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
                        c1.ConnectTo(c2)
                        doc.Create.NewElbowFitting(c1, c2)
                except:
                    pass
            return c_new
        
        # 1. Criar Stub 1 ligando na caixa 1
        cond1 = Conduit.Create(doc, conduit_type_id, pt1, pt_stub1, level_id)
        cond1.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
        for c in cond1.ConnectorManager.Connectors:
            if c.Origin.DistanceTo(pt1) < 0.05:
                c.ConnectTo(conn1)
                break
                
        last_cond = cond1
                
        # 2. Desenvolver o trecho central (MID)
        if is_flat:
            # Traçado Ortogonal Limitado: Apenas 1 L (Anda em X, depois em Y)
            pt_corner = XYZ(pt_stub2.X, pt_stub1.Y, pt_stub1.Z)
            
            # Se alinhado num eixo, a distancia para o corner sera 0
            last_cond = draw_conduit_and_connect(pt_stub1, pt_corner, last_cond)
            last_cond = draw_conduit_and_connect(pt_corner, pt_stub2, last_cond)
            
        else:
            # Traçado Adaptativo de Aclive (Sobe "fatiado" em N segmentos pra suportar cotovelos pequenos)
            max_seg = 0.5  # Fatias curtas de ~15cm
            n_segments = int(math.ceil(total_dist / max_seg))
            if n_segments < 1: n_segments = 1
            
            dx = (pt_stub2.X - pt_stub1.X) / float(n_segments)
            dy = (pt_stub2.Y - pt_stub1.Y) / float(n_segments)
            dz_step = (pt_stub2.Z - pt_stub1.Z) / float(n_segments) # named dz_step to not conflict with dz
            
            curr_pt = pt_stub1
            for i in range(n_segments):
                next_pt = XYZ(curr_pt.X + dx, curr_pt.Y + dy, curr_pt.Z + dz_step)
                # Ultimo ponto garante precisao exata
                if i == n_segments - 1:
                    next_pt = pt_stub2
                    
                last_cond = draw_conduit_and_connect(curr_pt, next_pt, last_cond)
                curr_pt = next_pt

        # 3. Criar Stub 2 chegando na caixa 2 (pt2 -> pt_stub2)
        cond2 = Conduit.Create(doc, conduit_type_id, pt2, pt_stub2, level_id)
        cond2.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
        for c in cond2.ConnectorManager.Connectors:
            if c.Origin.DistanceTo(pt2) < 0.05:
                c.ConnectTo(conn2)
                break
                
        # Conectar Mid com Stub2
        if last_cond and last_cond.Id != cond2.Id: # Evita tentar ligar nele mesmo
            try:
                c_mid = None
                for c in last_cond.ConnectorManager.Connectors:
                    if c.Origin.DistanceTo(pt_stub2) < 0.05:
                        c_mid = c; break
                c_last = None
                for c in cond2.ConnectorManager.Connectors:
                    if c.Origin.DistanceTo(pt_stub2) < 0.05:
                        c_last = c; break
                if c_last and c_mid:
                    c_last.ConnectTo(c_mid)
                    doc.Create.NewElbowFitting(c_last, c_mid)
            except:
                pass
            
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        uidoc.Application.Application.WriteJournalComment("Erro Conectar Eletroduto: " + str(e), True)
        TaskDialog.Show("Erro", "Erro ao desenhar: " + str(e))

create_sloped_conduits()
