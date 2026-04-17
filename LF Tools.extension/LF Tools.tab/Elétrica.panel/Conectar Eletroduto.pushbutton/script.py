#! python3
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

# ╔══════════════════════════════════════════════════════════════╗
# ║                    MODO DEBUG                                ║
# ║  True  = imprime tudo no console pyRevit + mostra erros      ║
# ║  False = silencioso (só mostra alert em caso de erro fatal)  ║
# ╚══════════════════════════════════════════════════════════════╝
DEBUG_MODE = False

# =====================================================================
#  IMPORTS
# =====================================================================
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB import BuiltInParameter, ElementId, StorageType, BuiltInCategory, XYZ
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import OperationCanceledException
import System
from System.Collections.Generic import List
from collections import OrderedDict
import math
import traceback

from pyrevit import forms           # script importado de forma lazy em load/save_config
from lf_utils import DebugLogger, get_revit_context, patch_forms, make_warning_swallower, get_script_config, save_script_config

patch_forms(forms)

# Instância global — usar `dbg` em todo o script
dbg = DebugLogger(DEBUG_MODE)

# Referências globais — preenchidas na primeira execução
uidoc = None
doc   = None

# =====================================================================
#  HELPERS DE NOME
# =====================================================================
def __get_name__(obj):
    try: return obj.Name
    except: pass
    try: return Element.Name.GetValue(obj)
    except: pass
    try:
        p = obj.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.HasValue: return p.AsString()
    except: pass
    return ""

def __get_family_name__(obj):
    try: return obj.FamilyName
    except: pass
    try: return FamilySymbol.FamilyName.GetValue(obj)
    except: pass
    try:
        p = obj.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if p and p.HasValue: return p.AsString()
    except: pass
    return ""

# =====================================================================
#  CONFIGURAÇÕES PERSISTENTES
# =====================================================================
def load_config():
    return get_script_config(__commandpath__, defaults={
        'selection_mode':   'conector',
        'conduit_type':     '',
        'default_diameter': '',
        'service_type':     '',
        'routing_strategy': 'auto',
    })

def save_config(settings):
    save_script_config(__commandpath__, settings)

def show_settings():
    settings = load_config()

    while True:
        mode_str  = "Conector Específico" if settings.get('selection_mode', 'conector') == 'conector' else "Caixa Inteira"
        cur_type  = settings.get('conduit_type', '')     or '(Usar Último Desenhado)'
        cur_diam  = settings.get('default_diameter', '') or '(Automático - Puxa do Último/Conector)'
        cur_serv  = settings.get('service_type', '')     or '(Nenhum)'
        strat     = settings.get('routing_strategy', 'auto')
        strat_str = "Auto (direto → calculado)" if strat == 'auto' else "Sempre Calculado"

        opcoes = OrderedDict([
            ("1. Modo de Seleção: "             + mode_str,  "selection_mode"),
            ("2. Tipo de Eletroduto Padrão: "   + cur_type,  "conduit_type"),
            ("3. Diâmetro Padrão (mm): "        + cur_diam,  "default_diameter"),
            ("4. Texto para Tipo de Serviço: "  + cur_serv,  "service_type"),
            ("5. Estratégia de Rota: "          + strat_str, "routing_strategy"),
            ("6. Salvar e Sair", "save"),
        ])

        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(),
            message="Configurações: Conectar Eletroduto",
            title="Shift+Click - Configurações"
        )

        if not escolha or opcoes.get(escolha) == "save":
            save_config(settings)
            forms.toast("Configurações salvas!", title="Conectar Eletroduto")
            break

        key = opcoes[escolha]

        if key == 'selection_mode':
            chosen = forms.CommandSwitchWindow.show(
                ['Conector Específico', 'Caixa Inteira'],
                message="Selecione o modo de seleção padrão:",
                title="Modo de Seleção"
            )
            if chosen:
                settings['selection_mode'] = 'conector' if 'Conector' in chosen else 'caixa'

        elif key == 'conduit_type':
            collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(ConduitType))
            types = {}
            for t in collector:
                try:
                    name = t.Name
                except Exception:
                    p = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    name = p.AsString() if (p and p.HasValue) else str(t.Id.IntegerValue)
                if name:
                    types[name] = name
            if not types:
                forms.alert("Nenhum tipo de eletroduto encontrado.")
                continue
            chosen = forms.CommandSwitchWindow.show(
                ["(Usar Último Desenhado)", "(Padrão do Revit)"] + sorted(types.keys()),
                message="Selecione o Tipo de Eletroduto:",
                title="Tipo de Eletroduto"
            )
            if chosen:
                settings['conduit_type'] = chosen

        elif key == 'default_diameter':
            val = forms.ask_for_string(
                default=settings.get('default_diameter', ''),
                prompt="Diâmetro padrão em mm (ex: 25, 32). Deixe EM BRANCO para (Automático):",
                title="Diâmetro"
            )
            if val is not None:
                settings['default_diameter'] = val.strip()

        elif key == 'service_type':
            val = forms.ask_for_string(
                default=settings.get('service_type', ''),
                prompt="Texto fixo para Tipo de Serviço (Deixe EM BRANCO para limpar):",
                title="Tipo de Serviço"
            )
            if val is not None:
                settings['service_type'] = val.strip()

        elif key == 'routing_strategy':
            chosen = forms.CommandSwitchWindow.show(
                ['Auto (direto → calculado)', 'Sempre Calculado'],
                message=(
                    "Auto: tenta 1 tubo direto com curva automática do Revit.\n"
                    "Se não encaixar, cai no caminho calculado (multi-segmento).\n\n"
                    "Sempre Calculado: sempre usa o caminho ortogonal planejado."
                ),
                title="Estratégia de Rota"
            )
            if chosen:
                settings['routing_strategy'] = 'auto' if 'Auto' in chosen else 'calculado'

# =====================================================================
#  FUNÇÕES AUXILIARES
# =====================================================================
def get_connectors(element, include_connected=False):
    """Retorna conectores do domínio Conduit, filtrando os já conectados por padrão."""
    connectors = []
    try:
        mgr = None
        if hasattr(element, "MEPModel") and element.MEPModel and element.MEPModel.ConnectorManager:
            mgr = element.MEPModel.ConnectorManager
        elif hasattr(element, "ConnectorManager") and element.ConnectorManager:
            mgr = element.ConnectorManager
        if mgr:
            for c in mgr.Connectors:
                if c.Domain != Domain.DomainCableTrayConduit:
                    continue
                if not include_connected and c.IsConnected:
                    continue
                connectors.append(c)
    except Exception as e:
        dbg.debug("get_connectors falhou: {}".format(e))
    return connectors


def get_connector_name_or_desc(connector):
    try:
        return str(connector.Description) if hasattr(connector, "Description") else ""
    except Exception:
        return ""


def get_last_conduit(doc):
    try:
        all_ids = FilteredElementCollector(doc).OfClass(Conduit).ToElementIds()
        if all_ids:
            max_val = max(eid.IntegerValue for eid in all_ids)
            return doc.GetElement(ElementId(System.Int64(max_val)))
    except Exception:
        pass
    return None


def get_default_conduit_type(doc):
    try:
        default_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ConduitType)
        if default_id != ElementId.InvalidElementId:
            return default_id
    except Exception:
        pass
    collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(ConduitType))
    return collector.FirstElementId()


def copy_conduit_parameters(source, target):
    """Copia parâmetros de instância (texto e inteiros) do source para o target."""
    if not source or not target:
        return
    for p_src in source.Parameters:
        if not p_src.HasValue or p_src.IsReadOnly:
            continue
        bip = BuiltInParameter.INVALID
        try:
            if hasattr(p_src.Definition, "BuiltInParameter"):
                bip = p_src.Definition.BuiltInParameter
        except Exception as e:
            dbg.debug("copy_conduit_parameters BIP: {}".format(e))
        if bip in [
            BuiltInParameter.ALL_MODEL_MARK,
            BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM,
            BuiltInParameter.CURVE_ELEM_LENGTH,
            BuiltInParameter.ELEM_PARTITION_PARAM,
        ]:
            continue
        st = p_src.StorageType
        if st not in [StorageType.String, StorageType.Integer]:
            continue
        try:
            p_tgt = None
            if p_src.IsShared:
                p_tgt = target.get_Parameter(p_src.GUID)
            else:
                if bip != BuiltInParameter.INVALID:
                    p_tgt = target.get_Parameter(bip)
                if not p_tgt:
                    p_tgt = target.LookupParameter(p_src.Definition.Name)
            if not p_tgt or p_tgt.IsReadOnly:
                continue
            if st == StorageType.String:
                val = p_src.AsString()
                if val:
                    p_tgt.Set(val)
            elif st == StorageType.Integer:
                p_tgt.Set(p_src.AsInteger())
        except Exception as e:
            dbg.debug("copy_conduit_parameters set: {}".format(e))

# =====================================================================
#  SELEÇÃO DE ELEMENTOS E CONECTORES
# =====================================================================
def find_best_connector_pair(el1, el2):
    conns1 = get_connectors(el1)
    conns2 = get_connectors(el2)
    if not conns1 or not conns2:
        return None, None

    best_c1 = best_c2 = None
    best_score = float('inf')
    is_same_type = False
    try:
        if el1.Symbol.Id == el2.Symbol.Id:
            is_same_type = True
    except Exception:
        pass

    for c1 in conns1:
        desc1 = get_connector_name_or_desc(c1)
        z1 = c1.Origin.Z
        for c2 in conns2:
            desc2 = get_connector_name_or_desc(c2)
            z2 = c2.Origin.Z
            dist = c1.Origin.DistanceTo(c2.Origin)
            score = dist
            if is_same_type:
                if desc1 and desc2 and desc1 == desc2:
                    score -= 1000
                elif abs(z1 - z2) < 0.1:
                    score -= 500
            else:
                if abs(z1 - z2) < 0.1:
                    score -= 50
            if score < best_score:
                best_score = score
                best_c1 = c1
                best_c2 = c2
    return best_c1, best_c2


def find_best_connector_pair_same_box(el):
    conns = get_connectors(el)
    if not conns or len(conns) < 2:
        return None, None
    best_c1 = best_c2 = None
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
    conns1 = get_connectors(el1)
    conns2 = get_connectors(el2)
    if not conns1 or not conns2:
        return None, None
    best_c1 = min(conns1, key=lambda c: c.Origin.DistanceTo(pt1))
    same_type = False
    try:
        if el1.Symbol.Id == el2.Symbol.Id:
            same_type = True
    except Exception:
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
    try:
        selected_ids = list(uidoc.Selection.GetElementIds())
        if len(selected_ids) == 2:
            return doc.GetElement(selected_ids[0]), doc.GetElement(selected_ids[1])
        uidoc.Selection.SetElementIds(System.Collections.Generic.List[ElementId]())
        ref1 = uidoc.Selection.PickObject(ObjectType.Element, "1. Selecione o primeiro elemento (ou caixa)")
        el1  = doc.GetElement(ref1.ElementId)
        ref2 = uidoc.Selection.PickObject(ObjectType.Element, "2. Selecione o segundo elemento (ou a MESMA caixa)")
        el2  = doc.GetElement(ref2.ElementId)
        return el1, el2
    except OperationCanceledException:
        return None, None


def pick_elements_with_points():
    try:
        ref1 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "1. Clique no CONECTOR do primeiro elemento")
        pt1  = ref1.GlobalPoint
        el1  = doc.GetElement(ref1.ElementId)
        ref2 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "2. Clique no CONECTOR do segundo elemento")
        pt2  = ref2.GlobalPoint
        el2  = doc.GetElement(ref2.ElementId)
        return el1, pt1, el2, pt2
    except OperationCanceledException:
        return None, None, None, None

# =====================================================================
#  LÓGICA DE TRAÇADO (45° e 90°)
# =====================================================================
def solve_chicane_2d(pt1, pt2, dir1, dir2, min_stub_len=0.25):
    dx = pt2.X - pt1.X
    dy = pt2.Y - pt1.Y
    D_vec = XYZ(dx, dy, 0)
    if D_vec.IsAlmostEqualTo(XYZ.Zero):
        return None, None
    cos45 = sin45 = 0.70710678
    U_a = XYZ(dir1.X*cos45 - dir1.Y*sin45, dir1.X*sin45 + dir1.Y*cos45, 0)
    U_b = XYZ(dir1.X*cos45 + dir1.Y*sin45, -dir1.X*sin45 + dir1.Y*cos45, 0)
    U_chosen = U_a if U_a.DotProduct(D_vec) > U_b.DotProduct(D_vec) else U_b

    def try_solve(U):
        denom_sym = (dir2.X - dir1.X)*U.Y - (dir2.Y - dir1.Y)*U.X
        if abs(denom_sym) > 0.01:
            s_sym = (dy*U.X - dx*U.Y) / denom_sym
            if s_sym >= min_stub_len - 0.01:
                p1_c = XYZ(pt1.X + dir1.X*s_sym, pt1.Y + dir1.Y*s_sym, pt1.Z)
                p2_c = XYZ(pt2.X + dir2.X*s_sym, pt2.Y + dir2.Y*s_sym, pt2.Z)
                k = (p2_c.X - p1_c.X)*U.X + (p2_c.Y - p1_c.Y)*U.Y
                if k > 0.05:
                    return s_sym, s_sym, k
        det = U.X * dir2.Y - U.Y * dir2.X
        if abs(det) < 0.01:
            return None, None, None
        bx = min_stub_len * dir1.X - dx
        by = min_stub_len * dir1.Y - dy
        s2 = (-U.Y * bx + U.X * by) / det
        k  = (dir2.X * by - dir2.Y * bx) / det
        if s2 >= min_stub_len - 0.01 and k > 0.05:
            return min_stub_len, s2, k
        bbx = dx + min_stub_len * dir2.X
        bby = dy + min_stub_len * dir2.Y
        det_alt = dir1.X * U.Y - dir1.Y * U.X
        if abs(det_alt) > 0.01:
            s1    = (U.Y * bbx - U.X * bby) / det_alt
            k_alt = (dir1.X * bby - dir1.Y * bbx) / det_alt
            if s1 >= min_stub_len - 0.01 and k_alt > 0.05:
                return s1, min_stub_len, k_alt
        return None, None, None

    res = try_solve(U_chosen)
    if not res[0]:
        U_other = U_b if U_chosen == U_a else U_a
        res = try_solve(U_other)
    if res[0]:
        s1, s2, k = res
        p1 = XYZ(pt1.X + dir1.X * s1, pt1.Y + dir1.Y * s1, pt1.Z)
        p2 = XYZ(pt2.X + dir2.X * s2, pt2.Y + dir2.Y * s2, pt2.Z)
        return p1, p2
    return None, None


def create_45_degree_path(p_stub1, p_stub2, dir1, dir2):
    dx = p_stub2.X - p_stub1.X
    dy = p_stub2.Y - p_stub1.Y
    if abs(dx) < 0.05 or abs(dy) < 0.05 or abs(abs(dx) - abs(dy)) < 0.1:
        return [(p_stub1, p_stub2)]
    if abs(dx) > abs(dy):
        diag   = abs(dy)
        sign_x = 1 if dx > 0 else -1
        reto   = abs(dx) - diag
        pt_mid = XYZ(p_stub1.X + sign_x * reto, p_stub1.Y, p_stub1.Z)
    else:
        diag   = abs(dx)
        sign_y = 1 if dy > 0 else -1
        reto   = abs(dy) - diag
        pt_mid = XYZ(p_stub1.X, p_stub1.Y + sign_y * reto, p_stub1.Z)
    return [(p_stub1, pt_mid), (pt_mid, p_stub2)]


def create_90_degree_path(p_stub1, p_stub2, dir1, dir2, force_vertical_drop=False):
    dz = p_stub2.Z - p_stub1.Z
    dx = p_stub2.X - p_stub1.X
    dy = p_stub2.Y - p_stub1.Y
    aligned_axes = sum([abs(dx) < 0.05, abs(dy) < 0.05, abs(dz) < 0.05])
    if aligned_axes >= 2:
        return [(p_stub1, p_stub2)]
    points = [p_stub1]
    if force_vertical_drop and abs(dz) > 0.3:
        axes = ['X', 'Y', 'Z'] if dz < 0 else ['Z', 'X', 'Y']
    else:
        axes = ['X', 'Y', 'Z'] if abs(dx) > abs(dy) else ['Y', 'X', 'Z']
    current_pt = p_stub1
    for axis in axes:
        next_pt = None
        if axis == 'X' and abs(p_stub2.X - current_pt.X) > 0.05:
            next_pt = XYZ(p_stub2.X, current_pt.Y, current_pt.Z)
        elif axis == 'Y' and abs(p_stub2.Y - current_pt.Y) > 0.05:
            next_pt = XYZ(current_pt.X, p_stub2.Y, current_pt.Z)
        elif axis == 'Z' and abs(p_stub2.Z - current_pt.Z) > 0.05:
            next_pt = XYZ(current_pt.X, current_pt.Y, p_stub2.Z)
        if next_pt:
            points.append(next_pt)
            current_pt = next_pt
    if points[-1].DistanceTo(p_stub2) > 0.01:
        points.append(p_stub2)
    segments = []
    for i in range(len(points) - 1):
        if points[i].DistanceTo(points[i+1]) > 0.05:
            segments.append((points[i], points[i+1]))
    if not segments and p_stub1.DistanceTo(p_stub2) > 0.05:
        segments.append((p_stub1, p_stub2))
    return segments

# =====================================================================
#  CRIAÇÃO DE ELETRODUTOS E CURVAS
# =====================================================================
def _connect_endpoint(doc, conduit, pt_near, target_conn, label="endpoint"):
    """Conecta a ponta do eletroduto mais próxima de pt_near ao target_conn.

    Tenta elbow físico primeiro (cria curva + conexão lógica).
    Fallback: ConnectTo lógico sem curva.
    Tolerância: 0.2 ft (~6 cm).
    """
    if target_conn is None or target_conn.IsConnected:
        return
    # Achar o conector livre mais próximo de pt_near
    c_near = None
    best_dist = float('inf')
    try:
        for c in conduit.ConnectorManager.Connectors:
            if c.IsConnected:
                continue
            d = c.Origin.DistanceTo(pt_near)
            if d < best_dist:
                best_dist = d
                c_near = c
    except Exception as e:
        dbg.debug("{}: erro ao iterar conectores — {}".format(label, e))
        return

    if c_near is None or best_dist > 0.2:
        dbg.debug("{}: nenhum conector livre dentro de 0.2 ft (melhor={:.4f})".format(
            label, best_dist))
        return

    # Tenta elbow (cria curva física + liga logicamente)
    try:
        len_p = conduit.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        length = len_p.AsDouble() if len_p else 1.0
        needs_elbow = c_near.Origin.DistanceTo(target_conn.Origin) >= 0.01
        if length >= 0.15 and needs_elbow:
            doc.Create.NewElbowFitting(c_near, target_conn)
            dbg.result(True, "{}: elbow criado".format(label))
            return
    except Exception as e:
        dbg.debug("{}: elbow falhou ({}) — tentando ConnectTo".format(label, e))

    # Fallback: conexão lógica sem curva física
    try:
        c_near.ConnectTo(target_conn)
        dbg.result(True, "{}: ConnectTo OK".format(label))
    except Exception as e:
        dbg.debug("{}: ConnectTo falhou — {}".format(label, e))


def draw_conduit_and_connect(doc, conduit_type_id, p_start, p_end, level_id, diameter,
                              prev_cond=None, last_ref_conduit=None):
    """Cria trecho de eletroduto e tenta conectar ao anterior."""
    if p_start.DistanceTo(p_end) < 0.01:
        return prev_cond
    c_new = Conduit.Create(doc, conduit_type_id, p_start, p_end, level_id)
    c_new.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
    if last_ref_conduit:
        copy_conduit_parameters(last_ref_conduit, c_new)
    if prev_cond:
        c1 = next((c for c in prev_cond.ConnectorManager.Connectors
                   if c.Origin.DistanceTo(p_start) < 0.05), None)
        c2 = next((c for c in c_new.ConnectorManager.Connectors
                   if c.Origin.DistanceTo(p_start) < 0.05), None)
        if c1 and c2:
            try:
                c1.ConnectTo(c2)
                len1 = prev_cond.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
                len2 = c_new.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
                l1 = len1.AsDouble() if len1 else 1.0
                l2 = len2.AsDouble() if len2 else 1.0
                if l1 >= 0.15 and l2 >= 0.15:
                    new_elbow = doc.Create.NewElbowFitting(c1, c2)
                    if new_elbow and last_ref_conduit:
                        copy_conduit_parameters(last_ref_conduit, new_elbow)
            except Exception as e:
                dbg.debug("draw_conduit fitting: {}".format(e))
        else:
            dbg.debug("draw_conduit: conectores não encontrados em p_start")
    return c_new


def create_terrain_segments(p_stub1, p_stub2, dist_meters):
    N = max(1, min(30, int(round(dist_meters * 1.5))))
    segments = []
    for i in range(N):
        t0 = i / float(N)
        t1 = (i + 1) / float(N)
        p_start = XYZ(
            p_stub1.X + (p_stub2.X - p_stub1.X) * t0,
            p_stub1.Y + (p_stub2.Y - p_stub1.Y) * t0,
            p_stub1.Z + (p_stub2.Z - p_stub1.Z) * t0,
        )
        p_end = XYZ(
            p_stub1.X + (p_stub2.X - p_stub1.X) * t1,
            p_stub1.Y + (p_stub2.Y - p_stub1.Y) * t1,
            p_stub1.Z + (p_stub2.Z - p_stub1.Z) * t1,
        )
        segments.append((p_start, p_end))
    return segments

# WarningSwallower vem de lf_utils.make_warning_swallower() — importado no topo


def _direct_route_compatible(pt1, pt2, conn1, conn2, tolerance=0.15):
    """
    Retorna True se a rota direta pt1→pt2 é geometricamente compatível
    com os conectores das caixas para criação de elbows.

    Um elbow padrão só funciona quando o eletroduto é paralelo (dot≈1)
    ou perpendicular (dot≈0) ao conector da caixa.
    Ângulos intermediários (~45°, dot≈0.5-0.7) causam o erro
    "conduite modificado para direção oposta".
    """
    if conn1 is None and conn2 is None:
        return True
    try:
        direct = pt2.Subtract(pt1)
        if direct.GetLength() < 1e-6:
            return True
        direct = direct.Normalize()
        for conn in [c for c in [conn1, conn2] if c is not None]:
            try:
                cd = conn.CoordinateSystem.BasisZ
                dot = abs(direct.DotProduct(cd))
                # incompatível se o ângulo é "diagonal" — nem paralelo nem perpendicular
                if tolerance < dot < (1.0 - tolerance):
                    dbg.debug("_direct_route_compatible: dot={:.3f} → incompatível".format(dot))
                    return False
            except Exception:
                pass
    except Exception:
        pass
    return True


def try_direct_with_fittings(doc, conduit_type_id, pt1, pt2, conn1, conn2,
                              level_id, diameter, last_ref_conduit):
    """Tenta conectar com 1 tubo direto pt1→pt2 + curvas automáticas."""
    if pt1.DistanceTo(pt2) < 0.05:
        return []
    orig_len = pt1.DistanceTo(pt2)
    if orig_len < 0.35:
        return []
    if not _direct_route_compatible(pt1, pt2, conn1, conn2):
        dbg.info("Rota direta descartada: ângulo incompatível com conectores.")
        return []
    sub = SubTransaction(doc)
    sub.Start()
    try:
        cond = Conduit.Create(doc, conduit_type_id, pt1, pt2, level_id)
        cond.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
        if last_ref_conduit:
            copy_conduit_parameters(last_ref_conduit, cond)
        conns_cond = list(cond.ConnectorManager.Connectors)
        c_start = min(conns_cond, key=lambda c: c.Origin.DistanceTo(pt1))
        c_end   = min(conns_cond, key=lambda c: c.Origin.DistanceTo(pt2))
        if not conn1.IsConnected:
            doc.Create.NewElbowFitting(c_start, conn1)
        if not conn2.IsConnected:
            doc.Create.NewElbowFitting(c_end, conn2)
        len_param = cond.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if len_param and len_param.HasValue:
            new_len = len_param.AsDouble()
            if new_len > orig_len + 0.05:
                raise Exception(
                    "Tubo invertido pelo fitting (cresceu {:.3f} > {:.3f})".format(
                        new_len, orig_len))
        sub.Commit()
        dbg.result(True, "Rota direta aceita pelo Revit.")
        return [cond]
    except Exception as e:
        dbg.debug("try_direct_with_fittings falhou: {}".format(e))
        try:
            sub.RollBack()
        except Exception:
            pass
        return []


def merge_collinear_segments(segments):
    if len(segments) <= 1:
        return segments
    merged = [segments[0]]
    for (pa, pb) in segments[1:]:
        prev_a, prev_b = merged[-1]
        if prev_b.DistanceTo(pa) > 0.01:
            merged.append((pa, pb))
            continue
        dir_prev = prev_b - prev_a
        dir_curr = pb - pa
        len_prev = dir_prev.GetLength()
        len_curr = dir_curr.GetLength()
        if len_prev < 0.01 or len_curr < 0.01:
            merged.append((pa, pb))
            continue
        if dir_prev.Normalize().DotProduct(dir_curr.Normalize()) > 0.9999:
            merged[-1] = (prev_a, pb)
        else:
            merged.append((pa, pb))
    return merged

# =====================================================================
#  EXECUÇÃO PRINCIPAL
# =====================================================================
def execute_connection():
    global uidoc, doc
    dbg.section("Conectar Eletroduto — Início")
    dbg.timer_start("total")

    # Inicializa contexto Revit aqui (não no módulo) para compatibilidade com
    # todas as versões de pythonnet/pyRevit
    try:
        uidoc, doc = get_revit_context()
    except RuntimeError as e:
        forms.alert(str(e), title="Conectar Eletroduto")
        return

    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        show_settings()
        return

    settings = load_config()
    dbg.dump("settings", settings)

    use_connector_mode = (settings.get('selection_mode', 'conector') == 'conector')
    dbg.info("Modo de seleção: {}".format("conector" if use_connector_mode else "caixa"))

    # ── Seleção ──────────────────────────────────────────────────
    dbg.section("Fase 1: Seleção")
    pt_click1 = pt_click2 = None
    if use_connector_mode:
        el1, pt_click1, el2, pt_click2 = pick_elements_with_points()
    else:
        el1, el2 = pick_elements_automatic()

    if not el1 or not el2:
        dbg.info("Seleção cancelada.")
        return

    same_box = (el1.Id == el2.Id)
    dbg.info("el1.Id={}  el2.Id={}  same_box={}".format(el1.Id, el2.Id, same_box))

    def _is_cabletray(el):
        try:
            if el and hasattr(el, "Category") and el.Category:
                if el.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray):
                    return True
        except Exception:
            pass
        return False

    if _is_cabletray(el1) or _is_cabletray(el2):
        dbg.warn("Elemento é CableTray — operação ignorada para eletroduto.")
        return

    # ── Conectores ────────────────────────────────────────────────
    dbg.section("Fase 2: Conectores")
    if use_connector_mode:
        if same_box:
            conns = get_connectors(el1)
            if len(conns) < 2:
                TaskDialog.Show("Erro", "Caixa precisa ter pelo menos 2 conectores.")
                return
            conn1 = min(conns, key=lambda c: c.Origin.DistanceTo(pt_click1))
            remaining = [c for c in conns if not c.Origin.IsAlmostEqualTo(conn1.Origin)]
            if remaining:
                conn2 = min(remaining, key=lambda c: c.Origin.DistanceTo(pt_click2))
            else:
                TaskDialog.Show("Erro", "Não foi possível identificar um segundo conector diferente.")
                return
        else:
            conn1, conn2 = find_symmetric_connector_pair(el1, pt_click1, el2, pt_click2)
    else:
        if same_box:
            conn1, conn2 = find_best_connector_pair_same_box(el1)
        else:
            conn1, conn2 = find_best_connector_pair(el1, el2)

    if not conn1 or not conn2:
        forms.alert("Não foi possível encontrar conectores para iniciar o traçado.", title="Erro")
        return

    pt1  = conn1.Origin
    pt2  = conn2.Origin
    dir1 = conn1.CoordinateSystem.BasisZ
    dir2 = conn2.CoordinateSystem.BasisZ

    dbg.xyz("pt1", pt1)
    dbg.xyz("pt2", pt2)
    dbg.debug("dir1=({:.3f},{:.3f},{:.3f})  dir2=({:.3f},{:.3f},{:.3f})".format(
        dir1.X, dir1.Y, dir1.Z, dir2.X, dir2.Y, dir2.Z))

    # ── Parâmetros ─────────────────────────────────────────────────
    dbg.section("Fase 3: Parâmetros")
    level_id = el1.LevelId
    if level_id == ElementId.InvalidElementId:
        view = doc.ActiveView
        level_id = (view.GenLevel.Id if hasattr(view, "GenLevel") and view.GenLevel
                    else FilteredElementCollector(doc).OfClass(Level).FirstElementId())
        dbg.warn("Elemento sem LevelId. Usando nível da view: {}".format(level_id))

    last_ref_conduit = get_last_conduit(doc)
    dbg.debug("last_ref_conduit: {}".format(last_ref_conduit.Id if last_ref_conduit else "None"))

    pref_conduit_type_name = settings.get('conduit_type', '')
    conduit_type_id = None
    if pref_conduit_type_name:
        for t in FilteredElementCollector(doc).OfClass(clr.GetClrType(ConduitType)):
            if __get_name__(t) == pref_conduit_type_name:
                conduit_type_id = t.Id
                break
    if not conduit_type_id:
        if pref_conduit_type_name == "(Padrão do Revit)":
            conduit_type_id  = get_default_conduit_type(doc)
            last_ref_conduit = None
        elif last_ref_conduit:
            conduit_type_id = last_ref_conduit.GetTypeId()
        else:
            conduit_type_id = get_default_conduit_type(doc)
    dbg.debug("conduit_type_id: {}".format(conduit_type_id))

    pref_diameter_str = settings.get('default_diameter', '')
    diameter_mm = 0
    try:
        diameter_mm = float(pref_diameter_str.replace("mm", "").strip())
    except Exception:
        pass

    if diameter_mm > 0:
        diameter = diameter_mm / 304.8
        dbg.debug("Diâmetro das configurações: {:.1f} mm".format(diameter_mm))
    else:
        diameter = 0.082021
        if last_ref_conduit:
            try:
                p_diam = last_ref_conduit.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                if p_diam and p_diam.HasValue:
                    diameter = p_diam.AsDouble()
                    dbg.debug("Diâmetro do último eletroduto: {:.4f} ft".format(diameter))
            except Exception:
                pass
    if conn1 and conn1.Shape == ConnectorProfileType.Round and diameter == 0.082021:
        diameter = conn1.Radius * 2
        dbg.debug("Diâmetro do conector: {:.4f} ft".format(diameter))

    # ── Contexto ───────────────────────────────────────────────────
    sym1 = el1.Symbol if hasattr(el1, "Symbol") else None
    sym2 = el2.Symbol if hasattr(el2, "Symbol") else None
    fam1_name  = __get_family_name__(sym1).lower() if sym1 else ""
    fam2_name  = __get_family_name__(sym2).lower() if sym2 else ""
    type1_name = __get_name__(sym1).lower() if sym1 else ""
    type2_name = __get_name__(sym2).lower() if sym2 else ""

    level_elevation = pt1.Z
    try:
        lvl = doc.GetElement(level_id)
        if lvl and hasattr(lvl, "Elevation"):
            level_elevation = lvl.Elevation
    except Exception as e:
        dbg.debug("level_elevation: {}".format(e))

    all_names = fam1_name + " " + fam2_name + " " + type1_name + " " + type2_name
    name_hint_piso   = ("piso" in all_names or "chao" in all_names or u"chão" in all_names)
    dir1_is_vertical = abs(dir1.Z) > 0.7
    dir2_is_vertical = abs(dir2.Z) > 0.7
    is_piso          = (dir1_is_vertical and dir2_is_vertical) or name_hint_piso

    force_vertical = False
    try:
        cat1 = el1.Category.Id.IntegerValue
        cat2 = el2.Category.Id.IntegerValue
        lum  = int(BuiltInCategory.OST_LightingFixtures)
        allowed = (
            int(BuiltInCategory.OST_ElectricalFixtures),
            int(BuiltInCategory.OST_LightingDevices),
            int(BuiltInCategory.OST_CommunicationDevices),
            int(BuiltInCategory.OST_DataDevices),
            int(BuiltInCategory.OST_TelephoneDevices),
            int(BuiltInCategory.OST_LightingFixtures),
        )
        if (cat1 == lum and cat2 in allowed) or (cat2 == lum and cat1 in allowed):
            force_vertical = True
    except Exception as e:
        dbg.debug("force_vertical: {}".format(e))

    dist_direct = pt1.DistanceTo(pt2)
    dz          = abs(pt1.Z - pt2.Z)
    is_flat     = dz < 0.25
    dist_metros = dist_direct * 0.3048
    is_same_family = (fam1_name and fam2_name and fam1_name == fam2_name)

    dbg.debug("dist={:.4f} ft ({:.3f} m)  dz={:.4f} ft  is_flat={}  is_piso={}  same_box={}  force_vertical={}".format(
        dist_direct, dist_metros, dz, is_flat, is_piso, same_box, force_vertical))

    vec_to_target = (pt2 - pt1).Normalize() if dist_direct > 0.01 else XYZ.BasisZ
    if dir1.DotProduct(vec_to_target) < -0.1 and not is_same_family:
        dir1 = dir1.Negate()
    if dir2.DotProduct(vec_to_target) > 0.1 and not is_same_family:
        dir2 = dir2.Negate()

    stub_len = max(0.05, min(0.25, dist_direct * 0.15))
    p_stub1  = pt1 + dir1 * stub_len
    p_stub2  = pt2 + dir2 * stub_len
    stub_ok  = (p_stub1.DistanceTo(pt2) > stub_len * 0.5 and
                p_stub2.DistanceTo(pt1) > stub_len * 0.5)
    if not stub_ok:
        p_stub1  = pt1 + dir1 * 0.05
        p_stub2  = pt2 + dir2 * 0.05
        stub_len = 0.05
        dbg.warn("Stubs inválidos — reduzido para 0.05 ft.")

    # ── Determinação da rota ───────────────────────────────────────
    dbg.section("Fase 4: Rota")
    segments = []

    if dist_direct < 0.5:
        dbg.info("Regra 0a: direto (<0.5 ft)")
        segments = [(pt1, pt2)]
    elif dist_direct < 1.0:
        facing  = dir1.DotProduct(dir2) < -0.7
        aligned = dir1.DotProduct((pt2 - pt1).Normalize()) > 0.6
        if facing or aligned:
            dbg.info("Regra 0b: direto (curto, alinhado/encarando)")
            segments = [(pt1, pt2)]
        else:
            dbg.info("Regra 0b: curto com ângulo — 90°")
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, False)
            segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    elif dir1.DotProduct((pt2 - pt1).Normalize()) > 0.85 and dir2.DotProduct((pt1 - pt2).Normalize()) > 0.85:
        dbg.info("Regra 0c: direto (conectores encarando em linha)")
        segments = [(pt1, pt2)]
    elif same_box and is_piso and dist_metros < 3.0:
        dbg.info("Regra 1: mesma caixa, piso, curto")
        segments = [(pt1, p_stub1), (p_stub1, p_stub2), (p_stub2, pt2)]
    elif same_box and not is_piso and is_flat:
        dbg.info("Regra 2: mesma caixa, parede/teto, mesmo nível")
        p1_c, p2_c = solve_chicane_2d(pt1, pt2, dir1, dir2, stub_len)
        if p1_c and p2_c:
            segments = [(pt1, p1_c), (p1_c, p2_c), (p2_c, pt2)]
        else:
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, False)
            segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    elif same_box and not is_flat:
        if is_piso:
            dbg.info("Regra 3a: mesma caixa, piso, desnível")
            mid_segs = create_terrain_segments(p_stub1, p_stub2, dist_metros)
        else:
            dbg.info("Regra 3b: mesma caixa, parede, desnível")
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, True)
        segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    elif not same_box and is_flat:
        if is_piso:
            dbg.info("Regra 4a: caixas diferentes, piso, mesmo nível")
            p1_c, p2_c = solve_chicane_2d(pt1, pt2, dir1, dir2, stub_len)
            if p1_c and p2_c:
                segments = [(pt1, p1_c), (p1_c, p2_c), (p2_c, pt2)]
            else:
                mid_segs = create_45_degree_path(p_stub1, p_stub2, dir1, dir2)
                segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
        else:
            dbg.info("Regra 4b: caixas diferentes, parede, mesmo nível")
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, False)
            segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    else:
        if force_vertical:
            dbg.info("Regra 5a: caixas diferentes, desnível, luminária")
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, True)
        elif is_piso:
            dbg.info("Regra 5b: caixas diferentes, piso, desnível")
            mid_segs = create_terrain_segments(p_stub1, p_stub2, dist_metros)
        else:
            dbg.info("Regra 5c: caixas diferentes, parede, desnível")
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, True)
        segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]

    segments = merge_collinear_segments(segments)
    dbg.debug("Segmentos após merge: {}".format(len(segments)))
    for idx, (pa, pb) in enumerate(segments):
        dbg.debug("  seg[{}]: {:.4f} ft".format(idx, pa.DistanceTo(pb)))

    # ── Execução da rota ───────────────────────────────────────────
    dbg.section("Fase 5: Criação dos Eletrodutos")
    t = Transaction(doc, "Conectar Eletrodutos Inteligente")
    ops = t.GetFailureHandlingOptions()
    ops.SetFailuresPreprocessor(make_warning_swallower())
    t.SetFailureHandlingOptions(ops)
    t.Start()

    try:
        created_conds = []
        routing_strategy = settings.get('routing_strategy', 'auto')
        is_multi_segment = len(segments) > 1

        if routing_strategy == 'auto' and is_multi_segment:
            dbg.info("Tentando rota direta...")
            dbg.timer_start("try_direct")
            created_conds = try_direct_with_fittings(
                doc, conduit_type_id, pt1, pt2,
                conn1, conn2, level_id, diameter, last_ref_conduit
            )
            dbg.timer_end("try_direct")

        if not created_conds:
            dbg.info("Usando rota calculada ({} segmentos).".format(len(segments)))
            dbg.timer_start("draw_segments")
            last_cond = None
            for idx, (p_a, p_b) in enumerate(segments):
                if p_a.DistanceTo(p_b) < 0.05:
                    dbg.warn("  seg[{}] muito curto — pulado.".format(idx))
                    continue
                c_new = draw_conduit_and_connect(
                    doc, conduit_type_id, p_a, p_b,
                    level_id, diameter, last_cond, last_ref_conduit
                )
                if c_new:
                    created_conds.append(c_new)
                    last_cond = c_new
                    dbg.result(True, "  seg[{}] Id={} len={:.4f} ft".format(
                        idx, c_new.Id, p_a.DistanceTo(p_b)))
            dbg.timer_end("draw_segments")

        # Conectar pontas com as caixas
        if created_conds:
            _connect_endpoint(doc, created_conds[0],  pt1, conn1, "ponta-início")
            _connect_endpoint(doc, created_conds[-1], pt2, conn2, "ponta-fim")

        # Aplicar Tipo de Serviço
        service_text = settings.get('service_type', '')
        if service_text:
            for c_elem in created_conds:
                try:
                    for p_name in ["Tipo de serviço", "Service Type", "Comentários", "Comments"]:
                        p = c_elem.LookupParameter(p_name)
                        if p and not p.IsReadOnly:
                            p.Set(service_text)
                except Exception:
                    pass

        t.Commit()
        dbg.section("Resultado")
        dbg.info("Eletrodutos criados: {}".format(len(created_conds)))
        dbg.timer_end("total")

    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        dbg.error("Exceção na transação: {}".format(e))
        raise e

def safe_execution():
    dbg.section("Conectar Eletroduto — BOOT")
    dbg.info("DEBUG_MODE = {}".format(DEBUG_MODE))
    try:
        execute_connection()
        dbg.section("Ferramenta Finalizada")
    except Exception as e:
        err_tb = traceback.format_exc()
        dbg.error("CRASH FATAL:\n{}".format(err_tb))
        if DEBUG_MODE:
            forms.alert("CRASH FATAL (DEBUG):\n\n" + err_tb, title="Conectar Eletroduto — Erro")
        else:
            forms.alert("Erro ao conectar eletrodutos:\n" + str(e), title="Aviso")

if __name__ == "__main__":
    safe_execution()
