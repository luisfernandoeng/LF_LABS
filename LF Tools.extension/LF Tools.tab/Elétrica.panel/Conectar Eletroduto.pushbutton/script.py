# -*- coding: utf-8 -*-
"""
Conectar Eletroduto - Conexão inteligente entre elementos MEP.

Shift-Click: Abre menu de configuração (Conector específico vs Caixa inteira)
Normal: Seleção automática (conector mais próximo)
"""

__title__ = 'Conectar\nEletroduto'
__author__ = 'Luis Fernando'

# =============================================================================
#  IMPORTS
# =============================================================================
import clr
import os
import re
import sys
import System
import json
import traceback
import math
from collections import OrderedDict

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.IO import StreamReader
from System.Windows.Markup import XamlReader
from System.Windows import Window, Thickness
from System.Windows.Interop import WindowInteropHelper
from System.Collections.Generic import List

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms
from lf_utils import DebugLogger, get_script_config, save_script_config, make_warning_swallower

# Instância global
dbg = DebugLogger(False)

# Referências globais nativas
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document if uidoc else None


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
        'selection_mode':        'conector',
        'multi_select':          False,
        'angle_mode_plan':       '90',
        'angle_mode_vertical':   '90',
        'conduit_type_plan':     '',
        'conduit_type_vertical': '',
        'default_diameter':      '',
        'service_type':          '',
        'routing_strategy':      'auto',
        'debug_mode':            False,
    })

def save_config(settings):
    save_script_config(__commandpath__, settings)

class SettingsWindow(object):
    def __init__(self, settings):
        self.settings = settings
        self._setup_window()
        self._load_conduit_types()
        self._load_current_values()

    def _setup_window(self):
        xaml_path = os.path.join(__commandpath__, 'settings.xaml')
        stream = StreamReader(xaml_path)
        self.win = XamlReader.Load(stream.BaseStream)
        stream.Close()

        WindowInteropHelper(self.win).Owner = __revit__.MainWindowHandle

        # Aba Plano / Piso
        self.rb_plan_90    = self.win.FindName("rb_plan_90")
        self.rb_plan_45    = self.win.FindName("rb_plan_45")
        self.rb_plan_livre = self.win.FindName("rb_plan_livre")
        self.cb_type_plan  = self.win.FindName("cb_type_plan")

        # Aba Vertical
        self.rb_vert_90    = self.win.FindName("rb_vert_90")
        self.rb_vert_45    = self.win.FindName("rb_vert_45")
        self.rb_vert_livre = self.win.FindName("rb_vert_livre")
        self.cb_type_vert  = self.win.FindName("cb_type_vert")

        # Aba Geral
        self.rb_qtd_2        = self.win.FindName("rb_qtd_2")
        self.rb_qtd_multi    = self.win.FindName("rb_qtd_multi")
        self.rb_sel_conector = self.win.FindName("rb_sel_conector")
        self.rb_sel_caixa    = self.win.FindName("rb_sel_caixa")
        self.tb_diameter     = self.win.FindName("tb_diameter")
        self.tb_service      = self.win.FindName("tb_service")
        self.rb_strat_auto   = self.win.FindName("rb_strat_auto")
        self.rb_strat_calc   = self.win.FindName("rb_strat_calc")
        self.chk_debug       = self.win.FindName("chk_debug")

        # Botões
        self.win.FindName("btn_save").Click   += self.save_click
        self.win.FindName("btn_cancel").Click += self.cancel_click

    def _load_conduit_types(self):
        collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(ConduitType))
        type_names = sorted([__get_name__(t) for t in collector if __get_name__(t)])
        self._conduit_options = ["(Usar Último Desenhado)", "(Padrão do Revit)"] + type_names
        self.cb_type_plan.ItemsSource = self._conduit_options
        self.cb_type_vert.ItemsSource = self._conduit_options

    def _load_current_values(self):
        # Ângulo plano
        angle_plan = self.settings.get('angle_mode_plan', '90')
        self.rb_plan_90.IsChecked    = (angle_plan == '90')
        self.rb_plan_45.IsChecked    = (angle_plan == '45')
        self.rb_plan_livre.IsChecked = (angle_plan == 'livre')

        # Tipo eletroduto plano
        pref_plan = self.settings.get('conduit_type_plan', '')
        if pref_plan in self._conduit_options:
            self.cb_type_plan.SelectedItem = pref_plan
        else:
            self.cb_type_plan.SelectedIndex = 0

        # Ângulo vertical
        angle_vert = self.settings.get('angle_mode_vertical', '90')
        self.rb_vert_90.IsChecked    = (angle_vert == '90')
        self.rb_vert_45.IsChecked    = (angle_vert == '45')
        self.rb_vert_livre.IsChecked = (angle_vert == 'livre')

        # Tipo eletroduto vertical
        pref_vert = self.settings.get('conduit_type_vertical', '')
        if pref_vert in self._conduit_options:
            self.cb_type_vert.SelectedItem = pref_vert
        else:
            self.cb_type_vert.SelectedIndex = 0

        multi = self.settings.get('multi_select', False)
        if self.rb_qtd_multi:
            self.rb_qtd_multi.IsChecked = multi
            if self.rb_qtd_2: self.rb_qtd_2.IsChecked = not multi

        # Modo seleção
        mode = self.settings.get('selection_mode', 'conector')
        self.rb_sel_conector.IsChecked = (mode == 'conector')
        self.rb_sel_caixa.IsChecked    = (mode == 'caixa')

        # Diâmetro e serviço
        self.tb_diameter.Text = self.settings.get('default_diameter', '')
        self.tb_service.Text  = self.settings.get('service_type', '')

        # Estratégia
        strat = self.settings.get('routing_strategy', 'auto')
        self.rb_strat_auto.IsChecked = (strat == 'auto')
        self.rb_strat_calc.IsChecked = (strat != 'auto')

        # Debug
        self.chk_debug.IsChecked = bool(self.settings.get('debug_mode', False))

    def save_click(self, sender, e):
        if self.rb_plan_45.IsChecked:
            self.settings['angle_mode_plan'] = '45'
        elif self.rb_plan_livre.IsChecked:
            self.settings['angle_mode_plan'] = 'livre'
        else:
            self.settings['angle_mode_plan'] = '90'

        if self.rb_vert_45.IsChecked:
            self.settings['angle_mode_vertical'] = '45'
        elif self.rb_vert_livre.IsChecked:
            self.settings['angle_mode_vertical'] = 'livre'
        else:
            self.settings['angle_mode_vertical'] = '90'

        sel_plan = self.cb_type_plan.SelectedItem
        self.settings['conduit_type_plan']     = sel_plan if sel_plan else ''
        sel_vert = self.cb_type_vert.SelectedItem
        self.settings['conduit_type_vertical'] = sel_vert if sel_vert else ''

        if self.rb_qtd_multi:
            self.settings['multi_select'] = bool(self.rb_qtd_multi.IsChecked)
        self.settings['selection_mode']   = 'caixa' if self.rb_sel_caixa.IsChecked else 'conector'
        self.settings['default_diameter'] = self.tb_diameter.Text or ''
        self.settings['service_type']     = self.tb_service.Text or ''
        self.settings['routing_strategy'] = 'calculado' if self.rb_strat_calc.IsChecked else 'auto'
        self.settings['debug_mode']       = bool(self.chk_debug.IsChecked)

        save_config(self.settings)
        self.win.DialogResult = True
        self.win.Close()

    def cancel_click(self, sender, e):
        self.win.Close()

    def show(self):
        return self.win.ShowDialog()

def show_settings():
    settings = load_config()
    xaml_path = os.path.join(__commandpath__, "settings.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("Arquivo settings.xaml não encontrado!", title="Erro")
        return
    win = SettingsWindow(settings)
    win.show()

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

    # ── FALLBACKS DE CONEXÃO ──
    # O Revit frequentemente falha ao criar conectores por falta de espaço (raio muito longo).
    # Vamos tentar uma cascata de 5 métodos diferentes para garantir a conexão.
    
    dist_conn = c_near.Origin.DistanceTo(target_conn.Origin)

    # Método 1: Estão grudados (distância < 0.01) - Conexão Lógica Direta
    if dist_conn < 0.01:
        try:
            c_near.ConnectTo(target_conn)
            dbg.result(True, "{}: ConnectTo direto OK (mesma coordenada)".format(label))
            return
        except Exception:
            pass

    # Método 2: NewElbowFitting (O padrão correto para curvas 90/45)
    try:
        doc.Create.NewElbowFitting(c_near, target_conn)
        dbg.result(True, "{}: Elbow criado com sucesso".format(label))
        return
    except Exception as e:
        dbg.debug("{}: Falha no Elbow: {}".format(label, e))

    # Método 3: NewUnionFitting (Usado quando o ângulo é quase 0°/180° e o Elbow falha)
    try:
        doc.Create.NewUnionFitting(c_near, target_conn)
        dbg.result(True, "{}: Union criado com sucesso".format(label))
        return
    except Exception as e:
        dbg.debug("{}: Falha no Union: {}".format(label, e))

    # Método 4: NewTransitionFitting (Caso haja diferença sutil de diâmetro que impossibilita Elbow/Union)
    try:
        doc.Create.NewTransitionFitting(c_near, target_conn)
        dbg.result(True, "{}: Transition criado com sucesso".format(label))
        return
    except Exception as e:
        dbg.debug("{}: Falha no Transition: {}".format(label, e))

    # Método 5: ConnectTo forçado (Conexão Lógica Pura)
    # Garante o circuito elétrico mesmo se a geometria de curva não couber no espaço.
    try:
        c_near.ConnectTo(target_conn)
        dbg.result(True, "{}: ConnectTo forçado (Conexão Lógica) OK".format(label))
        return
    except Exception as e:
        dbg.debug("{}: Falha Fatal, nenhuma conexão possível: {}".format(label, e))


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

# =====================================================================
#  ELETROCALHA / PERFILADO — FAMÍLIA DE UNIÃO
# =====================================================================
FAM_ELETROCALHA = "OFEletrico_Eletrocalha_Uniao_SaidaHorizontal"
FAM_PERFILADO   = "OFEletrico_Perfilado_Uniao_SaidaLateral"


def _is_perfilado(cable_tray):
    """Detecta perfilado por nome do tipo ou por seção quadrada pequena (≤ 60 mm)."""
    # 1. Por nome do tipo ou família — mais confiável
    try:
        ct_type = cable_tray.Document.GetElement(cable_tray.GetTypeId())
        if ct_type:
            names_to_check = []
            try:
                names_to_check.append((ct_type.Name or u"").lower())
            except Exception:
                pass
            try:
                p = ct_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if p and p.HasValue:
                    names_to_check.append((p.AsString() or u"").lower())
            except Exception:
                pass
            for n in names_to_check:
                if u"perfilado" in n or u"strut" in n or u"ladder" in n:
                    return True
    except Exception:
        pass
    # 2. Por dimensão: seção quadrada e pequena (≤ 60 mm, tolerância 10 mm)
    try:
        w_p = cable_tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        h_p = cable_tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        if w_p and h_p:
            w_mm = w_p.AsDouble() * 304.8
            h_mm = h_p.AsDouble() * 304.8
            if abs(w_mm - h_mm) < 10.0 and max(w_mm, h_mm) <= 60.0:
                return True
    except Exception:
        pass
    return False


def _find_symbol(doc, family_name):
    for fs in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if fs.Family and fs.Family.Name == family_name:
                return fs
        except Exception:
            pass
    return None


def _get_round_connector(inst):
    try:
        for c in inst.MEPModel.ConnectorManager.Connectors:
            if (c.Domain == Domain.DomainCableTrayConduit
                    and c.Shape == ConnectorProfileType.Round):
                return c
    except Exception:
        pass
    return None


def _get_ct_connectors(inst):
    """Retorna conectores retangulares (eletrocalha) do fitting."""
    result = []
    try:
        for c in inst.MEPModel.ConnectorManager.Connectors:
            if (c.Domain == Domain.DomainCableTrayConduit
                    and c.Shape != ConnectorProfileType.Round):
                result.append(c)
    except Exception:
        pass
    return result


def _split_cabletray(doc, cable_tray, split_pt, fallback_level_id=None):
    """
    Divide a eletrocalha em split_pt usando SubTransaction.
    Se CableTray.Create falhar, o rollback preserva a eletrocalha original.
    Retorna (conn_near_ct1, conn_near_ct2) para conectar ao fitting,
    ou (None, None) se o split não foi possível.
    """
    from Autodesk.Revit.DB import SubTransaction, StorageType

    sub = SubTransaction(doc)
    sub.Start()
    try:
        crv = cable_tray.Location.Curve
        p0  = crv.GetEndPoint(0)
        p1  = crv.GetEndPoint(1)

        proj = crv.Project(split_pt)
        s_pt = proj.XYZPoint if proj else XYZ(split_pt.X, split_pt.Y, p0.Z)

        ct_type_id = cable_tray.GetTypeId()

        # MEP curves usam RBS_START_LEVEL_PARAM; FAMILY_LEVEL_PARAM é fallback
        level_id = fallback_level_id
        for bip in [BuiltInParameter.RBS_START_LEVEL_PARAM,
                    BuiltInParameter.FAMILY_LEVEL_PARAM]:
            try:
                lp = cable_tray.get_Parameter(bip)
                if lp and lp.AsElementId() != ElementId.InvalidElementId:
                    level_id = lp.AsElementId()
                    break
            except Exception:
                continue

        # Salva todos os parâmetros não-readonly (inclui parâmetros de filtro)
        params_to_copy = []
        for p in cable_tray.Parameters:
            try:
                if p.IsReadOnly:
                    continue
                st = p.StorageType
                if st == StorageType.Double:
                    params_to_copy.append((p.Id, st, p.AsDouble()))
                elif st == StorageType.Integer:
                    params_to_copy.append((p.Id, st, p.AsInteger()))
                elif st == StorageType.String:
                    s = p.AsString()
                    if s is not None:
                        params_to_copy.append((p.Id, st, s))
                elif st == StorageType.ElementId:
                    params_to_copy.append((p.Id, st, p.AsElementId()))
            except Exception:
                continue

        # Salva vizinhos conectados em p0 e p1 para reconectar após o split
        p0_neighbors = []
        p1_neighbors = []
        try:
            for c in cable_tray.ConnectorManager.Connectors:
                near_p0 = c.Origin.DistanceTo(p0) < c.Origin.DistanceTo(p1)
                try:
                    for ref in c.AllRefs:
                        try:
                            if ref.Owner.Id == cable_tray.Id:
                                continue
                            info = (ref.Owner.Id, ref.Origin)
                            if near_p0:
                                p0_neighbors.append(info)
                            else:
                                p1_neighbors.append(info)
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

        doc.Delete(cable_tray.Id)
        doc.Regenerate()

        def _make_ct(pa, pb):
            if pa.DistanceTo(pb) < 0.1:
                return None
            ct = CableTray.Create(doc, ct_type_id, pa, pb, level_id)
            if ct:
                for (pid, st, val) in params_to_copy:
                    try:
                        p = ct.get_Parameter(pid)
                        if p is None or p.IsReadOnly:
                            continue
                        if st == StorageType.Double:
                            p.Set(val)
                        elif st == StorageType.Integer:
                            p.Set(val)
                        elif st == StorageType.String:
                            p.Set(val)
                        elif st == StorageType.ElementId:
                            p.Set(val)
                    except Exception:
                        pass
            return ct

        ct1 = _make_ct(p0, s_pt)
        ct2 = _make_ct(s_pt, p1)

        if ct1 is None and ct2 is None:
            dbg.warn("_split_cabletray: CableTray.Create retornou None — rollback")
            sub.RollBack()
            return None, None

        doc.Regenerate()

        def _conn_near(ct, pt):
            if ct is None:
                return None
            best, bd = None, float('inf')
            try:
                for c in ct.ConnectorManager.Connectors:
                    d = c.Origin.DistanceTo(pt)
                    if d < bd:
                        bd, best = d, c
            except Exception:
                pass
            return best

        # Reconecta os endpoints externos ao network original
        def _reconnect_endpoint(ct_seg, neighbor_infos, endpoint_pt):
            if ct_seg is None or not neighbor_infos:
                return
            seg_conn = _conn_near(ct_seg, endpoint_pt)
            if seg_conn is None:
                return
            for (el_id, ref_origin) in neighbor_infos:
                try:
                    el = doc.GetElement(el_id)
                    if el is None:
                        continue
                    cm_neighbor = None
                    try:
                        cm_neighbor = el.ConnectorManager
                    except Exception:
                        try:
                            cm_neighbor = el.MEPModel.ConnectorManager
                        except Exception:
                            pass
                    if cm_neighbor is None:
                        continue
                    best_nc, best_d = None, float('inf')
                    for nc in cm_neighbor.Connectors:
                        d = nc.Origin.DistanceTo(ref_origin)
                        if d < best_d:
                            best_d, best_nc = d, nc
                    if best_nc and best_d < 0.5:
                        try:
                            seg_conn.ConnectTo(best_nc)
                        except Exception:
                            pass
                except Exception:
                    pass

        _reconnect_endpoint(ct1, p0_neighbors, p0)
        _reconnect_endpoint(ct2, p1_neighbors, p1)

        sub.Commit()
        dbg.debug("_split_cabletray: ok")
        return _conn_near(ct1, s_pt), _conn_near(ct2, s_pt)

    except Exception as e:
        dbg.warn("_split_cabletray falhou ({}), eletrocalha preservada".format(e))
        try:
            sub.RollBack()
        except Exception:
            pass
        return None, None


def _place_union_on_cabletray(doc, cable_tray, click_pt, conn_dest, level_id):
    """
    Insere a família de união na eletrocalha, alinhada com conn_dest.

    1. Projeta conn_dest.Origin na curva da eletrocalha — posição onde o
       eletroduto sairá reto diretamente para a caixa.
    2. Rotaciona o conector de conduit para ser anti-paralelo a conn_dest.BasisZ
       (fitting aponta de volta para a caixa, alinhado com seu conector).
    3. Divide a eletrocalha e conecta os conectores CT do fitting a cada metade.

    Retorna (instância, conector_round) ou (None, None).
    """
    is_perf  = _is_perfilado(cable_tray)
    fam_name = FAM_PERFILADO if is_perf else FAM_ELETROCALHA
    dbg.info(u"União: {} ({})".format(fam_name, "perfilado" if is_perf else "eletrocalha"))

    sym = _find_symbol(doc, fam_name)
    if not sym:
        dbg.warn(u"Família não encontrada: {}".format(fam_name))
        return None, None

    if not sym.IsActive:
        sym.Activate()
        doc.Regenerate()

    # ── 1. Posição: projetar origem do conector destino na curva da eletrocalha ──
    place_pt = click_pt
    try:
        crv     = cable_tray.Location.Curve
        tray_z  = (crv.GetEndPoint(0).Z + crv.GetEndPoint(1).Z) / 2.0
        proj    = crv.Project(conn_dest.Origin)
        proj_xy = proj.XYZPoint if proj else None
        if proj_xy:
            place_pt = XYZ(proj_xy.X, proj_xy.Y, tray_z)
        else:
            place_pt = XYZ(click_pt.X, click_pt.Y, tray_z)
    except Exception as e:
        dbg.debug("place_union projection: {}".format(e))
        try:
            crv    = cable_tray.Location.Curve
            tray_z = (crv.GetEndPoint(0).Z + crv.GetEndPoint(1).Z) / 2.0
            place_pt = XYZ(click_pt.X, click_pt.Y, tray_z)
        except Exception:
            pass

    level = doc.GetElement(level_id)
    inst  = doc.Create.NewFamilyInstance(place_pt, sym, level, StructuralType.NonStructural)
    doc.Regenerate()

    try:
        # Corrigir Z caso o NewFamilyInstance ignore a coordenada Z
        curr_z = inst.Location.Point.Z
        if abs(curr_z - tray_z) > 0.001:
            ElementTransformUtils.MoveElement(doc, inst.Id, XYZ(0, 0, tray_z - curr_z))
            
        # Igualar tamanho da eletrocalha no fitting (isso afeta a posição do conector redondo)
        w_p = cable_tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        h_p = cable_tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        if w_p and h_p:
            p_comp = inst.LookupParameter("Comprimento")
            if p_comp and not p_comp.IsReadOnly:
                p_comp.Set(w_p.AsDouble())
            p_alt = inst.LookupParameter("Altura")
            if p_alt and not p_alt.IsReadOnly:
                p_alt.Set(h_p.AsDouble())
    except Exception as e:
        dbg.debug("ajuste de tamanho/altura da união falhou: {}".format(e))
        
    doc.Regenerate()

    conduit_conn = _get_round_connector(inst)

    # ── 2. Rotação: conduit connector anti-paralelo a conn_dest.BasisZ ──
    # conn_dest.BasisZ aponta da caixa para fora.
    # Queremos que o conector de conduit do fitting aponte DE VOLTA para a caixa
    # (anti-paralelo), para que o eletroduto saia alinhado com o conector da caixa.
    try:
        dir_conn     = conn_dest.CoordinateSystem.BasisZ
        target_angle = math.atan2(-dir_conn.Y, -dir_conn.X)

        if conduit_conn:
            bz         = conduit_conn.CoordinateSystem.BasisZ
            curr_angle = math.atan2(bz.Y, bz.X)
        else:
            v          = conn_dest.Origin - place_pt
            curr_angle = math.atan2(v.Y, v.X) if (abs(v.X) > 0.01 or abs(v.Y) > 0.01) else 0.0

        rot = target_angle - curr_angle
        while rot >  math.pi: rot -= 2 * math.pi
        while rot < -math.pi: rot += 2 * math.pi

        if abs(rot) > 0.001:
            axis = Line.CreateBound(place_pt, XYZ(place_pt.X, place_pt.Y, place_pt.Z + 1.0))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, rot)
            doc.Regenerate()
            conduit_conn = _get_round_connector(inst)
    except Exception as e:
        dbg.debug("_place_union rot: {}".format(e))

    # ── 3. Conectar à eletrocalha: dividir e ligar conectores CT ──
    try:
        ct_conns = _get_ct_connectors(inst)
        if ct_conns:
            c1_tray, c2_tray = _split_cabletray(doc, cable_tray, place_pt, level_id)
            remaining = list(ct_conns)
            for tray_conn in [c1_tray, c2_tray]:
                if tray_conn is None or not remaining:
                    continue
                # Pega o conector CT do fitting mais oposto ao conector da eletrocalha
                best = max(remaining,
                           key=lambda fc: fc.CoordinateSystem.BasisZ.DotProduct(
                               tray_conn.CoordinateSystem.BasisZ.Negate()))
                try:
                    best.ConnectTo(tray_conn)
                    remaining.remove(best)
                    dbg.debug("CT connector ligado: ok")
                except Exception as e:
                    dbg.debug("CT ConnectTo: {}".format(e))
    except Exception as e:
        dbg.debug("_place_union split/connect: {}".format(e))

    return inst, conduit_conn


def _execute_cabletray_connection(doc, settings, cable_tray_el, cable_tray_click,
                                   other_el, other_click, use_connector_mode,
                                   conduit_type_id, diameter, level_id, last_ref_conduit):
    """Conecta eletrocalha/perfilado → elemento elétrico via família de união."""
    dbg.section("Eletrocalha — Conexão")

    # Conector do elemento destino — pega o que "olha" para a eletrocalha
    other_conns = get_connectors(other_el)
    if not other_conns:
        forms.alert(u"Conector de eletroduto não encontrado no elemento de destino.",
                    title="Conectar Eletroduto")
        return

    def _facing_score(c, target_pt):
        """Dot product entre direção do conector e vetor para target_pt.
        Positivo = conector aponta para o alvo."""
        try:
            v = target_pt - c.Origin
            if v.GetLength() < 0.01:
                return 0.0
            return c.CoordinateSystem.BasisZ.DotProduct(v.Normalize())
        except Exception:
            return -1.0

    ref_pt = cable_tray_click if cable_tray_click else (other_click or XYZ.Zero)
    conn_other = max(other_conns, key=lambda c: _facing_score(c, ref_pt))

    pt_dest  = conn_other.Origin
    dir_dest = conn_other.CoordinateSystem.BasisZ

    dbg.xyz("pt_dest", pt_dest)

    t = Transaction(doc, u"Conectar Eletroduto — Eletrocalha")
    ops = t.GetFailureHandlingOptions()
    preprocessor = make_warning_swallower()
    if preprocessor:
        ops.SetFailuresPreprocessor(preprocessor)
    t.SetFailureHandlingOptions(ops)
    t.Start()

    try:
        # Passa conn_other para que _place_union alinhe o fitting com a direção do conector
        union_inst, union_conn = _place_union_on_cabletray(
            doc, cable_tray_el, cable_tray_click, conn_other, level_id
        )
        if not union_inst:
            t.RollBack()
            fam = FAM_PERFILADO if _is_perfilado(cable_tray_el) else FAM_ELETROCALHA
            forms.alert(
                u"Família não encontrada no projeto:\n{}\n\nVerifique se está carregada no template.".format(fam),
                title="Conectar Eletroduto"
            )
            return

        # Ponto e direção de partida = conector redondo da união
        if union_conn:
            pt_start  = union_conn.Origin
            dir_start = union_conn.CoordinateSystem.BasisZ
            vec = pt_dest - pt_start
            if vec.GetLength() > 0.01 and dir_start.DotProduct(vec.Normalize()) < -0.1:
                dir_start = dir_start.Negate()
        else:
            pt_start  = XYZ(cable_tray_click.X, cable_tray_click.Y,
                            union_inst.Location.Point.Z)
            v = pt_dest - pt_start
            dir_start = v.Normalize() if v.GetLength() > 0.01 else XYZ.BasisX

        # Corrigir dir_dest com pt_start real (mais preciso que cable_tray_click)
        vec_to_union = pt_start - pt_dest
        if vec_to_union.GetLength() > 0.01 and dir_dest.DotProduct(vec_to_union.Normalize()) < -0.1:
            dir_dest = dir_dest.Negate()

        dbg.xyz("pt_start (union)", pt_start)

        dist = pt_start.DistanceTo(pt_dest)
        stub_len = max(0.5, min(1.0, dist * 0.20))
        p_stub_s = pt_start + dir_start * stub_len
        p_stub_d = pt_dest  + dir_dest  * stub_len
        dz       = abs(pt_start.Z - pt_dest.Z)
        is_flat  = dz < 0.25
        angle    = settings.get('angle_mode_plan' if is_flat else 'angle_mode_vertical', '90')

        def _build_segments(ang):
            if dist < 0.5:
                return [(pt_start, pt_dest)]
            elif is_flat and ang == '45':
                mid = create_45_degree_path(p_stub_s, p_stub_d, dir_start, dir_dest)
                return [(pt_start, p_stub_s)] + mid + [(p_stub_d, pt_dest)]
            elif is_flat:
                mid = create_90_degree_path(p_stub_s, p_stub_d, dir_start, dir_dest, False)
                return [(pt_start, p_stub_s)] + mid + [(p_stub_d, pt_dest)]
            elif ang == '45':
                mid = create_45_degree_path(p_stub_s, p_stub_d, dir_start, dir_dest)
                return [(pt_start, p_stub_s)] + mid + [(p_stub_d, pt_dest)]
            else:
                mid = create_90_degree_path(p_stub_s, p_stub_d, dir_start, dir_dest, True)
                return [(pt_start, p_stub_s)] + mid + [(p_stub_d, pt_dest)]

        # Estratégias em cascata: ângulo configurado → alternativo → direto
        alt_angle  = '45' if angle == '90' else '90'
        strategies = [
            ('config',   _build_segments(angle)),
            ('alt_angle',_build_segments(alt_angle)),
            ('direto',   [(pt_start, pt_dest)]),
        ]

        def _draw_segments(segs):
            conds, last = [], None
            for pa, pb in merge_collinear_segments(segs):
                if pa.DistanceTo(pb) < 0.05:
                    continue
                c = draw_conduit_and_connect(
                    doc, conduit_type_id, pa, pb, level_id, diameter, last, last_ref_conduit
                )
                if c:
                    conds.append(c)
                    last = c
            return conds

        created_conds = []
        for strat_name, segs in strategies:
            dbg.debug("Tentando estratégia: {}  ({} seg)".format(strat_name, len(segs)))
            sub = SubTransaction(doc)
            sub.Start()
            conds = _draw_segments(segs)
            if conds:
                sub.Commit()
                created_conds = conds
                dbg.info("Estratégia '{}' OK ({} eletrodutos)".format(strat_name, len(conds)))
                break
            else:
                sub.RollBack()
                dbg.debug("Estratégia '{}' falhou, tentando próxima.".format(strat_name))

        # Último recurso: eletroduto bruto sem fittings
        if not created_conds and dist > 0.05:
            dbg.warn("Todas as estratégias falharam — criando eletroduto bruto.")
            try:
                raw = Conduit.Create(doc, conduit_type_id, pt_start, pt_dest, level_id)
                if raw:
                    try:
                        raw.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter)
                    except Exception:
                        pass
                    created_conds = [raw]
            except Exception as ex:
                dbg.warn("Eletroduto bruto também falhou: {}".format(ex))

        if created_conds:
            if union_conn:
                _connect_endpoint(doc, created_conds[0], pt_start, union_conn, "ponta-union")
            _connect_endpoint(doc, created_conds[-1], pt_dest, conn_other, "ponta-destino")

        t.Commit()
        dbg.info(u"Eletrodutos criados: {}".format(len(created_conds)))

    except Exception as e:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        dbg.error(u"_execute_cabletray_connection: {}".format(e))
        if dbg.enabled:
            forms.alert(u"Erro ao conectar eletrocalha:\n" + str(e), title="Conectar Eletroduto")


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
    global uidoc, doc, dbg
    
    # ── Configurações e Debug ──────────────────────────────────────
    settings = load_config()
    debug_active = settings.get('debug_mode', False)
    
    # Reinicia o logger com a configuração atual
    dbg = DebugLogger(debug_active)
    
    dbg.section("Conectar Eletroduto — Início")
    dbg.timer_start("total")
    dbg.dump("settings", settings)

    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        show_settings()
        return

    use_connector_mode = (settings.get('selection_mode', 'conector') == 'conector')
    dbg.info("Modo de seleção: {}".format("conector" if use_connector_mode else "caixa"))


    # ── Fase 1: Seleção em Lote ou Par ────────────────────────────
    multi_select = settings.get('multi_select', False)
    picked_elements = []
    points_list = []
    
    if multi_select:
        use_connector_mode = False  # Força seleção automática
        selected_ids = list(uidoc.Selection.GetElementIds())
        if selected_ids:
            for eid in selected_ids:
                el = doc.GetElement(eid)
                if el: picked_elements.append(el)
        else:
            try:
                refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecione os elementos a conectar em lote")
                for ref in refs:
                    el = doc.GetElement(ref)
                    if el: picked_elements.append(el)
            except OperationCanceledException:
                pass
                
        if len(picked_elements) < 2: return
        
        def sort_by_proximity(elements):
            if len(elements) <= 2: return elements
            pts = {}
            for el in elements:
                conns = get_connectors(el)
                if conns: pts[el.Id] = conns[0].Origin
                else: pts[el.Id] = el.Location.Point if hasattr(el.Location, 'Point') else XYZ.Zero
            max_d = -1
            start_el = elements[0]
            for e1 in elements:
                for e2 in elements:
                    if e1.Id != e2.Id:
                        d = pts[e1.Id].DistanceTo(pts[e2.Id])
                        if d > max_d: max_d, start_el = d, e1
            sorted_els = [start_el]
            remaining = [e for e in elements if e.Id != start_el.Id]
            current = start_el
            while remaining:
                closest, min_d = None, float('inf')
                for r in remaining:
                    d = pts[current.Id].DistanceTo(pts[r.Id])
                    if d < min_d: min_d, closest = d, r
                sorted_els.append(closest)
                remaining.remove(closest)
                current = closest
            return sorted_els
            
        picked_elements = sort_by_proximity(picked_elements)
        points_list = [None] * len(picked_elements)
    else:
        if use_connector_mode:
            el1, pt_click1, el2, pt_click2 = pick_elements_with_points()
            if el1 and el2:
                picked_elements = [el1, el2]
                points_list = [pt_click1, pt_click2]
        else:
            el1, el2 = pick_elements_automatic()
            if el1 and el2:
                picked_elements = [el1, el2]
                points_list = [None, None]
        if len(picked_elements) < 2:
            dbg.info("Seleção cancelada.")
            return

    tg = TransactionGroup(doc, "Conectar Eletrodutos em Lote")
    tg.Start()
    try:
        for i in range(len(picked_elements) - 1):
            el1 = picked_elements[i]
            el2 = picked_elements[i+1]
            pt1 = points_list[i]
            pt2 = points_list[i+1]
            same_box = (el1.Id == el2.Id)
            
            dbg.section("Processando Par {}/{}".format(i+1, len(picked_elements)-1))
            _process_pair(el1, el2, pt1, pt2, same_box, use_connector_mode, settings)
            
        tg.Assimilate()
    except Exception as e:
        tg.RollBack()
        raise e

def _process_pair(el1, el2, pt_click1, pt_click2, same_box, use_connector_mode, settings):
    global uidoc, doc, dbg
    def _is_cabletray(el):
        try:
            if el and hasattr(el, "Category") and el.Category:
                if el.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray):
                    return True
        except Exception:
            pass
        return False

    ct1 = _is_cabletray(el1)
    ct2 = _is_cabletray(el2)
    if ct1 and ct2:
        forms.alert(u"Selecione uma eletrocalha/perfilado e um ponto elétrico — não dois percursos.",
                    title="Conectar Eletroduto")
        return False
    if ct1 or ct2:
        cable_tray_el    = el1 if ct1 else el2
        cable_tray_click = (pt_click1 if ct1 else pt_click2)
        other_el         = el2 if ct1 else el1
        other_click      = (pt_click2 if ct1 else pt_click1)
        # Fallback se modo "caixa" (sem ponto de clique preciso): usar meio da eletrocalha
        if cable_tray_click is None:
            try:
                crv = cable_tray_el.Location.Curve
                cable_tray_click = crv.Evaluate(0.5, True)
            except Exception:
                forms.alert(u"Use o modo 'Conector' (Shift+Click → Configurações) para clicar no ponto exato da eletrocalha.",
                            title="Conectar Eletroduto")
                return False
        # Parâmetros de eletroduto (versão simplificada de Fase 3)
        _ct_level_id = cable_tray_el.LevelId
        if _ct_level_id == ElementId.InvalidElementId:
            _ct_level_id = other_el.LevelId
        if _ct_level_id == ElementId.InvalidElementId:
            view = doc.ActiveView
            _ct_level_id = (view.GenLevel.Id if hasattr(view, "GenLevel") and view.GenLevel
                           else FilteredElementCollector(doc).OfClass(Level).FirstElementId())
        _ct_last_ref   = get_last_conduit(doc)
        _ct_conduit_id = (_ct_last_ref.GetTypeId() if _ct_last_ref
                         else get_default_conduit_type(doc))
        _ct_diam = 0.082021
        try:
            _d = float(settings.get('default_diameter', '').replace("mm", "").strip())
            if _d > 0:
                _ct_diam = _d / 304.8
        except Exception:
            pass
        if _ct_diam <= 0.001 and _ct_last_ref:
            try:
                p = _ct_last_ref.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                if p:
                    _ct_diam = p.AsDouble()
            except Exception:
                pass
        _execute_cabletray_connection(
            doc, settings, cable_tray_el, cable_tray_click, other_el, other_click,
            use_connector_mode, _ct_conduit_id, _ct_diam, _ct_level_id, _ct_last_ref
        )
        return False

    # ── Conectores ────────────────────────────────────────────────
    dbg.section("Fase 2: Conectores")
    if use_connector_mode:
        if same_box:
            conns = get_connectors(el1)
            if len(conns) < 2:
                TaskDialog.Show("Erro", "Caixa precisa ter pelo menos 2 conectores.")
                return False
            conn1 = min(conns, key=lambda c: c.Origin.DistanceTo(pt_click1))
            remaining = [c for c in conns if not c.Origin.IsAlmostEqualTo(conn1.Origin)]
            if remaining:
                conn2 = min(remaining, key=lambda c: c.Origin.DistanceTo(pt_click2))
            else:
                TaskDialog.Show("Erro", "Não foi possível identificar um segundo conector diferente.")
                return False
        else:
            conn1, conn2 = find_symmetric_connector_pair(el1, pt_click1, el2, pt_click2)
    else:
        if same_box:
            conn1, conn2 = find_best_connector_pair_same_box(el1)
        else:
            conn1, conn2 = find_best_connector_pair(el1, el2)

    if not conn1 or not conn2:
        forms.alert("Não foi possível encontrar conectores para iniciar o traçado.", title="Erro")
        return False

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

    pref_plan = settings.get('conduit_type_plan', '')
    pref_vert = settings.get('conduit_type_vertical', '')

    def _resolve_conduit_id(pref_name):
        if pref_name and pref_name not in ("(Usar Último Desenhado)", "(Padrão do Revit)"):
            for t in FilteredElementCollector(doc).OfClass(clr.GetClrType(ConduitType)):
                if __get_name__(t) == pref_name:
                    return t.Id, False
        if pref_name == "(Padrão do Revit)":
            return get_default_conduit_type(doc), True
        if last_ref_conduit:
            return last_ref_conduit.GetTypeId(), False
        return get_default_conduit_type(doc), False

    conduit_type_id_plan, clear_ref_plan = _resolve_conduit_id(pref_plan)
    conduit_type_id_vert, clear_ref_vert = _resolve_conduit_id(pref_vert)
    dbg.debug("conduit_type_id_plan: {}  conduit_type_id_vert: {}".format(
        conduit_type_id_plan, conduit_type_id_vert))

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

    # Seleciona tipo de eletroduto e modos de ângulo conforme tipo de rota
    conduit_type_id = conduit_type_id_plan if is_flat else conduit_type_id_vert
    if (is_flat and clear_ref_plan) or (not is_flat and clear_ref_vert):
        last_ref_conduit = None
    angle_plan = settings.get('angle_mode_plan', '90')
    angle_vert = settings.get('angle_mode_vertical', '90')

    dbg.debug("dist={:.4f} ft ({:.3f} m)  dz={:.4f} ft  is_flat={}  is_piso={}  same_box={}  force_vertical={}".format(
        dist_direct, dist_metros, dz, is_flat, is_piso, same_box, force_vertical))
    dbg.debug("conduit_type_id: {}  angle_plan={}  angle_vert={}".format(
        conduit_type_id, angle_plan, angle_vert))

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
        dbg.info("Regra 2: mesma caixa, parede/teto, mesmo nível (angle_plan={})".format(angle_plan))
        if angle_plan == 'livre':
            segments = [(pt1, pt2)]
        else:
            p1_c, p2_c = solve_chicane_2d(pt1, pt2, dir1, dir2, stub_len)
            if p1_c and p2_c:
                segments = [(pt1, p1_c), (p1_c, p2_c), (p2_c, pt2)]
            elif angle_plan == '45':
                mid_segs = create_45_degree_path(p_stub1, p_stub2, dir1, dir2)
                segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
            else:
                mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, False)
                segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    elif same_box and not is_flat:
        if is_piso:
            dbg.info("Regra 3a: mesma caixa, piso, desnível")
            mid_segs = create_terrain_segments(p_stub1, p_stub2, dist_metros)
        else:
            dbg.info("Regra 3b: mesma caixa, parede, desnível (angle_vert={})".format(angle_vert))
            if angle_vert == 'livre':
                mid_segs = create_terrain_segments(p_stub1, p_stub2, dist_metros)
            else:
                mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, True)
        segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    elif not same_box and is_flat:
        if is_piso:
            dbg.info("Regra 4a: caixas diferentes, piso, mesmo nível (angle_plan={})".format(angle_plan))
            if angle_plan == 'livre':
                segments = [(pt1, pt2)]
            else:
                p1_c, p2_c = solve_chicane_2d(pt1, pt2, dir1, dir2, stub_len)
                if p1_c and p2_c:
                    segments = [(pt1, p1_c), (p1_c, p2_c), (p2_c, pt2)]
                elif angle_plan == '45':
                    mid_segs = create_45_degree_path(p_stub1, p_stub2, dir1, dir2)
                    segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
                else:
                    mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, False)
                    segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
        else:
            dbg.info("Regra 4b: caixas diferentes, parede, mesmo nível (angle_plan={})".format(angle_plan))
            if angle_plan == 'livre':
                segments = [(pt1, pt2)]
            elif angle_plan == '45':
                mid_segs = create_45_degree_path(p_stub1, p_stub2, dir1, dir2)
                segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
            else:
                mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, False)
                segments = [(pt1, p_stub1)] + mid_segs + [(p_stub2, pt2)]
    else:
        if force_vertical:
            dbg.info("Regra 5a: caixas diferentes, desnível, luminária")
            mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, True)
        elif is_piso:
            dbg.info("Regra 5b: caixas diferentes, piso, desnível (angle_vert={})".format(angle_vert))
            if angle_vert in ('livre', '45'):
                mid_segs = create_terrain_segments(p_stub1, p_stub2, dist_metros)
            else:
                mid_segs = create_90_degree_path(p_stub1, p_stub2, dir1, dir2, True)
        else:
            dbg.info("Regra 5c: caixas diferentes, parede, desnível (angle_vert={})".format(angle_vert))
            if angle_vert == 'livre':
                mid_segs = create_terrain_segments(p_stub1, p_stub2, dist_metros)
            elif angle_vert == '45':
                mid_segs = create_45_degree_path(p_stub1, p_stub2, dir1, dir2)
            else:
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
    preprocessor = make_warning_swallower()
    if preprocessor:
        ops.SetFailuresPreprocessor(preprocessor)
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
    # Carrega config inicial apenas para o BOOT do debug
    init_settings = load_config()
    is_debug = init_settings.get('debug_mode', False)
    
    # Inicia logger básico
    boot_dbg = DebugLogger(is_debug)
    boot_dbg.section("Conectar Eletroduto — BOOT")
    boot_dbg.info("DEBUG_MODE = {}".format(is_debug))
    
    try:
        execute_connection()
        boot_dbg.section("Ferramenta Finalizada")
    except OperationCanceledException:
        boot_dbg.info("Operação cancelada pelo usuário.")
    except Exception as e:
        err_tb = traceback.format_exc()
        boot_dbg.error("CRASH FATAL:\n{}".format(err_tb))
        if is_debug:
            forms.alert("CRASH FATAL (DEBUG):\n\n" + err_tb, title="Conectar Eletroduto — Erro")

if __name__ == "__main__":
    safe_execution()
