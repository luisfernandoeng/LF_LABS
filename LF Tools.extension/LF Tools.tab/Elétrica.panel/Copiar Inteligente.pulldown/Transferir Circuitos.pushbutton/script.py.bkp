# -*- coding: utf-8 -*-
"""Transferir Circuitos — Desconecta circuitos do quadro de origem
e reconecta no destino com o número de polos escolhido.

Fluxo:
  1. Tenta SelectPanel (rápido, quando tensão/fases são compatíveis).
  2. Se falhar, recria o circuito no destino preservando propriedades.
"""
__title__ = "Transferir\nCircuitos"
__author__ = "Luís Fernando"

# ╔══════════════════════════════════════════════════════════╗
# ║  DEBUG_MODE                                              ║
# ║  True  = imprime detalhes no console pyRevit             ║
# ║  False = silencioso                                      ║
# ╚══════════════════════════════════════════════════════════╝
DEBUG_MODE = False

import os
import re
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Collections')

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, TransactionStatus, SubTransaction, ElementSet, ElementId, StorageType, Domain
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType, DistributionSysType

from pyrevit import forms, script
from lf_utils import DebugLogger, make_warning_swallower

# ══════════════════════════════════════════════════════════════
#  INIT
# ══════════════════════════════════════════════════════════════

dbg   = DebugLogger(DEBUG_MODE)
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document


# ══════════════════════════════════════════════════════════════
#  SUPPRESS DIALOG (igual ao lf_electrical_core)
# ══════════════════════════════════════════════════════════════

def _elec_dialog_handler(sender, args):
    try:
        if hasattr(args, 'DialogId') and 'SpecifyCircuitInfo' in str(args.DialogId):
            args.OverrideResult(1)
    except Exception:
        pass

class suppress_elec_dialog(object):
    """Auto-confirma o dialog 'Specify Circuit Info' durante ElectricalSystem.Create.
    Quando tensão é read-only na família, o Revit exibe esse dialog.
    Confirmar com OK cria o circuito sem tensão fixada, permitindo que
    apliquemos o DistributionSysType do painel depois."""
    def __enter__(self):
        try: __revit__.DialogBoxShowing += _elec_dialog_handler
        except Exception: pass
        return self
    def __exit__(self, *args):
        try: __revit__.DialogBoxShowing -= _elec_dialog_handler
        except Exception: pass

_BUNDLE_DIR = os.path.dirname(__file__)


# ══════════════════════════════════════════════════════════════
#  FUNÇÕES DE APOIO
# ══════════════════════════════════════════════════════════════

def _safe_name(el):
    """Lê .Name de forma segura (IronPython e pythonnet)."""
    try:
        return el.Name
    except Exception:
        pass
    try:
        p = el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.HasValue:
            return p.AsString()
    except Exception:
        pass
    return "ID " + str(el.Id)


def get_electrical_panels():
    """Retorna lista de painéis (ElectricalEquipment) ordenados pelo nome."""
    collector = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                 .WhereElementIsNotElementType())
    panels = []
    for p in collector:
        name = _safe_name(p)
        panels.append({
            "element": p,
            "id": p.Id,
            "name": name,
            "display": "{} (ID: {})".format(name, p.Id)
        })
    return sorted(panels, key=lambda x: x["name"])


def get_circuits_from_panel(panel_element):
    """Pega os circuitos onde panel_element atua como painel (BaseEquipment)."""
    circuits = []
    try:
        mep = panel_element.MEPModel
        if mep is None:
            return circuits
        systems = mep.GetElectricalSystems()
        if systems:
            for sys in systems:
                try:
                    base_eq = sys.BaseEquipment
                    if base_eq and base_eq.Id == panel_element.Id:
                        circuits.append(sys)
                except Exception:
                    pass
    except Exception as e:
        dbg.error("get_circuits_from_panel: {}".format(e))
    return circuits


def get_unassigned_circuits():
    """Retorna circuitos (PowerCircuit) sem quadro atribuído."""
    circuits = []
    try:
        all_sys = FilteredElementCollector(doc).OfClass(ElectricalSystem).ToElements()
        for sys in all_sys:
            try:
                if sys.SystemType == ElectricalSystemType.PowerCircuit:
                    if not sys.BaseEquipment:
                        circuits.append(sys)
            except Exception:
                pass
    except Exception as e:
        dbg.error("get_unassigned_circuits: {}".format(e))
    return circuits


def _get_circuit_number(circ):
    """Retorna o numero do circuito como texto, tolerante a falhas."""
    try:
        num = circ.CircuitNumber
        return str(num or "")
    except Exception:
        return ""


def _get_circuit_start_slot(circ):
    """Retorna o slot inicial do circuito no quadro, quando disponivel."""
    try:
        slot = int(circ.StartSlot)
        if slot > 0:
            return slot
    except Exception:
        pass
    return None


def _circuit_sort_key(circ):
    """Ordenacao pela posicao real no quadro de origem.

    StartSlot e a ordem que aparece no painel. CircuitNumber entra apenas
    como fallback para casos em que a API nao exponha o slot.
    """
    slot = _get_circuit_start_slot(circ)
    try:
        eid = circ.Id.IntegerValue
    except Exception:
        eid = 0

    if slot is not None:
        return (0, slot, eid)

    cnum = _get_circuit_number(circ).strip()
    parts = re.split(r'(\d+)', cnum)
    natural = []
    for part in parts:
        if not part:
            continue
        if part.isdigit():
            natural.append((0, int(part)))
        else:
            natural.append((1, part.lower()))

    return (1, natural, eid)


# ══════════════════════════════════════════════════════════════
#  SNAPSHOT / RESTORE DE PROPRIEDADES
# ══════════════════════════════════════════════════════════════

# (label, BuiltInParameter ou None, [nomes lookup fallback])
_CIRCUIT_PROPS = [
    ("LoadName",        BuiltInParameter.RBS_ELEC_CIRCUIT_NAME, []),
    ("Rating",          BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM, []),
    ("Comments",        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, []),
    ("LoadClassification", None, ["Tipo de Carga"]),
    ("Description",     None, ["Descrição", "Description"]),
    ("WireSize",        None, ["Seção do Condutor Adotado (mm²)", "Condutor Adotado"]),
    ("L Considerado",   None, ["L Considerado", "L Considerado (m)"]),
    ("FCA",             None, ["FCA"]),
    ("FCT",             None, ["FCT"]),
]


def _read_param(elem, bip, names):
    """Lê o valor de um parâmetro (BuiltInParameter → LookupParameter)."""
    if bip is not None:
        try:
            p = elem.get_Parameter(bip)
            if p and p.HasValue:
                st = p.StorageType
                if st == StorageType.String:
                    return ("str", p.AsString())
                elif st == StorageType.Double:
                    return ("dbl", p.AsDouble())
                elif st == StorageType.Integer:
                    return ("int", p.AsInteger())
                elif st == StorageType.ElementId:
                    return ("eid", p.AsElementId())
        except Exception:
            pass
    for n in names:
        try:
            p = elem.LookupParameter(n)
            if p and p.HasValue:
                st = p.StorageType
                if st == StorageType.String:
                    return ("str", p.AsString())
                elif st == StorageType.Double:
                    return ("dbl", p.AsDouble())
                elif st == StorageType.Integer:
                    return ("int", p.AsInteger())
                elif st == StorageType.ElementId:
                    return ("eid", p.AsElementId())
        except Exception:
            pass
    return None


def _write_param(elem, bip, names, typed_val):
    """Escreve (tipo, valor) num parâmetro. Retorna True se conseguiu."""
    if typed_val is None:
        return False
    kind, val = typed_val

    def _do_set(p):
        if p is None or p.IsReadOnly:
            return False
        try:
            if kind == "str" and p.StorageType == StorageType.String:
                p.Set(str(val) if val else "")
                return True
            elif kind == "dbl" and p.StorageType == StorageType.Double:
                p.Set(float(val))
                return True
            elif kind == "int" and p.StorageType == StorageType.Integer:
                p.Set(int(val))
                return True
            elif kind == "eid" and p.StorageType == StorageType.ElementId:
                p.Set(val)
                return True
        except Exception:
            pass
        return False

    if bip is not None:
        try:
            if _do_set(elem.get_Parameter(bip)):
                return True
        except Exception:
            pass
    for n in names:
        try:
            if _do_set(elem.LookupParameter(n)):
                return True
        except Exception:
            pass
    return False


def snapshot_circuit(circ):
    """Captura todas as propriedades de um circuito elétrico."""
    snap = {}
    for label, bip, names in _CIRCUIT_PROPS:
        snap[label] = _read_param(circ, bip, names)
    try:
        snap["_CircuitNumber"] = circ.CircuitNumber
    except Exception:
        snap["_CircuitNumber"] = ""
    try:
        snap["_LoadName"] = circ.LoadName
    except Exception:
        snap["_LoadName"] = ""
    try:
        snap["_Poles"] = circ.PolesNumber
    except Exception:
        snap["_Poles"] = 1
    return snap


def restore_circuit(new_circ, snap):
    """Aplica as propriedades capturadas no circuito novo."""
    for label, bip, names in _CIRCUIT_PROPS:
        val = snap.get(label)
        if val is not None:
            _write_param(new_circ, bip, names, val)


_VOLTAGE_PARAM_NAMES = [
    "Tensão", "Tensao", "Voltagem", "Voltage",
    "Tensão Nominal", "Tensao Nominal", "Voltagem Nominal",
]

# Fator de conversão Revit: unidade interna = Volts × _VOLT_CONV
_VOLT_CONV = 10.7639104167

try:
    _TEXT_TYPE = unicode
except NameError:
    _TEXT_TYPE = str

_SHARED_VOLTAGE_GUID_PREFIXES = ["2bf202e8"]
_SHARED_PHASE_GUID_PREFIXES = ["d1d0c4b4-47d8-45bc-a138-06f76d6f0beb", "d1d0c4b4"]
_MEMBER_VOLTAGE_PARAM_NAMES = [
    u"Tensão (V)", u"Tensao (V)", u"Tensão", u"Tensao",
    u"Voltagem", "Voltage",
]
_MEMBER_PHASE_PARAM_NAMES = [
    u"N° de Fases", u"Nº de Fases", u"Número de Fases",
    u"Numero de Fases", u"Fases", u"Pólos", u"Polos",
    "Number of Phases", "Number of Poles",
]


def _circuit_voltage_volts(circ):
    try:
        return circ.Voltage / _VOLT_CONV
    except Exception:
        return None


def _norm_text(value):
    try:
        return _TEXT_TYPE(value or "").strip().lower()
    except Exception:
        try:
            return str(value or "").strip().lower()
        except Exception:
            return ""


def _param_guid_text(param):
    try:
        return _norm_text(param.GUID)
    except Exception:
        return ""


def _param_def_name(param):
    try:
        if param and param.Definition:
            return param.Definition.Name or ""
    except Exception:
        pass
    return ""


def _param_display_value(param):
    try:
        return param.AsValueString() or param.AsString()
    except Exception:
        pass
    try:
        if param.StorageType == StorageType.Double:
            return "{:.4f}".format(param.AsDouble())
        if param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
        if param.StorageType == StorageType.ElementId:
            return str(param.AsElementId().IntegerValue)
    except Exception:
        pass
    return "?"


def _param_label(param):
    guid = _param_guid_text(param)
    guid_part = " GUID={}".format(guid) if guid else ""
    return "[{}] RO={} ST={}{} val={}".format(
        _param_def_name(param), param.IsReadOnly, param.StorageType,
        guid_part, _param_display_value(param))


def _iter_params(elem):
    try:
        for param in elem.Parameters:
            yield param
    except Exception:
        return


def _find_param_by_name_or_guid(elem, names, guid_prefixes=None):
    """Acha parametro compartilhado preferindo o duplicado gravavel.

    LookupParameter pode devolver um parametro somente-leitura quando ha outro
    com o mesmo nome. Varrer Parameters evita esse falso bloqueio.
    """
    if elem is None:
        return None
    guid_prefixes = [_norm_text(g) for g in (guid_prefixes or []) if g]
    norm_names = [_norm_text(n) for n in names]
    matches = []

    for param in _iter_params(elem):
        pname = _norm_text(_param_def_name(param))
        pguid = _param_guid_text(param)
        by_guid = pguid and any(pguid.startswith(g) for g in guid_prefixes)
        by_name = pname and pname in norm_names
        if by_guid or by_name:
            score = 0
            if by_guid:
                score += 4
            if by_name:
                score += 2
            if not param.IsReadOnly:
                score += 1
            matches.append((score, param))

    if matches:
        matches.sort(key=lambda item: item[0], reverse=True)
        return matches[0][1]

    for name in names:
        try:
            param = elem.LookupParameter(name)
            if param:
                return param
        except Exception:
            pass
    return None


def _set_member_voltage_param(param, target_volts):
    val_internal = float(target_volts) * _VOLT_CONV
    val_text = "{} V".format(int(round(float(target_volts))))
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType == StorageType.Double:
            param.Set(val_internal)
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(int(round(float(target_volts))))
            return True
        if param.StorageType == StorageType.String:
            param.Set(val_text)
            return True
    except Exception:
        pass
    try:
        param.SetValueString(val_text)
        return True
    except Exception:
        return False


def _set_member_phase_param(param, target_poles):
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType == StorageType.Integer:
            param.Set(int(target_poles))
            return True
        if param.StorageType == StorageType.Double:
            param.Set(float(target_poles))
            return True
        if param.StorageType == StorageType.String:
            param.Set(str(int(target_poles)))
            return True
    except Exception:
        pass
    try:
        param.SetValueString(str(int(target_poles)))
        return True
    except Exception:
        return False


def _force_member_voltage(member, target_volts, target_poles):
    """Escreve Tensão (V) e N° de Fases antes de ElectricalSystem.Create."""
    errors = []
    changed = False

    if target_volts is not None:
        p_v = _find_param_by_name_or_guid(
            member, _MEMBER_VOLTAGE_PARAM_NAMES, _SHARED_VOLTAGE_GUID_PREFIXES)
        if p_v is None:
            errors.append(u"param 'Tensão (V)' não encontrado")
        elif p_v.IsReadOnly:
            errors.append(u"'Tensão (V)' somente-leitura {}".format(_param_label(p_v)))
        elif _set_member_voltage_param(p_v, target_volts):
            changed = True
            dbg.debug(u"force member voltage OK Id={} {}".format(member.Id, _param_label(p_v)))
        else:
            errors.append(u"falha ao setar Tensão (V) {}".format(_param_label(p_v)))

    if target_poles is not None:
        p_f = _find_param_by_name_or_guid(
            member, _MEMBER_PHASE_PARAM_NAMES, _SHARED_PHASE_GUID_PREFIXES)
        if p_f is None:
            errors.append(u"param 'N° de Fases' não encontrado")
        elif p_f.IsReadOnly:
            errors.append(u"'N° de Fases' somente-leitura {}".format(_param_label(p_f)))
        elif _set_member_phase_param(p_f, target_poles):
            changed = True
            dbg.debug(u"force member fases OK Id={} {}".format(member.Id, _param_label(p_f)))
        else:
            errors.append(u"falha ao setar N° de Fases {}".format(_param_label(p_f)))

    if errors:
        return False, "; ".join(errors)
    return changed, "OK"


def _get_bip(name):
    try:
        return getattr(BuiltInParameter, name)
    except Exception:
        return None


def _get_panel_dist_id(panel):
    """Retorna o ElementId do DistributionSysType lido diretamente do painel."""
    for bip_name in [
        "RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM",
        "RBS_ELEC_PANEL_DISTRIBUTION_SYSTEM",
    ]:
        try:
            bip = _get_bip(bip_name)
            if bip is None:
                continue
            p = panel.get_Parameter(bip)
            if p and p.HasValue and p.AsElementId() != ElementId.InvalidElementId:
                return p.AsElementId()
        except Exception:
            pass

    try:
        p = panel.LookupParameter(u"Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and p.HasValue and p.AsElementId() != ElementId.InvalidElementId:
            return p.AsElementId()
    except Exception:
        pass
    return None


def _apply_dist_to(elem, dist_id):
    """Aplica DistributionSysType ao elemento/circuito via LookupParameter."""
    if not dist_id:
        return False

    for bip_name in [
        "RBS_ELEC_CIRCUIT_DISTRIBUTION_SYSTEM_PARAM",
        "RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM",
        "RBS_ELEC_PANEL_DISTRIBUTION_SYSTEM",
    ]:
        try:
            bip = _get_bip(bip_name)
            if bip is None:
                continue
            p = elem.get_Parameter(bip)
            if p:
                dbg.debug("Distribution param {} RO={} ST={}".format(
                    bip_name, p.IsReadOnly, p.StorageType))
            if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                p.Set(dist_id)
                dbg.debug("Distribution system setado via {}".format(bip_name))
                return True
        except Exception as e:
            dbg.debug("Falha ao setar {}: {}".format(bip_name, e))

    for name in [u"Sistema de distribuição", "Distribution System", u"Sistema de Distribuição"]:
        try:
            p = elem.LookupParameter(name)
            if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                p.Set(dist_id)
                dbg.debug("Distribution system setado via [{}]".format(name))
                return True
        except Exception:
            pass
    # Fallback: busca qualquer param ElementId com nome de distribuição
    try:
        for p in elem.Parameters:
            if p.IsReadOnly or p.StorageType != StorageType.ElementId:
                continue
            pn = (p.Definition.Name or "").lower() if p.Definition else ""
            if any(kw in pn for kw in ["distribu", "sistema", "system"]):
                try:
                    p.Set(dist_id)
                    return True
                except Exception:
                    pass
    except Exception:
            pass
    return False


def _get_dist_name(dist_id):
    if not dist_id:
        return ""
    try:
        dist = doc.GetElement(dist_id)
        if dist:
            return _safe_name(dist)
    except Exception:
        pass
    return ""


def _apply_system_type_hint(elem, dist_id):
    """Preenche parametros customizados que dirigem formulas de tensao/fases."""
    dist_name = _get_dist_name(dist_id)
    if not dist_name:
        return False

    names = [
        u"Tipo de Sistema", u"Tipo de sistema",
        u"Sistema de Distribuição", u"Sistema de distribuição",
        "System Type", "Distribution System",
    ]

    def _try_target(target):
        changed = False
        for name in names:
            try:
                p = target.LookupParameter(name)
                if not p or p.IsReadOnly:
                    continue
                if p.StorageType == StorageType.String:
                    p.Set(dist_name)
                    changed = True
                elif p.StorageType == StorageType.ElementId:
                    p.Set(dist_id)
                    changed = True
            except Exception:
                pass
        return changed

    changed = _try_target(elem)
    try:
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            changed = _try_target(elem_type) or changed
    except Exception:
        pass
    return changed


def _configure_elec(elem, target_voltage, target_poles, dist_id):
    """Configura tensão (Volts), polos e DistributionSysType num elemento ou circuito.
    Tenta instância primeiro; se params forem RO (type-level), tenta o TYPE do elemento."""
    v_internal = float(target_voltage) * _VOLT_CONV if target_voltage else None
    v_str = (str(int(target_voltage)) + " V") if target_voltage else None

    _apply_system_type_hint(elem, dist_id)

    def _try_set_voltage(target_elem):
        if not v_internal:
            return
        try:
            p = target_elem.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
            if p and not p.IsReadOnly:
                try: p.SetValueString(v_str)
                except Exception: p.Set(v_internal)
                return
        except Exception:
            pass
        for p_name in _VOLTAGE_PARAM_NAMES + [u"Tensão (V)", u"Tensao (V)"]:
            try:
                p = target_elem.LookupParameter(p_name)
                if p and not p.IsReadOnly:
                    try: p.SetValueString(v_str)
                    except Exception:
                        if p.StorageType == StorageType.Double:
                            p.Set(v_internal)
                    return
            except Exception:
                pass

    def _try_set_poles(target_elem):
        try:
            p = target_elem.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
            if p and not p.IsReadOnly:
                p.Set(float(target_poles))
                return
        except Exception:
            pass
        for p_name in [u"Pólos", u"Polos", u"Fases", u"N° de Fases",
                       u"Número de polos", u"Number of Poles"]:
            try:
                p = target_elem.LookupParameter(p_name)
                if p and not p.IsReadOnly:
                    p.Set(int(target_poles))
                    return
            except Exception:
                pass

    # Instância
    _try_set_voltage(elem)
    _try_set_poles(elem)

    # TYPE — type-params aparecem como RO na instância, mas são RW no type
    try:
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            _try_set_voltage(elem_type)
            _try_set_poles(elem_type)
            _apply_dist_to(elem_type, dist_id)
    except Exception:
        pass

    # Conectores elétricos
    for conn in _electrical_connectors(elem):
        if v_internal:
            try: conn.Voltage = v_internal
            except Exception: pass
        try: conn.NumberOfPoles = target_poles
        except Exception: pass
        try: conn.Poles = target_poles
        except Exception: pass

    # DistributionSysType do painel (instância)
    _apply_dist_to(elem, dist_id)

def _force_circuit_voltage_and_poles(circuit, panel, target_voltage, target_poles):
    """Forca DistributionSysType/Poles no circuito recriado antes do SelectPanel."""
    changed = False
    dist_id = _get_panel_dist_id(panel)
    if not dist_id and target_voltage is not None:
        dist_id = _find_matching_dist_sys(target_voltage)

    if dist_id:
        changed = _apply_dist_to(circuit, dist_id) or changed
        try:
            doc.Regenerate()
        except Exception:
            pass
    else:
        dbg.warn("Nao encontrei DistributionSysType para aplicar ao circuito")

    if target_voltage is not None:
        try:
            _set_voltage_value(circuit, target_voltage)
        except Exception:
            pass

    if target_poles is not None:
        try:
            circuit.PolesNumber = int(target_poles)
            changed = True
        except Exception:
            try:
                p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
                if p and not p.IsReadOnly:
                    p.Set(int(target_poles))
                    changed = True
            except Exception as e:
                dbg.debug("Falha ao definir polos do circuito: {}".format(e))

    if changed:
        try:
            doc.Regenerate()
        except Exception:
            pass
    return changed


def _find_matching_dist_sys(target_voltage_volts):
    """Encontra DistributionSysType cujo range de tensão cobre target_voltage_volts."""
    v_internal = float(target_voltage_volts) * _VOLT_CONV
    try:
        for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
            for bip in [BuiltInParameter.RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_G_PARAM,
                        BuiltInParameter.RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_L_PARAM]:
                try:
                    vp = s.get_Parameter(bip)
                    if not vp or not vp.HasValue:
                        continue
                    vid = vp.AsElementId()
                    if vid == ElementId.InvalidElementId:
                        continue
                    vtype = doc.GetElement(vid)
                    if not vtype:
                        continue
                    min_p = vtype.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE_MIN_PARAM)
                    max_p = vtype.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE_MAX_PARAM)
                    if min_p and max_p:
                        if (min_p.AsDouble() - 55) <= v_internal <= (max_p.AsDouble() + 55):
                            return s.Id
                except Exception:
                    pass
    except Exception:
        pass
    return None


def _set_voltage_value(elem, target_voltage):
    """Forca a tensao em um elemento/circuito, quando o parametro permitir."""
    if elem is None or target_voltage is None:
        return False

    val_str = str(int(target_voltage)) + " V"
    val_compact = str(int(target_voltage)) + "V"

    def _try_param(p):
        if p is None or p.IsReadOnly:
            return False
        try:
            p.SetValueString(val_str)
            return True
        except Exception:
            pass
        try:
            if p.StorageType == StorageType.Double:
                p.Set(float(target_voltage) * _VOLT_CONV)
                return True
            if p.StorageType == StorageType.Integer:
                p.Set(int(target_voltage))
                return True
            if p.StorageType == StorageType.String:
                p.Set(val_compact)
                return True
        except Exception:
            pass
        return False

    changed = False
    try:
        changed = _try_param(elem.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)) or changed
    except Exception:
        pass

    for p_name in _VOLTAGE_PARAM_NAMES:
        try:
            changed = _try_param(elem.LookupParameter(p_name)) or changed
        except Exception:
            pass

    return changed


def _force_voltage_on_member(member, target_voltage):
    """Forca tensao na instancia e no tipo do membro do circuito."""
    changed = _set_voltage_value(member, target_voltage)
    try:
        member_type = doc.GetElement(member.GetTypeId())
        changed = _set_voltage_value(member_type, target_voltage) or changed
    except Exception:
        pass
    return changed


def _try_set_attr(obj, attr_name, value):
    try:
        setattr(obj, attr_name, value)
        return True
    except Exception:
        return False


def _debug_connector_capabilities(circ_num, conn, label):
    if not dbg.enabled:
        return
    try:
        names = []
        try:
            for prop in conn.GetType().GetProperties():
                names.append(prop.Name)
        except Exception:
            try:
                names = [n for n in dir(conn) if not n.startswith("_")]
            except Exception:
                names = []

        interesting = []
        tokens = ["volt", "pole", "phase", "system", "load", "power", "apparent", "domain", "type"]
        for name in sorted(set(names)):
            if any(t in name.lower() for t in tokens):
                interesting.append(name)
        dbg.debug("C{} {} CONN props: {}".format(circ_num, label, ", ".join(interesting[:80])))

        for name in interesting[:30]:
            try:
                dbg.debug("C{} {} CONN.{} = {}".format(circ_num, label, name, getattr(conn, name)))
            except Exception as e:
                dbg.debug("C{} {} CONN.{} ERR: {}".format(circ_num, label, name, e))

        try:
            info = conn.GetMEPConnectorInfo()
            info_names = []
            try:
                for prop in info.GetType().GetProperties():
                    info_names.append(prop.Name)
            except Exception:
                try:
                    info_names = [n for n in dir(info) if not n.startswith("_")]
                except Exception:
                    info_names = []
            interesting_info = []
            for name in sorted(set(info_names)):
                if any(t in name.lower() for t in tokens):
                    interesting_info.append(name)
            dbg.debug("C{} {} MEPConnectorInfo props: {}".format(
                circ_num, label, ", ".join(interesting_info[:80])))
            for name in interesting_info[:30]:
                try:
                    dbg.debug("C{} {} INFO.{} = {}".format(circ_num, label, name, getattr(info, name)))
                except Exception as e:
                    dbg.debug("C{} {} INFO.{} ERR: {}".format(circ_num, label, name, e))
        except Exception as e:
            dbg.debug("C{} {} GetMEPConnectorInfo ERR: {}".format(circ_num, label, e))
    except Exception as e:
        dbg.debug("C{} {} connector introspection ERR: {}".format(circ_num, label, e))


def _try_set_connector_obj(obj, target_voltage, target_poles, dist_id):
    changed = False
    if obj is None:
        return changed

    v_internal = float(target_voltage) * _VOLT_CONV if target_voltage is not None else None
    dist_elem = None
    if dist_id:
        try:
            dist_elem = doc.GetElement(dist_id)
        except Exception:
            dist_elem = None

    if v_internal is not None:
        for attr in ["Voltage", "AssignedVoltage", "NominalVoltage"]:
            changed = _try_set_attr(obj, attr, v_internal) or changed
        for attr in ["VoltageValue", "VoltageActualValue"]:
            changed = _try_set_attr(obj, attr, float(target_voltage)) or changed

    if target_poles is not None:
        for attr in ["Poles", "NumberOfPoles", "PolesNumber", "NumPoles", "PhaseNumber"]:
            changed = _try_set_attr(obj, attr, int(target_poles)) or changed

    for attr in ["DistributionSystem", "DistributionSysType", "DistributionSystemType"]:
        if dist_elem is not None:
            changed = _try_set_attr(obj, attr, dist_elem) or changed
        if dist_id:
            changed = _try_set_attr(obj, attr + "Id", dist_id) or changed

    return changed


def _get_member_connectors(member):
    connectors = []
    try:
        mep = member.MEPModel
        if mep and mep.ConnectorManager:
            for conn in mep.ConnectorManager.Connectors:
                connectors.append(conn)
    except Exception:
        pass
    try:
        if member.ConnectorManager:
            for conn in member.ConnectorManager.Connectors:
                connectors.append(conn)
    except Exception:
        pass
    return connectors


def _connector_domain_name(conn):
    try:
        return str(conn.Domain)
    except Exception:
        return ""


def _is_domain_electrical(conn):
    return _connector_domain_name(conn) == "DomainElectrical"


def _electrical_connectors(elem):
    connectors = []
    for conn in _get_member_connectors(elem):
        try:
            if _is_domain_electrical(conn):
                connectors.append(conn)
            elif dbg.enabled:
                dbg.debug("Ignorando conector nao eletrico: {}".format(_connector_domain_name(conn)))
        except Exception:
            pass
    return connectors


def _electrical_config_targets(elem):
    """Elemento + subcomponentes com conectores eletricos reais."""
    targets = []
    try:
        if _electrical_connectors(elem):
            targets.append(elem)
    except Exception:
        pass

    try:
        sub_ids = elem.GetSubComponentIds()
        if sub_ids:
            for sub_id in sub_ids:
                sub = doc.GetElement(sub_id)
                if sub and _electrical_connectors(sub):
                    targets.append(sub)
    except Exception:
        pass

    if not targets:
        targets.append(elem)
    return targets


def _get_first_electrical_connector(elem):
    for target in _electrical_config_targets(elem):
        conns = _electrical_connectors(target)
        if conns:
            try:
                dbg.debug("Create connector target Id={} Cat={}".format(
                    target.Id.IntegerValue,
                    target.Category.Name if target.Category else "?"))
            except Exception:
                pass
            return conns[0]
    return None


def _get_available_electrical_connector(elem):
    """Retorna conector DomainElectrical livre, preferindo MEPSystem=None."""
    fallback = None
    for target in _electrical_config_targets(elem):
        for conn in _electrical_connectors(target):
            if fallback is None:
                fallback = (conn, target)
            try:
                if conn.MEPSystem is None:
                    try:
                        dbg.debug("Create connector livre Id={} Cat={}".format(
                            target.Id.IntegerValue,
                            target.Category.Name if target.Category else "?"))
                    except Exception:
                        pass
                    return conn, target
            except Exception:
                pass
    if fallback:
        conn, target = fallback
        try:
            dbg.debug("Create connector fallback Id={} Cat={} MEPSystem={}".format(
                target.Id.IntegerValue,
                target.Category.Name if target.Category else "?",
                getattr(conn, "MEPSystem", None)))
        except Exception:
            pass
        return fallback
    return None, None


def _force_connector_electrical_config(member, target_voltage, target_poles):
    """Tenta alinhar os conectores eletricos que governam a tensao do circuito."""
    changed = False
    dist_id = _find_matching_dist_sys(target_voltage) if target_voltage is not None else None
    for conn in _electrical_connectors(member):
        changed = _try_set_connector_obj(conn, target_voltage, target_poles, dist_id) or changed
        try:
            info = conn.GetMEPConnectorInfo()
            changed = _try_set_connector_obj(info, target_voltage, target_poles, dist_id) or changed
        except Exception:
            pass

        if target_voltage is not None:
            # conn.Voltage usa unidades internas do Revit, não Volts
            changed = _try_set_attr(conn, "Voltage", float(target_voltage) * _VOLT_CONV) or changed
        try:
            if target_poles is not None:
                changed = _try_set_attr(conn, "NumberOfPoles", int(target_poles)) or changed
                changed = _try_set_attr(conn, "PolesNumber", int(target_poles)) or changed
        except Exception:
            pass
    return changed


def _force_member_electrical_config(member, target_voltage, target_poles):
    changed = False
    changed = _force_voltage_on_member(member, target_voltage) or changed
    changed = _force_connector_electrical_config(member, target_voltage, target_poles) or changed
    # Sem o DistributionSysType correto o Revit recusa SelectPanel mesmo com tensão certa
    if target_voltage is not None:
        dist_id = _find_matching_dist_sys(target_voltage)
        if dist_id:
            for name in [u"Sistema de distribuição", u"Distribution System",
                         u"Sistema de Distribuição"]:
                try:
                    p = member.LookupParameter(name)
                    if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                        p.Set(dist_id)
                        changed = True
                        break
                except Exception:
                    pass
    return changed


def _element_id_list(elements):
    ids = List[ElementId]()
    for elem in elements:
        ids.Add(elem.Id)
    return ids


def _refresh_elements_by_id(element_ids):
    refreshed = []
    for eid in element_ids:
        try:
            elem = doc.GetElement(eid)
            if elem is not None:
                refreshed.append(elem)
        except Exception:
            pass
    return refreshed


def _disconnect_panel_if_possible(circuit, circ_num):
    """Imita o primeiro passo manual: deixar o circuito sem painel antes do delete."""
    for method_name in ["DisconnectPanel", "DisconnectFromPanel", "RemovePanel"]:
        try:
            method = getattr(circuit, method_name, None)
            if method:
                method()
                dbg.debug("C{}: painel desconectado via {}".format(circ_num, method_name))
                return True
        except Exception as e:
            dbg.debug("C{}: {} falhou: {}".format(circ_num, method_name, e))

    try:
        circuit.SelectPanel(None)
        dbg.debug("C{}: painel desconectado via SelectPanel(None)".format(circ_num))
        return True
    except Exception as e:
        dbg.debug("C{}: SelectPanel(None) falhou: {}".format(circ_num, e))
    return False


def _circuit_base_equipment(circuit):
    try:
        fresh = doc.GetElement(circuit.Id)
        if fresh is not None:
            circuit = fresh
    except Exception:
        pass
    try:
        return circuit.BaseEquipment
    except Exception:
        return None


def _verify_circuit_on_panel(circuit, panel):
    base = _circuit_base_equipment(circuit)
    try:
        if base is not None and base.Id == panel.Id:
            return True
    except Exception:
        pass

    try:
        mep = panel.MEPModel
        systems = mep.GetElectricalSystems() if mep else None
        if systems:
            for sys in systems:
                try:
                    if sys.Id == circuit.Id:
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _create_power_circuit_from_members(members, circ_num, dist_id=None, allow_user_circuit_info=False):
    id_list = _element_id_list(members)
    sys_type = ElectricalSystemType.PowerCircuit

    # Em conversao de tensao, Revit 2025 respeita melhor o overload por Connector.
    # O overload de 4 args nao existe em algumas versoes e nao e o caminho do fix REVIT-180706.
    if dist_id and not allow_user_circuit_info:
        try:
            circ = ElectricalSystem.Create(doc, id_list, sys_type, dist_id)
            if circ is not None:
                dbg.debug("C{}: Create 4-arg com dist_id OK".format(circ_num))
                return circ, None
        except Exception as e:
            dbg.debug("C{}: Create 4-arg falhou (provavelmente Revit < 2022): {}".format(circ_num, e))

    def _create_3arg():
        if allow_user_circuit_info:
            try:
                conn = None
                conn_owner = None
                for m in members:
                    conn, conn_owner = _get_available_electrical_connector(m)
                    if conn:
                        break
                if conn:
                    dbg.debug("C{}: tentando Create por conector eletrico livre".format(circ_num))
                    new_circ = ElectricalSystem.Create(conn, sys_type)
                    if new_circ is None:
                        return None, "ElectricalSystem.Create(connector) retornou None"
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass

                    for m in members:
                        try:
                            if conn_owner is not None and m.Id == conn_owner.Id:
                                continue
                            add_ids = List[ElementId]()
                            add_ids.Add(m.Id)
                            new_circ.AddToCircuit(add_ids)
                        except Exception as add_err:
                            dbg.warn("C{}: membro {} nao adicionado: {}".format(
                                circ_num, m.Id, add_err))
                    return new_circ, None
            except Exception as conn_err:
                dbg.debug("C{}: Create por conector falhou: {}".format(circ_num, conn_err))

        try:
            return ElectricalSystem.Create(doc, id_list, sys_type), None
        except Exception as all_err:
            dbg.debug("C{}: Criacao com todos os membros falhou: {}".format(circ_num, all_err))

        try:
            first_ids = List[ElementId]()
            first_ids.Add(members[0].Id)
            new_circ = ElectricalSystem.Create(doc, first_ids, sys_type)
        except Exception as first_err:
            return None, first_err

        for m in members[1:]:
            try:
                add_ids = List[ElementId]()
                add_ids.Add(m.Id)
                new_circ.AddToCircuit(add_ids)
            except Exception as add_err:
                dbg.warn("C{}: membro {} nao adicionado: {}".format(circ_num, m.Id, add_err))
        return new_circ, None

    if allow_user_circuit_info:
        dbg.debug("C{}: criando sem suprimir Specify Circuit Information".format(circ_num))
        return _create_3arg()

    with suppress_elec_dialog():
        return _create_3arg()


def _debug_circuit_panel(circ_num, circuit, panel, dist_id):
    """Loga parâmetros relevantes do circuito e do painel para diagnóstico de mismatch."""
    try:
        dbg.debug("=== DEBUG mismatch C{} ===".format(circ_num))
        # Circuito
        for label, getter in [
            ("circ.Voltage", lambda: circuit.Voltage),
            ("circ.PolesNumber", lambda: circuit.PolesNumber),
        ]:
            try: dbg.debug("  {}: {}".format(label, getter()))
            except Exception as e: dbg.debug("  {}: ERR {}".format(label, e))
        for p in circuit.Parameters:
            try:
                if p.Definition is None: continue
                pn = (p.Definition.Name or "").lower()
                if not any(kw in pn for kw in ["distribu", "sistema", "system", "polo", "fases", "tens", "volt"]):
                    continue
                try: val = p.AsValueString() or p.AsString() or str(p.AsDouble())
                except Exception:
                    try: val = str(p.AsElementId().IntegerValue)
                    except Exception: val = "?"
                dbg.debug("  CIRC [{}] RO={} val={}".format(p.Definition.Name, p.IsReadOnly, val))
            except Exception: pass
        # Painel
        dbg.debug("  Panel dist_id buscado: {}".format(dist_id))
        for p in panel.Parameters:
            try:
                if p.Definition is None: continue
                pn = (p.Definition.Name or "").lower()
                if not any(kw in pn for kw in ["distribu", "sistema", "system"]):
                    continue
                try: val = p.AsValueString() or p.AsString() or str(p.AsElementId().IntegerValue)
                except Exception: val = "?"
                dbg.debug("  PANEL [{}] RO={} val={}".format(p.Definition.Name, p.IsReadOnly, val))
            except Exception: pass
        dbg.debug("=== END DEBUG ===")
    except Exception as e:
        dbg.debug("_debug_circuit_panel falhou: {}".format(e))


# ══════════════════════════════════════════════════════════════
#  TRANSFERÊNCIA INTELIGENTE
# ══════════════════════════════════════════════════════════════

def _debug_member_dist(circ_num, member, label):
    """Loga parâmetros de distribuição/tensão do membro para diagnóstico."""
    try:
        for p in member.Parameters:
            try:
                if not p.Definition:
                    continue
                pn = p.Definition.Name or ""
                if not any(kw in pn.lower() for kw in
                           ["distribu", "sistem", "tens", "volt", "polo", "fase"]):
                    continue
                try:
                    val = p.AsValueString() or p.AsString()
                    if not val:
                        val = str(p.AsElementId().IntegerValue)
                except Exception:
                    try:
                        val = "{:.2f}".format(p.AsDouble())
                    except Exception:
                        val = "?"
                dbg.debug("  C{} {} [{}] RO={} ST={} val={}".format(
                    circ_num, label, pn, p.IsReadOnly, p.StorageType, val))
            except Exception:
                pass
        try:
            mep = getattr(member, 'MEPModel', None)
            dbg.debug("  C{} {} MEPModel={}".format(circ_num, label, type(mep).__name__ if mep else "None"))
            mgr = getattr(mep, 'ConnectorManager', None) if mep else None
            if not mgr:
                mgr = getattr(member, 'ConnectorManager', None)
            if mgr:
                n = 0
                for c in mgr.Connectors:
                    n += 1
                    _debug_connector_capabilities(circ_num, c, "{} CONN[{}]".format(label, n))
                    try:
                        v = getattr(c, "Voltage", None)
                        v_text = "{:.2f}".format(v) if v is not None else "N/A"
                        dbg.debug("  C{} {} CONN[{}] Voltage={} Domain={} Poles={}".format(
                            circ_num, label, n, v_text, c.Domain,
                            getattr(c, 'NumberOfPoles', getattr(c, 'Poles', '?'))))
                    except Exception as ce:
                        dbg.debug("  C{} {} CONN[{}] ERR: {}".format(circ_num, label, n, ce))
                if n == 0:
                    dbg.debug("  C{} {} mgr={} mas 0 connectors".format(
                        circ_num, label, type(mgr).__name__))
            else:
                dbg.debug("  C{} {} ConnectorManager=None".format(circ_num, label))
        except Exception as e:
            dbg.debug("  C{} {} CONN access ERR: {}".format(circ_num, label, e))
        # TYPE params
        try:
            member_type = doc.GetElement(member.GetTypeId())
            if member_type:
                for p in member_type.Parameters:
                    try:
                        if not p.Definition:
                            continue
                        pn = p.Definition.Name or ""
                        if not any(kw in pn.lower() for kw in
                                   ["distribu", "sistem", "tens", "volt", "polo", "fase"]):
                            continue
                        try:
                            val = p.AsValueString() or p.AsString()
                            if not val:
                                val = str(p.AsElementId().IntegerValue)
                        except Exception:
                            try:
                                val = "{:.2f}".format(p.AsDouble())
                            except Exception:
                                val = "?"
                        dbg.debug("  C{} {} TYPE [{}] RO={} ST={} val={}".format(
                            circ_num, label, pn, p.IsReadOnly, p.StorageType, val))
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass


def transfer_one_circuit(circ, dest_panel, target_poles, target_voltage=None):
    """Transfere um circuito para dest_panel.

    Estratégia:
      1. SelectPanel direto (compatível sem mudança).
      2. Recriar: configura membros ANTES de deletar (igual Gerenciar Circuito),
         depois delete → create → SelectPanel.

    Retorna (sucesso, mensagem).
    """
    circ_num = ""
    try:
        circ_num = circ.CircuitNumber
    except Exception:
        pass

    # Tensão do circuito original em Volts (converter de unidade interna)
    circ_voltage_v = None
    try:
        circ_voltage_v = circ.Voltage / _VOLT_CONV
    except Exception:
        pass

    effective_voltage = target_voltage
    if effective_voltage is None:
        effective_voltage = circ_voltage_v

    # DistributionSysType do painel destino
    dist_id = _get_panel_dist_id(dest_panel)
    dbg.info("C{}: dist_id do painel = {}  tensao_circ={:.0f}V  target={}V".format(
        circ_num, dist_id,
        circ_voltage_v if circ_voltage_v else 0,
        effective_voltage if effective_voltage else "?"))
    if not dist_id and effective_voltage:
        dist_id = _find_matching_dist_sys(effective_voltage)
        dbg.debug("C{}: dist_id fallback por tensao = {}".format(circ_num, dist_id))
    if not dist_id:
        dbg.warn("C{}: nenhum DistributionSysType encontrado".format(circ_num))

    # Muda de tensão/polos?
    voltage_change = (target_voltage is not None and circ_voltage_v is not None
                      and abs(circ_voltage_v - target_voltage) > 5)
    poles_change = False
    try:
        poles_change = (circ.PolesNumber != target_poles)
    except Exception:
        pass
    need_change = voltage_change or poles_change

    # ── 1. SelectPanel direto (sem mudança de tensão/polos) ──
    if not need_change:
        try:
            circ.SelectPanel(dest_panel)
            dbg.info("C{}: SelectPanel OK direto".format(circ_num))
            return True, "SelectPanel OK"
        except Exception as e:
            dbg.debug("C{}: SelectPanel direto falhou ({})".format(circ_num, e))

    # ── 2. Recriar circuito ──
    snap = snapshot_circuit(circ)

    members = []
    try:
        if circ.Elements:
            for el in circ.Elements:
                members.append(el)
    except Exception:
        pass
    if not members:
        return False, u"Circuito sem membros"
    member_ids = [m.Id for m in members]

    # Configura membros ANTES de deletar (igual Gerenciar Circuito)
    # Regenerate após configure: o doc.Regenerate() após o delete reseta conectores
    # se a configuração for feita após o delete.
    dbg.debug("C{}: configurando {} membro(s) ANTES do delete".format(circ_num, len(members)))
    if dbg.enabled and members:
        _debug_member_dist(circ_num, members[0], "BEFORE")
    for m in members:
        ok, msg = _force_member_voltage(m, effective_voltage, target_poles)
        dbg.debug(u"C{}: forçar tensão/fases em {} -> {}".format(circ_num, m.Id, msg))
        _configure_elec(m, effective_voltage, target_poles, dist_id)
        for target in _electrical_config_targets(m):
            if target.Id != m.Id:
                ok, msg = _force_member_voltage(target, effective_voltage, target_poles)
                dbg.debug(u"C{}: forçar tensão/fases em sub {} -> {}".format(
                    circ_num, target.Id, msg))
                _configure_elec(target, effective_voltage, target_poles, dist_id)
    try:
        doc.Regenerate()
    except Exception:
        pass
    if dbg.enabled and members:
        _debug_member_dist(circ_num, members[0], "AFTER ")

    try:
        _disconnect_panel_if_possible(circ, circ_num)
        doc.Regenerate()
        doc.Delete(circ.Id)
        doc.Regenerate()
    except Exception as e:
        return False, u"Erro ao deletar: {}".format(e)

    members = _refresh_elements_by_id(member_ids)
    if not members:
        return False, u"Membros não encontrados após deletar circuito antigo"

    dbg.debug("C{}: configurando conectores APOS delete".format(circ_num))
    for m in members:
        ok, msg = _force_member_voltage(m, effective_voltage, target_poles)
        dbg.debug(u"C{}: pós-delete tensão/fases em {} -> {}".format(circ_num, m.Id, msg))
        _configure_elec(m, effective_voltage, target_poles, dist_id)
        _force_connector_electrical_config(m, effective_voltage, target_poles)
        for target in _electrical_config_targets(m):
            if target.Id != m.Id:
                ok, msg = _force_member_voltage(target, effective_voltage, target_poles)
                dbg.debug(u"C{}: pós-delete tensão/fases em sub {} -> {}".format(
                    circ_num, target.Id, msg))
                _configure_elec(target, effective_voltage, target_poles, dist_id)
                _force_connector_electrical_config(target, effective_voltage, target_poles)
    try:
        doc.Regenerate()
    except Exception:
        pass
    if dbg.enabled and members:
        _debug_member_dist(circ_num, members[0], "POSTDELETE")

    new_circ = None
    create_err = None
    try:
        new_circ, create_err = _create_power_circuit_from_members(
            members, circ_num, dist_id, allow_user_circuit_info=voltage_change)
    except Exception as e:
        create_err = e
    if create_err:
        return False, u"Erro ao criar: {}".format(create_err)
    if new_circ is None:
        return False, u"ElectricalSystem.Create retornou None"

    _force_circuit_voltage_and_poles(new_circ, dest_panel, effective_voltage, target_poles)

    try:
        doc.Regenerate()
    except Exception:
        pass

    circ_v_after = None
    try:
        circ_v_after = new_circ.Voltage / _VOLT_CONV
    except Exception:
        pass
    dbg.debug("C{}: circuito criado Voltage={} PolesNumber={}".format(
        circ_num,
        "{:.0f}V".format(circ_v_after) if circ_v_after is not None else "?",
        getattr(new_circ, 'PolesNumber', '?')))

    if voltage_change and circ_v_after is not None and abs(circ_v_after - target_voltage) > 5:
        return False, (
            u"O Revit recriou o circuito como {:.0f}V, não como {:.0f}V. "
            u"O circuito antigo foi preservado."
        ).format(circ_v_after, target_voltage)

    # Dump ALL parâmetros do circuito para encontrar qualquer param de distribuição oculto
    if dbg.enabled:
        dbg.debug("C{}: --- ALL circuit params ---".format(circ_num))
        try:
            for _p in new_circ.Parameters:
                try:
                    if not _p.Definition:
                        continue
                    _pn = _p.Definition.Name or ""
                    try:
                        _val = _p.AsValueString() or _p.AsString()
                        if not _val:
                            _eid = _p.AsElementId()
                            _val = str(_eid.IntegerValue) if _eid else "{:.4f}".format(_p.AsDouble())
                    except Exception:
                        _val = "?"
                    dbg.debug("  [{}] RO={} ST={} val={}".format(
                        _pn, _p.IsReadOnly, _p.StorageType, _val))
                except Exception:
                    pass
        except Exception:
            pass
        dbg.debug("C{}: --- END ALL params ---".format(circ_num))

    restore_circuit(new_circ, snap)
    _force_circuit_voltage_and_poles(new_circ, dest_panel, effective_voltage, target_poles)

    try:
        _force_circuit_voltage_and_poles(new_circ, dest_panel, effective_voltage, target_poles)
        new_circ.SelectPanel(dest_panel)
        doc.Regenerate()
        if not _verify_circuit_on_panel(new_circ, dest_panel):
            base = _circuit_base_equipment(new_circ)
            base_name = _safe_name(base) if base else "sem painel"
            return False, u"SelectPanel não persistiu. BaseEquipment atual: {}".format(base_name)
        dbg.info("C{}: Conectado ao painel com sucesso.".format(circ_num))
    except Exception as e:
        dbg.error("C{}: SelectPanel recusou: {}".format(circ_num, e))
        _debug_circuit_panel(circ_num, new_circ, dest_panel, dist_id)
        return False, u"Recriado mas falhou ao conectar: {}".format(e)

    restore_circuit(new_circ, snap)
    try:
        doc.Regenerate()
    except Exception:
        pass
    if not _verify_circuit_on_panel(new_circ, dest_panel):
        base = _circuit_base_equipment(new_circ)
        base_name = _safe_name(base) if base else "sem painel"
        return False, u"Circuito perdeu o painel após restaurar parâmetros. BaseEquipment atual: {}".format(base_name)
    dbg.info("C{}: Recriado e transferido".format(circ_num))
    return True, u"Recriado"


# ══════════════════════════════════════════════════════════════
#  INTERFACE (WPF / pyrevit.forms)
# ══════════════════════════════════════════════════════════════

class TransferCircuitsWindow(forms.WPFWindow):

    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.panels = get_electrical_panels()
        self.dest_panels = []
        # Cada item: (checkbox, combobox_polos, circuito)
        self.circuit_rows = []

        self._init_ui()
        self._bind_events()

    # ── Inicialização ──

    def _init_ui(self):
        if not self.panels:
            self.lbl_Info.Text = "Nenhum quadro elétrico encontrado no projeto."
            self.btn_Transfer.IsEnabled = False
            return

        self.cb_SourcePanel.Items.Add("--- Circuitos Sem Quadro ---")
        for p in self.panels:
            self.cb_SourcePanel.Items.Add(p["display"])

        self._update_dest_list(-1)
        self.lbl_Info.Text = "Selecione o quadro de origem."

        if self.panels:
            self.cb_SourcePanel.SelectedIndex = 0

    def _bind_events(self):
        self.btn_Cancel.Click    += self._on_cancel
        self.btn_Transfer.Click  += self._on_transfer
        self.btn_SelectAll.Click += self._on_select_all
        self.btn_SelectNone.Click += self._on_select_none
        self.cb_SourcePanel.SelectionChanged += self._on_source_changed

    # ── Eventos ──

    def _on_cancel(self, sender, args):
        self.Close()

    def _on_source_changed(self, sender, args):
        from System.Windows.Controls import CheckBox, ComboBox as WpfComboBox
        from System.Windows import Thickness

        self.sp_Circuits.Children.Clear()
        self.circuit_rows = []

        idx = self.cb_SourcePanel.SelectedIndex
        self._update_dest_list(idx)

        if idx < 0:
            return

        if idx == 0:
            circuits = get_unassigned_circuits()
        else:
            panel = self.panels[idx - 1]["element"]
            circuits = get_circuits_from_panel(panel)

        if not circuits:
            self.lbl_CircuitsCount.Text = "Nenhum circuito."
            return

        self.lbl_CircuitsCount.Text = "{} circuito(s)".format(len(circuits))

        circuits = sorted(circuits, key=_circuit_sort_key)

        for circ in circuits:
            self._add_circuit_row(circ)

        self.lbl_Info.Text = "{} circuito(s) — selecione e escolha o destino.".format(len(circuits))

    def _add_circuit_row(self, circ):
        """Cria uma linha: [CheckBox (info)] + [ComboBox (polos)]."""
        from System.Windows.Controls import (
            CheckBox, Grid as WpfGrid, ColumnDefinition,
            ComboBox as WpfComboBox
        )
        from System.Windows import Thickness, GridLength, GridUnitType

        # ── Dados ──
        c_name = ""
        try:
            c_name = circ.LoadName or ""
        except Exception:
            pass
        if not c_name:
            c_name = "Sem Nome"

        c_num = ""
        try:
            c_num = circ.CircuitNumber
        except Exception:
            pass

        curr_poles = 1
        try:
            curr_poles = circ.PolesNumber
        except Exception:
            pass

        load_str = ""
        try:
            load_va = circ.ApparentLoad
            load_str = "{:.0f}VA".format(load_va) if load_va < 1000 else "{:.1f}kVA".format(load_va / 1000)
        except Exception:
            pass

        voltage_str = ""
        try:
            voltage_v = _circuit_voltage_volts(circ)
            if voltage_v is not None:
                voltage_str = "{:.0f}V".format(voltage_v)
        except Exception:
            pass

        lbl = "C{} — {}".format(c_num, c_name)
        extras = []
        if load_str:
            extras.append(load_str)
        if voltage_str:
            extras.append(voltage_str)
        if curr_poles:
            extras.append("{}P".format(curr_poles))
        if extras:
            lbl += "  [{}]".format(" | ".join(extras))

        # ── Visual: Grid com CheckBox + ComboBox ──
        row = WpfGrid()
        col0 = ColumnDefinition()
        col0.Width = GridLength(1.0, GridUnitType.Star)
        col1 = ColumnDefinition()
        col1.Width = GridLength(90.0, GridUnitType.Pixel)
        col2 = ColumnDefinition()
        col2.Width = GridLength(80.0, GridUnitType.Pixel)
        row.ColumnDefinitions.Add(col0)
        row.ColumnDefinitions.Add(col1)
        row.ColumnDefinitions.Add(col2)
        row.Margin = Thickness(0, 2, 0, 2)

        cb = CheckBox()
        cb.Content = lbl
        cb.IsChecked = True
        cb.FontSize = 13
        WpfGrid.SetColumn(cb, 0)
        row.Children.Add(cb)
        
        volt_cb = WpfComboBox()
        volt_cb.Items.Add("Manter")
        volt_cb.Items.Add("220V")
        volt_cb.Items.Add("380V")
        volt_cb.Items.Add("127V")
        volt_cb.SelectedIndex = 0
        volt_cb.Width = 85
        volt_cb.Height = 22
        volt_cb.FontSize = 11
        WpfGrid.SetColumn(volt_cb, 1)
        row.Children.Add(volt_cb)

        poles_cb = WpfComboBox()
        poles_cb.Items.Add("1 Polo")
        poles_cb.Items.Add("2 Polos")
        poles_cb.Items.Add("3 Polos")
        if curr_poles == 3:
            poles_cb.SelectedIndex = 2
        elif curr_poles == 2:
            poles_cb.SelectedIndex = 1
        else:
            poles_cb.SelectedIndex = 0
        poles_cb.Width = 75
        poles_cb.Height = 22
        poles_cb.FontSize = 11
        WpfGrid.SetColumn(poles_cb, 2)
        row.Children.Add(poles_cb)

        self.sp_Circuits.Children.Add(row)
        self.circuit_rows.append((cb, volt_cb, poles_cb, circ))

    def _update_dest_list(self, source_idx):
        """Atualiza a ComboBox destino excluindo o quadro de origem."""
        prev_sel = self.cb_DestPanel.SelectedItem

        self.cb_DestPanel.Items.Clear()
        self.dest_panels = []

        for i, p in enumerate(self.panels):
            # i+1 is because cb_SourcePanel has index 0 as Unassigned
            if (i + 1) != source_idx:
                self.cb_DestPanel.Items.Add(p["display"])
                self.dest_panels.append(p)

        if prev_sel and prev_sel in [p["display"] for p in self.dest_panels]:
            self.cb_DestPanel.SelectedItem = prev_sel
        else:
            self.cb_DestPanel.SelectedIndex = -1

    def _on_select_all(self, sender, args):
        for cb, _, _, _ in self.circuit_rows:
            cb.IsChecked = True
        self._update_count()

    def _on_select_none(self, sender, args):
        for cb, _, _, _ in self.circuit_rows:
            cb.IsChecked = False
        self._update_count()

    def _update_count(self):
        count = sum(1 for cb, _, _, _ in self.circuit_rows if cb.IsChecked)
        total = len(self.circuit_rows)
        self.lbl_Info.Text = "{} de {} selecionado(s).".format(count, total)

    # ── Transferir ──

    def _on_transfer(self, sender, args):
        s_idx = self.cb_SourcePanel.SelectedIndex
        d_idx = self.cb_DestPanel.SelectedIndex

        if s_idx < 0:
            forms.alert("Escolha o quadro de origem.", title="Transferir Circuitos")
            return
        if d_idx < 0:
            forms.alert("Escolha o quadro de destino.", title="Transferir Circuitos")
            return

        # Aplicar config de debug do checkbox
        global dbg
        if hasattr(self, "cb_Debug"):
            dbg.enabled = bool(self.cb_Debug.IsChecked)

        dest_panel = self.dest_panels[d_idx]["element"]

        # Coletar selecionados + polos desejados + tensao
        selected = []
        for cb, volt_cb, poles_cb, circ in self.circuit_rows:
            if cb.IsChecked:
                pi = poles_cb.SelectedIndex
                target_poles = [1, 2, 3][pi] if 0 <= pi <= 2 else 1
                
                vi = volt_cb.SelectedIndex
                target_voltage = None
                if vi == 1: target_voltage = 220.0
                elif vi == 2: target_voltage = 380.0
                elif vi == 3: target_voltage = 127.0
                
                selected.append((circ, target_poles, target_voltage))

        selected = sorted(selected, key=lambda item: _circuit_sort_key(item[0]))

        if not selected:
            forms.alert("Selecione pelo menos um circuito.", title="Transferir Circuitos")
            return

        dest_name = _safe_name(dest_panel)
        voltage_changes = []
        for circ, _, target_voltage in selected:
            if target_voltage is None:
                continue
            curr_voltage = _circuit_voltage_volts(circ)
            if curr_voltage is None or abs(curr_voltage - target_voltage) > 5:
                circ_num = _get_circuit_number(circ) or "?"
                curr_text = "{:.0f}V".format(curr_voltage) if curr_voltage is not None else "?"
                voltage_changes.append("C{}: {} -> {:.0f}V".format(circ_num, curr_text, target_voltage))

        extra_note = ""
        if voltage_changes:
            extra_note = (
                "\n\nConversão de tensão em tentativa experimental:\n{}\n\n"
                "Se o Revit criar novamente em 127V ou recusar o quadro, o circuito antigo será preservado."
            ).format("\n".join(voltage_changes))

        confirma = forms.alert(
            "Transferir {} circuito(s) para '{}'?\n\n"
            "Os circuitos serão desconectados do quadro atual\n"
            "e reconectados no destino com os polos escolhidos.{}".format(
                len(selected), dest_name, extra_note),
            title="Transferir Circuitos",
            yes=True, no=True
        )
        if not confirma:
            return

        self.Close()

        # ── Executar transferência ──
        dbg.section("Transferir Circuitos")
        dbg.info("Destino: {} ({} circuitos)".format(dest_name, len(selected)))

        sucessos = 0
        erros = []

        from Autodesk.Revit.DB import IFailuresPreprocessor, FailureSeverity, FailureProcessingResult
        
        class TransferFailureLogger(IFailuresPreprocessor):
            def __init__(self):
                self.messages = []

            def PreprocessFailures(self, failuresAccessor):
                try:
                    failures = list(failuresAccessor.GetFailureMessages())
                    if not failures:
                        return FailureProcessingResult.Continue
                    for f in failures:
                        try:
                            desc = f.GetDescriptionText()
                            sev = f.GetSeverity()
                            self.messages.append("{}: {}".format(sev, desc))
                            if sev == FailureSeverity.Warning:
                                failuresAccessor.DeleteWarning(f)
                        except Exception as inner_e:
                            self.messages.append("Falha ao ler erro: {}".format(inner_e))
                except Exception as out_e:
                    self.messages.append("Falha no preprocessor: {}".format(out_e))
                return FailureProcessingResult.Continue

        failure_logger = TransferFailureLogger()
        t = Transaction(doc, "Transferir Circuitos")
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(failure_logger)
        t.SetFailureHandlingOptions(opts)
        t.Start()

        try:
            for circ, target_poles, target_voltage in selected:
                circ_num = ""
                try:
                    circ_num = circ.CircuitNumber
                except Exception:
                    pass

                sub = SubTransaction(doc)
                sub.Start()
                try:
                    ok, msg = transfer_one_circuit(circ, dest_panel, target_poles, target_voltage)
                except Exception as e:
                    ok, msg = False, str(e)

                if ok:
                    sub.Commit()
                    sucessos += 1
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                else:
                    try:
                        sub.RollBack()
                    except Exception:
                        pass
                    erros.append((circ_num, msg))
                    dbg.error("C{}: {}".format(circ_num, msg))

            status = t.Commit()
            if status != TransactionStatus.Committed and failure_logger.messages:
                erros.append(("GERAL", " | ".join(failure_logger.messages[:3])))
            if status != TransactionStatus.Committed:
                erros.append(("GERAL", "Revit cancelou a transação (Rollback silencioso). Status: {}".format(status)))
                dbg.error("Transação principal falhou: {}".format(status))
                sucessos = 0
        except Exception as e:
            try:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
            except Exception:
                pass
            erros.append(("GERAL", str(e)))
            dbg.error("Exceção geral: {}".format(e))

        # ── Resumo ──
        dbg.section("Resultado")
        dbg.info("Sucesso: {}  |  Falhas: {}".format(sucessos, len(erros)))

        msg = "Transferência Concluída!\n\n"
        msg += "Circuitos movidos com sucesso: {}\n".format(sucessos)

        if erros:
            msg += "\nFalhas: {}\n".format(len(erros))
            for num, err in erros:
                msg += " • C{}: {}\n".format(num, err[:120])

        forms.alert(msg, title="Resumo da Transferência")


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    win = TransferCircuitsWindow(os.path.join(_BUNDLE_DIR, 'ui.xaml'))
    win.ShowDialog()
