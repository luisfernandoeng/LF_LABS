# coding: utf-8
"""LF Electrical - Industrial
Criação de circuitos e interruptores (Industrial)"""

__title__ = "Industrial"
__author__ = "Luís Fernando"

from pyrevit import forms
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI.Selection import ObjectType
from collections import OrderedDict
import traceback
import re
from Autodesk.Revit.Exceptions import OperationCanceledException

import lf_electrical_core
from lf_electrical_core import (
    doc, uidoc, get_current_panel, set_current_panel, get_panel_name, set_param,
    ConnectorDomainFilter, get_valid_electrical_elements,
    is_element_connected_to_panel, ensure_element_is_free, get_room_name, get_family_name,
    configure_panel, PanelFilter, CategoryFilter, call_queda_tensao,
    VALID_SWITCH_LETTERS, _get_switch_label, load_config, dbg,
    suppress_elec_dialog
)

VOLTAGE_FACTOR = 10.7639104167
LOAD_NAME_PARAMS = ["Nome da carga", "Load Name", "Nome de carga", "Carga", "Load"]

INDUSTRIAL_PANEL_RULES = OrderedDict([
    ("127V monofasico", {
        "poles": 1,
        "voltage": 127,
        "patterns": ("127",),
        "reject_patterns": ("220/127", "380/220"),
        "circuit_options": OrderedDict([
            ("1 Polo - 127V", {"poles": 1, "voltage": 127}),
        ]),
    }),
    ("220/127V bifasico", {
        "poles": 2,
        "voltage": 220,
        "patterns": ("220/127",),
        "phase_hint": 2,
        "circuit_options": OrderedDict([
            ("1 Polo - 127V", {"poles": 1, "voltage": 127}),
            ("2 Polos - 220V", {"poles": 2, "voltage": 220}),
        ]),
    }),
    ("220/127V trifasico", {
        "poles": 3,
        "voltage": 220,
        "patterns": ("220/127",),
        "phase_hint": 3,
        "circuit_options": OrderedDict([
            ("1 Polo - 127V", {"poles": 1, "voltage": 127}),
            ("2 Polos - 220V", {"poles": 2, "voltage": 220}),
        ]),
    }),
    ("380/220V trifasico", {
        "poles": 3,
        "voltage": 380,
        "patterns": ("380/220",),
        "phase_hint": 3,
        "circuit_options": OrderedDict([
            ("1 Polo - 220V", {"poles": 1, "voltage": 220}),
            ("3 Polos - 380V", {"poles": 3, "voltage": 380}),
        ]),
    }),
])

def _norm_name(value):
    return (value or "").lower().replace(" ", "")

def _name_has_voltages(name, voltages):
    nums = set(re.findall(r"\d+", name or ""))
    return all(str(v) in nums for v in voltages)

def _dist_system_has_neutral_ground_name(system):
    name = _norm_name(_get_dist_system_name(system))
    has_neutral = any(token in name for token in ["neutro", "neutral", "+n", "-n", "fn", "f+n"])
    has_ground = any(token in name for token in ["terra", "ground", "aterr", "gnd", "+t", "-t", "pe", "+pe", "-pe"])
    return has_neutral and has_ground

def _dist_system_matches_phase_name(system, poles):
    name = _norm_name(_get_dist_system_name(system))
    if poles == 1:
        return any(token in name for token in ["mono", "monofas", "1f", "f+n"]) and not any(token in name for token in ["2f", "3f", "bif", "trif"])
    if poles == 2:
        return any(token in name for token in ["bif", "bifas", "2f"])
    if poles == 3:
        return any(token in name for token in ["trif", "trifas", "3f"])
    return False

def _get_dist_system_name(system):
    try:
        return system.Name
    except Exception:
        p = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        return p.AsString() if p else ""

def _get_bip(name):
    try:
        return getattr(BuiltInParameter, name)
    except Exception:
        return None

def _get_voltage_value_from_dist_system(system, bip_name):
    try:
        bip = _get_bip(bip_name)
        if bip is None:
            return None
        vp = system.get_Parameter(bip)
        if not vp or vp.AsElementId() == ElementId.InvalidElementId:
            return None
        vtype = doc.GetElement(vp.AsElementId())
        if not vtype:
            return None
        param_ids = []
        try: param_ids.append(BuiltInParameter.RBS_ELEC_VOLTAGE_VALUE)
        except Exception: pass
        param_ids.extend([
            BuiltInParameter.RBS_ELEC_VOLTAGE_MIN_PARAM,
            BuiltInParameter.RBS_ELEC_VOLTAGE_MAX_PARAM
        ])
        for param_id in param_ids:
            try:
                p = vtype.get_Parameter(param_id)
                if p and p.HasValue:
                    return int(round(p.AsDouble() / VOLTAGE_FACTOR))
            except Exception:
                pass
    except Exception:
        pass
    return None

def _dist_system_matches_rule(system, rule):
    name = _norm_name(_get_dist_system_name(system))
    if not _dist_system_has_neutral_ground_name(system):
        return False
    if not _dist_system_matches_phase_name(system, rule["poles"]):
        return False

    for reject in rule.get("reject_patterns", ()):
        if _norm_name(reject) in name:
            return False
    has_pattern = any(_norm_name(pattern) in name for pattern in rule.get("patterns", ()))
    if not has_pattern:
        if rule["voltage"] == 380:
            has_pattern = _name_has_voltages(name, (380, 220))
        elif rule["voltage"] == 220:
            has_pattern = _name_has_voltages(name, (220, 127))
        elif rule["voltage"] == 127:
            has_pattern = _name_has_voltages(name, (127,)) and not _name_has_voltages(name, (220,))
    if not has_pattern:
        return False

    hint = rule.get("phase_hint")
    if hint:
        try:
            phase_bip = _get_bip("RBS_ELEC_DISTRIBUTION_SYS_PHASE_PARAM")
            p = system.get_Parameter(phase_bip) if phase_bip is not None else None
            if p and p.HasValue and p.AsInteger() != hint:
                return False
        except Exception:
            pass

    ll = _get_voltage_value_from_dist_system(system, "RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_L_PARAM")
    lg = _get_voltage_value_from_dist_system(system, "RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_G_PARAM")
    expected = rule["voltage"]

    if "380/220" in name or _name_has_voltages(name, (380, 220)):
        return expected == 380 and (ll is None or ll == 380) and (lg is None or lg == 220)
    if "220/127" in name or "127/220" in name or _name_has_voltages(name, (220, 127)):
        return expected == 220 and (ll is None or ll == 220) and (lg is None or lg == 127)
    if expected == 127:
        return (ll == 127 or lg == 127 or ll is None and lg is None)
    return True

def _collect_industrial_distribution_options_for_rule(rule):
    systems = list(FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements())
    options = OrderedDict()
    matches = []
    for system in systems:
        if _dist_system_matches_rule(system, rule):
            matches.append(system)
    for system in sorted(matches, key=lambda s: _get_dist_system_name(s)):
        sys_name = _get_dist_system_name(system)
        options[sys_name] = system.Id
    return options

def _set_panel_voltage_and_poles(panel, rule):
    set_param(panel, ["N° de Fases", "Nº de Fases", "Número de Fases", "Número de polos", "Number of Poles", "Polos"], rule["poles"])
    set_param(panel, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], rule["voltage"])
    v_internal = float(rule["voltage"]) * VOLTAGE_FACTOR
    for bip, value in [(BuiltInParameter.RBS_ELEC_VOLTAGE, v_internal),
                       (BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES, float(rule["poles"]))]:
        try:
            p = panel.get_Parameter(bip)
            if p and not p.IsReadOnly:
                p.Set(value)
        except Exception:
            pass
    try:
        if hasattr(panel, "MEPModel") and panel.MEPModel:
            cm = getattr(panel.MEPModel, "ConnectorManager", None)
            if cm:
                for c in cm.Connectors:
                    if c.Domain == Domain.DomainElectrical:
                        try: c.Voltage = v_internal
                        except Exception: pass
                        try: c.Poles = rule["poles"]
                        except Exception: pass
    except Exception:
        pass

def _get_param_number(elem, names):
    for name in names:
        try:
            p = elem.get_Parameter(name) if isinstance(name, BuiltInParameter) else elem.LookupParameter(name)
            if p and p.HasValue:
                try:
                    val = p.AsDouble()
                    if val:
                        if isinstance(name, BuiltInParameter) and name == BuiltInParameter.RBS_ELEC_VOLTAGE:
                            return int(round(val / VOLTAGE_FACTOR))
                        return int(round(val))
                except Exception:
                    pass
                try:
                    text = p.AsValueString() or p.AsString() or ""
                    match = re.search(r"\d+", text)
                    if match:
                        return int(match.group())
                except Exception:
                    pass
        except Exception:
            pass
    return None

def _get_panel_rule_from_values(panel):
    poles = _get_param_number(panel, [
        BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
        "N° de Fases", "Nº de Fases", "Número de Fases", "Número de polos", "Number of Poles", "Polos"
    ])
    voltage = _get_param_number(panel, [
        BuiltInParameter.RBS_ELEC_VOLTAGE,
        "Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"
    ])
    for label, rule in INDUSTRIAL_PANEL_RULES.items():
        if poles == rule["poles"] and voltage == rule["voltage"]:
            return label, rule
    return None, None

def _get_panel_rule_from_distribution(panel):
    try:
        p = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and p.HasValue and p.AsElementId() != ElementId.InvalidElementId:
            system = doc.GetElement(p.AsElementId())
            if system:
                for label, rule in INDUSTRIAL_PANEL_RULES.items():
                    if _dist_system_matches_rule(system, rule):
                        return label, rule
    except Exception:
        pass
    return _get_panel_rule_from_values(panel)

def prompt_industrial_phase_voltage(panel):
    label, rule = _get_panel_rule_from_distribution(panel)
    if not rule:
        forms.alert("Nao consegui identificar o sistema industrial deste quadro. Configure o quadro novamente antes de criar circuitos.", title="Quadro sem regra industrial")
        return None
    options = rule["circuit_options"]
    escolha = forms.CommandSwitchWindow.show(
        options.keys(),
        message="Circuitos permitidos para {}:".format(label),
        title="Compatibilidade do Quadro"
    )
    if escolha:
        return options[escolha]
    return None

def prompt_industrial_tomada_voltage(panel):
    label, rule = _get_panel_rule_from_distribution(panel)
    if not rule:
        forms.alert("Nao consegui identificar o sistema industrial deste quadro. Configure o quadro novamente antes de criar circuitos.", title="Quadro sem regra industrial")
        return None

    options = rule["circuit_options"]
    if not options:
        forms.alert("Nenhuma tensao permitida para tomadas neste quadro.", title="Tomadas")
        return None

    if len(options) == 1:
        key = list(options.keys())[0]
        dbg.step('Tomadas: unica opcao permitida {}'.format(key))
        return options[key]

    escolha = forms.CommandSwitchWindow.show(
        options.keys(),
        message="Circuitos de tomadas permitidos para {}:".format(label),
        title="Compatibilidade do Quadro"
    )
    if escolha:
        return options[escolha]
    return None

def _get_panel_distribution_id(panel):
    try:
        p = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and p.HasValue and p.AsElementId() != ElementId.InvalidElementId:
            return p.AsElementId()
    except Exception:
        pass
    return None

def _set_param_value(p, value=None, value_string=None):
    if p is None or p.IsReadOnly:
        return False
    try:
        if value_string:
            try:
                if p.SetValueString(value_string):
                    return True
            except Exception:
                pass
        st = p.StorageType
        if st == StorageType.Integer:
            p.Set(int(round(float(value))))
            return True
        if st == StorageType.Double:
            p.Set(float(value))
            return True
        if st == StorageType.ElementId:
            if isinstance(value, ElementId):
                p.Set(value)
                return True
            p.Set(ElementId(int(value)))
            return True
        if st == StorageType.String:
            p.Set(str(value))
            return True
    except Exception:
        pass
    return False

def _apply_panel_distribution(elem, panel):
    dist_id = _get_panel_distribution_id(panel)
    if not dist_id:
        return False
    for name in ["Sistema de distribuição", "Distribution System", "Sistema de Distribuição"]:
        try:
            p = elem.LookupParameter(name)
            if p and not p.IsReadOnly:
                p.Set(dist_id)
                dbg.ok('Sistema do quadro aplicado ao elemento Id={}'.format(elem.Id.IntegerValue))
                return True
        except Exception as ex:
            dbg.warn('Falha ao aplicar sistema do quadro em Id={}: {}'.format(elem.Id.IntegerValue, ex))
    return False

def _get_industrial_config_targets(elem):
    targets = []
    try:
        parent = getattr(elem, "SuperComponent", None)
        if parent and parent.Id != elem.Id:
            targets.append(parent)
            dbg.step('Config alvo externo via SuperComponent: Id={} Cat={}'.format(
                parent.Id.IntegerValue,
                parent.Category.Name if parent.Category else "?"))
    except Exception as ex:
        dbg.step('Sem SuperComponent configuravel para Id={}: {}'.format(elem.Id.IntegerValue, ex))
    targets.append(elem)
    return targets

def _get_first_electrical_connector(elem):
    try:
        mgr = None
        mep = getattr(elem, "MEPModel", None)
        if mep:
            mgr = getattr(mep, "ConnectorManager", None)
        if not mgr:
            mgr = getattr(elem, "ConnectorManager", None)
        if mgr:
            for c in mgr.Connectors:
                if c.Domain == Domain.DomainElectrical:
                    return c
    except Exception as ex:
        dbg.warn('Falha ao obter conector eletrico Id={}: {}'.format(elem.Id.IntegerValue, ex))
    return None

def _configure_industrial_target_for_circuit(elem, phase_config, panel):
    voltage_internal = float(phase_config["voltage"]) * VOLTAGE_FACTOR
    voltage_text = str(int(round(phase_config["voltage"]))) + " V"
    poles = int(phase_config["poles"])
    dist_id = _get_panel_distribution_id(panel)

    try:
        _set_param_value(elem.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE), voltage_internal, voltage_text)
    except Exception:
        pass
    try:
        _set_param_value(elem.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES), poles)
    except Exception:
        pass

    try:
        mgr = None
        mep = getattr(elem, "MEPModel", None)
        if mep:
            mgr = getattr(mep, "ConnectorManager", None)
        if not mgr:
            mgr = getattr(elem, "ConnectorManager", None)
        if mgr:
            for c in mgr.Connectors:
                if c.Domain == Domain.DomainElectrical:
                    try:
                        c.Voltage = voltage_internal
                        dbg.ok('Conector Id={} tensao setada para {}'.format(elem.Id.IntegerValue, voltage_text))
                    except Exception as ex:
                        dbg.warn('Conector Id={} nao aceitou tensao: {}'.format(elem.Id.IntegerValue, ex))
                    try:
                        c.Poles = poles
                        dbg.ok('Conector Id={} polos setado para {}'.format(elem.Id.IntegerValue, poles))
                    except Exception as ex:
                        dbg.warn('Conector Id={} nao aceitou polos: {}'.format(elem.Id.IntegerValue, ex))
                    try:
                        read_v = c.Voltage / VOLTAGE_FACTOR
                        read_p = c.Poles
                        dbg.step('Conector Id={} apos set: {:.0f}V / {} polo(s)'.format(
                            elem.Id.IntegerValue, read_v, read_p))
                    except Exception:
                        pass
    except Exception:
        pass

    if dist_id:
        for name in [u"Sistema de distribuição", "Distribution System", u"Sistema de Distribuição"]:
            try:
                if _set_param_value(elem.LookupParameter(name), dist_id):
                    break
            except Exception:
                pass

    p_names_poles = [u"Pólos", "Polos", u"Número de Polos", "Poles",
                     u"Número de polos", "Fases", u"N° de Fases",
                     u"Nº de Fases", u"Número de Fases",
                     u"Número de polos", "Number of Poles"]
    p_names_volt = [u"Tensão", u"Tensão Numérica", "Voltagem", "Voltage",
                    u"Tensão (V)", "Volts"]

    for p_name in p_names_poles:
        try:
            if _set_param_value(elem.LookupParameter(p_name), poles):
                break
        except Exception:
            pass
    for p_name in p_names_volt:
        try:
            if _set_param_value(elem.LookupParameter(p_name), voltage_internal, voltage_text):
                break
        except Exception:
            pass

    try:
        elem_type = doc.GetElement(elem.GetTypeId()) if hasattr(elem, "GetTypeId") else None
        if elem_type:
            dbg.step('Tentando modificar parametros no TYPE: {}'.format(elem_type.Id.IntegerValue))
            for p_name in p_names_poles:
                p = elem_type.LookupParameter(p_name)
                if p:
                    dbg.step('TYPE param [{}] RO={}'.format(p_name, p.IsReadOnly))
                    try:
                        if not p.IsReadOnly:
                            p.Set(poles)
                            dbg.ok('TYPE param [{}] alterado para {}'.format(p_name, poles))
                            break
                    except Exception as ex:
                        dbg.warn('Erro ao setar TYPE param [{}]: {}'.format(p_name, ex))

            for p_name in p_names_volt:
                p = elem_type.LookupParameter(p_name)
                if p:
                    dbg.step('TYPE param [{}] RO={}'.format(p_name, p.IsReadOnly))
                    try:
                        if not p.IsReadOnly:
                            try: p.Set(voltage_internal)
                            except: p.Set(voltage_text)
                            dbg.ok('TYPE param [{}] alterado para {}'.format(p_name, voltage_text))
                            break
                    except Exception as ex:
                        dbg.warn('Erro ao setar TYPE param [{}]: {}'.format(p_name, ex))
    except Exception as ex:
        dbg.warn('Erro ao acessar TYPE: {}'.format(ex))

def _configure_industrial_element_for_circuit(elem, phase_config, panel):
    for target in _get_industrial_config_targets(elem):
        try:
            dbg.step('Configurando alvo Id={} para {} polo(s), {}V'.format(
                target.Id.IntegerValue, phase_config["poles"], phase_config["voltage"]))
            _configure_industrial_target_for_circuit(target, phase_config, panel)
        except Exception as ex:
            dbg.warn('Falha ao configurar alvo Id={}: {}'.format(target.Id.IntegerValue, ex))

def _parse_first_int(value):
    if value is None:
        return None
    try:
        return int(round(float(value)))
    except Exception:
        pass
    try:
        m = re.search(r"\d+", str(value))
        if m:
            return int(m.group(0))
    except Exception:
        pass
    return None

def _param_display_value(param):
    if not param:
        return None
    for getter in ("AsValueString", "AsString"):
        try:
            val = getattr(param, getter)()
            if val not in (None, ""):
                return val
        except Exception:
            pass
    try:
        return param.AsInteger()
    except Exception:
        pass
    try:
        return param.AsDouble()
    except Exception:
        pass
    return None

def _read_poles_param(elem):
    names = [
        u"Número de polos", u"Numero de polos", u"Número de Polos",
        u"Número de Fases", u"Numero de Fases", u"N° de Fases",
        u"Nº de Fases", "Number of Poles", "Poles", "Polos",
        u"Pólos", "Fases"
    ]
    try:
        bip = elem.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
        val = _parse_first_int(_param_display_value(bip))
        if val:
            return val, getattr(bip, "IsReadOnly", False), "RBS_ELEC_NUMBER_OF_POLES"
    except Exception:
        pass
    for name in names:
        try:
            p = elem.LookupParameter(name)
            val = _parse_first_int(_param_display_value(p))
            if val:
                return val, getattr(p, "IsReadOnly", False), name
        except Exception:
            pass
    return None, False, None

def _assert_elements_match_phase_config(elems, phase_config):
    expected_poles = int(phase_config["poles"])
    bad = []
    for elem in elems:
        val, ro, pname = _read_poles_param(elem)
        if val and val != expected_poles:
            bad.append((elem.Id.IntegerValue, val, pname or "parametro", ro))
    if not bad:
        return
    details = []
    for eid, val, pname, ro in bad:
        details.append("Id {}: {} = {}{}".format(
            eid, pname, val, " (somente leitura)" if ro else ""))
    msg = (
        "A familia/tipo ainda esta configurada como {} polo(s), mas voce escolheu {} polo(s).\n"
        "O Revit criaria o circuito com a configuracao da familia e o quadro recusaria a ligacao.\n\n{}"
    ).format(bad[0][1], expected_poles, "\n".join(details))
    dbg.warn(msg)
    return False

def _circuit_panel_matches(circuit, panel):
    try:
        base = circuit.BaseEquipment
        if base and base.Id == panel.Id:
            return True
    except Exception:
        pass
    try:
        panel_name = get_panel_name(panel)
        for pname in ["Painel", "Panel"]:
            p = circuit.LookupParameter(pname)
            if p and p.HasValue:
                val = p.AsString() or p.AsValueString() or ""
                if val == panel_name:
                    return True
    except Exception:
        pass
    return False

def _force_circuit_distribution(circuit, panel):
    """Tenta forcar o sistema de distribuicao do circuito para coincidir com o do quadro."""
    dist_id = _get_panel_distribution_id(panel)
    if not dist_id:
        dbg.warn('Panel nao tem sistema de distribuicao')
        return False
    for name in [u"Sistema de distribuição", "Distribution System", u"Sistema de Distribuição"]:
        try:
            p = circuit.LookupParameter(name)
            if p:
                dbg.step('Circuit dist [{}] RO={} ST={}'.format(name, p.IsReadOnly, p.StorageType))
                if not p.IsReadOnly:
                    p.Set(dist_id)
                    dbg.ok('Distribution system forced via [{}]'.format(name))
                    return True
        except Exception as ex:
            dbg.step('Dist [{}] failed: {}'.format(name, ex))
    try:
        for p in circuit.Parameters:
            if p.IsReadOnly:
                continue
            pdef = p.Definition
            if not pdef:
                continue
            pn = (pdef.Name or "").lower()
            if any(kw in pn for kw in ['distribu', 'sistema', 'system']):
                dbg.step('Writable dist param: [{}] ST={}'.format(pdef.Name, p.StorageType))
                try:
                    if p.StorageType == StorageType.ElementId:
                        p.Set(dist_id)
                        dbg.ok('Dist set via [{}]'.format(pdef.Name))
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    dbg.warn('Nao foi possivel forcar dist system no circuito')
    return False

def _configure_circuit_for_phase(circuit, panel, phase_config):
    if not phase_config:
        return
    set_param(circuit, [u"N° de Fases", u"Nº de Fases", u"Número de Fases", u"Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
    set_param(circuit, [u"Tensão (V)", u"Tensão", "Voltage", "Voltagem", u"Tensão Nominal", "Volts"], phase_config["voltage"])
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
        if p and not p.IsReadOnly:
            p.Set(float(phase_config["poles"]))
            dbg.ok('Circuit BIP poles={}'.format(phase_config["poles"]))
        else:
            dbg.step('Circuit BIP poles RO={}'.format(p.IsReadOnly if p else 'N/A'))
    except Exception as ex:
        dbg.step('Circuit BIP poles err: {}'.format(ex))
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
        if p and not p.IsReadOnly:
            p.Set(float(phase_config["voltage"]) * VOLTAGE_FACTOR)
            dbg.ok('Circuit BIP voltage={}'.format(phase_config["voltage"]))
        else:
            dbg.step('Circuit BIP voltage RO={}'.format(p.IsReadOnly if p else 'N/A'))
    except Exception as ex:
        dbg.step('Circuit BIP voltage err: {}'.format(ex))
    _apply_panel_distribution(circuit, panel)

def _select_panel_checked(circuit, panel, phase_config=None):
    panel_name = get_panel_name(panel)
    _configure_circuit_for_phase(circuit, panel, phase_config)
    _force_circuit_distribution(circuit, panel)
    try:
        doc.Regenerate()
    except Exception:
        pass
    dbg.step('Conectando circuito Id={} ao quadro Id={} ({})'.format(
        circuit.Id.IntegerValue, panel.Id.IntegerValue, panel_name))
    try:
        circuit.SelectPanel(panel)
    except Exception as ex:
        if "do not match" not in str(ex).lower():
            raise
        dbg.warn('SelectPanel dist mismatch: {}'.format(ex))
        # Dump circuit params for diagnostics
        try:
            for p in circuit.Parameters:
                pdef = p.Definition
                if pdef:
                    pn = (pdef.Name or "").lower()
                    if any(kw in pn for kw in ['distribu', 'sistema', 'system', 'polo',
                                                'fase', 'tens', 'volt', 'painel', 'panel']):
                        val = ""
                        try:
                            val = p.AsValueString() or p.AsString() or str(p.AsDouble())
                        except Exception:
                            try:
                                val = str(p.AsElementId().IntegerValue)
                            except Exception:
                                val = "?"
                        dbg.step('  CIRC [{}] RO={} val={}'.format(pdef.Name, p.IsReadOnly, val))
        except Exception:
            pass
        try:
            doc.Regenerate()
            circuit.SelectPanel(panel)
            dbg.ok('SelectPanel OK no retry')
        except Exception as ex2:
            dbg.fail('SelectPanel retry falhou: {}'.format(ex2))
            raise
    try:
        doc.Regenerate()
    except Exception as ex:
        dbg.warn('Regenerate apos SelectPanel falhou: {}'.format(ex))
    if not _circuit_panel_matches(circuit, panel):
        raise Exception("Circuito criado, mas o Revit nao conectou ao quadro '{}'.".format(panel_name))
    dbg.ok('Circuito Id={} conectado ao quadro {}'.format(circuit.Id.IntegerValue, panel_name))

def _get_assigned_panel_name(elem):
    try:
        for pname in ["Painel", "Panel"]:
            p = elem.LookupParameter(pname)
            if p and p.HasValue:
                val = p.AsString() or p.AsValueString() or ""
                if val and val.strip():
                    return val.strip()
    except Exception:
        pass
    return ""

def _should_skip_connected_element(elem, panel, reconnect_state):
    assigned_panel = _get_assigned_panel_name(elem)
    if not assigned_panel:
        return False

    target_panel = get_panel_name(panel)
    if assigned_panel == target_panel:
        dbg.step('Id={} ja esta ligado ao quadro selecionado "{}" -- pulado'.format(
            elem.Id.IntegerValue, target_panel))
        return True

    if reconnect_state.get("choice") is None:
        reconnect_state["choice"] = forms.alert(
            "Existem elementos ligados ao quadro '{}'.\n\nDeseja remover o circuito existente e recriar no quadro '{}' ?".format(
                assigned_panel, target_panel),
            title="Religar circuitos",
            yes=True,
            no=True
        )

    if reconnect_state.get("choice"):
        dbg.step('Id={} estava no painel "{}" e sera religado para "{}"'.format(
            elem.Id.IntegerValue, assigned_panel, target_panel))
        return False

    dbg.step('Id={} ligado ao painel "{}" -- pulado'.format(elem.Id.IntegerValue, assigned_panel))
    return True

def _ask_circuit_description(default_text=""):
    value = forms.ask_for_string(
        default=default_text,
        prompt="Descricao do circuito:\nDeixe em branco para usar a descricao automatica.",
        title="Descricao do Circuito"
    )
    if value is None:
        return None
    return value.strip()

def _circuit_description(custom_text, fallback_text):
    return custom_text if custom_text else fallback_text

def _read_voltage_param(elem):
    for name in [u"Tensão (V)", u"Tensão", "Voltage", "Voltagem", u"Tensão Nominal", "Volts"]:
        try:
            p = elem.LookupParameter(name)
            if p and p.HasValue:
                text = p.AsValueString() or p.AsString() or ""
                match = re.search(r"[-+]?\d*\.\d+|\d+", str(text).replace(",", "."))
                if match:
                    return int(round(float(match.group()))), p.IsReadOnly, name
                try:
                    raw = p.AsDouble()
                    if raw:
                        value = raw / VOLTAGE_FACTOR if raw > 1000 else raw
                        return int(round(value)), p.IsReadOnly, name
                except Exception:
                    pass
        except Exception:
            pass
    return None, False, None

def _voltage_locked_incompatible(elem, target_voltage):
    current_voltage, readonly, pname = _read_voltage_param(elem)
    if current_voltage is None:
        return False, None
    if abs(current_voltage - int(target_voltage)) <= 2:
        return False, current_voltage
    if readonly:
        dbg.warn('Id={} tem {}={}V somente leitura; esperado {}V'.format(
            elem.Id.IntegerValue, pname, current_voltage, target_voltage))
        return True, current_voltage
    return False, current_voltage


def _get_element_wattage(elem):
    for pname in ["Potência Aparente (VA)", "Potência Aparente", "Apparent Load", "Potência", "Power", "Wattage", "Potencia Aparente", "Potencia", "Carga Aparente"]:
        p = elem.LookupParameter(pname)
        if p and p.HasValue:
            try:
                val_str = p.AsValueString()
                if val_str:
                    s = str(val_str).replace(',', '.')
                    match = re.search(r"[-+]?\d*\.\d+|\d+", s)
                    if match: return float(match.group())
            except Exception: pass
            try: return p.AsDouble()
            except Exception: pass
    try:
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            cm = elem.MEPModel.ConnectorManager
            if cm:
                for c in cm.Connectors:
                    if c.Domain == Domain.DomainElectrical:
                        try: return c.Voltage * c.Current
                        except Exception: pass
    except Exception: pass
    return 0.0

def _detect_amperage_from_family(elem):
    fname = get_family_name(elem).lower()
    tname = ""
    try:
        if hasattr(elem, 'Symbol') and elem.Symbol:
            tname = elem.Symbol.Name.lower() if hasattr(elem.Symbol, 'Name') else ""
    except Exception: pass
    full_name = fname + " " + tname
    if "20a" in full_name or "20 a" in full_name: return 20
    elif "10a" in full_name or "10 a" in full_name: return 10
    return 10

def configure_panel_industrial(panel):
    messages = []

    new_name = forms.ask_for_string(default=get_panel_name(panel), prompt="Nome do Quadro Industrial (ex: QGBT, CCM-01):", title="Quadro Industrial")
    if not new_name: return False, "Cancelado"

    with Transaction(doc, "Nome do Quadro Industrial") as t:
        t.Start()
        set_param(panel, ["Nome do painel", "Panel Name", "Mark"], new_name)
        t.Commit()
    messages.append("Nome: " + new_name)

    chosen_rule = forms.CommandSwitchWindow.show(
        INDUSTRIAL_PANEL_RULES.keys(),
        message="Defina primeiro a tensao e a quantidade de fases do quadro:",
        title="Regra do Quadro Industrial"
    )
    if not chosen_rule:
        return False, "Cancelado"

    rule = INDUSTRIAL_PANEL_RULES[chosen_rule]
    messages.append("Regra: " + chosen_rule)

    with Transaction(doc, "Tensao e Polos Quadro Industrial") as t:
        t.Start()
        _set_panel_voltage_and_poles(panel, rule)
        if set_param(panel, [u"Op\u00e7\u00e3o de numera\u00e7\u00e3o do circuito", "Circuit Numbering Option"], 0):
            messages.append("Nomenclatura: Por projeto")
        if set_param(panel, [u"Nomenclatura do circuito", "Circuit Naming"], "Por projeto"):
            messages.append("Nomenclatura do circuito: Por projeto")
        if set_param(panel, [u"N\u00famero do circuito", "Numero do circuito", "Circuit Number"], 0):
            messages.append("Numero do circuito: 0")
        t.Commit()

    dist_options = _collect_industrial_distribution_options_for_rule(rule)
    if not dist_options:
        forms.alert(
            "Nenhum sistema de distribuicao compativel encontrado.\n\n"
            "Crie/renomeie um sistema com fase(s) + neutro + terra para {}.".format(chosen_rule),
            title="Sistema de Distribuicao"
        )
        return False, "Sistema de distribuicao compativel nao encontrado"

    chosen_dist = forms.CommandSwitchWindow.show(
        dist_options.keys(),
        message="Sistema compativel com {} (fase(s) + neutro + terra):".format(chosen_rule),
        title="Sistema de Distribuicao"
    )
    if not chosen_dist:
        return False, "Cancelado"

    sys_id = dist_options[chosen_dist]
    with Transaction(doc, "Sistema Quadro Industrial") as t:
        t.Start()
        p = panel.LookupParameter(u"Sistema de distribui\u00e7\u00e3o") or panel.LookupParameter("Distribution System")
        if p and not p.IsReadOnly and sys_id:
            try:
                p.Set(sys_id)
                messages.append("Sistema: " + chosen_dist)
            except Exception as e:
                t.RollBack()
                forms.alert(
                    u"Sistema selecionado e incompativel com este quadro.\n\n"
                    u"Detalhe: {}".format(e),
                    title="Sistema Invalido"
                )
                return False, "Sistema de distribuicao invalido"
        messages.append("Quadro: {} polo(s), {}V".format(rule["poles"], rule["voltage"]))
        t.Commit()

    return True, "\n".join(messages)

def select_and_configure_panel_industrial():
    dbg.section('INDUSTRIAL - Selecionar Quadro')
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, PanelFilter(), "Selecione o QUADRO INDUSTRIAL")
        panel = doc.GetElement(ref.ElementId)
        dbg.elem_info(panel, 'Quadro selecionado')

        p_sys = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        has_sys = p_sys and p_sys.HasValue and p_sys.AsElementId() != ElementId.InvalidElementId
        rule_label, rule = _get_panel_rule_from_distribution(panel)
        has_rule = rule is not None
        
        dbg.step('Params validos: sys={}, regra={}'.format(has_sys, rule_label or "nenhuma"))

        if has_sys and has_rule:
            set_current_panel(panel.Id)
            dbg.ok('Quadro já configurado')
            forms.alert("Quadro Industrial Selecionado: " + get_panel_name(panel))
        else:
            dbg.step('Quadro precisa de configuração')
            success, msg = configure_panel_industrial(panel)
            if success:
                set_current_panel(panel.Id)
                dbg.ok('Quadro configurado')
                forms.alert("Quadro Industrial Configurado!\n" + msg)
            else:
                dbg.warn('Configuração cancelada')
                forms.alert(msg)
    except OperationCanceledException:
        dbg.step('ESC - Cancelado')
        pass
    except Exception as e:
        dbg.fail('Erro: {}'.format(e))
        forms.alert("Erro:\n" + str(e))

def create_ilum_circuit_industrial():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    phase_config = prompt_industrial_phase_voltage(panel)
    if not phase_config: return
    refs = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_LightingFixtures), "Selecione as luminarias (2400W cada circuito)")
    if not refs: return
    circuit_desc = _ask_circuit_description()
    if circuit_desc is None: return

    refs_list = list(refs)
    refs_list.reverse()

    batches = []
    current_batch = List[ElementId]()
    current_watts = 0.0
    current_rooms = set()
    skipped = []
    reconnect_state = {"choice": None}

    for r in refs_list:
        elem = doc.GetElement(r.ElementId)
        if _should_skip_connected_element(elem, panel, reconnect_state):
            skipped.append(str(r.ElementId.IntegerValue))
            continue
            
        watts = _get_element_wattage(elem)
        rm = get_room_name(elem)
        if current_watts + watts > 2400 and current_batch.Count > 0:
            batches.append((current_batch, current_watts, set(current_rooms)))
            current_batch = List[ElementId]()
            current_watts = 0.0
            current_rooms = set()
            
        current_batch.Add(r.ElementId)
        current_watts += watts
        if rm: current_rooms.add(rm)
        
    if current_batch.Count > 0:
        batches.append((current_batch, current_watts, set(current_rooms)))

    dbg.step('Batches: {} | Elementos totais: {}'.format(len(batches), sum(b[0].Count for b in batches)))
    created = 0
    with Transaction(doc, "Circuitos Ilum Industrial") as t:
        t.Start()
        for bi, (b_ids, b_watts, b_rooms) in enumerate(batches):
            dbg.step('Batch {}: {} elemento(s) {:.0f}W rooms={}'.format(bi+1, b_ids.Count, b_watts, list(b_rooms)))
            for eid in b_ids:
                elem = doc.GetElement(eid)
                ensure_element_is_free(elem)
                _configure_industrial_element_for_circuit(elem, phase_config, panel)

            doc.Regenerate()
            try:
                with suppress_elec_dialog():
                    circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                _select_panel_checked(circuit, panel, phase_config)
                room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Ilum Industrial"
                set_param(circuit, LOAD_NAME_PARAMS, _circuit_description(circuit_desc, "{} ({:.0f}W)".format(room_str, b_watts)))
                dbg.ok('Circuito ilum criado Id={}'.format(circuit.Id.IntegerValue))
                created += 1
            except Exception as ex:
                dbg.fail('ElectricalSystem.Create falhou batch {}: {}'.format(bi+1, ex))
                raise
        t.Commit()

    if created > 0: forms.toast("Sucesso! {} circuito(s) criado(s).".format(created), title="Industrial")

class _FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

def _unlock_family_electrical_params(unique_types):
    unlocked_count = 0
    p_names = ["P\xf3los", "Polos", "N\xfamero de Polos", "Poles",
               "N\xfamero de polos", "Fases", "N\xb0 de Fases",
               "N\xba de Fases", "N\xfamero de Fases", "Number of Poles",
               "Tens\xe3o", "Tens\xe3o Num\xe9rica", "Voltagem", "Voltage",
               "Tens\xe3o (V)", "Volts"]
    p_names_lower = [n.lower() for n in p_names]
    
    for elem_type in unique_types:
        family = getattr(elem_type, "Family", None)
        if not family or not family.IsEditable:
            continue
            
        # Check if it actually has locked params in the project first to save time
        needs_unlock = False
        for p_name in p_names:
            p = elem_type.LookupParameter(p_name)
            if p and p.IsReadOnly:
                needs_unlock = True
                break
        if not needs_unlock:
            continue
            
        dbg.step('Tentando DESTRAVAR formulas na familia: {}'.format(family.Name))
        try:
            fam_doc = doc.EditFamily(family)
            if not fam_doc:
                continue
            
            changed = False
            with Transaction(fam_doc, "Unlock Electrical Params") as t:
                t.Start()
                for p in fam_doc.FamilyManager.Parameters:
                    pdef_name = (p.Definition.Name or "").lower()
                    if pdef_name in p_names_lower and p.IsDeterminedByFormula:
                        try:
                            fam_doc.FamilyManager.SetFormula(p, "")
                            dbg.ok('Formula apagada do parametro: {}'.format(p.Definition.Name))
                            changed = True
                        except Exception as ex:
                            dbg.warn('Falha ao apagar formula de {}: {}'.format(p.Definition.Name, ex))
                t.Commit()
                
            if changed:
                fam_doc.LoadFamily(doc, _FamilyLoadOptions())
                dbg.ok('Familia {} recarregada e destravada!'.format(family.Name))
                unlocked_count += 1
            fam_doc.Close(False)
        except Exception as ex:
            dbg.warn('Erro ao editar familia {}: {}'.format(family.Name, ex))
            
    return unlocked_count

def create_tomada_circuit_industrial(selection_category=BuiltInCategory.OST_ElectricalFixtures,
                                     selection_prompt="Selecione as tomadas",
                                     transaction_name="Circuitos Tomadas Industrial"):
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    phase_config = prompt_industrial_tomada_voltage(panel)
    if not phase_config: return
    refs = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(selection_category), selection_prompt)
    if not refs: return
    circuit_desc = _ask_circuit_description()
    if circuit_desc is None: return

    refs_list = list(refs)
    refs_list.reverse()

    detected_amp = 10
    for r in refs_list:
        amp = _detect_amperage_from_family(doc.GetElement(r.ElementId))
        if amp > detected_amp: detected_amp = amp
    
    limite_w = 4000 if detected_amp >= 20 else 2200

    batches = []
    current_batch = List[ElementId]()
    current_watts = 0.0
    current_rooms = set()
    reconnect_state = {"choice": None}

    # Unlock read-only family parameters BEFORE starting the main transaction
    unique_types = set()
    for r in refs_list:
        host = doc.GetElement(r.ElementId)
        if hasattr(host, "GetTypeId"):
            elem_type = doc.GetElement(host.GetTypeId())
            if elem_type:
                unique_types.add(elem_type)
        parent = getattr(host, "SuperComponent", None)
        if parent and hasattr(parent, "GetTypeId"):
            parent_type = doc.GetElement(parent.GetTypeId())
            if parent_type:
                unique_types.add(parent_type)
        valid_pairs = get_valid_electrical_elements(host, (Domain.DomainElectrical,))
        for sub_id, _ in valid_pairs:
            sub = doc.GetElement(sub_id)
            if hasattr(sub, "GetTypeId"):
                sub_type = doc.GetElement(sub.GetTypeId())
                if sub_type:
                    unique_types.add(sub_type)
            parent_sub = getattr(sub, "SuperComponent", None)
            if parent_sub and hasattr(parent_sub, "GetTypeId"):
                parent_sub_type = doc.GetElement(parent_sub.GetTypeId())
                if parent_sub_type:
                    unique_types.add(parent_sub_type)
    
    if unique_types:
        dbg.step('Iniciando rotina de destravamento de familias ({} tipos encontrados)'.format(len(unique_types)))
        _unlock_family_electrical_params(unique_types)

    for r in refs_list:
        host = doc.GetElement(r.ElementId)
        dbg.elem_full(host, 'HOST')

        # Resolve os elementos com conector elétrico real.
        # Para famílias com sub-componentes (ex: FAST-ELE-CONDULETE-TOMADA), o host é
        # ConduitFitting sem conector elétrico — o conector fica no sub-componente.
        valid_pairs = get_valid_electrical_elements(host, (Domain.DomainElectrical,))
        dbg.step('Host Id={} -> {} par(es) com conector eletrico'.format(
            r.ElementId.IntegerValue, len(valid_pairs)))

        if not valid_pairs:
            dbg.warn('Sem conector eletrico Id={}'.format(r.ElementId.IntegerValue))
            continue

        watts = _get_element_wattage(host)
        rm = get_room_name(host)

        sub_ids_to_add = []
        for sub_id, conns in valid_pairs:
            sub = doc.GetElement(sub_id)
            dbg.elem_full(sub, 'SUB Id={}'.format(sub_id.IntegerValue))
            if _should_skip_connected_element(sub, panel, reconnect_state):
                dbg.step('  Sub Id={} ja ligado ao painel — pulado'.format(sub_id.IntegerValue))
                continue
            dbg.step('  Sub Id={} sera configurado para {} polo(s), {}V'.format(
                sub_id.IntegerValue, phase_config["poles"], phase_config["voltage"]))
            sub_ids_to_add.append(sub_id)

        dbg.step('sub_ids_para_batch: {}'.format([s.IntegerValue for s in sub_ids_to_add]))
        if not sub_ids_to_add:
            continue

        if current_watts + watts > limite_w and current_batch.Count > 0:
            batches.append((current_batch, current_watts, set(current_rooms)))
            current_batch = List[ElementId]()
            current_watts = 0.0
            current_rooms = set()

        for sub_id in sub_ids_to_add:
            current_batch.Add(sub_id)
        current_watts += watts
        if rm: current_rooms.add(rm)

    if current_batch.Count > 0:
        batches.append((current_batch, current_watts, set(current_rooms)))

    dbg.step('Batches tomada: {} | Limite {}W'.format(len(batches), limite_w))
    created = 0
    unassigned_ids = []
    with Transaction(doc, transaction_name) as t:
        t.Start()
        try:
            for bi, (b_ids, b_watts, b_rooms) in enumerate(batches):
                dbg.step('Batch {}: {} elemento(s) {:.0f}W'.format(bi+1, b_ids.Count, b_watts))
                for eid in b_ids:
                    elem = doc.GetElement(eid)
                    ensure_element_is_free(elem)
                    _configure_industrial_element_for_circuit(elem, phase_config, panel)
                dbg.step('Loop elementos batch {} concluido -- Regenerate...'.format(bi+1))
                try:
                    doc.Regenerate()
                    dbg.ok('doc.Regenerate OK')
                except Exception as ex_regen:
                    dbg.fail('doc.Regenerate falhou: {}'.format(ex_regen))
                    dbg.fail(traceback.format_exc())
                _assert_elements_match_phase_config([doc.GetElement(eid) for eid in b_ids], phase_config)
                ids_list = [eid.IntegerValue for eid in b_ids]
                dbg.step('ElectricalSystem.Create IDs: {}'.format(ids_list))

                circuit = None
                assigned_to_panel = False
                # Tentativa 1: criar com suppress_elec_dialog e SelectPanel
                sub = SubTransaction(doc)
                sub.Start()
                try:
                    with suppress_elec_dialog():
                        if b_ids.Count == 1:
                            first_elem = doc.GetElement(b_ids[0])
                            first_conn = _get_first_electrical_connector(first_elem)
                            if first_conn:
                                dbg.step('ElectricalSystem.Create por Connector Id={}'.format(first_elem.Id.IntegerValue))
                                circuit = ElectricalSystem.Create(first_conn, ElectricalSystemType.PowerCircuit)
                            else:
                                circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                        else:
                            circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                    dbg.ok('Create OK Id={}'.format(circuit.Id.IntegerValue))
                    _select_panel_checked(circuit, panel, phase_config)
                    sub.Commit()
                    dbg.ok('SubTransaction 1 COMMITTED')
                    assigned_to_panel = True
                except Exception as ex1:
                    dbg.warn('Tentativa 1 falhou: {}'.format(ex1))
                    try:
                        sub.RollBack()
                    except Exception:
                        pass
                    circuit = None

                    if "do not match" in str(ex1).lower():
                        # Tentativa 2: SEM suprimir dialog — Revit pergunta ao usuario
                        dbg.step('Tentativa 2: sem suppress_elec_dialog (dialog do Revit)')
                        sub2 = SubTransaction(doc)
                        sub2.Start()
                        try:
                            if b_ids.Count == 1:
                                first_elem = doc.GetElement(b_ids[0])
                                first_conn = _get_first_electrical_connector(first_elem)
                                if first_conn:
                                    dbg.step('ElectricalSystem.Create por Connector Id={} (sem suppress)'.format(first_elem.Id.IntegerValue))
                                    circuit = ElectricalSystem.Create(first_conn, ElectricalSystemType.PowerCircuit)
                                else:
                                    circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                            else:
                                circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                            dbg.ok('Create (sem suppress) OK Id={}'.format(circuit.Id.IntegerValue))
                            _select_panel_checked(circuit, panel, phase_config)
                            sub2.Commit()
                            dbg.ok('SubTransaction 2 COMMITTED')
                            assigned_to_panel = True
                        except Exception as ex2:
                            dbg.warn('Tentativa 2 falhou: {}'.format(ex2))
                            try:
                                sub2.RollBack()
                            except Exception:
                                pass
                            circuit = None

                            # Tentativa 3: criar circuito SEM atribuir ao quadro
                            dbg.step('Tentativa 3: criar circuito SEM atribuir quadro')
                            sub3 = SubTransaction(doc)
                            sub3.Start()
                            try:
                                if b_ids.Count == 1:
                                    first_elem = doc.GetElement(b_ids[0])
                                    first_conn = _get_first_electrical_connector(first_elem)
                                    if first_conn:
                                        circuit = ElectricalSystem.Create(first_conn, ElectricalSystemType.PowerCircuit)
                                    else:
                                        circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                                else:
                                    circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                                dbg.ok('Create (sem panel) OK Id={}'.format(circuit.Id.IntegerValue))
                                dbg.warn('Circuito Id={} criado SEM quadro (atribuir manualmente)'.format(circuit.Id.IntegerValue))
                                sub3.Commit()
                                assigned_to_panel = False
                            except Exception as ex3:
                                dbg.fail('Tentativa 3 falhou: {}'.format(ex3))
                                try:
                                    sub3.RollBack()
                                except Exception:
                                    pass
                                circuit = None
                    else:
                        raise

                if circuit:
                    room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Tomada {}A".format(detected_amp)
                    set_param(circuit, LOAD_NAME_PARAMS, _circuit_description(circuit_desc, "{} ({:.0f}W)".format(room_str, b_watts)))
                    dbg.ok('Circuito tomada criado Id={}'.format(circuit.Id.IntegerValue))
                    created += 1
                    if not assigned_to_panel:
                        unassigned_ids.append(circuit.Id.IntegerValue)

            t.Commit()
            dbg.ok('Transaction COMMITTED')
        except Exception as ex_outer:
            dbg.fail('EXCECAO NAO CAPTURADA na Transaction: {}'.format(ex_outer))
            dbg.fail(traceback.format_exc())
            try: t.RollBack()
            except Exception: pass
            forms.alert(u"Circuito de tomada nao criado:\n\n{}".format(ex_outer))

    if created > 0:
        if unassigned_ids:
            forms.alert(
                u"Criado(s) {} circuito(s), porém {} NÃO foi(ram) conectado(s) ao quadro.\n\n"
                u"IDs sem quadro: {}\n\n"
                u"Motivo: A família tem tensão/polos travados no conector que não "
                u"são compatíveis com o sistema de distribuição do quadro selecionado.\n\n"
                u"Solução: Selecione o(s) circuito(s) no projeto e atribua o quadro "
                u"manualmente via 'Selecionar Painel' ou edite a família para "
                u"ajustar a tensão do conector.".format(
                    created, len(unassigned_ids),
                    ", ".join(str(x) for x in unassigned_ids)),
                title="Circuitos criados com pendencia"
            )
        else:
            forms.toast("Sucesso! {} circuito(s) de tomada criado(s).".format(created), title="Industrial")

def create_tomada_conduit_circuit_industrial():
    return create_tomada_circuit_industrial(
        BuiltInCategory.OST_ConduitFitting,
        "Selecione as conexoes do conduite/caixas com tomadas",
        "Circuitos Tomadas por Conexao do Conduite"
    )

def name_switch_industrial():
    dbg.section('INDUSTRIAL - Nomear Interruptores')
    start_str = forms.ask_for_string(
        default="a",
        prompt="Letra inicial (minúscula):\n(o e s são puladas automaticamente)\n(Após z: aa, bb, cc...)",
        title="Nomear Interruptores Industrial"
    )
    if not start_str:
        dbg.exit('name_switch_industrial', 'CANCELADO')
        return
    start_char = start_str.lower().strip()
    if not start_char or not start_char[0].isalpha(): return

    counter = 0
    target = start_char
    if len(target) == 1:
        if target in VALID_SWITCH_LETTERS: counter = VALID_SWITCH_LETTERS.index(target)
        else:
            for i, l in enumerate(VALID_SWITCH_LETTERS):
                if l >= target:
                    counter = i
                    break
    else:
        n = len(VALID_SWITCH_LETTERS)
        if target[0] in VALID_SWITCH_LETTERS:
            idx = VALID_SWITCH_LETTERS.index(target[0])
            counter = n + idx
        else: counter = n

    dbg.step('Counter inicial: {} -> label: {}'.format(counter, _get_switch_label(counter)))

    while True:
        try:
            label = _get_switch_label(counter)
            dbg.step('Aguardando seleção para label: {}'.format(label))
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                CategoryFilter(BuiltInCategory.OST_LightingDevices),
                "Selecione o INTERRUPTOR para nomear como: " + label + " (ESC para sair)"
            )
            interruptor = doc.GetElement(ref.ElementId)
            dbg.elem_info(interruptor, 'Interruptor selecionado')

            with Transaction(doc, "Nomear Interruptor Industrial") as t:
                t.Start()
                success = set_param(interruptor, ["ID do comando"], label)
                t.Commit()

            if success:
                dbg.ok('Nomeado: {} -> Id={}'.format(label, ref.ElementId.IntegerValue))
                counter += 1
            else:
                dbg.fail('Parâmetro ID do comando NÃO encontrado')
                break
        except OperationCanceledException:
            dbg.step('ESC pressionado - saindo')
            break
        except Exception as e:
            dbg.fail('Erro: {}'.format(e))
            forms.alert("Erro ao nomear interruptor:\n" + str(e))
            break
    dbg.exit('name_switch_industrial', 'FIM')

def create_individual_circuits_industrial():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return
    phase_config = prompt_industrial_phase_voltage(panel)
    if not phase_config: return
    forms.toast("Selecione os equipamentos (1 circuito por conector)...", title="Industrial")
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        ConnectorDomainFilter((Domain.DomainElectrical,)),
        "Selecione os equipamentos — cada conector elétrico será um circuito individual"
    )
    if not refs: return
    circuit_desc = _ask_circuit_description()
    if circuit_desc is None: return

    refs_list = list(refs)
    refs_list.reverse()

    created_count = 0
    reconnect_state = {"choice": None}
    with Transaction(doc, "Circuitos Individuais Industrial") as t:
        t.Start()
        try:
            for r in refs_list:
                host_elem = doc.GetElement(r.ElementId)
                valid_pairs = get_valid_electrical_elements(host_elem, (Domain.DomainElectrical,))
                if not valid_pairs:
                    continue

                for sub_id, conns in valid_pairs:
                    sub_elem = doc.GetElement(sub_id)
                    if _should_skip_connected_element(sub_elem, panel, reconnect_state):
                        continue

                    ensure_element_is_free(sub_elem)
                    _configure_industrial_element_for_circuit(sub_elem, phase_config, panel)
                    doc.Regenerate()

                    rm = get_room_name(host_elem)
                    for c in conns:
                        if c.IsConnected:
                            continue
                        try:
                            with suppress_elec_dialog():
                                circuit = ElectricalSystem.Create(c, ElectricalSystemType.PowerCircuit)
                            _select_panel_checked(circuit, panel, phase_config)
                            set_param(circuit, LOAD_NAME_PARAMS, _circuit_description(circuit_desc, rm if rm else "Carga Industrial"))
                            dbg.ok('Circuito individual criado Id={}'.format(circuit.Id.IntegerValue))
                            created_count += 1
                        except Exception as e:
                            dbg.fail('Falha ao criar circuito Id={}: {}'.format(sub_id.IntegerValue, e))
                            raise

            t.Commit()
            if created_count > 0:
                forms.toast("Industrial: {} circuito(s) individual(ais) criado(s).".format(created_count), title="Industrial")
            else:
                forms.alert("Nenhum circuito individual criado.")
        except Exception as e:
            t.RollBack()
            dbg.fail('Erro: {}'.format(e))
            forms.alert("Erro. Detalhes no console do pyRevit.")

def create_grouped_circuit_industrial():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    cat_choice = forms.CommandSwitchWindow.show(
        ["Luminárias", "Tomadas/Dispositivos"], message="Tipo de elementos a agrupar:", title="Circuito Agrupado Industrial"
    )
    if not cat_choice: return

    cat_id = BuiltInCategory.OST_LightingFixtures if "Lumin" in cat_choice else BuiltInCategory.OST_ElectricalFixtures
    phase_config = prompt_industrial_phase_voltage(panel)
    if not phase_config: return
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, CategoryFilter(cat_id), "Selecione os elementos para agrupar em 1 circuito"
    )
    if not refs: return
    circuit_desc = _ask_circuit_description()
    if circuit_desc is None: return

    ids = List[ElementId]()
    rooms = set()
    reconnect_state = {"choice": None}

    for r in refs:
        host = doc.GetElement(r.ElementId)
        valid_pairs = get_valid_electrical_elements(host, (Domain.DomainElectrical,))
        if not valid_pairs:
            continue
        for sub_id, _ in valid_pairs:
            sub = doc.GetElement(sub_id)
            if not _should_skip_connected_element(sub, panel, reconnect_state):
                ids.Add(sub_id)
        rm = get_room_name(host)
        if rm: rooms.add(rm)

    if ids.Count == 0:
        forms.toast("Nenhum elemento válido.", title="Agrupado Industrial")
        return

    with Transaction(doc, "Circuito Agrupado Industrial") as t:
        t.Start()
        for eid in ids:
            child_elem = doc.GetElement(eid)
            ensure_element_is_free(child_elem)
            _configure_industrial_element_for_circuit(child_elem, phase_config, panel)

        doc.Regenerate()
        try:
            with suppress_elec_dialog():
                circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
            _select_panel_checked(circuit, panel, phase_config)
            room_str = " / ".join(sorted(rooms)) if rooms else "Agrupado Industrial"
            set_param(circuit, LOAD_NAME_PARAMS, _circuit_description(circuit_desc, room_str))
            t.Commit()
            forms.toast("Circuito agrupado criado com {} elemento(s).".format(ids.Count), title="Industrial")
        except Exception as e:
            dbg.fail('ElectricalSystem.Create falhou: {}'.format(e))
            t.RollBack()
            forms.alert("Erro ao criar circuito agrupado:\n" + str(e))

def main_menu():
    while True:
        quadro = get_current_panel()
        status = "Quadro: " + (get_panel_name(quadro) if quadro else "NENHUM")

        opcoes = OrderedDict([
            ("1. Selecionar/Configurar Quadro", select_and_configure_panel_industrial),
            ("2. Criar Circuito Iluminação (max 2400W)", create_ilum_circuit_industrial),
            ("3. Comando Interruptor (a, b, c...)", name_switch_industrial),
            ("4. Criar Circuito Tomadas (220V, 10A/20A)", create_tomada_circuit_industrial),
            ("5. Criar Circuitos Individuais (1 por elemento)", create_individual_circuits_industrial),
            ("6. Criar Circuito Agrupado (1 para muitos)", create_grouped_circuit_industrial),
            ("7. Queda de Tensão", call_queda_tensao),
            ("8. Sair", lambda: None),
        ])

        escolha = forms.CommandSwitchWindow.show(opcoes.keys(), message=status, title="🏭 Industrial - " + status)
        if not escolha or "Sair" in escolha: break
        try: opcoes[escolha]()
        except Exception as e:
            if "aborted" not in str(e).lower() and "cancel" not in str(e).lower():
                forms.alert("Erro. Veja o console do pyRevit.")

if __name__ == "__main__":
    try: is_shift = __shiftclick__
    except NameError: is_shift = False

    if is_shift:
        lf_electrical_core.show_settings()
    else:
        main_menu()
