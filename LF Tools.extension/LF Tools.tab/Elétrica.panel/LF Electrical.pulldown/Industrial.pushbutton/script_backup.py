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
            ("3 Polos - 220V", {"poles": 3, "voltage": 220}),
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

    options = OrderedDict()
    for text, config in rule["circuit_options"].items():
        if config["poles"] == 1:
            options[text] = config

    if not options:
        forms.alert("Nenhuma tensao monofasica permitida para tomadas neste quadro.", title="Tomadas")
        return None

    if len(options) == 1:
        key = list(options.keys())[0]
        dbg.step('Tomadas: usando automaticamente {}'.format(key))
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

def _configure_industrial_element_for_circuit(elem, phase_config, panel):
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
                    try: c.Voltage = voltage_internal
                    except Exception: pass
                    try: c.Poles = poles
                    except Exception: pass
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
        prompt="Descricao do circuito:\nDeixe em branco para usar a descricao automatica.",
        title="Descricao do Circuito"
    )
    if value is None:
        return None
    return value.strip()

def _circuit_description(custom_text, fallback_text):
    return custom_text if custom_text else fallback_text

def _read_voltage_param(elem):
    for name in ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"]:
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

def create_tomada_circuit_industrial():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    phase_config = prompt_industrial_tomada_voltage(panel)
    if not phase_config: return
    refs = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_ElectricalFixtures), "Selecione as tomadas")
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
    with Transaction(doc, "Circuitos Tomadas Industrial") as t:
        t.Start()
        try:
            for bi, (b_ids, b_watts, b_rooms) in enumerate(batches):
                dbg.step('Batch {}: {} elemento(s) {:.0f}W'.format(bi+1, b_ids.Count, b_watts))
                for eid in b_ids:
                    elem = doc.GetElement(eid)
                    ensure_element_is_free(elem)
                    _configure_industrial_element_for_circuit(elem, phase_config, panel)
                dbg.step('Loop elementos batch {} concluido — Regenerate...'.format(bi+1))
                try:
                    doc.Regenerate()
                    dbg.ok('doc.Regenerate OK')
                except Exception as ex_regen:
                    dbg.fail('doc.Regenerate falhou: {}'.format(ex_regen))
                    dbg.fail(traceback.format_exc())
                ids_list = [eid.IntegerValue for eid in b_ids]
                dbg.step('ElectricalSystem.Create IDs: {}'.format(ids_list))
                try:
                    with suppress_elec_dialog():
                        circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                    dbg.ok('Create OK Id={}'.format(circuit.Id.IntegerValue))
                    _select_panel_checked(circuit, panel, phase_config)
                    room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Tomada {}A".format(detected_amp)
                    set_param(circuit, LOAD_NAME_PARAMS, _circuit_description(circuit_desc, "{} ({:.0f}W)".format(room_str, b_watts)))
                    dbg.ok('Circuito tomada criado Id={}'.format(circuit.Id.IntegerValue))
                    created += 1
                except Exception as ex:
                    dbg.fail('ElectricalSystem.Create falhou batch {} IDs={}: {}'.format(
                        bi+1, ids_list, ex))
                    dbg.fail(traceback.format_exc())
                    raise
            t.Commit()
            dbg.ok('Transaction COMMITTED')
        except Exception as ex_outer:
            dbg.fail('EXCECAO NAO CAPTURADA na Transaction: {}'.format(ex_outer))
            dbg.fail(traceback.format_exc())
            try: t.RollBack()
            except Exception: pass

    if created > 0: forms.toast("Sucesso! {} circuito(s) de tomada criado(s).".format(created), title="Industrial")

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
