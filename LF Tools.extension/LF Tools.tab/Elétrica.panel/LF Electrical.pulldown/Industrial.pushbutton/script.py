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
    prompt_phase_voltage, ConnectorDomainFilter, get_valid_electrical_elements,
    is_element_connected_to_panel, ensure_element_is_free, get_room_name, get_family_name,
    configure_panel, PanelFilter, CategoryFilter, call_queda_tensao,
    VALID_SWITCH_LETTERS, _get_switch_label, load_config, dbg
)

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

    # Filtrar sistemas compativeis com o quadro
    sys_id = None
    compatible_systems = lf_electrical_core._get_compatible_systems(panel)

    if compatible_systems:
        chosen_sys = forms.CommandSwitchWindow.show(sorted(compatible_systems.keys()), message="Sistema de Distribuição para " + new_name + " ({} compatíveis):".format(len(compatible_systems)), title="Sistema de Distribuição")
        if chosen_sys:
            sys_id = compatible_systems.get(chosen_sys)
            messages.append("Sistema: " + chosen_sys)
    else:
        forms.alert("Nenhum sistema de distribuição compatível com este quadro.")

    all_namings = {}
    for s in FilteredElementCollector(doc).OfClass(CircuitNamingScheme).ToElements():
        all_namings[s.Name] = s.Id

    nam_id = None
    if all_namings:
        chosen_naming = forms.CommandSwitchWindow.show(sorted(all_namings.keys()), message="Nomenclatura do Circuito:", title="Nomenclatura")
        if chosen_naming:
            nam_id = all_namings.get(chosen_naming)
            messages.append("Nomenclatura: " + chosen_naming)

    with Transaction(doc, "Config Quadro Industrial") as t:
        t.Start()
        p = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and not p.IsReadOnly and sys_id:
            try: p.Set(sys_id)
            except Exception: messages.append("Sistema: FALHA")

        p = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")
        if p and not p.IsReadOnly and nam_id: p.Set(nam_id)
        t.Commit()

    return True, "\n".join(messages)

def select_and_configure_panel_industrial():
    dbg.section('INDUSTRIAL - Selecionar Quadro')
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, PanelFilter(), "Selecione o QUADRO INDUSTRIAL")
        panel = doc.GetElement(ref.ElementId)
        dbg.elem_info(panel, 'Quadro selecionado')

        p_sys = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        p_nam = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")
        
        has_sys = p_sys and p_sys.HasValue and p_sys.AsElementId() != ElementId.InvalidElementId
        has_nam = p_nam and p_nam.HasValue and p_nam.AsElementId() != ElementId.InvalidElementId
        
        dbg.step('Params validos: sys={}, nam={}'.format(has_sys, has_nam))

        if has_sys and has_nam:
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

    phase_config = prompt_phase_voltage()
    if not phase_config: return

    refs = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_LightingFixtures), "Selecione as luminarias (2400W cada circuito)")
    if not refs: return

    refs_list = list(refs)
    refs_list.reverse()

    batches = []
    current_batch = List[ElementId]()
    current_watts = 0.0
    current_rooms = set()
    skipped = []

    for r in refs_list:
        elem = doc.GetElement(r.ElementId)
        if is_element_connected_to_panel(elem):
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

    created = 0
    with Transaction(doc, "Circuitos Ilum Industrial") as t:
        t.Start()
        for b_ids, b_watts, b_rooms in batches:
            for eid in b_ids:
                elem = doc.GetElement(eid)
                set_param(elem, ["N° de Fases", "Nº de Fases", "Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
                set_param(elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                ensure_element_is_free(elem)
            
            doc.Regenerate()
            try:
                circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                circuit.SelectPanel(panel)
                room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Ilum Industrial"
                set_param(circuit, ["Descrição", "Comments"], "{} ({:.0f}W)".format(room_str, b_watts))
                created += 1
            except Exception: pass
        t.Commit()

    if created > 0: forms.toast("Sucesso! {} circuito(s) criado(s).".format(created), title="Industrial")

def create_tomada_circuit_industrial():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    phase_config = prompt_phase_voltage()
    if not phase_config: return

    refs = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_ElectricalFixtures), "Selecione as tomadas")
    if not refs: return

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

    for r in refs_list:
        elem = doc.GetElement(r.ElementId)
        if is_element_connected_to_panel(elem): continue
            
        watts = _get_element_wattage(elem)
        rm = get_room_name(elem)
        if current_watts + watts > limite_w and current_batch.Count > 0:
            batches.append((current_batch, current_watts, set(current_rooms)))
            current_batch = List[ElementId]()
            current_watts = 0.0
            current_rooms = set()
        
        current_batch.Add(r.ElementId)
        current_watts += watts
        if rm: current_rooms.add(rm)

    if current_batch.Count > 0:
        batches.append((current_batch, current_watts, set(current_rooms)))

    created = 0
    with Transaction(doc, "Circuitos Tomadas Industrial") as t:
        t.Start()
        for b_ids, b_watts, b_rooms in batches:
            for eid in b_ids:
                elem = doc.GetElement(eid)
                set_param(elem, ["N° de Fases", "Nº de Fases", "Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
                set_param(elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                ensure_element_is_free(elem)
            doc.Regenerate()
            try:
                circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                circuit.SelectPanel(panel)
                room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Tomada {}A".format(detected_amp)
                set_param(circuit, ["Descrição", "Comments"], "{} ({:.0f}W)".format(room_str, b_watts))
                created += 1
            except Exception: pass
        t.Commit()

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
    phase_config = prompt_phase_voltage()
    if not phase_config: return

    forms.toast("Selecione os equipamentos (1 circuito por elemento)...", title="Industrial")
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        CategoryFilter(BuiltInCategory.OST_ElectricalFixtures),
        "Selecione os equipamentos — cada um será um circuito individual"
    )
    if not refs: return

    refs_list = list(refs)
    refs_list.reverse()

    created_count, contador = 0, 1
    with Transaction(doc, "Circuitos Individuais Industrial") as t:
        t.Start()
        try:
            for r in refs_list:
                elem = doc.GetElement(r.ElementId)
                if is_element_connected_to_panel(elem): continue

                has_connector = False
                try:
                    if hasattr(elem, 'MEPModel') and elem.MEPModel:
                        cm = elem.MEPModel.ConnectorManager
                        if cm:
                            for c in cm.Connectors:
                                if c.Domain == Domain.DomainElectrical:
                                    has_connector = True; break
                except Exception: pass

                if not has_connector: continue

                ensure_element_is_free(elem)
                set_param(elem, ["N° de Fases", "Nº de Fases", "Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
                set_param(elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                doc.Regenerate()

                ids = List[ElementId]()
                ids.Add(elem.Id)
                try: circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
                except Exception: continue

                circuit.SelectPanel(panel)
                set_param(circuit, ["Nome da carga", "Load Name"], str(contador))
                rm = get_room_name(elem)
                set_param(circuit, ["Descrição", "Comments"], rm if rm else "Carga Industrial")
                contador += 1
                created_count += 1
            t.Commit()
            forms.toast("Industrial: {} circuito(s) individual(ais) criado(s).".format(created_count), title="Industrial")
        except Exception as e:
            t.RollBack()
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
    phase_config = prompt_phase_voltage()
    if not phase_config: return

    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, CategoryFilter(cat_id), "Selecione os elementos para agrupar em 1 circuito"
    )
    if not refs: return

    ids = List[ElementId]()
    rooms = set()

    for r in refs:
        elem = doc.GetElement(r.ElementId)
        if is_element_connected_to_panel(elem): continue

        try:
            if hasattr(elem, 'MEPModel') and elem.MEPModel:
                existing = elem.MEPModel.ElectricalSystems
                if existing and existing.Count > 0:
                    for es in existing:
                        try: doc.Delete(es.Id)
                        except Exception: pass
        except Exception: pass

        has_connector = False
        try:
            if hasattr(elem, 'MEPModel') and elem.MEPModel:
                cm = elem.MEPModel.ConnectorManager
                if cm:
                    for c in cm.Connectors:
                        if c.Domain == Domain.DomainElectrical:
                            has_connector = True; break
        except Exception: pass

        if has_connector:
            ids.Add(r.ElementId)
            rm = get_room_name(elem)
            if rm: rooms.add(rm)

    if ids.Count == 0:
        forms.toast("Nenhum elemento valido.", title="Agrupado Industrial")
        return

    with Transaction(doc, "Circuito Agrupado Industrial") as t:
        t.Start()
        for eid in ids:
            child_elem = doc.GetElement(eid)
            set_param(child_elem, ["N° de Fases", "Nº de Fases", "Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
            set_param(child_elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])

        doc.Regenerate()
        try:
            circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
            circuit.SelectPanel(panel)
            room_str = " / ".join(sorted(rooms)) if rooms else "Agrupado Industrial"
            set_param(circuit, ["Descrição", "Comments"], room_str)
            t.Commit()
            forms.toast("Circuito agrupado criado com {} elemento(s).".format(ids.Count), title="Industrial")
        except Exception as e:
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
