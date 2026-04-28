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
    VALID_SWITCH_LETTERS, _get_switch_label, load_config, dbg,
    configure_element_for_voltage, get_element_poles,
    suppress_elec_dialog, CIRCUIT_DESC_PARAMS
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

    # Obter TODOS os sistemas sem filtro (deixar o Revit validar na hora do Set)
    all_systems = {}
    for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
        n = ""
        try: n = s.Name
        except Exception:
            pp = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if pp: n = pp.AsString()
        if n: all_systems[n] = s.Id

    sys_id = None
    chosen_sys = None
    if all_systems:
        chosen_sys = forms.CommandSwitchWindow.show(
            sorted(all_systems.keys()),
            message="Sistema de Distribuição para {} ({} disponíveis):".format(new_name, len(all_systems)),
            title="Sistema de Distribuição"
        )
        if chosen_sys:
            sys_id = all_systems.get(chosen_sys)
    else:
        forms.alert("Nenhum sistema de distribuição encontrado no modelo.")

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
            try:
                p.Set(sys_id)
                messages.append("Sistema: " + chosen_sys)
            except Exception as e:
                t.RollBack()
                forms.alert(
                    u"Sistema '{}' é incompatível com este quadro.\n"
                    u"Escolha um sistema com tensão compatível.\n\n"
                    u"Detalhe: {}".format(chosen_sys, e),
                    title="Sistema Inválido"
                )
                return False, "Sistema de distribuição inválido"

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

    dbg.step('Batches: {} | Elementos totais: {}'.format(len(batches), sum(b[0].Count for b in batches)))
    created = 0
    with Transaction(doc, "Circuitos Ilum Industrial") as t:
        t.Start()
        for bi, (b_ids, b_watts, b_rooms) in enumerate(batches):
            dbg.step('Batch {}: {} elemento(s) {:.0f}W rooms={}'.format(bi+1, b_ids.Count, b_watts, list(b_rooms)))
            for eid in b_ids:
                elem = doc.GetElement(eid)
                ensure_element_is_free(elem)
                set_param(elem, ["N\xb0 de Fases", "N\xba de Fases", "N\xfamero de Fases", "N\xfamero de polos", "Number of Poles", "Polos"], phase_config["poles"])
                set_param(elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                configure_element_for_voltage(elem, phase_config["voltage"], phase_config["poles"])

            doc.Regenerate()
            try:
                with suppress_elec_dialog():
                    circuit = ElectricalSystem.Create(doc, b_ids, ElectricalSystemType.PowerCircuit)
                circuit.SelectPanel(panel)
                room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Ilum Industrial"
                set_param(circuit, CIRCUIT_DESC_PARAMS, "{} ({:.0f}W)".format(room_str, b_watts))
                dbg.ok('Circuito ilum criado Id={}'.format(circuit.Id.IntegerValue))
                created += 1
            except Exception as ex:
                dbg.fail('ElectricalSystem.Create falhou batch {}: {}'.format(bi+1, ex))
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
    skipped_poles = []

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
            if is_element_connected_to_panel(sub):
                dbg.step('  Sub Id={} ja ligado ao painel — pulado'.format(sub_id.IntegerValue))
                continue
            sub_poles = get_element_poles(sub) or get_element_poles(host)
            dbg.step('  Sub Id={} poles={} config={}'.format(sub_id.IntegerValue, sub_poles, phase_config["poles"]))
            if sub_poles is not None and sub_poles != phase_config["poles"]:
                dbg.warn('Sub Id={} tem {} fase(s), esperado {}. Pulado.'.format(
                    sub_id.IntegerValue, sub_poles, phase_config["poles"]))
                skipped_poles.append(str(sub_id.IntegerValue))
                continue
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

    if skipped_poles:
        forms.alert(u"{} elemento(s) ignorado(s) por ter fases incompatíveis com a configuração ({} fase(s)):\n{}".format(
            len(skipped_poles), phase_config["poles"], ", ".join(skipped_poles)), title="Incompatibilidade de Fases")

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
                    set_param(elem, ["N\xb0 de Fases", "N\xba de Fases", "N\xfamero de Fases", "N\xfamero de polos", "Number of Poles", "Polos"], phase_config["poles"])
                    set_param(elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                    configure_element_for_voltage(elem, phase_config["voltage"], phase_config["poles"])
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
                    circuit.SelectPanel(panel)
                    room_str = " / ".join(sorted(b_rooms)) if b_rooms else "Tomada {}A".format(detected_amp)
                    set_param(circuit, CIRCUIT_DESC_PARAMS, "{} ({:.0f}W)".format(room_str, b_watts))
                    dbg.ok('Circuito tomada criado Id={}'.format(circuit.Id.IntegerValue))
                    created += 1
                except Exception as ex:
                    dbg.fail('ElectricalSystem.Create falhou batch {} IDs={}: {}'.format(
                        bi+1, ids_list, ex))
                    dbg.fail(traceback.format_exc())
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
    phase_config = prompt_phase_voltage()
    if not phase_config: return

    forms.toast("Selecione os equipamentos (1 circuito por conector)...", title="Industrial")
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        ConnectorDomainFilter((Domain.DomainElectrical,)),
        "Selecione os equipamentos — cada conector elétrico será um circuito individual"
    )
    if not refs: return

    refs_list = list(refs)
    refs_list.reverse()

    created_count = 0
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
                    if is_element_connected_to_panel(sub_elem):
                        continue

                    ensure_element_is_free(sub_elem)
                    set_param(sub_elem, ["N\xb0 de Fases", "N\xba de Fases", "N\xfamero de Fases", "N\xfamero de polos", "Number of Poles", "Polos"], phase_config["poles"])
                    set_param(sub_elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                    configure_element_for_voltage(sub_elem, phase_config["voltage"], phase_config["poles"])
                    doc.Regenerate()

                    rm = get_room_name(host_elem)
                    for c in conns:
                        if c.IsConnected:
                            continue
                        try:
                            with suppress_elec_dialog():
                                circuit = ElectricalSystem.Create(c, ElectricalSystemType.PowerCircuit)
                            circuit.SelectPanel(panel)
                            set_param(circuit, CIRCUIT_DESC_PARAMS, rm if rm else "Carga Industrial")
                            dbg.ok('Circuito individual criado Id={}'.format(circuit.Id.IntegerValue))
                            created_count += 1
                        except Exception as e:
                            dbg.fail('Falha ao criar circuito Id={}: {}'.format(sub_id.IntegerValue, e))

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
    phase_config = prompt_phase_voltage()
    if not phase_config: return

    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, CategoryFilter(cat_id), "Selecione os elementos para agrupar em 1 circuito"
    )
    if not refs: return

    ids = List[ElementId]()
    rooms = set()

    for r in refs:
        host = doc.GetElement(r.ElementId)
        valid_pairs = get_valid_electrical_elements(host, (Domain.DomainElectrical,))
        if not valid_pairs:
            continue
        for sub_id, _ in valid_pairs:
            sub = doc.GetElement(sub_id)
            if not is_element_connected_to_panel(sub):
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
            set_param(child_elem, ["N\xb0 de Fases", "N\xba de Fases", "N\xfamero de Fases", "N\xfamero de polos", "Number of Poles", "Polos"], phase_config["poles"])
            set_param(child_elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
            configure_element_for_voltage(child_elem, phase_config["voltage"], phase_config["poles"])

        doc.Regenerate()
        try:
            with suppress_elec_dialog():
                circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
            circuit.SelectPanel(panel)
            room_str = " / ".join(sorted(rooms)) if rooms else "Agrupado Industrial"
            set_param(circuit, CIRCUIT_DESC_PARAMS, room_str)
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
