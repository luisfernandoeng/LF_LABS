# coding: utf-8
"""LF Electrical - Residencial
Criação de circuitos e interruptores (Residencial)"""

__title__ = "Residencial"
__author__ = "Luís Fernando"

from pyrevit import forms
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI.Selection import ObjectType
from collections import OrderedDict
import traceback
from Autodesk.Revit.Exceptions import OperationCanceledException

import lf_electrical_core
from lf_electrical_core import (
    doc, uidoc, get_current_panel, get_panel_name, set_param, 
    prompt_phase_voltage, ConnectorDomainFilter, get_valid_electrical_elements,
    is_element_connected_to_panel, ensure_element_is_free, get_room_name,
    select_and_configure_panel, call_queda_tensao,
    CategoryFilter, next_valid_letter, dbg
)

def create_grouped_circuit(load_name, target_voltage=None):
    dbg.section('RESIDENCIAL - Criar Circuito Agrupado')
    dbg.enter('create_grouped_circuit', load_name=load_name)
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        dbg.exit('create_grouped_circuit', 'SEM QUADRO')
        return

    phase_config = prompt_phase_voltage()
    if not phase_config:
        dbg.exit('create_grouped_circuit', 'CANCELADO (fase)')
        return
    dbg.step('Fase/Tensao: polos={}, voltage={}'.format(phase_config['poles'], phase_config['voltage']))

    refs = uidoc.Selection.PickObjects(ObjectType.Element, ConnectorDomainFilter((Domain.DomainElectrical,)), "Selecione os elementos -> " + load_name)
    if not refs:
        dbg.exit('create_grouped_circuit', 'CANCELADO (selecao)')
        return
    dbg.step('Elementos selecionados: {}'.format(len(list(refs))))

    with Transaction(doc, "Criar Circuito " + load_name) as t:
        t.Start()
        dbg.step('Transaction STARTED')

        ids = List[ElementId]()
        skipped_no_connector = []
        skipped_already_connected = []
        disconnected = 0
        rooms = set()

        for r in refs:
            host_elem = doc.GetElement(r.ElementId)
            dbg.elem_info(host_elem, 'Processando')
            valid_pairs = get_valid_electrical_elements(host_elem, (Domain.DomainElectrical,))
            
            if not valid_pairs:
                skipped_no_connector.append(str(r.ElementId.IntegerValue))
                dbg.warn('Sem conector: Id={}'.format(r.ElementId.IntegerValue))
                continue

            for sub_id, conns in valid_pairs:
                elem = doc.GetElement(sub_id)
                if is_element_connected_to_panel(elem):
                    skipped_already_connected.append(str(sub_id.IntegerValue))
                    continue
                if ensure_element_is_free(elem):
                    disconnected += 1

                ids.Add(sub_id)
                rm = get_room_name(elem)
                if rm: rooms.add(rm)

        dbg.step('IDs para circuito: {} | Pulados(sem conector): {} | Pulados(ja ligados): {} | Desconectados: {}'.format(
            ids.Count, len(skipped_no_connector), len(skipped_already_connected), disconnected))

        if ids.Count == 0:
            t.RollBack()
            dbg.fail('Nenhum elemento disponivel - RollBack')
            msg = "Nenhum elemento disponivel para criar circuito.\n"
            if skipped_already_connected: msg += "\nJa ligados: " + ", ".join(skipped_already_connected)
            if skipped_no_connector: msg += "\nSem conector: " + ", ".join(skipped_no_connector)
            forms.alert(msg)
            return

        dbg.step('Configurando polos/tensao nos elementos...')
        for eid in ids:
            child_elem = doc.GetElement(eid)
            set_param(child_elem, ["N° de Fases", "Nº de Fases", "Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
            set_param(child_elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])

        if disconnected > 0:
            dbg.step('Regenerando documento...')
            doc.Regenerate()

        dbg.step('Criando ElectricalSystem...')
        try:
            circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
            dbg.ok('Circuito criado: Id={}'.format(circuit.Id.IntegerValue))
        except Exception as e:
            dbg.fail('ElectricalSystem.Create falhou: {}'.format(e))
            if "electComponents" in str(e):
                t.RollBack()
                forms.alert("Nenhum dos elementos selecionados pôde criar circuito de força. Verifique conectores.")
                return
            else:
                t.RollBack()
                forms.alert("ERRO de Compatibilidade. Detalhes no console do pyRevit.\n\n" + str(e))
                return

        dbg.step('SelectPanel...')
        circuit.SelectPanel(panel)
        set_param(circuit, ["Nome da carga", "Load Name"], load_name)
        room_str = " / ".join(sorted(rooms)) if rooms else "Automático"
        set_param(circuit, ["Descrição", "Comments"], room_str)

        t.Commit()
        dbg.ok('Transaction COMMITTED')
        dbg.exit('create_grouped_circuit', 'SUCESSO')
        forms.toast("Circuito criado: " + load_name)

def create_individual_circuits():
    dbg.section('RESIDENCIAL - Criar Circuitos Individuais')
    dbg.enter('create_individual_circuits')
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        dbg.exit('create_individual_circuits', 'SEM QUADRO')
        return

    prefixo = forms.ask_for_string(default="AC", prompt="Prefixo (ex: AC, CH):", title="Prefixo")
    if not prefixo:
        dbg.exit('create_individual_circuits', 'CANCELADO (prefixo)')
        return

    phase_config = prompt_phase_voltage()
    if not phase_config:
        dbg.exit('create_individual_circuits', 'CANCELADO (fase)')
        return
    dbg.step('Fase/Tensao: polos={}, voltage={}'.format(phase_config['poles'], phase_config['voltage']))

    inicio_str = forms.ask_for_string(default="1", prompt="Início (ex: 1):", title="Contador")
    try: contador = int(inicio_str)
    except Exception: contador = 1
    dbg.step('Contador inicial: {}'.format(contador))

    forms.toast("Selecione os equipamentos...")
    refs = uidoc.Selection.PickObjects(ObjectType.Element, ConnectorDomainFilter((Domain.DomainElectrical,)), "Selecione os equipamentos")
    if not refs:
        dbg.exit('create_individual_circuits', 'CANCELADO (selecao)')
        return
    dbg.step('Elementos selecionados: {}'.format(len(list(refs))))

    refs_list = list(refs)
    refs_list.reverse()

    created_count = 0
    skipped_already_connected = []

    with Transaction(doc, "Criar Circuitos " + prefixo) as t:
        t.Start()
        dbg.step('Transaction STARTED')
        try:
            for r in refs_list:
                host_elem = doc.GetElement(r.ElementId)
                dbg.elem_info(host_elem, 'Processando')
                valid_pairs = get_valid_electrical_elements(host_elem, (Domain.DomainElectrical,))
                if not valid_pairs: continue

                for sub_id, conns in valid_pairs:
                    sub_elem = doc.GetElement(sub_id)
                    if is_element_connected_to_panel(sub_elem):
                        skipped_already_connected.append(str(sub_id.IntegerValue))
                        dbg.warn('Ja ligado: Id={}'.format(sub_id.IntegerValue))
                        continue

                    ensure_element_is_free(sub_elem)
                    set_param(sub_elem, ["N° de Fases", "Nº de Fases", "Número de polos", "Number of Poles", "Polos"], phase_config["poles"])
                    set_param(sub_elem, ["Tensão (V)", "Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], phase_config["voltage"])
                    doc.Regenerate()
                    
                    rm = get_room_name(sub_elem)
                    for c in conns:
                        if c.IsConnected: continue
                        try:
                            circuit = ElectricalSystem.Create(c, ElectricalSystemType.PowerCircuit)
                            circuit.SelectPanel(panel)
                            nome = prefixo + str(contador)
                            set_param(circuit, ["Nome da carga", "Load Name"], nome)
                            set_param(circuit, ["Descrição", "Comments"], rm if rm else "Carga Específica")
                            dbg.ok('Circuito criado: {} (Id={})'.format(nome, circuit.Id.IntegerValue))
                            contador += 1
                            created_count += 1
                        except Exception as e:
                            dbg.fail('Falha ao criar circuito: {}'.format(e))
                            pass

            t.Commit()
            dbg.ok('Transaction COMMITTED')
            dbg.exit('create_individual_circuits', 'SUCESSO (count={})'.format(created_count))
            if created_count > 0:
                forms.toast("AC/CH: {} circuito(s) criado(s).".format(created_count))
            else:
                forms.alert("Nenhum circuito AC/CH foi criado.")
        except Exception as e:
            t.RollBack()
            dbg.fail('Erro fatal: {}'.format(e))
            forms.alert("Erro ao criar circuitos individuais. Detalhes no console do pyRevit.\n" + str(e))

def command_name_switch():
    dbg.section('RESIDENCIAL - Nomear Interruptores')
    start_letter_str = forms.ask_for_string(
        default="A",
        prompt="Letra inicial (ex: A, C, T):\n(O e S são puladas automaticamente)",
        title="Nomear Interruptores"
    )
    if not start_letter_str or len(start_letter_str) != 1:
        dbg.exit('command_name_switch', 'CANCELADO')
        return

    start_letter = start_letter_str.upper()
    if not start_letter.isalpha():
        forms.alert("Por favor, insira uma única letra do alfabeto.")
        return

    counter = ord(start_letter) - ord('A')
    next_letter, counter = next_valid_letter(counter)
    dbg.step('Letra inicial: {} (counter={})'.format(next_letter, counter))

    while True:
        try:
            next_letter, counter = next_valid_letter(counter)
            dbg.step('Aguardando seleção do interruptor para letra: {}'.format(next_letter))
            ref = uidoc.Selection.PickObject(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_LightingDevices), "Selecione o INTERRUPTOR para nomear como: " + next_letter + " (ESC para sair)")
            interruptor = doc.GetElement(ref.ElementId)
            dbg.elem_info(interruptor, 'Interruptor selecionado')

            with Transaction(doc, "Nomear Interruptor") as t:
                t.Start()
                success = set_param(interruptor, ["ID do comando"], next_letter)
                t.Commit()
                dbg.step('Transaction committed para letra {}'.format(next_letter))

            if success:
                dbg.ok('Interruptor nomeado: {} -> Id={}'.format(next_letter, ref.ElementId.IntegerValue))
                counter += 1
            else:
                dbg.fail('Parâmetro ID do comando NÃO encontrado')
                forms.alert("Parâmetro 'ID do comando' não encontrado.")
                break
        except OperationCanceledException:
            dbg.step('ESC pressionado - saindo do loop')
            break
        except Exception as e:
            dbg.fail('Erro no loop: {}'.format(e))
            forms.alert("Erro ao nomear interruptor:\n" + str(e))
            break
    dbg.exit('command_name_switch', 'FIM')

def main_menu():
    while True:
        quadro = get_current_panel()
        status = "Quadro: " + (get_panel_name(quadro) if quadro else "NENHUM")

        def call_ilum():
            nome = forms.ask_for_string(default="1", prompt="Nome:", title="Iluminação")
            if nome: create_grouped_circuit(nome)

        def call_tomada():
            nome = forms.ask_for_string(default="T", prompt="Nome:", title="Tomadas Gerais")
            if nome: create_grouped_circuit(nome)

        opcoes = OrderedDict([
            ("1. Selecionar/Configurar Quadro", select_and_configure_panel),
            ("2. Criar Circuito Iluminação (Geral)", call_ilum),
            ("3. Comando Interruptor", command_name_switch),
            ("4. Criar Circuito Tomadas (Geral)", call_tomada),
            ("5. Criar Circuitos Específicos (AC/CH)", create_individual_circuits),
            ("6. Queda de Tensão", call_queda_tensao),
            ("7. Sair", lambda: None),
        ])

        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(), message=status, title="🏠 Residencial - " + status
        )
        if not escolha or "Sair" in escolha: break
        try: opcoes[escolha]()
        except Exception as e:
            if "aborted" not in str(e).lower() and "cancel" not in str(e).lower():
                forms.alert("Erro: " + str(e))

if __name__ == "__main__":
    try: is_shift = __shiftclick__
    except NameError: is_shift = False

    if is_shift:
        lf_electrical_core.show_settings()
    else:
        main_menu()
