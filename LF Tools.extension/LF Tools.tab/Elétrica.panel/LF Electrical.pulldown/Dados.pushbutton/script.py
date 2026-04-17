# coding: utf-8
"""LF Electrical - Dados / Telecom
Criação de circuitos e redes de telecom"""

__title__ = "Telecom"
__author__ = "Luís Fernando"

from pyrevit import forms
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI.Selection import ObjectType
from collections import OrderedDict

import lf_electrical_core
from lf_electrical_core import (
    doc, uidoc, get_current_panel, set_current_panel, get_panel_name, set_param, 
    ConnectorDomainFilter, get_valid_electrical_elements,
    ensure_element_is_free, get_room_name
)

def get_item_description(element, room_name):
    """Gera a descrição baseada no tipo de elemento (Camera, Wifi, etc)"""
    if not element:
        return room_name if room_name else "Ponto de Dados"
        
    el_type = doc.GetElement(element.GetTypeId())
    family_name = el_type.FamilyName.lower() if el_type else ""
    
    # Nome do tipo
    type_name = ""
    if el_type:
        tn_param = el_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = tn_param.AsString().lower() if tn_param else ""
    
    prefix = ""
    if "camera" in family_name or "camera" in type_name:
        prefix = "Camera "
    elif "wifi" in family_name or "wifi" in type_name:
        prefix = "Ponto Wifi Teto "
        
    if prefix:
        return (prefix + (room_name if room_name else "")).strip()
    
    return room_name if room_name else "Ponto de Dados"

def select_and_configure_panel_data():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, lf_electrical_core.PanelFilter(), "Selecione o Quadro/Rack de Telecom")
    except Exception:
        return
    panel = doc.GetElement(ref.ElementId)
    if not panel: return
    set_current_panel(panel.Id)
    forms.toast("Rack configurado: " + get_panel_name(panel), title="Dados")

def create_grouped_circuit_data():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o rack de telecom primeiro!")
        return
        
    load_name = forms.ask_for_string(prompt="Nome para a rede (Ex: RACK-1):", title="Rede Agrupada")
    if not load_name: return

    domains = (Domain.DomainData, Domain.DomainTelephone, Domain.DomainCommunication)
    refs = uidoc.Selection.PickObjects(ObjectType.Element, ConnectorDomainFilter(domains), "Selecione os pontos de rede -> " + load_name)
    if not refs: return

    with Transaction(doc, "Criar Rede " + load_name) as t:
        t.Start()
        ids = List[ElementId]()
        rooms = set()
        is_camera = False
        is_wifi = False

        for r in refs:
            host = doc.GetElement(r.ElementId)
            vp = get_valid_electrical_elements(host, domains)
            for sub_id, conns in vp:
                sub = doc.GetElement(sub_id)
                ensure_element_is_free(sub)
                ids.Add(sub_id)
                
                # Detectar tipo para prefixo
                el_type = doc.GetElement(sub.GetTypeId())
                if el_type:
                    fn = el_type.FamilyName.lower()
                    tn_param = el_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    tn = tn_param.AsString().lower() if tn_param else ""
                    if "camera" in fn or "camera" in tn: is_camera = True
                    if "wifi" in fn or "wifi" in tn: is_wifi = True

                rm = get_room_name(sub)
                if rm: rooms.add(rm)
        
        if ids.Count > 0:
            circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.DataCircuit)
            circuit.SelectPanel(panel)
            set_param(circuit, ["Nome da carga", "Load Name"], load_name)
            
            prefix = ""
            if is_camera: prefix = "Camera "
            elif is_wifi: prefix = "Ponto Wifi Teto "
            
            room_str = " / ".join(sorted(rooms)) if rooms else ("Rede de Dados" if not prefix else "")
            desc = (prefix + room_str).strip()
            set_param(circuit, ["Descrição", "Comments"], desc)
            t.Commit()
            forms.toast("Rede de dados criada: {}".format(load_name))
        else:
            t.RollBack()
            forms.alert("Nenhum conector válido.")

def create_individual_circuits_data():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o rack de telecom primeiro!")
        return

    prefixo = forms.ask_for_string(default="PT", prompt="Prefixo do Ponto (ex: PT, VZ):", title="Prefixo")
    if not prefixo: return
    try: contador = int(forms.ask_for_string(default="1", prompt="Início (ex: 1):", title="Contador"))
    except: contador = 1
    
    domains = (Domain.DomainData, Domain.DomainTelephone, Domain.DomainCommunication)
    refs = uidoc.Selection.PickObjects(ObjectType.Element, ConnectorDomainFilter(domains), "Selecione os pontos de rede")
    if not refs: return
    
    refs_list = list(refs)
    refs_list.reverse()
    
    created = 0
    with Transaction(doc, "Pontos Individuais") as t:
        t.Start()
        for r in refs_list:
            host = doc.GetElement(r.ElementId)
            vp = get_valid_electrical_elements(host, domains)
            for sub_id, conns in vp:
                sub = doc.GetElement(sub_id)
                ensure_element_is_free(sub)
                rm = get_room_name(sub)
                for c in conns:
                    if c.IsConnected: continue
                    try:
                        circuit = ElectricalSystem.Create(c, ElectricalSystemType.DataCircuit)
                        circuit.SelectPanel(panel)
                        nome = prefixo + str(contador)
                        set_param(circuit, ["Nome da carga", "Load Name"], nome)
                        desc = get_item_description(sub, rm)
                        set_param(circuit, ["Descrição", "Comments"], desc)
                        contador += 1
                        created += 1
                    except: pass
        t.Commit()
        forms.toast("{} ponto(s) criado(s).".format(created))

def main_menu():
    while True:
        quadro = get_current_panel()
        status = "Rack/Quadro: " + (get_panel_name(quadro) if quadro else "NENHUM")
        opcoes = OrderedDict([
            ("1. Selecionar Rack/Painel de Telecom", select_and_configure_panel_data),
            ("2. Criar Circuitos Individuais (1 Fio Exclusivo p/ cada porta/ponto)", create_individual_circuits_data),
            ("3. Criar Circuito Agrupado (Agrupar pontos num conduíte/rede único)", create_grouped_circuit_data),
            ("4. Sair", lambda: None),
        ])
        escolha = forms.CommandSwitchWindow.show(opcoes.keys(), message=status, title="📡 Telecom/Dados - " + status)
        if not escolha or "Sair" in escolha: break
        try: opcoes[escolha]()
        except Exception as _err:
            import traceback
            forms.alert("Erro:\n" + traceback.format_exc())

if __name__ == "__main__":
    try: is_shift = __shiftclick__
    except NameError: is_shift = False

    if is_shift:
        lf_electrical_core.show_settings()
    else:
        main_menu()
