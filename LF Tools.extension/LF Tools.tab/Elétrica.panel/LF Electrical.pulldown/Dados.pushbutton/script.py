# coding: utf-8
"""LF Electrical - Dados
Criação de circuitos e redes de dados"""

__title__ = "Dados"
__author__ = "Luís Fernando"

from pyrevit import forms
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from collections import OrderedDict
import traceback

import lf_electrical_core
from lf_electrical_core import (
    doc, uidoc, get_current_panel, set_current_panel, get_panel_name, set_param,
    ConnectorDomainFilter, get_valid_electrical_elements,
    ensure_element_is_free, get_room_name, dbg,
    suppress_elec_dialog, CIRCUIT_DESC_PARAMS
)

# Categorias de dados/telecom usadas em toda a seleção.
# OST_ConduitFitting (-2008128) = "Conexões do conduite" (caixas de passagem FAST-ELE-CAIXA...)
# Incluída no filtro visual para o usuário poder clicar, mas será descartada no processamento
# pois essas famílias NÃO possuem conector DomainElectrical — só DomainCableTrayConduit.
DATA_CATS = {
    int(BuiltInCategory.OST_DataDevices),           # Dispositivos de dados (RJ45, condulete)
    int(BuiltInCategory.OST_CommunicationDevices),  # Dispositivos de comunicação
    int(BuiltInCategory.OST_TelephoneDevices),      # Dispositivos de telefone
    int(BuiltInCategory.OST_ConduitFitting),        # Conexões do conduite (caixas passagem)
}

# ConnectorDomainFilter já inclui essas categorias no seu __init__ padrão.
# Passar allowed_categories= causava TypeError — bug corrigido: não existe esse kwarg.
DATA_DOMAIN = (Domain.DomainElectrical,)

# ElectricalSystemType.DataCircuit só existe no Revit 2022+ como membro nomeado.
# Em versões anteriores, o valor inteiro 3 pode ainda existir na DLL sem estar exposto via nome.
# PowerCircuit (0) é incompatível com conectores de dados e sempre falha.
import System as _SYS

def _resolve_data_sys_type():
    # Tenta 1: nomes conhecidos por versão de Revit
    for attr_name in ['DataCircuit', 'Data']:
        try:
            v = getattr(ElectricalSystemType, attr_name)
            dbg.ok('ElectricalSystemType.{} disponivel = {}'.format(attr_name, int(v)))
            return v
        except AttributeError:
            pass

    # Tenta 2: enumera todos os valores disponíveis e procura por nome "data" ou "circuit"
    try:
        available = []
        data_candidate = None
        for v in _SYS.Enum.GetValues(ElectricalSystemType):
            name = _SYS.Enum.GetName(ElectricalSystemType, v)
            available.append('{}={}'.format(name, int(v)))
            if name and 'data' in name.lower() and data_candidate is None:
                data_candidate = v
        dbg.warn('ElectricalSystemType disponíveis: {}'.format(', '.join(available)))
        if data_candidate is not None:
            dbg.ok('Encontrado por nome: {}'.format(data_candidate))
            return data_candidate
    except Exception as _ex:
        dbg.warn('Nao foi possivel enumerar ElectricalSystemType: {}'.format(_ex))

    # Tenta 3: inteiro via Enum.ToObject — somente valores que existem na DLL
    # (evita retornar inteiro inválido como 3 quando o tipo não existe)
    for candidate in [5, 3, 4]:
        try:
            v = _SYS.Enum.ToObject(ElectricalSystemType, candidate)
            name = _SYS.Enum.GetName(ElectricalSystemType, v)
            if name:
                dbg.warn('ElectricalSystemType({}) = {} (nome={})'.format(candidate, int(v), name))
                return v
            else:
                dbg.warn('ElectricalSystemType({}) sem nome definido - ignorando'.format(candidate))
        except Exception as _ex:
            dbg.warn('ElectricalSystemType({}) falhou: {}'.format(candidate, _ex))

    dbg.fail('Nenhum ElectricalSystemType compatível com dados encontrado nesta versao do Revit')
    return None

_DATA_SYS_TYPE = _resolve_data_sys_type()


class DataCategoryFilter(ISelectionFilter):
    """Filtro de seleção restrito às categorias de dados/telecom."""
    def AllowElement(self, e):
        if not e.Category:
            return False
        return e.Category.Id.IntegerValue in DATA_CATS
    def AllowReference(self, ref, pos):
        return False


def _debug_connector(c, label=""):
    """Loga detalhes de um conector."""
    try:
        dbg.step('  Conector {}: domain={} connected={} IsConnected={}'.format(
            label, c.Domain, c.IsConnected, c.IsConnected))
    except Exception as ex:
        dbg.warn('  Conector {} erro ao inspecionar: {}'.format(label, ex))


def _debug_element_connectors(elem, label=""):
    """Loga todos os conectores de um elemento."""
    try:
        mgr = None
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            mgr = getattr(elem.MEPModel, 'ConnectorManager', None)
        if not mgr and hasattr(elem, 'ConnectorManager'):
            mgr = elem.ConnectorManager

        if not mgr:
            dbg.warn('  {} Id={} -> SEM ConnectorManager'.format(label, elem.Id.IntegerValue))
            return

        conns = list(mgr.Connectors)
        dbg.step('  {} Id={} -> {} conectores totais'.format(label, elem.Id.IntegerValue, len(conns)))
        for i, c in enumerate(conns):
            try:
                dbg.step('    [{}] domain={} connected={}'.format(i, c.Domain, c.IsConnected))
            except Exception as ex:
                dbg.warn('    [{}] erro: {}'.format(i, ex))
    except Exception as ex:
        dbg.warn('  _debug_element_connectors erro: {}'.format(ex))


def _get_circuit_prefix(elem):
    """Retorna o prefixo de circuito baseado no nome da família/tipo."""
    try:
        el_type = doc.GetElement(elem.GetTypeId())
        family_name = el_type.FamilyName.lower() if el_type else ""
        tn_param = el_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) if el_type else None
        type_name = tn_param.AsString().lower() if tn_param else ""
        full = family_name + " " + type_name
        if "cftv" in full or "camera" in full:
            return "Camera"
        if "wifi" in full:
            return "Wifi"
        if "telefon" in full or "phone" in full:
            return "Tel"
    except Exception:
        pass
    return "Dados"


def get_item_description(element, room_name):
    if not element:
        return room_name if room_name else "Ponto de Dados"

    try:
        el_type = doc.GetElement(element.GetTypeId())
        family_name = el_type.FamilyName.lower() if el_type else ""

        type_name = ""
        if el_type:
            tn_param = el_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = tn_param.AsString().lower() if tn_param else ""

        full = family_name + " " + type_name
        prefix = ""
        if "cftv" in full or "camera" in full:
            prefix = "Camera "
        elif "wifi" in full:
            prefix = "Ponto Wifi Teto "
        elif "rj45" in full or "dados" in full or "data" in full or "telecom" in full:
            prefix = "Ponto RJ45 "
        elif "telefon" in full or "phone" in full:
            prefix = "Ponto Tel "

        if prefix:
            return (prefix + (room_name if room_name else "")).strip()
    except Exception as ex:
        dbg.warn('get_item_description erro: {}'.format(ex))

    return room_name if room_name else "Ponto de Dados"


def select_and_configure_panel_data():
    dbg.section('DADOS - Selecionar Rack/Painel')
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            lf_electrical_core.PanelFilter(),
            "Selecione o Quadro/Rack de Telecom"
        )
    except OperationCanceledException:
        dbg.step('ESC - Cancelado')
        return
    except Exception as ex:
        dbg.fail('PickObject erro: {}'.format(ex))
        return

    panel = doc.GetElement(ref.ElementId)
    if not panel:
        dbg.fail('Elemento nao encontrado Id={}'.format(ref.ElementId.IntegerValue))
        return

    dbg.elem_info(panel, 'Rack selecionado')
    set_current_panel(panel.Id)
    forms.toast("Rack configurado: " + get_panel_name(panel), title="Dados")
    dbg.ok('Rack configurado: ' + get_panel_name(panel))


def create_grouped_circuit_data():
    dbg.section('DADOS - Criar Circuito Agrupado')
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o rack de telecom primeiro!")
        dbg.exit('create_grouped_circuit_data', 'SEM RACK')
        return

    dbg.elem_info(panel, 'Rack ativo')

    load_name = forms.ask_for_string(
        prompt="Nome para a rede (Ex: RACK-1):",
        title="Rede Agrupada"
    )
    if not load_name:
        dbg.step('Cancelado no nome')
        return

    dbg.step('Nome da rede: {}'.format(load_name))

    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            DataCategoryFilter(),
            "Selecione os pontos de rede -> " + load_name
        )
    except OperationCanceledException:
        dbg.step('ESC - Cancelado na selecao')
        return
    except Exception as ex:
        dbg.fail('PickObjects falhou: {}'.format(ex))
        forms.alert("Erro na seleção:\n" + str(ex))
        return

    if not refs:
        dbg.warn('Nenhum elemento selecionado')
        return

    refs_list = list(refs)
    dbg.step('{} elementos selecionados'.format(len(refs_list)))

    with Transaction(doc, "Criar Rede " + load_name) as t:
        t.Start()
        dbg.step('Transaction STARTED')

        ids = List[ElementId]()
        rooms = set()
        is_camera = False
        is_wifi = False
        skipped_no_connector = []
        freed_count = 0

        for r in refs_list:
            host = doc.GetElement(r.ElementId)
            dbg.elem_info(host, 'Processando host')
            _debug_element_connectors(host, 'host')

            vp = get_valid_electrical_elements(host, DATA_DOMAIN)
            dbg.step('  get_valid_electrical_elements -> {} par(es)'.format(len(vp)))

            if not vp:
                skipped_no_connector.append(str(r.ElementId.IntegerValue))
                dbg.warn('  Sem conectores validos: Id={}'.format(r.ElementId.IntegerValue))
                continue

            for sub_id, conns in vp:
                sub = doc.GetElement(sub_id)
                dbg.step('  Sub Id={} conns={}'.format(sub_id.IntegerValue, len(conns)))

                freed = ensure_element_is_free(sub)
                if freed:
                    freed_count += 1
                    dbg.ok('  Elemento liberado Id={}'.format(sub_id.IntegerValue))

                ids.Add(sub_id)

                try:
                    el_type = doc.GetElement(sub.GetTypeId())
                    if el_type:
                        fn = el_type.FamilyName.lower()
                        tn_param = el_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        tn = tn_param.AsString().lower() if tn_param else ""
                        if "camera" in fn or "camera" in tn:
                            is_camera = True
                        if "wifi" in fn or "wifi" in tn:
                            is_wifi = True
                except Exception as ex:
                    dbg.warn('  Erro ao detectar tipo: {}'.format(ex))

                rm = get_room_name(sub)
                if rm:
                    rooms.add(rm)
                    dbg.step('  Sala: {}'.format(rm))

        dbg.step('Total ids={} | sem_conector={} | liberados={}'.format(
            ids.Count, len(skipped_no_connector), freed_count))

        if ids.Count == 0:
            t.RollBack()
            msg = "Nenhum conector válido encontrado."
            if skipped_no_connector:
                msg += "\nElementos sem conector elétrico: {}".format(", ".join(skipped_no_connector))
            dbg.fail(msg)
            forms.alert(msg)
            return

        doc.Regenerate()
        dbg.step('doc.Regenerate() OK')

        try:
            with suppress_elec_dialog():
                circuit = ElectricalSystem.Create(doc, ids, _DATA_SYS_TYPE)
            dbg.ok('ElectricalSystem.Create OK tipo={} Id={}'.format(_DATA_SYS_TYPE, circuit.Id.IntegerValue))
        except Exception as ex:
            t.RollBack()
            dbg.fail('ElectricalSystem.Create FALHOU: {}'.format(ex))
            forms.alert("Falha ao criar circuito de dados:\n{}\n\n{}".format(ex, traceback.format_exc()))
            return

        try:
            circuit.SelectPanel(panel)
            dbg.ok('SelectPanel OK')
        except Exception as ex:
            dbg.warn('SelectPanel falhou: {}'.format(ex))

        set_param(circuit, ["Nome da carga", "Load Name"], load_name)

        prefix = ""
        if is_camera:
            prefix = "Camera "
        elif is_wifi:
            prefix = "Ponto Wifi Teto "

        room_str = " / ".join(sorted(rooms)) if rooms else ("Rede de Dados" if not prefix else "")
        desc = (prefix + room_str).strip()
        set_param(circuit, CIRCUIT_DESC_PARAMS, desc)

        t.Commit()
        dbg.ok('Transaction COMMITTED')
        forms.toast("Rede de dados criada: {} ({} elemento(s))".format(load_name, ids.Count), title="Dados")


def create_individual_circuits_data():
    dbg.section('DADOS - Criar Circuitos Individuais')
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o rack de telecom primeiro!")
        dbg.exit('create_individual_circuits_data', 'SEM RACK')
        return

    dbg.elem_info(panel, 'Rack ativo')

    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            DataCategoryFilter(),
            "Selecione os pontos de rede"
        )
    except OperationCanceledException:
        dbg.step('ESC - Cancelado na selecao')
        return
    except Exception as ex:
        dbg.fail('PickObjects falhou: {}'.format(ex))
        forms.alert("Erro na seleção:\n" + str(ex))
        return

    if not refs:
        dbg.warn('Nenhum elemento selecionado')
        return

    refs_list = list(refs)
    refs_list.reverse()
    dbg.step('{} elementos selecionados'.format(len(refs_list)))

    created = 0
    errors = []

    with Transaction(doc, "Pontos de Dados") as t:
        t.Start()
        dbg.step('Transaction STARTED')

        for r in refs_list:
            host = doc.GetElement(r.ElementId)
            dbg.elem_info(host, 'Processando host')
            _debug_element_connectors(host, 'host')

            vp = get_valid_electrical_elements(host, DATA_DOMAIN)
            dbg.step('  get_valid_electrical_elements -> {} par(es)'.format(len(vp)))

            if not vp:
                dbg.warn('  Sem conectores: Id={}'.format(r.ElementId.IntegerValue))
                errors.append("Id={} sem conector eletrico".format(r.ElementId.IntegerValue))
                continue

            for sub_id, conns in vp:
                sub = doc.GetElement(sub_id)
                dbg.step('  Sub Id={} conectores eletricos={}'.format(sub_id.IntegerValue, len(conns)))

                freed = ensure_element_is_free(sub)
                dbg.step('  ensure_element_is_free -> {}'.format(freed))

                rm = get_room_name(sub)
                dbg.step('  Sala: {}'.format(rm or '(sem sala)'))
                prefix = _get_circuit_prefix(sub)
                desc = (prefix + " " + rm).strip() if rm else prefix
                dbg.step('  Prefixo: {} -> desc: {}'.format(prefix, desc))

                for ci, c in enumerate(conns):
                    if c.IsConnected:
                        dbg.step('  Conector[{}] ja conectado, pulando'.format(ci))
                        continue

                    doc.Regenerate()
                    circuit = None

                    try:
                        with suppress_elec_dialog():
                            circuit = ElectricalSystem.Create(c, _DATA_SYS_TYPE)
                        dbg.ok('  Create(connector[{}]) OK Id={}'.format(ci, circuit.Id.IntegerValue))
                    except Exception as ex1:
                        dbg.warn('  Create(connector[{}]) falhou: {} — tentando via ElementId'.format(ci, ex1))
                        try:
                            ids_single = List[ElementId]()
                            ids_single.Add(sub_id)
                            with suppress_elec_dialog():
                                circuit = ElectricalSystem.Create(doc, ids_single, _DATA_SYS_TYPE)
                            dbg.ok('  Create(elem[{}]) OK Id={}'.format(ci, circuit.Id.IntegerValue))
                        except Exception as ex2:
                            dbg.fail('  Create FALHOU conector[{}] Id={}: {}'.format(ci, sub_id.IntegerValue, ex2))
                            errors.append('Id={} porta{}: {}'.format(sub_id.IntegerValue, ci + 1, ex2))
                            continue

                    try:
                        circuit.SelectPanel(panel)
                        dbg.ok('  SelectPanel OK')
                    except Exception as ex:
                        dbg.warn('  SelectPanel falhou: {}'.format(ex))

                    set_param(circuit, ["Nome da carga", "Load Name"], desc)
                    set_param(circuit, [BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                                        "Comentários", "Observações", "Comments"], desc)
                    dbg.ok('  Ponto criado: {} (conector[{}])'.format(desc, ci))
                    created += 1

        t.Commit()
        dbg.ok('Transaction COMMITTED')

    dbg.step('Resultado: {} criado(s), {} erro(s)'.format(created, len(errors)))

    if created > 0:
        msg = "{} ponto(s) criado(s).".format(created)
        if errors:
            msg += "\n{} falha(s): {}".format(len(errors), "; ".join(errors[:3]))
        forms.toast(msg, title="Dados")
    else:
        msg = "Nenhum circuito criado."
        if errors:
            msg += "\n\nErros:\n" + "\n".join(errors[:5])
        forms.alert(msg, title="Dados - Falha")


def main_menu():
    while True:
        quadro = get_current_panel()
        status = "Rack/Quadro: " + (get_panel_name(quadro) if quadro else "NENHUM")
        opcoes = OrderedDict([
            ("1. Selecionar Rack/Painel de Dados", select_and_configure_panel_data),
            ("2. Criar Circuitos de Dados (1 por porta)", create_individual_circuits_data),
            ("3. Criar Circuito Agrupado (todos num conduíte)", create_grouped_circuit_data),
            ("4. Sair", lambda: None),
        ])
        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(),
            message=status,
            title=u"\U0001f4e1 Dados - " + status
        )
        if not escolha or "Sair" in escolha:
            break
        try:
            opcoes[escolha]()
        except OperationCanceledException:
            pass
        except Exception as ex:
            dbg.fail('main_menu excecao: {}'.format(ex))
            forms.alert("Erro:\n" + traceback.format_exc())


if __name__ == "__main__":
    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        lf_electrical_core.show_settings()
    else:
        main_menu()
