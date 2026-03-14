# coding: utf-8
"""LF Electrical - Automacao BIM (v2 - Com Sistema de Comando e Configs)
Autor: Luis Fernando
Features v2:
- Seleção de sistema de distribuição (não fixo)
- Shift+Click para configurações persistentes
- Pula letras O e S na nomenclatura (parecem 0 e 5)
- Auto-correção de tensão/fases pelo nome da família
- Sistema de Comando (Luminária → Interruptor)
- AC/CH: Ordem de clique respeitada via refs_list.reverse()
"""

__title__ = "LF Electrical"
__author__ = "Luís Fernando"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, script
import traceback

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Sessão persistente
output = script.get_output()
if not hasattr(output, 'lf_panel_id'):
    output.lf_panel_id = None

# ==================== CONSTANTES ====================

# Letras proibidas na nomenclatura (parecem 0 e 5 no desenho)
SKIP_LETTERS = {'O', 'S'}

# ==================== CONFIGURAÇÕES PERSISTENTES ====================

def load_config():
    """Carrega configurações salvas."""
    config = script.get_config()
    defaults = {
        'dist_system': '',
        'naming_scheme': '',
        'circuit_prefix': '',
    }
    result = {}
    for key, default in defaults.items():
        try:
            result[key] = config.get_option(key, default)
        except:
            result[key] = default
    return result

def save_config(settings):
    """Salva configurações."""
    config = script.get_config()
    for key, value in settings.items():
        config.set_option(key, value)
    script.save_config()

def show_settings():
    """Tela de configurações via Shift+Click (não-modal, sem WPF)."""
    settings = load_config()

    while True:
        current_dist = settings.get('dist_system', '') or '(nenhum)'
        current_naming = settings.get('naming_scheme', '') or '(nenhum)'
        current_prefix = settings.get('circuit_prefix', '') or '(nenhum)'

        opcoes = {
            "1. Sistema de Distribuição Padrão: " + current_dist: "dist_system",
            "2. Nomenclatura de Circuito Padrão: " + current_naming: "naming_scheme",
            "3. Prefixo de Circuito Padrão: " + current_prefix: "circuit_prefix",
            "4. Salvar e Sair": "save",
        }

        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(),
            message="Configurações LF Electrical",
            title="Shift+Click - Configurações"
        )

        if not escolha or opcoes.get(escolha) == "save":
            save_config(settings)
            forms.toast("Configurações salvas!", title="LF Electrical")
            break

        key = opcoes[escolha]

        if key == 'dist_system':
            # Lista todos os sistemas de distribuição do modelo
            all_systems = []
            for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
                try:
                    name = s.Name
                except:
                    p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    name = p.AsString() if p else "?"
                all_systems.append(name)

            if not all_systems:
                forms.alert("Nenhum sistema de distribuição encontrado no modelo.")
                continue

            chosen = forms.CommandSwitchWindow.show(
                sorted(all_systems),
                message="Selecione o sistema padrão:",
                title="Sistema de Distribuição"
            )
            if chosen:
                settings['dist_system'] = chosen

        elif key == 'naming_scheme':
            all_schemes = []
            for s in FilteredElementCollector(doc).OfClass(CircuitNamingScheme).ToElements():
                all_schemes.append(s.Name)

            if not all_schemes:
                forms.alert("Nenhuma nomenclatura de circuito encontrada.")
                continue

            chosen = forms.CommandSwitchWindow.show(
                sorted(all_schemes),
                message="Selecione a nomenclatura padrão:",
                title="Nomenclatura de Circuito"
            )
            if chosen:
                settings['naming_scheme'] = chosen

        elif key == 'circuit_prefix':
            val = forms.ask_for_string(
                default=settings.get('circuit_prefix', ''),
                prompt="Prefixo (ex: QDL, QFL):",
                title="Prefixo de Circuito"
            )
            if val is not None:
                settings['circuit_prefix'] = val

# ==================== FUNÇÕES AUXILIARES ====================

def get_panel_name(panel):
    try:
        for n in ["Nome do painel", "Panel Name", "Mark"]:
            p = panel.LookupParameter(n)
            if p and p.HasValue:
                return p.AsString()
        return panel.Name
    except:
        return "Quadro"

def get_current_panel():
    if output.lf_panel_id:
        try:
            panel = doc.GetElement(output.lf_panel_id)
            if panel and panel.IsValidObject:
                return panel
        except:
            pass
    output.lf_panel_id = None
    return None

def set_param(elem, names, value):
    """Tenta definir o parâmetro procurando por vários nomes."""
    for n in names:
        p = elem.LookupParameter(n)
        if p and not p.IsReadOnly:
            try:
                if isinstance(value, (int, float)):
                    p.Set(float(value))
                else:
                    p.Set(str(value))
                return True
            except:
                continue
    return False

def is_element_connected_to_panel(elem):
    """
    Verifica se o elemento já está ligado a algum quadro usando o parâmetro incorporado 'Painel'/'Panel'.
    Se esse parâmetro estiver preenchido, considera o elemento como já conectado e não deve ser ligado a outro quadro.
    """
    try:
        for n in ["Painel", "Panel"]:
            p = elem.LookupParameter(n)
            if p and p.HasValue:
                val = p.AsString()
                if val and val.strip():
                    return True
    except:
        pass
    return False

def ensure_element_is_free(elem):
    """Remove circuitos antigos para liberar o conector (compatível com antigo.py)."""
    try:
        if hasattr(elem, "MEPModel") and elem.MEPModel:
            sistemas = elem.MEPModel.ElectricalSystems
            if sistemas and not sistemas.IsEmpty:
                for sys in sistemas:
                    doc.Delete(sys.Id)
                return True
    except:
        pass
    return False

# REMOVED: auto_correct_voltage
def get_family_name(elem):
    """Retorna o nome da família do elemento."""
    try:
        if hasattr(elem, 'Symbol') and elem.Symbol:
            return elem.Symbol.FamilyName or ""
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            fam_param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            if fam_param and fam_param.HasValue:
                return fam_param.AsString() or ""
    except:
        pass
    return ""

# ==================== CONFIGURAÇÃO DO QUADRO ====================

def configure_panel(panel):
    from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInParameter
    from Autodesk.Revit.DB.Electrical import DistributionSysType, CircuitNamingScheme

    messages = []
    settings = load_config()

    # Nome do quadro
    new_name = forms.ask_for_string(
        default=get_panel_name(panel),
        prompt="Nome do Quadro (ex: QDL-01):",
        title="Configurar Quadro"
    )
    if not new_name: return False, "Cancelado"

    with Transaction(doc, "Nome do Quadro") as t:
        t.Start()
        set_param(panel, ["Nome do painel", "Panel Name", "Mark"], new_name)
        t.Commit()
    messages.append("Nome: " + new_name)

    # Coleta todos os sistemas de distribuição
    all_systems = {}
    for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
        n = ""
        try: n = s.Name
        except:
            p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p: n = p.AsString()
        if n:
            all_systems[n] = s.Id

    # Seleciona sistema de distribuição (usa config padrão como destaque)
    default_sys = settings.get('dist_system', '')
    sys_id = None

    if all_systems:
        # Ordena com o padrão primeiro se existir
        sys_names = sorted(all_systems.keys())
        if default_sys and default_sys in sys_names:
            sys_names.remove(default_sys)
            sys_names.insert(0, default_sys + " ★")
            all_systems[default_sys + " ★"] = all_systems[default_sys]

        chosen_sys = forms.CommandSwitchWindow.show(
            sys_names,
            message="Sistema de Distribuição para " + new_name + ":",
            title="Tipo de Quadro"
        )
        if chosen_sys:
            sys_id = all_systems.get(chosen_sys)
            messages.append("Sistema: " + chosen_sys.replace(" ★", ""))

    # Nomenclatura de circuito
    default_naming = settings.get('naming_scheme', '')
    nam_id = None

    all_namings = {}
    for s in FilteredElementCollector(doc).OfClass(CircuitNamingScheme).ToElements():
        all_namings[s.Name] = s.Id

    if all_namings:
        naming_names = sorted(all_namings.keys())
        if default_naming and default_naming in naming_names:
            naming_names.remove(default_naming)
            naming_names.insert(0, default_naming + " ★")
            all_namings[default_naming + " ★"] = all_namings[default_naming]

        chosen_naming = forms.CommandSwitchWindow.show(
            naming_names,
            message="Nomenclatura do Circuito:",
            title="Nomenclatura"
        )
        if chosen_naming:
            nam_id = all_namings.get(chosen_naming)
            messages.append("Nomenclatura: " + chosen_naming.replace(" ★", ""))

    with Transaction(doc, "Sistema e Nomenclatura") as t:
        t.Start()
        p = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and not p.IsReadOnly and sys_id:
            try: p.Set(sys_id)
            except: messages.append("Sistema: FALHA (Verifique fases)")

        p = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")
        if p and not p.IsReadOnly and nam_id:
            p.Set(nam_id)

        # Prefixo de circuito (se configurado)
        prefix = settings.get('circuit_prefix', '')
        if prefix:
            set_param(panel, ["Prefixo do circuito", "Circuit Prefix"], prefix)

        t.Commit()

    return True, "\n".join(messages)

# ==================== FILTROS ====================

class CategoryFilter(ISelectionFilter):
    def __init__(self, cat_id):
        self.cat_id = cat_id
    def AllowElement(self, e):
        return e.Category and e.Category.Id.IntegerValue == int(self.cat_id)
    def AllowReference(self, ref, pos): return False

# ==================== 1. CIRCUITOS AGRUPADOS (TUG/ILUM) ====================

def create_grouped_circuit(cat_id, load_name, target_voltage=None):
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    refs = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(cat_id), "Selecione os elementos -> " + load_name)
    if not refs: return

    with Transaction(doc, "Criar Circuito " + load_name) as t:
        t.Start()

        ids = List[ElementId]()
        skipped_no_connector = []
        skipped_already_connected = []
        disconnected = 0

        for r in refs:
            elem = doc.GetElement(r.ElementId)

            # 1. Verifica se o elemento já pertence a circuito ligado a algum quadro
            # usando o parâmetro incorporado "Painel"/"Panel"
            if is_element_connected_to_panel(elem):
                skipped_already_connected.append(str(r.ElementId.IntegerValue))
                continue

            # Se tiver circuitos sem painel, limpa para poder recircuitar
            try:
                if hasattr(elem, 'MEPModel') and elem.MEPModel:
                    existing = elem.MEPModel.ElectricalSystems
                    if existing and existing.Count > 0:
                        for es in existing:
                            try:
                                doc.Delete(es.Id)
                            except:
                                pass
                        disconnected += 1
            except:
                pass

            # 2. Verifica se tem conector elétrico válido
            has_connector = False
            try:
                if hasattr(elem, 'MEPModel') and elem.MEPModel:
                    cm = elem.MEPModel.ConnectorManager
                    if cm:
                        for c in cm.Connectors:
                            if c.Domain == Domain.DomainElectrical:
                                has_connector = True
                                break
            except:
                pass

            if has_connector:
                ids.Add(r.ElementId)
            else:
                skipped_no_connector.append(str(r.ElementId.IntegerValue))

        if ids.Count == 0:
            t.RollBack()
            msg = "Nenhum elemento disponível para criar circuito.\n"
            if skipped_already_connected:
                msg += "\nElementos já ligados a quadro (não alterados): " + ", ".join(skipped_already_connected)
            if skipped_no_connector:
                msg += "\nElementos sem conector elétrico: " + ", ".join(skipped_no_connector)
            output.print_md("### ⚠️ Nenhum elemento válido para novo circuito")
            output.print_md(msg)
            # Apenas log no console; sem prompt modal
            return

        if disconnected > 0:
            doc.Regenerate()

        try:
            circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
        except Exception as e:
            # Tratamento específico para erro "electComponents" (nenhum componente aceito)
            if "electComponents" in str(e):
                # Log detalhado por elemento para debug
                output.print_md("### ⚠️ Detalhes do erro electComponents (circuito agrupado)")
                for elem_id in ids:
                    elem = doc.GetElement(elem_id)
                    try:
                        cat_name = elem.Category.Name if elem.Category else "Sem categoria"
                    except:
                        cat_name = "Erro ao ler categoria"
                    fam_name = get_family_name(elem)
                    type_name = ""
                    try:
                        if hasattr(elem, "Name") and elem.Name:
                            type_name = elem.Name
                        else:
                            et = doc.GetElement(elem.GetTypeId())
                            if et:
                                type_name = et.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                    except:
                        type_name = ""

                    has_connector = False
                    try:
                        if hasattr(elem, "MEPModel") and elem.MEPModel:
                            cm = elem.MEPModel.ConnectorManager
                            if cm:
                                for c in cm.Connectors:
                                    if c.Domain == Domain.DomainElectrical:
                                        has_connector = True
                                        break
                    except:
                        pass

                    systems_info = []
                    try:
                        if hasattr(elem, "MEPModel") and elem.MEPModel:
                            existing = elem.MEPModel.ElectricalSystems
                            if existing and existing.Count > 0:
                                for es in existing:
                                    panel_name = ""
                                    try:
                                        if es.PanelId and es.PanelId != ElementId.InvalidElementId:
                                            pnl = doc.GetElement(es.PanelId)
                                            if pnl:
                                                panel_name = get_panel_name(pnl)
                                    except:
                                        panel_name = ""
                                    systems_info.append("SysId={} Tipo={} PanelId={} PanelName='{}'".format(
                                        es.Id.IntegerValue,
                                        es.SystemType,
                                        es.PanelId.IntegerValue if es.PanelId else -1,
                                        panel_name
                                    ))
                    except:
                        pass

                    output.print_md(
                        "- Elem ID {} | Cat='{}' | Fam='{}' | Tipo='{}' | ConectorElétrico={} | Sistemas=[{}]".format(
                            elem_id.IntegerValue,
                            cat_name,
                            fam_name,
                            type_name,
                            has_connector,
                            "; ".join(systems_info) if systems_info else "nenhum"
                        )
                    )

                t.RollBack()
                msg = "Nenhum dos elementos selecionados pôde criar circuito de força.\n"
                msg += "Verifique se as famílias são do tipo correto e possuem conectores elétricos válidos.\n"
                msg += "Detalhes completos foram registrados no console do pyRevit."
                output.print_md("### ⚠️ Circuito não criado (electComponents)")
                output.print_md("Elementos candidatos: {} | Já ligados a quadro: {} | Sem conector: {}".format(
                    ids.Count, len(skipped_already_connected), len(skipped_no_connector)))
                # Apenas log no console; sem prompt modal
                return
            else:
                t.RollBack()
                err_msg = traceback.format_exc()
                output.print_md("### ❌ ERRO ao criar circuito")
                output.print_md("Elementos candidatos: {} | Já ligados a quadro: {} | Sem conector: {}".format(
                    ids.Count, len(skipped_already_connected), len(skipped_no_connector)))
                if disconnected > 0:
                    output.print_md("Circuitos anteriores desconectados (sem quadro): {}".format(disconnected))
                output.print_md("```python\n{}\n```".format(err_msg))
                forms.alert("ERRO de Compatibilidade. Detalhes no console do pyRevit.\n\n" + str(e))
                return

        circuit.SelectPanel(panel)
        set_param(circuit, ["Nome da carga", "Load Name"], load_name)
        set_param(circuit, ["Descrição", "Comments"], "Automático")

        t.Commit()
        msg = "Circuito criado: " + load_name
        if skipped_already_connected:
            msg += "\n\nElementos já ligados a quadro (não alterados): " + ", ".join(skipped_already_connected)
        if skipped_no_connector:
            msg += "\nElementos sem conector elétrico: " + ", ".join(skipped_no_connector)
        forms.toast(msg)

# ==================== 2. CIRCUITOS INDIVIDUAIS (AC/CH) ====================

def create_individual_circuits():
    panel = get_current_panel()
    if not panel:
        forms.alert("Selecione o quadro primeiro!")
        return

    prefixo = forms.ask_for_string(default="AC", prompt="Prefixo (ex: AC, CH):", title="Prefixo")
    if not prefixo: return

    inicio_str = forms.ask_for_string(default="1", prompt="Início (ex: 1):", title="Contador")
    try: contador = int(inicio_str)
    except: contador = 1

    forms.toast("Selecione os equipamentos...")
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        CategoryFilter(BuiltInCategory.OST_ElectricalFixtures),
        "Selecione os equipamentos"
    )
    if not refs: return

    # FIX DA ORDEM: inverter lista selecionada para respeitar ordem de clique
    refs_list = list(refs)
    refs_list.reverse()

    created_count = 0
    skipped_no_connector = []
    skipped_already_connected = []

    with Transaction(doc, "Criar Circuitos " + prefixo) as t:
        t.Start()
        try:
            for r in refs_list:
                elem = doc.GetElement(r.ElementId)

                # Verifica se já está ligado a algum quadro (parâmetro Painel/Panel)
                if is_element_connected_to_panel(elem):
                    skipped_already_connected.append(str(r.ElementId.IntegerValue))
                    continue

                # Verifica conector elétrico
                has_connector = False
                try:
                    if hasattr(elem, 'MEPModel') and elem.MEPModel:
                        cm = elem.MEPModel.ConnectorManager
                        if cm:
                            for c in cm.Connectors:
                                if c.Domain == Domain.DomainElectrical:
                                    has_connector = True
                                    break
                except:
                    pass

                if not has_connector:
                    skipped_no_connector.append(str(r.ElementId.IntegerValue))
                    output.print_md("⚠️ Elemento ID {} sem conector elétrico, ignorado (AC/CH).".format(r.ElementId.IntegerValue))
                    continue

                # Libera de circuitos antigos sem quadro (se houver)
                ensure_element_is_free(elem)

                # Configura 220V (mantendo comportamento estável do antigo.py)
                set_param(elem, ["Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], 220)

                # Força o Revit a atualizar conectores após mudança de tensão
                doc.Regenerate()

                ids = List[ElementId]()
                ids.Add(elem.Id)

                try:
                    circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
                except Exception as ce:
                    # Tratamento específico para erro "electComponents"
                    if "electComponents" in str(ce):
                        output.print_md("⚠️ Elemento ID {} não pôde criar circuito elétrico (familia não reconhecida como componente elétrico válido).".format(r.ElementId.IntegerValue))
                        continue
                    else:
                        raise

                circuit.SelectPanel(panel)

                nome = prefixo + str(contador)
                set_param(circuit, ["Nome da carga", "Load Name"], nome)
                set_param(circuit, ["Descrição", "Comments"], "Carga Específica")

                contador += 1
                created_count += 1

            t.Commit()

            # Mensagens não-críticas: apenas console/toast, sem prompt modal
            if created_count > 0:
                msg = "Sucesso! " + str(created_count) + " circuitos criados."
                if skipped_already_connected:
                    msg += "\nElementos já ligados a quadro (não alterados): " + ", ".join(skipped_already_connected)
                if skipped_no_connector:
                    msg += "\nElementos sem conector elétrico: " + ", ".join(skipped_no_connector)
                output.print_md("### ✅ Circuitos AC/CH criados")
                output.print_md(msg)
                forms.toast("AC/CH: {} circuito(s) criado(s).".format(created_count))
            else:
                msg = "Nenhum circuito AC/CH foi criado."
                if skipped_already_connected:
                    msg += "\nElementos já ligados a quadro (não alterados): " + ", ".join(skipped_already_connected)
                if skipped_no_connector:
                    msg += "\nElementos sem conector elétrico: " + ", ".join(skipped_no_connector)
                output.print_md("### ⚠️ Nenhum circuito AC/CH criado")
                output.print_md(msg)

        except Exception as e:
            t.RollBack()
            err_msg = traceback.format_exc()
            output.print_md("### ❌ Erro ao criar circuitos Específicos (AC/CH)")
            output.print_md("```python\n{}\n```".format(err_msg))
            forms.alert("Erro ao criar circuitos individuais. Detalhes no console do pyRevit.")

def next_valid_letter(counter):
    """Retorna a próxima letra válida, pulando O e S."""
    while counter < 26:
        letter = chr(ord('A') + counter)
        if letter not in SKIP_LETTERS:
            return letter, counter
        counter += 1
    # Se passou de Z, volta pra A (segurança)
    return 'A', 0

# ==================== NOMEAR INTERRUPTORES (A, B, C... sem O, S) ====================

def name_switch():
    """
    Permite nomear interruptores em sequência alfabética.
    Pula as letras O e S (parecem 0 e 5 no desenho).
    """
    start_letter_str = forms.ask_for_string(
        default="A",
        prompt="Letra inicial (ex: A, C, T):\n(O e S são puladas automaticamente)",
        title="Nomear Interruptores"
    )
    if not start_letter_str or len(start_letter_str) != 1:
        return

    start_letter = start_letter_str.upper()
    if not start_letter.isalpha():
        forms.alert("Por favor, insira uma única letra do alfabeto.")
        return

    # Calcula o contador inicial baseado na letra (A=0, B=1, etc.)
    counter = ord(start_letter) - ord('A')

    # Se a letra inicial é proibida, pula pra próxima válida
    next_letter, counter = next_valid_letter(counter)

    while True:
        try:
            next_letter, counter = next_valid_letter(counter)
            ref = uidoc.Selection.PickObject(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_LightingDevices), "Selecione o INTERRUPTOR para nomear como: " + next_letter + " (ESC para sair)")
            interruptor = doc.GetElement(ref.ElementId)

            with Transaction(doc, "Nomear Interruptor") as t:
                t.Start()
                success = set_param(interruptor, ["ID do comando"], next_letter)
                t.Commit()

            if success:
                # Removido Toast
                counter += 1
            else:
                forms.alert("Não foi possível encontrar o parâmetro 'ID do comando' no interruptor selecionado.")
                break

        except Exception:
            break

# ==================== SISTEMA DE COMANDO (REMOVIDO A PEDIDO DO USUÁRIO) ====================

# ==================== MENU ====================

class PanelFilter(ISelectionFilter):
    def AllowElement(self, e):
        return e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment)
    def AllowReference(self, ref, pos): return False

def select_and_configure_panel():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, PanelFilter(), "Selecione o QUADRO")
        panel = doc.GetElement(ref.ElementId)
        # Se o quadro já tiver sistema de distribuição e nomenclatura definidos,
        # apenas seleciona sem reabrir a configuração completa.
        p_sys = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        p_nam = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")

        if p_sys and p_sys.HasValue and p_nam and p_nam.HasValue:
            output.lf_panel_id = panel.Id
            forms.alert("Quadro selecionado: " + get_panel_name(panel))
        else:
            success, msg = configure_panel(panel)
            if success:
                output.lf_panel_id = panel.Id
                forms.alert("Quadro Configurado!\n" + msg)
            else:
                forms.alert(msg)
    except: pass

def main_menu():
    while True:
        quadro = get_current_panel()
        status = "Quadro: " + (get_panel_name(quadro) if quadro else "NENHUM")

        def call_ilum():
            nome = forms.ask_for_string(default="1", prompt="Nome:", title="Iluminação")
            if nome: create_grouped_circuit(BuiltInCategory.OST_LightingFixtures, nome)

        def call_tomada():
            nome = forms.ask_for_string(default="T", prompt="Nome:", title="Tomadas Gerais")
            if nome:
                create_grouped_circuit(BuiltInCategory.OST_ElectricalFixtures, nome, target_voltage=None)

        opcoes = {
            "1. Selecionar/Configurar Quadro": select_and_configure_panel,
            "2. Criar Circuito Iluminação (Geral)": call_ilum,
            "3. Comando interruptor": name_switch,
            "4. Criar Circuito Tomadas (Geral 1F/2F)": call_tomada,
            "5. Criar Circuitos Específicos (AC/CH)": create_individual_circuits,
            "6. Sair": lambda: None
        }

        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(),
            message=status,
            title="LF Electrical - Automação"
        )

        if not escolha or "Sair" in escolha: break
        try: opcoes[escolha]()
        except Exception as e:
            if "aborted" not in str(e).lower() and "cancel" not in str(e).lower():
                err_msg = traceback.format_exc()
                output.print_md("### ❌ Erro na ferramenta: {}".format(escolha))
                output.print_md("```python\n{}\n```".format(err_msg))
                forms.alert("Um erro ocorreu. Veja a janela do pyRevit para detalhes.")

if __name__ == "__main__":
    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        show_settings()
    else:
        main_menu()