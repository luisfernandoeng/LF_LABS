# -*- coding: utf-8 -*-
"""LF Electrical - Automação BIM (Versão Final + Regeneração Forçada)
Autor: Luís Fernando + Gemini
Correção: Força o Revit a ler a voltagem 220V antes de criar o circuito."""

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

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Sessão persistente
output = script.get_output()
if not hasattr(output, 'lf_panel_id'):
    output.lf_panel_id = None

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
                # Tenta definir como número (voltagem interna) ou string
                if isinstance(value, (int, float)):
                    p.Set(float(value)) 
                else:
                    p.Set(str(value))
                return True
            except:
                continue
    return False

def ensure_element_is_free(elem):
    """Remove circuitos antigos para liberar o conector."""
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

# ==================== CONFIGURAÇÃO DO QUADRO ====================

def configure_panel(panel):
    from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInParameter
    from Autodesk.Revit.DB.Electrical import DistributionSysType, CircuitNamingScheme

    messages = []
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

    target_sys = "127/220V Bifásico (2F+N+T)"
    target_nam = "Descrição/Nome da Carga"
    sys_id = None
    nam_id = None

    for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
        n = ""
        try: n = s.Name
        except: 
            p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p: n = p.AsString()
        if n == target_sys:
            sys_id = s.Id
            break

    for s in FilteredElementCollector(doc).OfClass(CircuitNamingScheme).ToElements():
        if s.Name == target_nam:
            nam_id = s.Id
            break

    with Transaction(doc, "Sistema e Nomenclatura") as t:
        t.Start()
        p = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and not p.IsReadOnly and sys_id:
            try: p.Set(sys_id)
            except: messages.append("Sistema: FALHA (Verifique fases)")

        p = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")
        if p and not p.IsReadOnly and nam_id:
            p.Set(nam_id)
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
        
        # 1. Configura a Voltagem e Limpa Conexões
        for r in refs:
            elem = doc.GetElement(r.ElementId)
            ensure_element_is_free(elem) # Limpa circuitos antigos
            
            if target_voltage:
                # Lista expandida de nomes possíveis para o parâmetro
                set_param(elem, ["Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], target_voltage)
            
            ids.Add(elem.Id)

        # 2. O PULO DO GATO: Força o Revit a atualizar os conectores AGORA
        if target_voltage:
            doc.Regenerate()

        try:
            # 3. Cria o Circuito (Agora o Revit sabe que é 220V)
            circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
            circuit.SelectPanel(panel)
            set_param(circuit, ["Nome da carga", "Load Name"], load_name)
            set_param(circuit, ["Descrição", "Comments"], "Automático")
            
            t.Commit()
            forms.toast("Circuito criado: " + load_name)
            
        except Exception as e:
            t.RollBack()
            # Se falhar, é porque a familia nao aceita mudar voltagem por instancia
            forms.alert("ERRO: O Revit não aceitou a voltagem.\n\n" + str(e) + "\n\nVerifique se a familia da tomada tem o parâmetro 'Tensão' ou 'Voltagem' editável na aba Propriedades.")

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

    created_count = 0
    
    with Transaction(doc, "Criar Circuitos " + prefixo) as t:
        t.Start()
        try:
            for r in refs:
                elem = doc.GetElement(r.ElementId)
                ensure_element_is_free(elem)
                
                # A. Configura 220V
                set_param(elem, ["Tensão", "Voltage", "Voltagem", "Tensão Nominal", "Volts"], 220)
                
                # B. Regenera (IMPORTANTE para atualizar conector)
                doc.Regenerate()
                
                # C. Cria Circuito
                ids = List[ElementId]()
                ids.Add(elem.Id)
                circuit = ElectricalSystem.Create(doc, ids, ElectricalSystemType.PowerCircuit)
                circuit.SelectPanel(panel)
                
                nome = prefixo + str(contador)
                set_param(circuit, ["Nome da carga", "Load Name"], nome)
                set_param(circuit, ["Descrição", "Comments"], "Carga Específica")
                
                contador += 1
                created_count += 1
            
            t.Commit()
            forms.alert("Sucesso! " + str(created_count) + " circuitos criados.")
            
        except Exception as e:
            t.RollBack()
            forms.alert("Erro: " + str(e))

# ==================== ASSOCIAR INTERRUPTORES ====================

def link_switches():
    SWITCHING_SYSTEM_TYPE_VALUE = 3 
    try:
        forms.toast("Selecione LUMINÁRIAS...")
        refs_lum = uidoc.Selection.PickObjects(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_LightingFixtures), "Selecione Luminárias")
        if not refs_lum: return
            
        lum_conns = List[Connector]()
        for r in refs_lum:
            lum = doc.GetElement(r.ElementId)
            if hasattr(lum, "MEPModel") and lum.MEPModel:
                for c in lum.MEPModel.ConnectorManager.Connectors:
                    if c.Domain == Domain.DomainElectrical:
                        lum_conns.Add(c)
                        break 
        
        if not lum_conns:
             forms.alert("Sem conectores nas luminárias.")
             return

        forms.toast("Selecione INTERRUPTOR...")
        ref_int = uidoc.Selection.PickObject(ObjectType.Element, CategoryFilter(BuiltInCategory.OST_LightingDevices), "Clique no interruptor")
        interruptor = doc.GetElement(ref_int.ElementId)
        
        switch_conn = None
        if hasattr(interruptor, "MEPModel") and interruptor.MEPModel:
            for c in interruptor.MEPModel.ConnectorManager.Connectors:
                if c.Domain == Domain.DomainElectrical and c.Direction == FlowDirectionType.Out:
                    switch_conn = c
                    break
        
        if not switch_conn:
            forms.alert("Interruptor sem saída.")
            return

        with Transaction(doc, "Associar Interruptor") as t:
            t.Start()
            sys = ElectricalSystem.Create(doc, lum_conns, SWITCHING_SYSTEM_TYPE_VALUE)
            sys.ConnectTo(switch_conn) 
            set_param(sys, ["Comando", "Switch ID", "Marca"], get_panel_name(interruptor))
            t.Commit()
            forms.alert("Sucesso! Comando: " + get_panel_name(interruptor))

    except Exception as e:
        forms.alert("Erro: " + str(e))

# ==================== MENU ====================

class PanelFilter(ISelectionFilter):
    def AllowElement(self, e):
        return e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment)
    def AllowReference(self, ref, pos): return False

def select_and_configure_panel():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, PanelFilter(), "Selecione o QUADRO")
        panel = doc.GetElement(ref.ElementId)
        success, msg = configure_panel(panel)
        if success:
            output.lf_panel_id = panel.Id
            forms.alert("Quadro Configurado!\n" + msg)
        else: forms.alert(msg)
    except: pass

def main_menu():
    while True:
        quadro = get_current_panel()
        status = "Quadro: " + (get_panel_name(quadro) if quadro else "NENHUM")

        def call_ilum():
            nome = forms.ask_for_string(default="Iluminação", prompt="Nome:", title="Iluminação")
            if nome: create_grouped_circuit(BuiltInCategory.OST_LightingFixtures, nome)

        def call_tomada():
            nome = forms.ask_for_string(default="Tomadas", prompt="Nome:", title="Tomadas Gerais")
            if not nome: return
            res = forms.CommandSwitchWindow.show(
                ["1 Polo (127V)", "2 Polos (220V)"], message="Selecione Fases:"
            )
            if res:
                # Corrigido: Garante que o valor é passado como número inteiro
                volt = 127 if "1 Polo" in res else 220
                create_grouped_circuit(BuiltInCategory.OST_ElectricalFixtures, nome, target_voltage=volt)

        opcoes = {
            "1. Selecionar/Configurar Quadro": select_and_configure_panel,
            "2. Criar Circuito Iluminação (Geral)": call_ilum,
            "3. Criar Circuito Tomadas (Geral 1F/2F)": call_tomada,
            "4. Criar Circuitos Específicos (AC/CH)": create_individual_circuits,
            "5. Associar Interruptores": link_switches,
            "6. Sair": lambda: None
        }

        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(),
            message=status,
            title="LF Electrical - Automação"
        )

        if not escolha or "Sair" in escolha: break
        try: opcoes[escolha]()
        except Exception as e: forms.alert(str(e))

if __name__ == "__main__":
    main_menu()