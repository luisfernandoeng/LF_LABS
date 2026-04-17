# coding: utf-8
"""Gerenciar Circuito - Adicionar/Remover elementos de circuitos
Autor: Luís Fernando
Comando standalone para atalho de teclado no Revit.
Selecione um elemento → membros ficam vermelhos → adicione ou remova.
"""

__title__ = "Gerenciar\nCircuito"
__author__ = "Luís Fernando"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, script

class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        failures = failuresAccessor.GetFailureMessages()
        for f in failures:
            if f.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(f)
        return FailureProcessingResult.Continue

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()


def set_red_override(view, element_ids, apply=True):
    if apply:
        # Busca o padrão de preenchimento sólido no documento
        solid_fill = None
        fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
        for fp in fill_patterns:
            try:
                if fp.GetFillPattern().IsSolidFill:
                    solid_fill = fp
                    break
            except: continue
        
        ogs = OverrideGraphicSettings()
        red = Color(255, 0, 0)
        
        # Configura as linhas em vermelho e com espessura maior
        ogs.SetProjectionLineColor(red)
        try: ogs.SetProjectionLineWeight(8)
        except: pass
        
        # Se encontrou o padrão sólido, aplica preenchimento (Foreground e Background)
        if solid_fill:
            # Foreground
            try:
                ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                ogs.SetSurfaceForegroundPatternColor(red)
                ogs.SetSurfaceForegroundPatternVisible(True)
            except: pass
            
            # Background (importante se houver outro padrão sobreposto)
            try:
                ogs.SetSurfaceBackgroundPatternId(solid_fill.Id)
                ogs.SetSurfaceBackgroundPatternColor(red)
                ogs.SetSurfaceBackgroundPatternVisible(True)
            except: pass

            # Fallback para versões mais antigas do Revit
            try:
                ogs.SetProjectionFillPatternId(solid_fill.Id)
                ogs.SetProjectionFillColor(red)
            except: pass
            
        for eid in element_ids:
            try: view.SetElementOverrides(eid, ogs)
            except: pass
    else:
        # Limpa todos os overrides aplicados
        blank = OverrideGraphicSettings()
        for eid in element_ids:
            try: view.SetElementOverrides(eid, blank)
            except: pass


def find_circuit(elem):
    if isinstance(elem, ElectricalSystem):
        return elem

    circuit = None

    # Método 1: MEPModel e métodos diretos
    try:
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            mep = elem.MEPModel
            # Tenta diversos métodos de obtenção de sistemas elétricos
            for method_name in ['GetElectricalSystems', 'GetAssignedElectricalSystems', 'ElectricalSystems']:
                if hasattr(mep, method_name):
                    res = getattr(mep, method_name)
                    # No caso de ElectricalSystems é propriedade, outros são métodos
                    systems = res() if method_name != 'ElectricalSystems' else res
                    if systems:
                        for sys in systems:
                            return sys
        
        # Tenta também no próprio elemento (caso de fios/wires)
        if hasattr(elem, 'ElectricalSystems'):
            for sys in elem.ElectricalSystems:
                return sys
    except: pass

    # Método 2: Conectores elétricos
    try:
        cm = None
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            cm = elem.MEPModel.ConnectorManager
        elif hasattr(elem, 'ConnectorManager'):
            cm = elem.ConnectorManager
        if cm:
            for conn in cm.Connectors:
                if conn.Domain == Domain.DomainElectrical:
                    if conn.IsConnected:
                        for ref_conn in conn.AllRefs:
                            if isinstance(ref_conn.Owner, ElectricalSystem):
                                return ref_conn.Owner
                    if hasattr(conn, 'MEPSystem') and conn.MEPSystem:
                        if isinstance(conn.MEPSystem, ElectricalSystem):
                            return conn.MEPSystem
    except: pass

    # Método 3: Busca global (Fallback lento)
    try:
        for es in FilteredElementCollector(doc).OfClass(ElectricalSystem).ToElements():
            try:
                if es.Elements:
                    for member in es.Elements:
                        if member.Id == elem.Id:
                            return es
            except: continue
    except: pass

    return None


def manage_circuit():
    elem = None
    selected_ids = list(uidoc.Selection.GetElementIds())
    if selected_ids:
        elem = doc.GetElement(selected_ids[0])
        
    if not elem:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                "Selecione um elemento que pertence a um circuito"
            )
            elem = doc.GetElement(ref.ElementId)
        except:
            return

    circuit = find_circuit(elem)
    if not circuit:
        forms.alert("Elemento não pertence a nenhum circuito elétrico.")
        return

    view = doc.ActiveView
    member_ids = []
    old_member_ids = []

    try:
        while True:
            # Atualiza a lista de membros a cada iteração
            member_ids = []
            try:
                for el in circuit.Elements:
                    member_ids.append(el.Id)
            except:
                break

            # Aplica/Atualiza o destaque vermelho
            with Transaction(doc, "Highlight Circuito") as t:
                t.Start()
                # Limpa destaque apenas de quem SAÍU do circuito (comparando com a lista anterior)
                to_clear = [eid for eid in old_member_ids if eid not in member_ids]
                if to_clear:
                    set_red_override(view, to_clear, apply=False)
                # Re-aplica no estado atual (garante que todos os membros atuais fiquem vermelhos)
                set_red_override(view, member_ids, apply=True)
                t.Commit()
            
            old_member_ids = list(member_ids)

            circ_num = "?"
            try: circ_num = circuit.CircuitNumber
            except: pass

            action = forms.CommandSwitchWindow.show(
                ['➕ Adicionar elementos', '➖ Remover elementos', '🗑️ Deletar Circuito', '◀ Sair'],
                message="Circuito: {} | {} membro(s)".format(circ_num, len(member_ids)),
                title="Gerenciar Circuito"
            )

            if not action or 'Sair' in action:
                break

            if 'Adicionar' in action:
                try:
                    add_refs = uidoc.Selection.PickObjects(
                        ObjectType.Element, "Selecione elementos para ADICIONAR"
                    )
                    if add_refs:
                        with Transaction(doc, "Adicionar ao Circuito") as t:
                            t.Start()
                            opts = t.GetFailureHandlingOptions()
                            opts.SetFailuresPreprocessor(WarningSwallower())
                            t.SetFailureHandlingOptions(opts)
                            
                            added = 0
                            for r in add_refs:
                                el = doc.GetElement(r.ElementId)
                                single_set = ElementSet()
                                single_set.Insert(el)
                                try:
                                    circuit.AddToCircuit(single_set)
                                    added += 1
                                except:
                                    # Fallback: tentar remover de outro circuito primeiro
                                    try:
                                        ex_circ = find_circuit(el)
                                        if ex_circ and ex_circ.Id != circuit.Id:
                                            ex_circ.RemoveFromCircuit(single_set)
                                            circuit.AddToCircuit(single_set)
                                            added += 1
                                    except: pass
                            t.Commit()
                            if added > 0:
                                forms.toast("✅ {} adicionado(s)".format(added))
                except: pass

            elif 'Remover' in action:
                try:
                    rem_refs = uidoc.Selection.PickObjects(
                        ObjectType.Element, "Selecione elementos para REMOVER"
                    )
                    if rem_refs:
                        with Transaction(doc, "Remover do Circuito") as t:
                            t.Start()
                            removed = 0
                            for r in rem_refs:
                                el = doc.GetElement(r.ElementId)
                                single_set = ElementSet()
                                single_set.Insert(el)
                                try:
                                    circuit.RemoveFromCircuit(single_set)
                                    removed += 1
                                except: pass
                            t.Commit()
                            if removed > 0:
                                forms.toast("✅ {} removido(s)".format(removed))
                except: pass

            elif 'Deletar' in action:
                if forms.alert("Deseja realmente deletar o circuito?", yes=True, no=True):
                    with Transaction(doc, "Deletar Circuito") as t:
                        t.Start()
                        # Limpa highlight antes de deletar
                        set_red_override(view, member_ids, apply=False)
                        try:
                            doc.Delete(circuit.Id)
                            forms.toast("✅ Circuito deletado!")
                            t.Commit()
                            return # Sai da função pois o circuito foi deletado
                        except:
                            t.RollBack()

    finally:
        # Garante que o destaque seja removido ao sair do comando
        try:
            with Transaction(doc, "Limpar Highlight") as t:
                t.Start()
                # Limpa tudo que foi destacado durante a sessão (membros atuais e anteriores)
                all_to_clear = list(set(member_ids + old_member_ids))
                set_red_override(view, all_to_clear, apply=False)
                t.Commit()
        except: pass


if __name__ == "__main__":
    manage_circuit()
