# coding: utf-8
"""Gerenciar Circuito - Adicionar/Remover elementos de circuitos
Autor: Luís Fernando
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
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


class ElectricalConnectorFilter(ISelectionFilter):
    """Só aceita elementos que possuem pelo menos um conector elétrico."""
    def AllowElement(self, elem):
        try:
            cm = None
            mep = getattr(elem, 'MEPModel', None)
            if mep:
                cm = getattr(mep, 'ConnectorManager', None)
            if cm is None:
                cm = getattr(elem, 'ConnectorManager', None)
            if cm:
                for conn in cm.Connectors:
                    if conn.Domain == Domain.DomainElectrical:
                        return True
        except:
            pass
        return False

    def AllowReference(self, ref, point):
        return True


class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for f in failuresAccessor.GetFailureMessages():
            if f.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(f)
        return FailureProcessingResult.Continue


def with_warning_swallower(t):
    opts = t.GetFailureHandlingOptions()
    opts.SetFailuresPreprocessor(WarningSwallower())
    t.SetFailureHandlingOptions(opts)


def set_red_override(view, element_ids, apply=True):
    if apply:
        solid_fill = None
        for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
            try:
                if fp.GetFillPattern().IsSolidFill:
                    solid_fill = fp
                    break
            except:
                continue

        ogs = OverrideGraphicSettings()
        red = Color(255, 0, 0)
        ogs.SetProjectionLineColor(red)
        try:
            ogs.SetProjectionLineWeight(8)
        except:
            pass
        if solid_fill:
            for setter in [
                lambda: (ogs.SetSurfaceForegroundPatternId(solid_fill.Id),
                         ogs.SetSurfaceForegroundPatternColor(red),
                         ogs.SetSurfaceForegroundPatternVisible(True)),
                lambda: (ogs.SetSurfaceBackgroundPatternId(solid_fill.Id),
                         ogs.SetSurfaceBackgroundPatternColor(red),
                         ogs.SetSurfaceBackgroundPatternVisible(True)),
                lambda: (ogs.SetProjectionFillPatternId(solid_fill.Id),
                         ogs.SetProjectionFillColor(red)),
            ]:
                try:
                    setter()
                except:
                    pass
        for eid in element_ids:
            try:
                view.SetElementOverrides(eid, ogs)
            except:
                pass
    else:
        blank = OverrideGraphicSettings()
        for eid in element_ids:
            try:
                view.SetElementOverrides(eid, blank)
            except:
                pass


def find_circuit(elem):
    if isinstance(elem, ElectricalSystem):
        return elem

    # Método 1: MEPModel e ElectricalSystems direto
    try:
        mep = getattr(elem, 'MEPModel', None)
        sources = [mep] if mep else []
        sources.append(elem)
        for src in sources:
            if src is None:
                continue
            for attr in ['GetElectricalSystems', 'GetAssignedElectricalSystems']:
                if hasattr(src, attr):
                    try:
                        result = getattr(src, attr)()
                        if result:
                            for sys in result:
                                return sys
                    except:
                        pass
            if hasattr(src, 'ElectricalSystems'):
                try:
                    for sys in src.ElectricalSystems:
                        return sys
                except:
                    pass
    except:
        pass

    # Método 2: Conectores elétricos
    try:
        cm = None
        mep = getattr(elem, 'MEPModel', None)
        if mep:
            cm = getattr(mep, 'ConnectorManager', None)
        if cm is None:
            cm = getattr(elem, 'ConnectorManager', None)
        if cm:
            for conn in cm.Connectors:
                if conn.Domain != Domain.DomainElectrical:
                    continue
                for attr in ['MEPSystem']:
                    sys = getattr(conn, attr, None)
                    if isinstance(sys, ElectricalSystem):
                        return sys
                if conn.IsConnected:
                    for ref in conn.AllRefs:
                        if isinstance(ref.Owner, ElectricalSystem):
                            return ref.Owner
    except:
        pass

    # Método 3: Varredura global (fallback lento)
    try:
        for es in FilteredElementCollector(doc).OfClass(ElectricalSystem).ToElements():
            try:
                if es.Elements:
                    for member in es.Elements:
                        if member.Id == elem.Id:
                            return es
            except:
                continue
    except:
        pass

    return None


def get_member_ids(circuit):
    ids = []
    try:
        for el in circuit.Elements:
            ids.append(el.Id)
    except:
        pass
    return ids


def is_in_circuit(circuit, elem_id):
    """Verifica se elemento ainda está no circuito (re-query após transação)."""
    try:
        # Recarrega o circuito do doc para ter estado atualizado
        fresh = doc.GetElement(circuit.Id)
        if fresh and fresh.Elements:
            for member in fresh.Elements:
                if member.Id == elem_id:
                    return True
    except:
        pass
    return False


def try_remove(circuit, el):
    """
    Tenta remover elemento do circuito com múltiplos métodos.
    Deve ser chamada DENTRO de uma transação aberta.
    Retorna True se removido com sucesso.
    """
    # Verifica se o elemento a ser removido é o Quadro (Panel) do circuito
    try:
        if circuit.BaseEquipment and circuit.BaseEquipment.Id == el.Id:
            circuit.SelectPanel(None)
            doc.Regenerate()
            return True
    except:
        pass

    single_set = ElementSet()
    single_set.Insert(el)

    # Método 1: RemoveFromCircuit padrão
    try:
        circuit.RemoveFromCircuit(single_set)
        doc.Regenerate()  # Força o Revit a atualizar a lista circuit.Elements
    except:
        pass

    # Verifica se foi removido de verdade
    still_in = False
    try:
        fresh_circuit = doc.GetElement(circuit.Id)
        for member in fresh_circuit.Elements:
            if member.Id == el.Id:
                still_in = True
                break
    except:
        pass

    if not still_in:
        return True

    # Método 3: Desconectar via conectores
    try:
        cm = None
        mep = getattr(el, 'MEPModel', None)
        if mep:
            cm = getattr(mep, 'ConnectorManager', None)
        if cm is None:
            cm = getattr(el, 'ConnectorManager', None)
        if cm:
            for conn in cm.Connectors:
                if conn.Domain != Domain.DomainElectrical or not conn.IsConnected:
                    continue
                to_disconnect = []
                for ref in conn.AllRefs:
                    if isinstance(ref.Owner, ElectricalSystem) and ref.Owner.Id == circuit.Id:
                        to_disconnect.append(ref)
                for ref in to_disconnect:
                    try:
                        conn.DisconnectFrom(ref)
                    except:
                        pass
        # Verifica
        for member in circuit.Elements:
            if member.Id == el.Id:
                return False
        return True
    except:
        pass

    return False


def _pick_one_by_one(message):
    """Loop de PickObject — retorna lista de Elements; ESC encerra."""
    elec_filter = ElectricalConnectorFilter()
    elements = []
    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                elec_filter,
                u"{} (ESC para finalizar)".format(message)
            )
            el = doc.GetElement(ref.ElementId)
            if el:
                elements.append(el)
        except:
            break
    return elements


def manage_circuit():
    elec_filter = ElectricalConnectorFilter()
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            elec_filter,
            u"Selecione um elemento que pertence a um circuito"
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
            member_ids = get_member_ids(circuit)

            with Transaction(doc, "Highlight Circuito") as t:
                t.Start()
                to_clear = [eid for eid in old_member_ids if eid not in member_ids]
                if to_clear:
                    set_red_override(view, to_clear, apply=False)
                set_red_override(view, member_ids, apply=True)
                t.Commit()

            old_member_ids = list(member_ids)

            circ_num = "?"
            try:
                circ_num = circuit.CircuitNumber
            except:
                pass

            action = forms.CommandSwitchWindow.show(
                [u'➕ Adicionar elementos', u'➖ Remover elementos',
                 u'🗑️ Deletar Circuito', u'◀ Sair'],
                message=u"Circuito: {} | {} membro(s)".format(circ_num, len(member_ids)),
                title=u"Gerenciar Circuito"
            )

            if not action or 'Sair' in action:
                break

            if 'Adicionar' in action:
                add_elements = _pick_one_by_one(u"Selecione elemento para ADICIONAR")
                if not add_elements:
                    continue

                added = 0
                failed = 0
                with Transaction(doc, "Adicionar ao Circuito") as t:
                    t.Start()
                    with_warning_swallower(t)
                    for el in add_elements:
                        single_set = ElementSet()
                        single_set.Insert(el)
                        ok = False
                        try:
                            circuit.AddToCircuit(single_set)
                            ok = True
                        except:
                            pass
                        if not ok:
                            # Tenta remover de outro circuito primeiro
                            try:
                                ex_circ = find_circuit(el)
                                if ex_circ and ex_circ.Id != circuit.Id:
                                    ex_circ.RemoveFromCircuit(single_set)
                                    doc.Regenerate()
                                    circuit.AddToCircuit(single_set)
                                    doc.Regenerate()
                                    ok = True
                            except:
                                pass
                        if ok:
                            added += 1
                        else:
                            failed += 1
                    t.Commit()

                if added > 0:
                    msg = u"✅ {} adicionado(s)".format(added)
                    if failed > 0:
                        msg += u" | ⚠️ {} falhou".format(failed)
                    forms.toast(msg)
                elif failed > 0:
                    forms.alert(u"Não foi possível adicionar {} elemento(s).".format(failed))

            elif 'Remover' in action:
                rem_elements = _pick_one_by_one(u"Selecione elemento para REMOVER")
                if not rem_elements:
                    continue

                removed = 0
                failed = 0
                rem_ids = [el.Id for el in rem_elements]
                with Transaction(doc, "Remover do Circuito") as t:
                    t.Start()
                    with_warning_swallower(t)
                    for el in rem_elements:
                        if try_remove(circuit, el):
                            removed += 1
                        else:
                            failed += 1
                    t.Commit()

                # Valida novamente fora da transação (estado real do documento)
                actually_removed = 0
                still_there = 0
                for eid in rem_ids:
                    if is_in_circuit(circuit, eid):
                        still_there += 1
                    else:
                        actually_removed += 1

                if actually_removed > 0:
                    msg = u"✅ {} removido(s)".format(actually_removed)
                    if still_there > 0:
                        msg += u" | ⚠️ {} não removido(s)".format(still_there)
                    forms.toast(msg)
                else:
                    forms.alert(
                        u"Nenhum elemento foi removido.\n"
                        u"Verifique se os elementos selecionados pertencem a este circuito."
                    )

            elif 'Deletar' in action:
                if forms.alert(u"Deseja realmente deletar o circuito?", yes=True, no=True):
                    with Transaction(doc, "Deletar Circuito") as t:
                        t.Start()
                        set_red_override(view, member_ids, apply=False)
                        try:
                            doc.Delete(circuit.Id)
                            t.Commit()
                            forms.toast(u"✅ Circuito deletado!")
                            return
                        except Exception as ex:
                            t.RollBack()
                            forms.alert(u"Não foi possível deletar o circuito:\n{}".format(ex))

    finally:
        try:
            with Transaction(doc, "Limpar Highlight") as t:
                t.Start()
                all_ids = list(set(member_ids + old_member_ids))
                set_red_override(view, all_ids, apply=False)
                t.Commit()
        except:
            pass


if __name__ == "__main__":
    manage_circuit()
