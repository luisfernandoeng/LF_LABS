# coding: utf-8
"""Gerenciar Circuito - Adicionar/Remover elementos de circuitos elétricos.

CAUSA RAIZ DO PROBLEMA ANTERIOR:
  AddToCircuit() no Revit percorre o GRAFO de conectividade elétrica inteiro
  quando os elementos não têm conexão física (Conectado:False).
  Resultado: 70+ elementos adicionados ao clicar em 1.

SOLUÇÃO:
  Em vez de AddToCircuit, RECRIAMOS o circuito com ElectricalSystem.Create()
  passando exatamente os conectores desejados — sem traversal de rede.
"""

__title__ = "Gerenciar\nCircuito"
__author__ = "Luís Fernando"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import System
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS DE CONECTORES
# ─────────────────────────────────────────────────────────────────────────────

def _get_cm(elem):
    try:
        mep = getattr(elem, 'MEPModel', None)
        if mep:
            cm = getattr(mep, 'ConnectorManager', None)
            if cm: return cm
    except Exception: pass
    try:
        cm = getattr(elem, 'ConnectorManager', None)
        if cm: return cm
    except Exception: pass
    return None


def _elec_connectors(elem):
    result = []
    cm = _get_cm(elem)
    if cm is None: return result
    try:
        for c in cm.Connectors:
            if c.Domain == Domain.DomainElectrical:
                result.append(c)
    except Exception: pass
    return result


def _has_elec_conn(elem):
    return len(_elec_connectors(elem)) > 0


# ─────────────────────────────────────────────────────────────────────────────
#  FILTRO DE SELEÇÃO
# ─────────────────────────────────────────────────────────────────────────────

class ElecFilter(ISelectionFilter):
    def AllowElement(self, elem): return _has_elec_conn(elem)
    def AllowReference(self, ref, point): return True


# ─────────────────────────────────────────────────────────────────────────────
#  BUSCA DE CIRCUITO
# ─────────────────────────────────────────────────────────────────────────────

def _find_circuits(elem):
    found = {}
    # Via AllRefs dos conectores (mais confiável)
    for conn in _elec_connectors(elem):
        try:
            for ref in conn.AllRefs:
                owner = ref.Owner
                if isinstance(owner, ElectricalSystem):
                    found[owner.Id.IntegerValue] = owner
        except Exception: pass
    if found: return list(found.values())
    # Via MEPModel
    try:
        sources = []
        mep = getattr(elem, 'MEPModel', None)
        if mep: sources.append(mep)
        sources.append(elem)
        for src in sources:
            for attr in ('GetElectricalSystems', 'GetAssignedElectricalSystems'):
                fn = getattr(src, attr, None)
                if fn:
                    try:
                        for es in (fn() or []):
                            found[es.Id.IntegerValue] = es
                    except Exception: pass
    except Exception: pass
    return list(found.values())


def find_circuit(elem):
    circuits = _find_circuits(elem)
    if not circuits: return None
    if len(circuits) == 1: return circuits[0]
    opts = [u"#{} (ID:{})".format(
        getattr(es, 'CircuitNumber', '?'), es.Id.IntegerValue
    ) for es in circuits]
    chosen = forms.SelectFromList.show(opts, title=u"Múltiplos circuitos", multiselect=False)
    if chosen is None: return None
    return circuits[opts.index(chosen)]


def _get_member_ids(circuit):
    ids = []
    try:
        fresh = doc.GetElement(circuit.Id)
        for el in (fresh.Elements or []):
            ids.append(el.Id)
    except Exception: pass
    return ids


# ─────────────────────────────────────────────────────────────────────────────
#  NÚCLEO: RECRIAR CIRCUITO COM MEMBROS EXATOS
#
#  Por que recriar em vez de AddToCircuit?
#  AddToCircuit faz traversal do grafo elétrico inteiro quando os elementos
#  não têm conexão física — adiciona 70+ elementos ao invés de 1.
#  ElectricalSystem.Create(IList[ElementId]) cria circuito com EXATAMENTE os
#  elementos passados, sem nenhum traversal de rede.
# ─────────────────────────────────────────────────────────────────────────────

def _rebuild_circuit(circuit, desired_elements):
    """
    Recria o circuito com exatamente desired_elements.
    Preserva SystemType, tensão (VoltageType), polos e painel.
    Deve ser chamado DENTRO de uma transação aberta.
    Retorna (novo_ElectricalSystem | None, msg_erro).
    """
    from System.Collections.Generic import List as CsList

    if not desired_elements:
        try:
            doc.Delete(circuit.Id)
            doc.Regenerate()
            return None, "circuito deletado (sem membros)"
        except Exception as e:
            return None, str(e)

    # ── Captura estado original ANTES de deletar ─────────────────────────
    sys_type = ElectricalSystemType.PowerCircuit
    try: sys_type = circuit.SystemType
    except Exception: pass

    panel = None
    try: panel = circuit.BaseEquipment
    except Exception: pass

    # Tensão: ElementId que aponta para a definição de VoltageType
    old_voltage_id = ElementId.InvalidElementId
    try:
        vt = circuit.VoltageType
        if vt: old_voltage_id = vt.Id
    except Exception: pass

    # Número de polos: 1=monofásico, 2=bifásico, 3=trifásico
    old_poles = 1
    try: old_poles = circuit.PolesNumber
    except Exception: pass

    # ── Monta IList[ElementId] — esta versão do Revit exige IList, não ConnectorSet ──
    id_list = CsList[ElementId]()
    for elem in desired_elements:
        if elem and elem.Id != ElementId.InvalidElementId:
            id_list.Add(elem.Id)

    if id_list.Count == 0:
        return None, u"Nenhum elemento válido encontrado"

    # ── Deleta circuito antigo ───────────────────────────────────────────
    try:
        doc.Delete(circuit.Id)
        doc.Regenerate()
    except Exception as e:
        return None, u"Falha ao deletar circuito antigo: {}".format(e)

    # ── Cria novo circuito com exatamente os IDs ────────────────────────
    try:
        new_circuit = ElectricalSystem.Create(doc, id_list, sys_type)
        doc.Regenerate()
    except Exception as e:
        return None, u"Falha ao criar novo circuito: {}".format(e)

    if new_circuit is None:
        return None, u"ElectricalSystem.Create retornou None"

    # ── Restaura tensão e polos do circuito original ─────────────────────
    if old_voltage_id != ElementId.InvalidElementId:
        try:
            vt_elem = doc.GetElement(old_voltage_id)
            if vt_elem:
                new_circuit.VoltageType = vt_elem
        except Exception: pass

    try:
        new_circuit.PolesNumber = old_poles
    except Exception: pass

    doc.Regenerate()

    # ── Re-associa ao mesmo painel ───────────────────────────────────────
    if panel:
        try:
            new_circuit.SelectPanel(panel)
            doc.Regenerate()
        except Exception: pass

    return new_circuit, u""


# ─────────────────────────────────────────────────────────────────────────────
#  ADD / REMOVE via recriação
# ─────────────────────────────────────────────────────────────────────────────

def _add_elements(circuit, elems_to_add):
    """
    Adiciona uma lista de elementos ao circuito de uma vez (1 rebuild só).
    Retorna (novo_circuit, ok, msg).
    """
    current_ids = _get_member_ids(circuit)
    current_ids_int = set(eid.IntegerValue for eid in current_ids)
    current_elems = [doc.GetElement(eid) for eid in current_ids if doc.GetElement(eid)]

    novos = [e for e in elems_to_add if e and e.Id.IntegerValue not in current_ids_int]
    ja_membros = len(elems_to_add) - len(novos)

    if not novos:
        return circuit, True, u"todos já eram membros"

    desired = current_elems + novos
    new_c, msg = _rebuild_circuit(circuit, desired)
    info = u"{} adicionado(s)".format(len(novos))
    if ja_membros:
        info += u" ({} já eram membros)".format(ja_membros)
    if msg:
        info += u" | " + msg
    if new_c is not None:
        return new_c, True, info
    return None, False, msg


def _remove_elements(circuit, elems_to_rem):
    """
    Remove uma lista de elementos do circuito de uma vez (1 rebuild só).
    Retorna (novo_circuit, ok, msg).
    """
    current_ids = _get_member_ids(circuit)
    current_ids_int = set(eid.IntegerValue for eid in current_ids)
    current_elems = [doc.GetElement(eid) for eid in current_ids if doc.GetElement(eid)]

    rem_ids_int = set(e.Id.IntegerValue for e in elems_to_rem if e)
    nao_membros = len([e for e in elems_to_rem if e and e.Id.IntegerValue not in current_ids_int])
    a_remover   = len(rem_ids_int) - nao_membros

    desired = [e for e in current_elems if e.Id.IntegerValue not in rem_ids_int]
    new_c, msg = _rebuild_circuit(circuit, desired)
    info = u"{} removido(s)".format(a_remover)
    if nao_membros:
        info += u" ({} não eram membros)".format(nao_membros)
    if msg:
        info += u" | " + msg
    ok = (new_c is not None) or ("deletado" in msg) or ("sem membros" in msg)
    return new_c, ok, info


# ─────────────────────────────────────────────────────────────────────────────
#  HIGHLIGHT
# ─────────────────────────────────────────────────────────────────────────────

def _solid_fill_id():
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill: return fp.Id
        except Exception: continue
    return ElementId.InvalidElementId

_FILL_ID = None

def _apply_highlight(view, eids, apply=True):
    global _FILL_ID
    if _FILL_ID is None: _FILL_ID = _solid_fill_id()
    if apply:
        ogs = OverrideGraphicSettings()
        red = Color(220, 50, 50)
        ogs.SetProjectionLineColor(red)
        try: ogs.SetProjectionLineWeight(6)
        except Exception: pass
        if _FILL_ID != ElementId.InvalidElementId:
            try:
                ogs.SetSurfaceForegroundPatternId(_FILL_ID)
                ogs.SetSurfaceForegroundPatternColor(red)
                ogs.SetSurfaceForegroundPatternVisible(True)
            except Exception: pass
    else:
        ogs = OverrideGraphicSettings()
    for eid in eids:
        try: view.SetElementOverrides(eid, ogs)
        except Exception: pass


# ─────────────────────────────────────────────────────────────────────────────
#  SELEÇÃO INTERATIVA (acumula cliques, ESC confirma)
# ─────────────────────────────────────────────────────────────────────────────

def _pick_one(message):
    """Retorna um único Element ou None (ESC cancela). Usado no seed inicial."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, ElecFilter(),
            u"{} — ESC para cancelar".format(message)
        )
        return doc.GetElement(ref.ElementId)
    except Exception:
        return None


def _pick_many(message):
    """
    Loop de PickObject — acumula elementos a cada clique.
    Clicar no mesmo elemento 2x o desfaz (toggle).
    ESC encerra e retorna a lista acumulada.
    """
    sel_filter = ElecFilter()
    collected  = {}   # int_id → Element

    while True:
        count = len(collected)
        hint  = u" [{} selecionado(s) — ESC para confirmar]".format(count) if count else u" [ESC para cancelar]"
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, sel_filter,
                u"{}{}".format(message, hint)
            )
        except Exception:
            break   # ESC ou qualquer erro encerra o loop

        try:
            el  = doc.GetElement(ref.ElementId)
            key = ref.ElementId.IntegerValue
            if el:
                if key in collected:
                    del collected[key]   # segundo clique = deseleciona
                    forms.toast(u"↩ Removido da seleção")
                else:
                    collected[key] = el
                    forms.toast(u"✔ {} selecionado(s)".format(len(collected)))
        except Exception:
            pass

    return list(collected.values())


# ─────────────────────────────────────────────────────────────────────────────
#  LOOP PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

class _WarnSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        for f in fa.GetFailureMessages():
            if f.GetSeverity() == FailureSeverity.Warning:
                fa.DeleteWarning(f)
        return FailureProcessingResult.Continue


def manage_circuit():
    # 1. Seleciona elemento base (deve ter circuito)
    seed = _pick_one(u"Clique num elemento que pertence a um circuito")
    if not seed: return

    circuit = find_circuit(seed)
    if not circuit:
        forms.alert(u"O elemento não pertence a nenhum circuito elétrico.",
                    title=u"Gerenciar Circuito")
        return

    view = doc.ActiveView
    highlighted = []

    try:
        while True:
            if circuit is None:
                forms.toast(u"Circuito deletado — encerrando.")
                break

            member_ids = _get_member_ids(circuit)

            # Atualiza highlight
            with Transaction(doc, u"Highlight Circuito") as t:
                t.Start()
                to_clear = [eid for eid in highlighted if eid not in member_ids]
                _apply_highlight(view, to_clear, apply=False)
                _apply_highlight(view, member_ids, apply=True)
                t.Commit()
            highlighted = list(member_ids)

            circ_num = u"?"
            try: circ_num = circuit.CircuitNumber or u"?"
            except Exception: pass

            panel_info = u""
            try:
                base = circuit.BaseEquipment
                if base: panel_info = u" | {}".format(base.Name)
            except Exception: pass

            action = forms.CommandSwitchWindow.show(
                [u"➕  Adicionar elemento",
                 u"➖  Remover elemento",
                 u"🗑️  Deletar circuito",
                 u"✖  Sair"],
                message=u"Circuito: {}{} | {} membro(s)".format(
                    circ_num, panel_info, len(member_ids)),
                title=u"Gerenciar Circuito"
            )

            if not action or u"Sair" in action or u"✖" in action:
                break

            # ── ADICIONAR ─────────────────────────────────────────────────
            if u"Adicionar" in action:
                # Loop: clique para acumular, ESC para confirmar lote
                elems = _pick_many(
                    u"ADICIONAR ao circuito {} — clique nos elementos".format(circ_num)
                )
                if not elems: continue

                with Transaction(doc, u"Adicionar ao Circuito") as t:
                    opts = t.GetFailureHandlingOptions()
                    opts.SetFailuresPreprocessor(_WarnSwallower())
                    t.SetFailureHandlingOptions(opts)
                    t.Start()
                    new_c, ok, msg = _add_elements(circuit, elems)
                    if ok:
                        t.Commit()
                        if new_c: circuit = new_c
                        forms.toast(u"✅ {}".format(msg))
                    else:
                        t.RollBack()
                        forms.alert(u"Não foi possível adicionar:\n{}".format(msg),
                                    title=u"Erro")

            # ── REMOVER ────────────────────────────────────────────────────
            elif u"Remover" in action:
                elems = _pick_many(
                    u"REMOVER do circuito {} — clique nos elementos".format(circ_num)
                )
                if not elems: continue

                with Transaction(doc, u"Remover do Circuito") as t:
                    opts = t.GetFailureHandlingOptions()
                    opts.SetFailuresPreprocessor(_WarnSwallower())
                    t.SetFailureHandlingOptions(opts)
                    t.Start()
                    new_c, ok, msg = _remove_elements(circuit, elems)
                    if ok:
                        t.Commit()
                        circuit = new_c
                        forms.toast(u"✅ {}".format(msg))
                        if circuit is None: break
                    else:
                        t.RollBack()
                        forms.alert(u"Não foi possível remover:\n{}".format(msg),
                                    title=u"Erro")

            # ── DELETAR ────────────────────────────────────────────────────
            elif u"Deletar" in action:
                if not forms.alert(u"Deletar circuito {}?".format(circ_num),
                                   title=u"Confirmar", yes=True, no=True):
                    continue
                with Transaction(doc, u"Deletar Circuito") as t:
                    t.Start()
                    _apply_highlight(view, highlighted, apply=False)
                    try:
                        doc.Delete(circuit.Id)
                        t.Commit()
                        forms.toast(u"✅ Circuito {} deletado!".format(circ_num))
                        highlighted = []
                        circuit = None
                        break
                    except Exception as ex:
                        t.RollBack()
                        forms.alert(u"Erro ao deletar:\n{}".format(ex), title=u"Erro")

    finally:
        try:
            with Transaction(doc, u"Limpar Highlight") as t:
                t.Start()
                _apply_highlight(view, list(set(highlighted)), apply=False)
                t.Commit()
        except Exception:
            pass


if __name__ == "__main__":
    manage_circuit()
