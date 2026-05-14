# coding: utf-8
"""Desacoplar - desfaz conexoes MEP entre elementos selecionados."""

__title__ = "Desacoplar"
__author__ = "Luis Fernando"

from pyrevit import forms, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

MODE_CHAIN = "Corrente"
MODE_ALL = "Tudo"
SELECT_ESC = "ESC"
SELECT_FINISH = "Concluir"


class AnyFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, ref, pos): return False


def _ct_connectors(elem):
    for getter in [lambda e: e.ConnectorManager,
                   lambda e: e.MEPModel.ConnectorManager]:
        try:
            cm = getter(elem)
            return [c for c in cm.Connectors
                    if c.Domain == Domain.DomainCableTrayConduit]
        except Exception:
            continue
    return []


def _pick_chain():
    elems = []
    while True:
        msg = (u"Clique no 1o elemento (ESC para cancelar)"
               if not elems else
               u"Clique no proximo elemento (ESC para finalizar)")
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, AnyFilter(), msg)
            elem = doc.GetElement(ref.ElementId)
            if elem and (not elems or elem.Id != elems[-1].Id):
                elems.append(elem)
        except OperationCanceledException:
            break
    return elems


def _pick_chain_finish():
    elems = []
    try:
        ref_first = uidoc.Selection.PickObject(
            ObjectType.Element, AnyFilter(),
            u"Clique no 1o elemento")
        elems.append(doc.GetElement(ref_first.ElementId))
    except OperationCanceledException:
        return []

    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, AnyFilter(),
            u"Selecione os proximos elementos e clique em Concluir")
    except OperationCanceledException:
        return []

    for ref in refs:
        elem = doc.GetElement(ref.ElementId)
        if elem and (not elems or elem.Id != elems[-1].Id):
            elems.append(elem)
    return elems


def _pick_configured():
    if _load_selection_mode() == SELECT_FINISH:
        return _pick_chain_finish()
    return _pick_chain()


def _other_refs(conn):
    refs = []
    try:
        for ref in conn.AllRefs:
            try:
                if ref.Owner and ref.Owner.Id != conn.Owner.Id:
                    refs.append(ref)
            except Exception:
                pass
    except Exception:
        pass
    return refs


def _disconnect_pair(el_a, el_b):
    count = 0
    for ca in _ct_connectors(el_a):
        for rb in _other_refs(ca):
            try:
                if rb.Owner.Id == el_b.Id:
                    ca.DisconnectFrom(rb)
                    count += 1
            except Exception:
                pass
    return count


def _disconnect_all(elem):
    count = 0
    for c in _ct_connectors(elem):
        refs = list(_other_refs(c))
        for r in refs:
            try:
                c.DisconnectFrom(r)
                count += 1
            except Exception:
                pass
    return count


def _load_mode():
    try:
        cfg = script.get_config()
        mode = getattr(cfg, "mode", MODE_CHAIN)
        if mode in [MODE_CHAIN, MODE_ALL]:
            return mode
    except Exception:
        pass
    return MODE_CHAIN


def _load_selection_mode():
    try:
        cfg = script.get_config()
        mode = getattr(cfg, "selection_mode", SELECT_ESC)
        if mode in [SELECT_ESC, SELECT_FINISH]:
            return mode
    except Exception:
        pass
    return SELECT_ESC


def _save_settings(mode, selection_mode):
    try:
        cfg = script.get_config()
        cfg.mode = mode
        cfg.selection_mode = selection_mode
        script.save_config()
    except Exception:
        pass


def show_settings():
    current_mode = _load_mode()
    current_selection = _load_selection_mode()

    mode_choices = [
        MODE_CHAIN + u" - desconecta apenas pares consecutivos clicados",
        MODE_ALL + u" - desconecta tudo dos elementos clicados",
    ]
    mode_choices = [
        (u"[Atual] " + c) if c.startswith(current_mode) else c
        for c in mode_choices
    ]
    chosen_mode = forms.SelectFromList.show(
        mode_choices,
        title=u"Desacoplar - O que desconectar",
        button_name=u"Salvar")
    if not chosen_mode:
        return

    selection_choices = [
        u"Selecao por ESC - clique um a um e finalize com ESC",
        u"Selecao por Concluir - selecao multipla do Revit, permite janela/caixa",
    ]
    selection_choices = [
        (u"[Atual] " + c) if (
            (current_selection == SELECT_ESC and c.startswith(u"Selecao por ESC")) or
            (current_selection == SELECT_FINISH and c.startswith(u"Selecao por Concluir"))
        ) else c
        for c in selection_choices
    ]
    chosen_selection = forms.SelectFromList.show(
        selection_choices,
        title=u"Desacoplar - Modo de selecao",
        button_name=u"Salvar")
    if not chosen_selection:
        return

    clean = chosen_mode.replace(u"[Atual] ", u"")
    mode = MODE_ALL if clean.startswith(MODE_ALL) else MODE_CHAIN
    clean_selection = chosen_selection.replace(u"[Atual] ", u"")
    selection_mode = (SELECT_FINISH if clean_selection.startswith(u"Selecao por Concluir")
                      else SELECT_ESC)
    _save_settings(mode, selection_mode)
    forms.toast(u"Modo salvo: {} / {}".format(mode, selection_mode),
                title=u"Desacoplar")


def main():
    mode = _load_mode()
    elems = _pick_configured()
    min_count = 1 if mode == MODE_ALL else 2
    if len(elems) < min_count:
        forms.alert(u"Selecao insuficiente para desacoplar.", title=u"Desacoplar")
        return

    t = Transaction(doc, u"Desacoplar")
    total = 0
    errors = []
    t.Start()
    try:
        if mode == MODE_ALL:
            for elem in elems:
                total += _disconnect_all(elem)
        else:
            for i in range(len(elems) - 1):
                count = _disconnect_pair(elems[i], elems[i + 1])
                if count:
                    total += count
                else:
                    errors.append(u"Id={} <-> Id={}: sem conexao direta".format(
                        elems[i].Id.IntegerValue,
                        elems[i + 1].Id.IntegerValue))
        t.Commit()
    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        forms.alert(u"Falha ao desacoplar:\n{}".format(e), title=u"Desacoplar")
        return

    if total:
        msg = u"{} conexao(oes) desfeita(s).".format(total)
        if errors:
            msg += u"\n{} aviso(s):\n{}".format(len(errors), u"\n".join(errors[:3]))
        forms.toast(msg, title=u"Desacoplar")
    else:
        forms.alert(u"Nenhuma conexao encontrada.\n\n" + u"\n".join(errors[:5]),
                    title=u"Desacoplar")


if __name__ == "__main__":
    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        show_settings()
    else:
        main()
