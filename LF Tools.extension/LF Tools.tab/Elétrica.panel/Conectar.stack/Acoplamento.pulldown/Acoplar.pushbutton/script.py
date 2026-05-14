# coding: utf-8
"""Acoplar — conecta um elemento base a múltiplos via conector MEP de bandeja/eletroduto.

Casos tratados:
  A)  Target sem conector CT                     → insere adaptador no target, conecta
  A') Base sem conector CT                       → insere adaptador na base, conecta
  B)  Ambos têm conector CT do mesmo perfil      → ConnectTo direto, sem adaptador
  C)  Ambos têm conector CT de perfis diferentes → tenta ConnectTo; se falhar, usa adaptador ponte
  D)  Nenhum tem conector CT                     → erro
"""

__title__ = "Acoplar"
__author__ = "Luís Fernando"

from pyrevit import forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

import sys
import os

_ADAPTER_CATS = {
    int(BuiltInCategory.OST_CableTrayFitting),
    int(BuiltInCategory.OST_ConduitFitting),
    int(BuiltInCategory.OST_ElectricalEquipment),
    int(BuiltInCategory.OST_ElectricalFixtures),
    int(BuiltInCategory.OST_LightingFixtures),
    int(BuiltInCategory.OST_LightingDevices),
    int(BuiltInCategory.OST_DataDevices),
    int(BuiltInCategory.OST_CommunicationDevices),
    int(BuiltInCategory.OST_FireAlarmDevices),
    int(BuiltInCategory.OST_SecurityDevices),
    int(BuiltInCategory.OST_GenericModel),
    int(BuiltInCategory.OST_SpecialityEquipment),
}


class AnyFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, ref, pos): return False


# ── Helpers básicos ───────────────────────────────────────────────────────────

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


def _profile(c):
    try:
        return c.Shape
    except Exception:
        return None


def _location(elem):
    try:
        loc = elem.Location
        if hasattr(loc, "Point"):
            return loc.Point
        if hasattr(loc, "Curve"):
            return loc.Curve.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        bb = elem.get_BoundingBox(doc.ActiveView)
        if bb:
            return XYZ((bb.Min.X + bb.Max.X) / 2.0,
                       (bb.Min.Y + bb.Max.Y) / 2.0,
                       bb.Min.Z)
    except Exception:
        pass
    return XYZ.Zero


def _level(elem):
    try:
        lid = elem.LevelId
        if lid != ElementId.InvalidElementId:
            return doc.GetElement(lid)
    except Exception:
        pass
    return None


def _best_conn(conns, ref_pt):
    """Conector livre mais próximo de ref_pt; fallback para qualquer um."""
    free = [c for c in conns if not c.IsConnected]
    pool = free if free else list(conns)
    return min(pool, key=lambda c: c.Origin.DistanceTo(ref_pt)) if pool else None


def _conn_by_profile(conns, profile_type):
    """Primeiro conector livre com o profile pedido; fallback sem filtro de profile."""
    for c in conns:
        if not c.IsConnected and _profile(c) == profile_type:
            return c
    for c in conns:
        if _profile(c) == profile_type:
            return c
    return None


def _try_connect(conn_a, conn_b):
    for a, b in [(conn_a, conn_b), (conn_b, conn_a)]:
        try:
            a.ConnectTo(b)
            return True
        except Exception:
            continue
    return False


def _is_cable_tray(elem):
    try:
        return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray)
    except Exception:
        return False


def _find_ct_for_point(type_id, pt):
    """Acha o segmento de eletrocalha (mesmo tipo) cuja curva XY é mais próxima de pt.

    Necessário após splits: a eletrocalha original é deletada e substituída por segmentos.
    Projetamos pt no plano Z da eletrocalha para ignorar diferença de altura.
    """
    best_ct, best_d = None, float('inf')
    for ct in FilteredElementCollector(doc).OfClass(CableTray).ToElements():
        try:
            if ct.GetTypeId() != type_id:
                continue
            crv = ct.Location.Curve
            z = crv.GetEndPoint(0).Z
            flat_pt = XYZ(pt.X, pt.Y, z)
            try:
                proj = crv.Project(flat_pt)
                d = proj.Distance if proj else float('inf')
            except Exception:
                mid = crv.Evaluate(0.5, True)
                d = XYZ(mid.X - pt.X, mid.Y - pt.Y, 0.0).GetLength()
            if d < best_d:
                best_d, best_ct = d, ct
        except Exception:
            pass
    return best_ct


class _AggressiveSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        failuresAccessor.DeleteAllWarnings()
        has_error = False
        fail_list = failuresAccessor.GetFailureMessages()
        for f in fail_list:
            try:
                if f.GetSeverity() == FailureSeverity.Error:
                    failuresAccessor.ResolveFailure(f)
                    has_error = True
            except Exception:
                pass
        if has_error:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue

def _make_t(name=u"Acoplar"):
    t = Transaction(doc, name)
    try:
        ops = t.GetFailureHandlingOptions()
        ops.SetClearAfterRollback(True)
        try:
            ops.SetFailuresPreprocessor(_AggressiveSwallower())
        except Exception:
            pass
        t.SetFailureHandlingOptions(ops)
    except Exception:
        pass
    return t


def _adapter_families():
    result = {}
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements():
        if sym.Category and sym.Category.Id.IntegerValue in _ADAPTER_CATS:
            key = u"{} — {}".format(sym.FamilyName, sym.Name)
            result[key] = sym
    return result


def _place_instance(sym, pt, lv):
    if lv:
        return doc.Create.NewFamilyInstance(pt, sym, lv, StructuralType.NonStructural)
    return doc.Create.NewFamilyInstance(pt, sym, StructuralType.NonStructural)


def _pick_adapter(title_extra=u""):
    fam_dict = _adapter_families()
    if not fam_dict:
        forms.alert(
            u"Nenhuma família adequada encontrada no modelo.\n"
            u"Carregue uma família de fitting de bandeja/eletroduto.",
            title="Acoplar — Sem Família")
        return None, None
    title = u"Família do Conector Falso" + (u" — " + title_extra if title_extra else u"")
    chosen = forms.SelectFromList.show(sorted(fam_dict.keys()),
                                       title=title,
                                       button_name=u"Usar como Adaptador")
    if not chosen:
        return None, None
    return chosen, fam_dict[chosen]


def _copy_parameters(source_elem, target_inst):
    if not source_elem or not target_inst:
        return
    try:
        w_p = source_elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        h_p = source_elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        if w_p and h_p:
            tgt_w = target_inst.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
            if tgt_w and not tgt_w.IsReadOnly:
                tgt_w.Set(w_p.AsDouble())
            else:
                for p_name in ["Largura", "Largura 1", "Width"]:
                    p_comp = target_inst.LookupParameter(p_name)
                    if p_comp and not p_comp.IsReadOnly:
                        p_comp.Set(w_p.AsDouble())
                        break

            tgt_h = target_inst.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
            if tgt_h and not tgt_h.IsReadOnly:
                tgt_h.Set(h_p.AsDouble())
            else:
                p_alt = target_inst.LookupParameter("Altura")
                if p_alt and not p_alt.IsReadOnly:
                    p_alt.Set(h_p.AsDouble())
    except Exception:
        pass

    for p_src in source_elem.Parameters:
        if p_src.IsReadOnly or not p_src.HasValue:
            continue
        try:
            p_tgt = target_inst.LookupParameter(p_src.Definition.Name)
            if not p_tgt or p_tgt.IsReadOnly:
                continue
            st = p_src.StorageType
            if st == StorageType.String:
                p_tgt.Set(p_src.AsString())
            elif st == StorageType.Integer:
                p_tgt.Set(p_src.AsInteger())
            elif st == StorageType.Double:
                if p_src.Definition.Name not in ["Largura", "Altura", "Comprimento"]:
                    p_tgt.Set(p_src.AsDouble())
            elif st == StorageType.ElementId:
                p_tgt.Set(p_src.AsElementId())
        except Exception:
            pass

def _orient_adapter(inst, conn_dest):
    try:
        import math
        target_dir = conn_dest.CoordinateSystem.BasisZ.Negate()

        my_conns = []
        try:
            my_conns = list(inst.MEPModel.ConnectorManager.Connectors)
        except Exception:
            try:
                my_conns = list(inst.ConnectorManager.Connectors)
            except Exception:
                pass

        if not my_conns:
            return

        my_conn = None
        for c in my_conns:
            if c.Shape == conn_dest.Shape:
                my_conn = c
                break
        if not my_conn:
            my_conn = my_conns[0]

        my_dir = my_conn.CoordinateSystem.BasisZ

        angle_my = math.atan2(my_dir.Y, my_dir.X)
        angle_target = math.atan2(target_dir.Y, target_dir.X)
        rot = angle_target - angle_my

        while rot > math.pi: rot -= 2 * math.pi
        while rot < -math.pi: rot += 2 * math.pi

        if abs(rot) > 0.001:
            pt = inst.Location.Point
            axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, rot)
            doc.Regenerate()
    except Exception:
        pass


# ── Split helper ──────────────────────────────────────────────────────────────

def _split_cabletray(doc, cable_tray, split_pt, fallback_level_id=None):
    from Autodesk.Revit.DB import SubTransaction, StorageType
    sub = SubTransaction(doc)
    sub.Start()
    try:
        crv = cable_tray.Location.Curve
        p0  = crv.GetEndPoint(0)
        p1  = crv.GetEndPoint(1)
        proj = crv.Project(split_pt)
        s_pt = proj.XYZPoint if proj else XYZ(split_pt.X, split_pt.Y, p0.Z)
        ct_type_id = cable_tray.GetTypeId()
        level_id = fallback_level_id
        for bip in [BuiltInParameter.RBS_START_LEVEL_PARAM, BuiltInParameter.FAMILY_LEVEL_PARAM]:
            try:
                lp = cable_tray.get_Parameter(bip)
                if lp and lp.AsElementId() != ElementId.InvalidElementId:
                    level_id = lp.AsElementId()
                    break
            except Exception:
                continue
        params_to_copy = []
        for p in cable_tray.Parameters:
            try:
                if p.IsReadOnly: continue
                st = p.StorageType
                if st == StorageType.Double: params_to_copy.append((p.Id, st, p.AsDouble()))
                elif st == StorageType.Integer: params_to_copy.append((p.Id, st, p.AsInteger()))
                elif st == StorageType.String:
                    s = p.AsString()
                    if s is not None: params_to_copy.append((p.Id, st, s))
                elif st == StorageType.ElementId: params_to_copy.append((p.Id, st, p.AsElementId()))
            except Exception: continue
        p0_neighbors = []
        p1_neighbors = []
        try:
            for c in cable_tray.ConnectorManager.Connectors:
                near_p0 = c.Origin.DistanceTo(p0) < c.Origin.DistanceTo(p1)
                try:
                    for ref in c.AllRefs:
                        if ref.Owner.Id != cable_tray.Id:
                            if near_p0: p0_neighbors.append((ref.Owner.Id, ref.Origin))
                            else: p1_neighbors.append((ref.Owner.Id, ref.Origin))
                except Exception: pass
        except Exception: pass
        doc.Delete(cable_tray.Id)
        doc.Regenerate()
        def _make_ct(pa, pb):
            if pa.DistanceTo(pb) < 0.1: return None
            ct = CableTray.Create(doc, ct_type_id, pa, pb, level_id)
            if ct:
                for (pid, st, val) in params_to_copy:
                    try:
                        p = ct.get_Parameter(pid)
                        if p and not p.IsReadOnly:
                            p.Set(val)
                    except Exception: pass
            return ct
        ct1 = _make_ct(p0, s_pt)
        ct2 = _make_ct(s_pt, p1)
        if not ct1 and not ct2:
            sub.RollBack()
            return None, None
        doc.Regenerate()
        def _conn_near(ct, pt):
            if not ct: return None
            best, bd = None, float('inf')
            try:
                for c in ct.ConnectorManager.Connectors:
                    d = c.Origin.DistanceTo(pt)
                    if d < bd: bd, best = d, c
            except Exception: pass
            return best
        def _reconnect_endpoint(ct_seg, neighbor_infos, endpoint_pt):
            if not ct_seg or not neighbor_infos: return
            seg_conn = _conn_near(ct_seg, endpoint_pt)
            if not seg_conn: return
            for (el_id, ref_origin) in neighbor_infos:
                try:
                    el = doc.GetElement(el_id)
                    cm_neighbor = el.ConnectorManager if hasattr(el, "ConnectorManager") else el.MEPModel.ConnectorManager
                    best_nc, best_d = None, float('inf')
                    for nc in cm_neighbor.Connectors:
                        d = nc.Origin.DistanceTo(ref_origin)
                        if d < best_d: best_d, best_nc = d, nc
                    if best_nc and best_d < 0.5: seg_conn.ConnectTo(best_nc)
                except Exception: pass
        _reconnect_endpoint(ct1, p0_neighbors, p0)
        _reconnect_endpoint(ct2, p1_neighbors, p1)
        sub.Commit()
        return _conn_near(ct1, s_pt), _conn_near(ct2, s_pt)
    except Exception:
        try: sub.RollBack()
        except Exception: pass
        return None, None


# ── Operações de conexão ──────────────────────────────────────────────────────

def _insert_adapter_simple(sym, insert_pt, lv, conn_target, pt_target, source_elem=None):
    """
    CASO A — insere adaptador em insert_pt e conecta ao conn_target.
    Retorna (instância, conectado:bool, conn_adapter).
    """
    t = _make_t(u"Acoplar — Adaptador Simples")
    t.Start()
    try:
        if not sym.IsActive:
            sym.Activate()
            doc.Regenerate()
        inst = _place_instance(sym, insert_pt, lv)
        doc.Regenerate()

        if source_elem:
            _copy_parameters(source_elem, inst)

        if conn_target:
            _orient_adapter(inst, conn_target)

        adapter_conns = _ct_connectors(inst)

        is_ct = (conn_target and conn_target.Owner and
                 conn_target.Owner.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray))
        connected = False
        conn_adapter = None

        if is_ct:
            c1_tray, c2_tray = _split_cabletray(doc, conn_target.Owner, insert_pt, lv)
            remaining = list(adapter_conns)
            for tray_conn in [c1_tray, c2_tray]:
                if tray_conn and remaining:
                    best = max(remaining, key=lambda fc: fc.CoordinateSystem.BasisZ.DotProduct(
                        tray_conn.CoordinateSystem.BasisZ.Negate()))
                    try:
                        best.ConnectTo(tray_conn)
                        remaining.remove(best)
                        connected = True
                    except Exception:
                        pass
            if remaining:
                conn_adapter = remaining[0]
        else:
            conn_adapter = _best_conn(adapter_conns, pt_target)
            if conn_adapter and conn_target:
                connected = _try_connect(conn_adapter, conn_target)

        t.Commit()
        return inst, connected, conn_adapter
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        raise


def _insert_bridge_adapter(sym, insert_pt, lv,
                            round_target_conn, rect_target_conn, source_elem=None):
    """
    CASO C — adaptador ponte: conecta lado round ao conduíte, lado rectangular à bandeja.
    Retorna (instância, round_ok:bool, rect_ok:bool).
    """
    t = _make_t(u"Acoplar — Adaptador Ponte")
    t.Start()
    try:
        if not sym.IsActive:
            sym.Activate()
            doc.Regenerate()
        inst = _place_instance(sym, insert_pt, lv)
        doc.Regenerate()

        if source_elem:
            _copy_parameters(source_elem, inst)

        if rect_target_conn:
            _orient_adapter(inst, rect_target_conn)

        adapter_conns = _ct_connectors(inst)
        adapter_round = _conn_by_profile(adapter_conns, ConnectorProfileType.Round)

        is_ct = (rect_target_conn and rect_target_conn.Owner and
                 rect_target_conn.Owner.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray))

        round_ok = False
        rect_ok  = False

        if adapter_round and round_target_conn:
            round_ok = _try_connect(adapter_round, round_target_conn)

        if is_ct:
            c1_tray, c2_tray = _split_cabletray(doc, rect_target_conn.Owner, insert_pt, lv)
            rect_conns = [c for c in adapter_conns if c.Shape != ConnectorProfileType.Round]
            for tray_conn in [c1_tray, c2_tray]:
                if tray_conn and rect_conns:
                    best = max(rect_conns, key=lambda fc: fc.CoordinateSystem.BasisZ.DotProduct(
                        tray_conn.CoordinateSystem.BasisZ.Negate()))
                    try:
                        best.ConnectTo(tray_conn)
                        rect_conns.remove(best)
                        rect_ok = True
                    except Exception:
                        pass
        else:
            adapter_rect = _conn_by_profile(adapter_conns, ConnectorProfileType.Rectangular)
            if adapter_rect and rect_target_conn:
                rect_ok = _try_connect(adapter_rect, rect_target_conn)

        t.Commit()
        return inst, round_ok, rect_ok
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        raise


def _connect_direct(conn_a, conn_b):
    """Tenta ConnectTo entre dois conectores CT."""
    t = _make_t(u"Acoplar — Conexão Direta")
    t.Start()
    try:
        ok = _try_connect(conn_a, conn_b)
        t.Commit()
        return ok
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return False


# ── Main ─────────────────────────────────────────────────────────────────────

def _pick_chain():
    """Coleta elementos em cadeia: o usuário clica um a um e pressiona ESC para encerrar.
    Retorna lista de elementos ou [] se cancelado antes do segundo clique.
    """
    elements = []
    prompt_msgs = [
        u"Clique no 1º elemento da cadeia (ESC para cancelar)",
        u"Clique no próximo elemento (ESC para encerrar a seleção)",
    ]
    while True:
        msg = prompt_msgs[0] if not elements else prompt_msgs[1]
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, AnyFilter(), msg)
            elem = doc.GetElement(ref.ElementId)
            # Evita duplicatas consecutivas
            if not elements or elem.Id != elements[-1].Id:
                elements.append(elem)
        except OperationCanceledException:
            break
    return elements


def _connect_pair(el_a, el_b, sym):
    """Executa a lógica de conexão entre dois elementos.
    Retorna (ok:bool, msg_erro:str|None).
    """
    pt_a = _location(el_a)
    pt_b = _location(el_b)
    c_a  = _ct_connectors(el_a)
    c_b  = _ct_connectors(el_b)

    # Caso D: nenhum tem conector CT
    if not c_a and not c_b:
        return False, u"Id={} ↔ Id={}: sem conector CT".format(
            el_a.Id.IntegerValue, el_b.Id.IntegerValue)

    # Caso A: el_b sem CT → adaptador em el_b
    if not c_b:
        conn_tray = _best_conn(c_a, pt_b)
        try:
            _, connected, _ = _insert_adapter_simple(
                sym, pt_b, _level(el_b),
                conn_tray, pt_a, source_elem=el_a)
            if connected:
                return True, None
            return False, u"Id={}: adaptador inserido sem conexão".format(el_b.Id.IntegerValue)
        except Exception as e:
            return False, u"Id={}: {}".format(el_b.Id.IntegerValue, e)

    # Caso A': el_a sem CT → adaptador em el_a
    if not c_a:
        conn_tray = _best_conn(c_b, pt_a)
        try:
            _, connected, _ = _insert_adapter_simple(
                sym, pt_a, _level(el_a),
                conn_tray, pt_b, source_elem=el_b)
            if connected:
                return True, None
            return False, u"Id={}: adaptador inserido sem conexão".format(el_a.Id.IntegerValue)
        except Exception as e:
            return False, u"Id={}: {}".format(el_a.Id.IntegerValue, e)

    # Casos B/C: ambos têm CT → tenta direto
    conn_a = _best_conn(c_a, pt_b)
    conn_b = _best_conn(c_b, pt_a)
    if _connect_direct(conn_a, conn_b):
        return True, None

    # Direto falhou — perfis diferentes → adaptador ponte
    pa, pb = _profile(conn_a), _profile(conn_b)
    if pa != pb:
        if sym is None:
            return False, u"Id={}: perfis distintos mas sem adaptador".format(el_b.Id.IntegerValue)
        if pa == ConnectorProfileType.Round:
            round_conn, rect_conn = conn_a, conn_b
            round_elem, rect_elem = el_a, el_b
        else:
            round_conn, rect_conn = conn_b, conn_a
            round_elem, rect_elem = el_b, el_a
        try:
            _, round_ok, rect_ok = _insert_bridge_adapter(
                sym, _location(round_elem), _level(round_elem),
                round_conn, rect_conn, source_elem=rect_elem)
            if round_ok or rect_ok:
                return True, None
            return False, u"Id={}: ponte sem conexão".format(el_b.Id.IntegerValue)
        except Exception as e:
            return False, u"Id={}: {}".format(el_b.Id.IntegerValue, e)

    return False, u"Id={}: ConnectTo rejeitado (mesmo perfil, já conectado?)".format(
        el_b.Id.IntegerValue)


def main():
    # Coleta elementos em cadeia (ESC encerra)
    chain = _pick_chain()

    if len(chain) < 2:
        forms.alert(u"Selecione ao menos 2 elementos para acoplar.",
                    title=u"Acoplar")
        return

    # Verifica se algum par precisará de adaptador
    need_adapter = any(
        not _ct_connectors(chain[i]) or not _ct_connectors(chain[i + 1])
        for i in range(len(chain) - 1)
    )
    sym = None
    if need_adapter:
        _, sym = _pick_adapter()
        if sym is None:
            return

    ok_count = 0
    errors   = []

    # Conecta em cadeia: 1→2, 2→3, 3→4 ...
    for i in range(len(chain) - 1):
        el_a = chain[i]
        el_b = chain[i + 1]

        # Se el_a é eletrocalha, re-localiza o segmento correto após splits anteriores
        if _is_cable_tray(el_a):
            el_a = _find_ct_for_point(el_a.GetTypeId(), _location(el_b)) or el_a

        ok, err = _connect_pair(el_a, el_b, sym)
        if ok:
            ok_count += 1
        elif err:
            errors.append(err)

    if ok_count > 0:
        msg = u"{} conexão(ões) realizada(s) em cadeia.".format(ok_count)
        if errors:
            msg += u"\n{} falha(s):\n{}".format(len(errors), u"\n".join(errors[:3]))
        forms.toast(msg, title=u"Acoplar")
    else:
        forms.alert(u"Nenhum acoplamento realizado.\n\n" + u"\n".join(errors[:5]),
                    title=u"Acoplar — Falha")


if __name__ == "__main__":
    main()
