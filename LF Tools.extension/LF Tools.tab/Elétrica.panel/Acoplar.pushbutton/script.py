# coding: utf-8
"""Acoplar — conecta dois elementos via conector MEP de bandeja/eletroduto.

Casos tratados:
  A) Um elemento tem conector CT, o outro não   → insere adaptador no sem-conector, conecta
  B) Ambos têm conector CT do mesmo perfil      → ConnectTo direto, sem adaptador
  C) Ambos têm conector CT de perfis diferentes → tenta ConnectTo; se falhar, usa adaptador
     ponte (família com conector round + rectangular)
  D) Nenhum tem conector CT                     → erro
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

_LUMINARIA_CATS = {
    int(BuiltInCategory.OST_LightingFixtures),
    int(BuiltInCategory.OST_LightingDevices),
}

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
    # Alinha perfeitamente o conector do adaptador com o conector destino
    try:
        import math
        # Queremos apontar para o lado oposto do conector destino
        target_dir = conn_dest.CoordinateSystem.BasisZ.Negate()
        
        # Acha o conector correspondente no adaptador
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
        
        # Mantém a rotação no intervalo seguro
        while rot > math.pi: rot -= 2 * math.pi
        while rot < -math.pi: rot += 2 * math.pi

        if abs(rot) > 0.001:
            pt = inst.Location.Point
            axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, rot)
            doc.Regenerate()
    except Exception:
        pass

# ── Descida ──────────────────────────────────────────────────────────────────

def _is_luminaria(elem):
    try:
        return elem.Category.Id.IntegerValue in _LUMINARIA_CATS
    except Exception:
        return False


def _level_id_of(elem):
    try:
        lid = elem.LevelId
        if lid != ElementId.InvalidElementId:
            return lid
    except Exception:
        pass
    try:
        v = doc.ActiveView
        if hasattr(v, "GenLevel") and v.GenLevel:
            return v.GenLevel.Id
    except Exception:
        pass
    return FilteredElementCollector(doc).OfClass(Level).FirstElementId()


def _top_face_z(elem):
    """Z da face superior do elemento (bounding box max Z)."""
    for view in [None, doc.ActiveView]:
        try:
            bb = elem.get_BoundingBox(view)
            if bb:
                return bb.Max.Z
        except Exception:
            continue
    return _location(elem).Z


def _is_cable_tray(elem):
    try:
        return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray)
    except Exception:
        return False


def _draw_descida(elem_bottom, elem_top):
    """
    Cria bandeja ou eletroduto descendo verticalmente de elem_top até a face
    superior de elem_bottom. Usa o mesmo tipo e categoria de elem_top.
    """
    type_id  = elem_top.GetTypeId()
    pt_bot   = _location(elem_bottom)
    z_top    = _location(elem_top).Z
    z_face   = _top_face_z(elem_bottom)

    pt_start = XYZ(pt_bot.X, pt_bot.Y, z_top)  # diretamente acima de elem_bottom
    pt_end   = XYZ(pt_bot.X, pt_bot.Y, z_face)  # face superior de elem_bottom

    if pt_start.DistanceTo(pt_end) < 0.1:
        forms.alert(u"Os elementos estão praticamente no mesmo nível — descida não necessária.",
                    title="Acoplar")
        return False

    lv_id          = _level_id_of(elem_bottom)
    use_cable_tray = _is_cable_tray(elem_top)

    t = _make_t(u"Acoplar — Descida")
    t.Start()
    try:
        if use_cable_tray:
            segment = CableTray.Create(doc, type_id, pt_start, pt_end, lv_id)
            # Copia Largura e Altura de elem_top (parâmetros de instância)
            for bip in [BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
                        BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM]:
                try:
                    src_p = elem_top.get_Parameter(bip)
                    tgt_p = segment.get_Parameter(bip)
                    if src_p and tgt_p and not tgt_p.IsReadOnly:
                        tgt_p.Set(src_p.AsDouble())
                except Exception:
                    pass
        else:
            segment = Conduit.Create(doc, type_id, pt_start, pt_end, lv_id)

        doc.Regenerate()

        # Conecta ponta superior do segmento a elem_top.
        # NewTakeoffFitting cria um tê no meio da curva MEP — não precisa de conector livre na ponta.
        # Fallback: ConnectTo ponta-a-ponta caso elem_top tenha conector livre nessa posição.
        seg_conns = list(segment.ConnectorManager.Connectors)
        top_conn  = min(seg_conns, key=lambda c: c.Origin.DistanceTo(pt_start))
        try:
            doc.Create.NewTakeoffFitting(top_conn, elem_top)
        except Exception:
            try:
                free_top = [c for c in _ct_connectors(elem_top) if not c.IsConnected]
                if free_top:
                    nearest_top = min(free_top, key=lambda c: c.Origin.DistanceTo(pt_start))
                    _try_connect(top_conn, nearest_top)
            except Exception:
                pass

        t.Commit()
        return True
    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        raise


def _offer_descida(elem_a, elem_b):
    """Pergunta e desenha descida entre os dois elementos se não forem luminárias."""
    if _is_luminaria(elem_a) or _is_luminaria(elem_b):
        return

    pt_a = _location(elem_a)
    pt_b = _location(elem_b)
    if abs(pt_a.Z - pt_b.Z) < 0.1:
        return

    elem_bottom = elem_a if pt_a.Z <= pt_b.Z else elem_b
    elem_top    = elem_b if pt_a.Z <= pt_b.Z else elem_a

    resp = forms.alert(
        u"Deseja desenhar a descida (eletroduto vertical) até o elemento?",
        title="Acoplar — Descida",
        options=[u"Sim", u"Não"])
    if resp != u"Sim":
        return

    try:
        ok = _draw_descida(elem_bottom, elem_top)
        if ok:
            forms.toast(u"Descida desenhada.", title="Acoplar")
    except Exception as e:
        forms.alert(u"Erro ao desenhar descida:\n" + str(e), title="Acoplar — Erro")


# ── Operações de conexão ──────────────────────────────────────────────────────

def _connect_direct(conn_a, conn_b):
    """Tenta ConnectTo entre dois conectores CT (mesmo Domain, qualquer perfil)."""
    t = _make_t(u"Acoplar — Conexão Direta")
    t.Start()
    try:
        ok = _try_connect(conn_a, conn_b)
        t.Commit()
        return ok
    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        return False


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


def _insert_adapter_simple(sym, insert_pt, lv, conn_target, pt_target, source_elem=None):
    """
    CASO A — insere adaptador em insert_pt e conecta ao conn_target.
    Retorna (instância, conectado:bool).
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
        
        # Se for eletrocalha, divide e conecta dos dois lados
        is_ct = conn_target and conn_target.Owner and conn_target.Owner.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray)
        connected = False
        conn_adapter = None
        
        if is_ct:
            c1_tray, c2_tray = _split_cabletray(doc, conn_target.Owner, insert_pt, lv)
            remaining = list(adapter_conns)
            for tray_conn in [c1_tray, c2_tray]:
                if tray_conn and remaining:
                    best = max(remaining, key=lambda fc: fc.CoordinateSystem.BasisZ.DotProduct(tray_conn.CoordinateSystem.BasisZ.Negate()))
                    try:
                        best.ConnectTo(tray_conn)
                        remaining.remove(best)
                        connected = True
                    except Exception:
                        pass
            if remaining:
                conn_adapter = remaining[0]
        else:
            conn_adapter  = _best_conn(adapter_conns, pt_target)
            if conn_adapter and conn_target:
                connected = _try_connect(conn_adapter, conn_target)

        t.Commit()
        return inst, connected, conn_adapter
    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        raise


def _insert_bridge_adapter(sym, insert_pt, lv,
                            round_target_conn, rect_target_conn, source_elem=None):
    """
    CASO C — adaptador ponte: conecta lado round ao elemento com conduite,
    lado rectangular ao elemento com bandeja.
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
        
        # O adaptador_rect pode ser conectado a duas pontas se dividirmos a bandeja
        is_ct = rect_target_conn and rect_target_conn.Owner and rect_target_conn.Owner.Category.Id.IntegerValue == int(BuiltInCategory.OST_CableTray)
        
        round_ok = False
        rect_ok  = False

        if adapter_round and round_target_conn:
            round_ok = _try_connect(adapter_round, round_target_conn)
            
        if is_ct:
            c1_tray, c2_tray = _split_cabletray(doc, rect_target_conn.Owner, insert_pt, lv)
            # Pega conectores retangulares que sobraram
            rect_conns = [c for c in adapter_conns if c.Shape != ConnectorProfileType.Round]
            for tray_conn in [c1_tray, c2_tray]:
                if tray_conn and rect_conns:
                    best = max(rect_conns, key=lambda fc: fc.CoordinateSystem.BasisZ.DotProduct(tray_conn.CoordinateSystem.BasisZ.Negate()))
                    try:
                        best.ConnectTo(tray_conn)
                        rect_conns.remove(best)
                        rect_ok = True
                    except Exception:
                        pass
        else:
            adapter_rect  = _conn_by_profile(adapter_conns, ConnectorProfileType.Rectangular)
            if adapter_rect and rect_target_conn:
                rect_ok  = _try_connect(adapter_rect,  rect_target_conn)

        t.Commit()
        return inst, round_ok, rect_ok
    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        raise


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    # 1. Selecionar os dois elementos
    try:
        r1 = uidoc.Selection.PickObject(ObjectType.Element, AnyFilter(),
             u"1/2 — Clique no 1º elemento")
    except OperationCanceledException:
        return
    try:
        r2 = uidoc.Selection.PickObject(ObjectType.Element, AnyFilter(),
             u"2/2 — Clique no 2º elemento")
    except OperationCanceledException:
        return

    el1 = doc.GetElement(r1.ElementId)
    el2 = doc.GetElement(r2.ElementId)
    pt1 = _location(el1)
    pt2 = _location(el2)
    c1  = _ct_connectors(el1)
    c2  = _ct_connectors(el2)

    # ── CASO D ───────────────────────────────────────────────────────
    if not c1 and not c2:
        forms.alert(
            u"Nenhum dos dois elementos tem conector de bandeja/eletroduto.\n"
            u"Pelo menos um deve ter (ex: eletrocalha, eletroduto, perfilado).",
            title="Acoplar")
        return

    # ── CASO A: um sem conector ───────────────────────────────────────
    if not c1 or not c2:
        elem_no_conn  = el2 if c1 else el1
        elem_has_conn = el1 if c1 else el2
        conns_has     = c1 if c1 else c2
        pt_no_conn    = _location(elem_no_conn)
        pt_has_conn   = _location(elem_has_conn)
        conn_target   = _best_conn(conns_has, pt_no_conn)

        _, sym = _pick_adapter()
        if sym is None:
            return

        try:
            inst, connected, conn_adapter = _insert_adapter_simple(
                sym, pt_no_conn, _level(elem_no_conn), conn_target, pt_has_conn, source_elem=elem_has_conn)
        except Exception as e:
            forms.alert(u"Erro ao criar adaptador:\n" + str(e), title="Acoplar — Erro")
            return

        if connected:
            forms.toast(u"Adaptador inserido e conectado.", title="Acoplar")
            _offer_descida(elem_no_conn, elem_has_conn)
        else:
            if not _ct_connectors(inst):
                detail = u"A família não tem conector de bandeja/eletroduto."
            elif conn_adapter and conn_target and _profile(conn_adapter) != _profile(conn_target):
                detail = (u"Perfis incompatíveis: adaptador={}, destino={}.".format(
                    _profile(conn_adapter), _profile(conn_target)))
            else:
                detail = u"ConnectTo rejeitado (tamanho ou tipo incompatível)."
            forms.alert(u"Adaptador inserido mas sem conexão física.\n\n" + detail,
                        title="Acoplar — Aviso")
        return

    # ── CASO B/C: ambos têm conector CT ──────────────────────────────
    conn_a = _best_conn(c1, pt2)
    conn_b = _best_conn(c2, pt1)
    pa = _profile(conn_a)
    pb = _profile(conn_b)

    # Tenta direto — a API do Revit exige apenas mesmo Domain, não mesmo perfil
    ok = _connect_direct(conn_a, conn_b)
    if ok:
        forms.toast(u"Conectado diretamente (sem adaptador).", title="Acoplar")
        _offer_descida(el1, el2)
        return

    # ConnectTo falhou — perfis diferentes? Usa adaptador ponte
    if pa != pb:
        # Identifica qual conector é round e qual é rectangular
        if pa == ConnectorProfileType.Round:
            round_elem, round_conn, round_pt = el1, conn_a, pt1
            rect_elem,  rect_conn,  rect_pt  = el2, conn_b, pt2
        else:
            round_elem, round_conn, round_pt = el2, conn_b, pt2
            rect_elem,  rect_conn,  rect_pt  = el1, conn_a, pt1

        result = forms.alert(
            u"ConnectTo direto falhou (perfis diferentes: conduíte vs bandeja).\n\n"
            u"Posso inserir um adaptador ponte com os dois tipos de conector.\n"
            u"O adaptador será colocado sobre o elemento com conduíte (round).\n\n"
            u"Deseja continuar?",
            title="Acoplar — Adaptador Ponte",
            options=[u"Sim, escolher adaptador", u"Cancelar"])
        if not result or result == u"Cancelar":
            return

        _, sym = _pick_adapter(u"deve ter conector round + rectangular")
        if sym is None:
            return

        insert_pt = _location(round_elem)
        lv        = _level(round_elem)

        try:
            inst, round_ok, rect_ok = _insert_bridge_adapter(
                sym, insert_pt, lv,
                round_conn, rect_conn, source_elem=rect_elem)
        except Exception as e:
            forms.alert(u"Erro ao criar adaptador ponte:\n" + str(e), title="Acoplar — Erro")
            return

        if round_ok and rect_ok:
            forms.toast(u"Adaptador ponte inserido e conectado nos dois lados.", title="Acoplar")
            _offer_descida(round_elem, rect_elem)
        elif round_ok or rect_ok:
            lado = u"conduíte" if round_ok else u"bandeja"
            forms.alert(
                u"Adaptador inserido mas conectado apenas no lado {}.\n\n"
                u"O outro lado falhou — verifique se a família tem os dois tipos "
                u"de conector (round + rectangular).".format(lado),
                title="Acoplar — Parcial")
        else:
            forms.alert(
                u"Adaptador inserido mas nenhuma conexão foi estabelecida.\n\n"
                u"A família selecionada precisa ter um conector round E um rectangular "
                u"(ambos do domínio bandeja/eletroduto).",
                title="Acoplar — Aviso")
    else:
        # Mesmo perfil mas ConnectTo falhou mesmo assim
        forms.alert(
            u"ConnectTo rejeitado pelo Revit mesmo com conectores do mesmo perfil.\n"
            u"Verifique se os conectores já estão conectados a outro elemento.",
            title="Acoplar — Erro")


if __name__ == "__main__":
    main()
