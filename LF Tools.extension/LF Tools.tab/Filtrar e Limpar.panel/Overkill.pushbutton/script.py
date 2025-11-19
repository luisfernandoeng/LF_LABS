# -*- coding: utf-8 -*-
import clr
import os
import math
import datetime
import codecs
from Autodesk.Revit.DB import Transaction

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Element, ElementId, ViewSheet,
    XYZ, Transaction, Line, Arc, Options, SpatialElementGeometryCalculator
)
from Autodesk.Revit.DB import FamilyInstance, LocationPoint, LocationCurve
from Autodesk.Revit.UI import TaskDialog

# pyRevit
from pyrevit import revit, forms, script

# ====== Contexto Revit / pyRevit ======
uiapp = __revit__
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document if uidoc else None
output = script.get_output()
logger = script.get_logger()

# REMOVIDA: A lógica de log em arquivo não será mais utilizada
def log_line(msg):
    # Usaremos apenas o logger do pyRevit, que já é exibido na janela
    pass 

# ====== Utilidades ======
MM_PER_FT = 304.8

def quantize_feet_to_mm_int(val_ft, step_mm):
    """Converte feet -> mm e 'snap' para grade inteira de step_mm."""
    return int(round((val_ft * MM_PER_FT) / float(step_mm)))

def qxyz(pt, step_mm):
    """XYZ em grade inteira de mm (tuple)."""
    return (
        quantize_feet_to_mm_int(pt.X, step_mm),
        quantize_feet_to_mm_int(pt.Y, step_mm),
        quantize_feet_to_mm_int(pt.Z, step_mm)
    )

def safe_level_id(el):
    """Tenta obter o LevelId. Se não houver, retorna ElementId.InvalidElementId."""
    try:
        lid = getattr(el, 'LevelId', None)
        if lid and lid.IntegerValue > 0:
            return lid
    except:
        pass
    # tenta por parâmetro comum
    try:
        p = el.LookupParameter("Level") or el.LookupParameter("Nível") or el.LookupParameter("Reference Level")
        if p and p.AsElementId() and p.AsElementId().IntegerValue > 0:
            return p.AsElementId()
    except:
        pass
    try:
        return ElementId.InvalidElementId
    except:
        return None

def is_pinned(el):
    try:
        return bool(el.Pinned)
    except:
        return False

def get_location_signature(el, step_mm_curve=1.0, step_mm_point=2.0):
    """
    Retorna assinatura geométrica normalizada do elemento:
      - Para curvas (Conduit/Duct/Pipe/CableTray): endpoints normalizados (ordem independente).
      - Para arcs: centro + raio + ângulos quantizados (fallback se precisar).
      - Para FamilyInstance com ponto: posição quantizada.
      - Fallback: centro do bounding box quantizado.
    """
    try:
        loc = el.Location
        # Curvas (LocationCurve)
        if isinstance(loc, LocationCurve):
            c = loc.Curve
            try:
                line = c if isinstance(c, Line) else None
                arc = c if isinstance(c, Arc) else None

                if line:
                    p1 = line.GetEndPoint(0)
                    p2 = line.GetEndPoint(1)
                    a = qxyz(p1, step_mm_curve)
                    b = qxyz(p2, step_mm_curve)
                    # ordem independente
                    if a <= b:
                        ends = (a, b)
                    else:
                        ends = (b, a)
                    return ("LINE", ends)

                if arc:
                    center = arc.Center
                    r_mm = quantize_feet_to_mm_int(arc.Radius, step_mm_curve)
                    # ângulos em 0.1 rad (grosseiro, mas suficiente p/ detectar igualdade)
                    a1 = int(round(arc.StartAngle*10.0))
                    a2 = int(round(arc.EndAngle*10.0))
                    cxyz = qxyz(center, step_mm_curve)
                    # normalizar ordem dos ângulos
                    if a1 <= a2:
                        aa = (a1, a2)
                    else:
                        aa = (a2, a1)
                    return ("ARC", cxyz, r_mm, aa)
            except:
                pass

        # Ponto (FamilyInstance com LocationPoint)
        if isinstance(el, FamilyInstance):
            try:
                lp = el.Location
                if isinstance(lp, LocationPoint):
                    pt = lp.Point
                    return ("POINT", qxyz(pt, step_mm_point))
            except:
                pass

        # Fallback: centro do bounding box
        try:
            bb = el.get_BoundingBox(None)
            if bb:
                center = (bb.Min + bb.Max) * 0.5
                return ("BBOX", qxyz(center, step_mm_point))
        except:
            pass

    except Exception as ex:
        logger.warning(u"Falha em get_location_signature para {}: {}".format(el.Id, ex))

    return None

def default_categories():
    """Categorias padrão do Overkill (inclui equipamentos mecânicos)."""
    return [
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_ConduitFitting,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_CableTray,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_MechanicalEquipment
    ]

def pick_scope_active_view():
    """Pergunta se quer limitar à vista ativa."""
    return forms.alert(
        "Limitar a busca de duplicados APENAS à vista ativa?",
        yes=True, no=True, warn_icon=False
    )

# ====== Principal ======
def run_overkill():
    if not doc or doc.IsFamilyDocument:
        forms.alert("Abra um PROJETO (RVT) antes de rodar o Overkill.")
        return

    # Escolha de escopo
    only_active_view = pick_scope_active_view()
    active_view_id = uidoc.ActiveView.Id if only_active_view else None

    # Escolha de categorias (lista simples)
    cat_map = {
        u"Conduítes": BuiltInCategory.OST_Conduit,
        u"Conexões de Conduíte": BuiltInCategory.OST_ConduitFitting,
        u"Dutos": BuiltInCategory.OST_DuctCurves,
        u"Conexões de Dutos": BuiltInCategory.OST_DuctFitting,
        u"Eletrocalhas": BuiltInCategory.OST_CableTray,
        u"Tubulações": BuiltInCategory.OST_PipeCurves,
        u"Conexões de Tubulações": BuiltInCategory.OST_PipeFitting,
        u"Equipamentos Mecânicos": BuiltInCategory.OST_MechanicalEquipment
    }
    items = forms.SelectFromList.show(
        sorted(cat_map.keys()),
        title="Categorias para Overkill",
        multiselect=True,
        button_name="OK",
        width=450,
        height=400
    )
    if not items:
        forms.alert("Nenhuma categoria selecionada. Cancelado.")
        return
    selected_bics = [cat_map[n] for n in items]

    # Tolerâncias
    try:
        tol_pt = forms.ask_for_string(
            default="2.0",
            title="Tolerância de coincidência para PONTO/BBox (mm)",
            prompt="Valor em mm (ex.: 2.0)"
        )
        tol_curve = forms.ask_for_string(
            default="1.0",
            title="Tolerância de coincidência para CURVAS (mm)",
            prompt="Valor em mm (ex.: 1.0)"
        )
        tol_pt = float(tol_pt.replace(",", ".")) if tol_pt else 2.0
        tol_curve = float(tol_curve.replace(",", ".")) if tol_curve else 1.0
    except:
        tol_pt, tol_curve = 2.0, 1.0

    # Coleta
    to_check = []
    for bic in selected_bics:
        col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        elems = list(col)
        if only_active_view:
            # filtra por vista ativa (visíveis na vista)
            vid = active_view_id
            elems = [e for e in elems if e.OwnerViewId.IntegerValue == vid.IntegerValue or e.OwnerViewId.IntegerValue == -1]
        to_check.extend(elems)

    total = len(to_check)
    if total == 0:
        forms.alert("Nenhum elemento encontrado nas categorias selecionadas.")
        return

    # Indexação
    duplicates_to_delete = []  # lista de ElementId para deletar
    duplicates_info = [] # Nova lista para armazenar informações detalhadas
    groups = {}  # signature -> id mantido

    kept_count = 0
    pinned_skipped = 0
    no_sig = 0

    for el in to_check:
        try:
            if is_pinned(el):
                pinned_skipped += 1
                continue

            # Chaves de agrupamento fortes:
            cat = el.Category.Id.IntegerValue if el.Category else -1
            typ = el.GetTypeId().IntegerValue if el.GetTypeId() else -1
            lvl = safe_level_id(el)
            lvl_id = lvl.IntegerValue if lvl else -1

            sig_geo = get_location_signature(el, step_mm_curve=tol_curve, step_mm_point=tol_pt)
            if sig_geo is None:
                no_sig += 1
                continue

            signature = (cat, typ, lvl_id, sig_geo)

            if signature in groups:
                # duplicado -> marcar para deletar
                # mantém o de menor Id
                keep_id = groups[signature]
                
                # Obtém o nome e nível de forma segura
                level_name = ""
                level = doc.GetElement(lvl)
                if level:
                    level_name = level.Name
                
                element_name = ""
                try:
                    element_type = doc.GetElement(el.GetTypeId())
                    if element_type:
                        element_name = element_type.Name
                except:
                    # fallback para o nome da categoria se o tipo não tiver nome
                    if el.Category:
                        element_name = el.Category.Name
                
                if not element_name:
                    element_name = "Elemento ID: {}".format(el.Id.IntegerValue)

                if el.Id.IntegerValue < keep_id.IntegerValue:
                    # troca: mantém o atual, apaga o antigo
                    duplicates_to_delete.append(keep_id)
                    groups[signature] = el.Id
                    duplicates_info.append((element_name, level_name)) # Adiciona à nova lista
                else:
                    duplicates_to_delete.append(el.Id)
                    duplicates_info.append((element_name, level_name)) # Adiciona à nova lista
            else:
                groups[signature] = el.Id
                kept_count += 1

        except Exception as ex:
            logger.warning(u"Falha analisando elemento {}: {}".format(el.Id, ex))

    dup_count = len(duplicates_to_delete)

    # Confirmação
    msg_preview = u"Encontrados {} elementos potenciais duplicados.\n" \
                  u"Serão deletados: {}\n" \
                  u"Manter: {}\n" \
                  u"(Ignorados por estarem 'pinned'): {}\n" \
                  u"(Sem assinatura geométrica confiável): {}\n\n" \
                  u"Deseja prosseguir com a exclusão?".format(total, dup_count, kept_count, pinned_skipped, no_sig)

    if not forms.alert(msg_preview, yes=True, no=True):
        output.print_md("**Overkill cancelado pelo usuário.**")
        return

    # Deletar vs. Selecionar
    selection_mode = not forms.alert(
        msg_preview + "\n\nSe clicar 'Não', o script irá selecionar os elementos duplicados em vez de deletá-los.",
        yes=True, no=True, title="Modo de Operação"
    )

    deleted_info = [] # Lista para armazenar informações dos elementos removidos
    failed_info = []  # Lista para armazenar informações dos elementos que falharam

    if not selection_mode:
        with revit.TransactionGroup("Overkill", doc) as tg:
            with revit.Transaction("Deletar duplicados", doc):
                seen = set()
                for eid in duplicates_to_delete:
                    if eid.IntegerValue in seen:
                        continue
                    seen.add(eid.IntegerValue)
                    try:
                        # Obtem o elemento antes de deletar
                        el = doc.GetElement(eid)
                        level_name = ""
                        level_id = safe_level_id(el)
                        if level_id and level_id.IntegerValue > 0:
                            level = doc.GetElement(level_id)
                            if level:
                                level_name = level.Name
                        
                        element_name = ""
                        try:
                            element_type = doc.GetElement(el.GetTypeId())
                            if element_type:
                                element_name = element_type.Name
                        except:
                            if el.Category:
                                element_name = el.Category.Name

                        if not element_name:
                             element_name = "Elemento ID: {}".format(el.Id.IntegerValue)

                        res = doc.Delete(eid)
                        if res and len(res) > 0:
                            deleted_info.append((element_name, level_name))
                        else:
                            failed_info.append((element_name, level_name))
                            logger.warning(u"Não foi possível deletar elemento {}.".format(eid))
                    except Exception as ex:
                        failed_info.append((element_name, level_name))
                        logger.warning(u"Erro ao deletar {}: {}".format(eid, ex))

    else:
        # Modo Seleção: apenas seleciona os elementos
        if duplicates_to_delete:
            clr.AddReference("System.Collections")
            from System.Collections.Generic import List
            
            element_ids_to_select = List[ElementId](duplicates_to_delete)
            uidoc.Selection.SetElementIds(element_ids_to_select)
            
        logger.info(u"[MODO SELEÇÃO] Elementos duplicados selecionados.")


    # Cria a string de saída
    out = u"### Overkill finalizado\n"
    out += u"- Analisados: {}\n".format(total)
    out += u"- Duplicados detectados: {}\n".format(len(duplicates_to_delete))

    if not selection_mode:
        out += u"- Removidos: {}\n".format(len(deleted_info))
        out += u"- Falharam: {}\n".format(len(failed_info))
        if deleted_info:
            out += u"\n**Elementos Removidos:**\n"
            for name, level in deleted_info:
                out += u"  - **{0}** (Nível: {1})\n".format(name, level)

        if failed_info:
            out += u"\n**Falha na Remoção:**\n"
            for name, level in failed_info:
                out += u"  - **{0}** (Nível: {1})\n".format(name, level)
    
    out += u"\n- Mantidos: {}\n".format(kept_count)
    out += u"- Pinned (ignorados): {}\n".format(pinned_skipped)
    out += u"- Sem assinatura: {}\n".format(no_sig)
    out += u"- Modo de operação: {}\n".format("SELEÇÃO" if selection_mode else "DELETAR")
    
    output.print_md(out)


if __name__ == "__main__":
    try:
        run_overkill()
    except Exception as e:
        logger.error(u"Erro fatal no Overkill:\n{}".format(e))
        forms.alert(u"Erro fatal no Overkill:\n{}".format(e))