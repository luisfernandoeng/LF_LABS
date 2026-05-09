# -*- coding: utf-8 -*-
import clr
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory, ElementId,
    LocationCurve, LocationPoint, Line, Arc, FamilyInstance, TextNote,
    LinePatternElement, View, GraphicsStyleType,
)
import System

from pyrevit import revit, forms, script

doc  = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

MM_PER_FT = 304.8


# ====== Classes de dados ======

class CategoryOption(object):
    def __init__(self, name, bic, checked=False):
        self.name       = name
        self.bic        = bic
        self.is_checked = checked


class CleanupItem(object):
    """Representa um único item a ser removido, de qualquer operação."""
    def __init__(self, el_id, item_type, name, detail=""):
        self.ElementId    = el_id
        self.ElementIdStr = str(el_id.IntegerValue)
        self.ItemType     = item_type
        self.Name         = name
        self.Detail       = detail
        self._is_selected = True

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value


# ====== Funções geométricas (detecção de sobrepostos) ======

def _quantize(val_ft, step_mm):
    return int(round((val_ft * MM_PER_FT) / float(step_mm)))


def _qxyz(pt, step_mm):
    return (_quantize(pt.X, step_mm), _quantize(pt.Y, step_mm), _quantize(pt.Z, step_mm))


def _safe_level_id(el):
    try:
        lid = getattr(el, 'LevelId', None)
        if lid and lid.IntegerValue > 0:
            return lid.IntegerValue
    except Exception:
        pass
    try:
        for pname in ("Level", "Nível", "Reference Level"):
            p = el.LookupParameter(pname)
            if p and p.AsElementId().IntegerValue > 0:
                return p.AsElementId().IntegerValue
    except Exception:
        pass
    return -1


def _location_signature(el, step_curve, step_point):
    try:
        if isinstance(el, TextNote):
            return ("TEXT", _qxyz(el.Coord, step_point), el.Text.strip())

        loc = el.Location
        if isinstance(loc, LocationCurve):
            c = loc.Curve
            if isinstance(c, Line):
                p1 = _qxyz(c.GetEndPoint(0), step_curve)
                p2 = _qxyz(c.GetEndPoint(1), step_curve)
                return ("LINE", tuple(sorted((p1, p2))))
            if isinstance(c, Arc):
                return ("ARC", _qxyz(c.Center, step_curve), _quantize(c.Radius, step_curve))

        if isinstance(el, FamilyInstance) and isinstance(loc, LocationPoint):
            return ("POINT", _qxyz(loc.Point, step_point))

        bb = el.get_BoundingBox(None)
        if bb:
            center = (bb.Min + bb.Max) * 0.5
            return ("BBOX", _qxyz(center, step_point))
    except Exception:
        pass
    return None


def _element_description(el):
    try:
        if isinstance(el, TextNote):
            t = el.Text.strip()
            return t[:67] + "..." if len(t) > 70 else t
        if isinstance(el, FamilyInstance):
            sym = el.Symbol
            fam = sym.Family.Name if sym and sym.Family else "?"
            try:
                return "{}: {}".format(fam, el.Name)
            except Exception:
                return fam
        loc = el.Location
        if isinstance(loc, LocationCurve):
            mm = loc.Curve.Length * MM_PER_FT
            try:
                return "{} ({:.0f} mm)".format(el.Name, mm)
            except Exception:
                return "{:.0f} mm".format(mm)
        try:
            return el.Name
        except Exception:
            return "El. {}".format(el.Id.IntegerValue)
    except Exception:
        return "El. {}".format(el.Id.IntegerValue)


# ====== Funções de análise por operação ======

def find_duplicate_elements(doc, selected_bics, is_active_view, tol_curve, tol_point):
    items = []
    elements = []
    for bic in selected_bics:
        col = (FilteredElementCollector(doc, doc.ActiveView.Id)
               if is_active_view
               else FilteredElementCollector(doc))
        elements.extend(list(col.OfCategory(bic).WhereElementIsNotElementType()))

    seen = {}
    for el in elements:
        try:
            cat_id  = el.Category.Id.IntegerValue if el.Category else -1
            cat_name = el.Category.Name if el.Category else "Desconhecido"
            type_id = el.GetTypeId().IntegerValue if el.GetTypeId() else -1
            lvl_id  = _safe_level_id(el)
            sig     = _location_signature(el, tol_curve, tol_point)
            if not sig:
                continue
            full_sig = (cat_id, type_id, lvl_id, sig)
            if full_sig in seen:
                existing_id = seen[full_sig]
                # mantém o ID menor (elemento mais antigo); marca o mais novo como duplicado
                if el.Id.IntegerValue > existing_id.IntegerValue:
                    dup_id = el.Id
                else:
                    dup_id = existing_id
                    seen[full_sig] = el.Id
                dup_el = doc.GetElement(dup_id)
                desc = _element_description(dup_el) if dup_el else "?"
                items.append(CleanupItem(dup_id, "Sobreposto", cat_name, desc))
            else:
                seen[full_sig] = el.Id
        except Exception:
            pass
    return items


def find_unused_view_templates(doc):
    items = []
    try:
        all_views = list(FilteredElementCollector(doc).OfClass(View))
        templates = [v for v in all_views if v.IsTemplate]
        used_ids  = set()
        for v in all_views:
            if not v.IsTemplate:
                try:
                    tid = v.ViewTemplateId
                    if tid and tid.IntegerValue > 0:
                        used_ids.add(tid.IntegerValue)
                except Exception:
                    pass
        for t in templates:
            if t.Id.IntegerValue not in used_ids:
                items.append(CleanupItem(
                    t.Id, "Modelo de Vista", t.Name,
                    "Não aplicado a nenhuma vista"
                ))
    except Exception as ex:
        logger.warning("Modelos de vista: {}".format(ex))
    return items


def find_unused_text_types(doc):
    items = []
    try:
        used_ids = set()
        for tn in FilteredElementCollector(doc).OfClass(TextNote):
            try:
                tid = tn.GetTypeId()
                if tid and tid.IntegerValue > 0:
                    used_ids.add(tid.IntegerValue)
            except Exception:
                pass
        col = (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_TextNotes)
               .WhereElementIsElementType())
        for tt in col:
            if tt.Id.IntegerValue not in used_ids:
                try:
                    name = tt.Name
                except Exception:
                    name = "Tipo {}".format(tt.Id.IntegerValue)
                items.append(CleanupItem(
                    tt.Id, "Tipo de Texto", name,
                    "Nenhuma nota de texto utiliza este tipo"
                ))
    except Exception as ex:
        logger.warning("Tipos de texto: {}".format(ex))
    return items


def find_unused_line_patterns(doc):
    """
    Coleta padrões de linha não referenciados pelas configurações
    de categoria (projeção e corte). Não detecta overrides por elemento.
    """
    items = []
    try:
        used_ids = set()
        try:
            for cat in doc.Settings.Categories:
                for gs_type in (GraphicsStyleType.Projection, GraphicsStyleType.Cut):
                    try:
                        pid = cat.GetLinePatternId(gs_type)
                        if pid and pid.IntegerValue > 0:
                            used_ids.add(pid.IntegerValue)
                    except Exception:
                        pass
                try:
                    for subcat in cat.SubCategories:
                        for gs_type in (GraphicsStyleType.Projection, GraphicsStyleType.Cut):
                            try:
                                pid = subcat.GetLinePatternId(gs_type)
                                if pid and pid.IntegerValue > 0:
                                    used_ids.add(pid.IntegerValue)
                            except Exception:
                                pass
                except Exception:
                    pass
        except Exception:
            pass

        for lp in FilteredElementCollector(doc).OfClass(LinePatternElement):
            if lp.Id.IntegerValue not in used_ids:
                try:
                    name = lp.GetLinePattern().Name
                except Exception:
                    try:
                        name = lp.Name
                    except Exception:
                        name = "Padrão {}".format(lp.Id.IntegerValue)
                items.append(CleanupItem(
                    lp.Id, "Padrão de Linha", name,
                    "Não referenciado em estilos de categoria"
                ))
    except Exception as ex:
        logger.warning("Padrões de linha: {}".format(ex))
    return items


def _legacy_find_unused_load_classifications(doc):
    """ElectricalLoadClassification não referenciada por nenhum ElectricalSystem."""
    items = []
    try:
        from Autodesk.Revit.DB.Electrical import ElectricalLoadClassification, ElectricalSystem

        all_clf = list(FilteredElementCollector(doc).OfClass(ElectricalLoadClassification))
        if not all_clf:
            return items

        # .LoadClassifications retorna o nome da classificação como string.
        # Um circuito pode ter múltiplas; dividimos por vírgula para garantir.
        used_names = set()
        for circ in FilteredElementCollector(doc).OfClass(ElectricalSystem):
            try:
                raw = circ.LoadClassifications  # ex: "Iluminação" ou "Iluminação, Motor"
                if raw:
                    for part in raw.split(","):
                        used_names.add(part.strip())
            except Exception:
                pass

        for clf in all_clf:
            try:
                name = clf.Name
            except Exception:
                name = "Classificação {}".format(clf.Id.IntegerValue)
            if name not in used_names:
                items.append(CleanupItem(
                    clf.Id, "Tipo de Carga", name,
                    "Nenhum circuito usa esta classificação"
                ))
    except Exception:
        pass  # Projeto sem elétrica — silencioso
    return items


def _find_load_classification_references(doc, clf_id, clf_name):
    """Retorna descricoes curtas de referencias que impedem exclusao segura."""
    refs = []
    clf_int = clf_id.IntegerValue
    try:
        from Autodesk.Revit.DB.Electrical import ElectricalSystem
        for circ in FilteredElementCollector(doc).OfClass(ElectricalSystem):
            try:
                raw = circ.LoadClassifications
                names = [p.strip() for p in raw.replace(";", ",").split(",")] if raw else []
                if clf_name in names:
                    refs.append("Circuito {}".format(circ.Id.IntegerValue))
            except Exception:
                pass
    except Exception:
        pass

    try:
        from Autodesk.Revit.DB import StorageType
        for el in FilteredElementCollector(doc):
            try:
                if el.Id.IntegerValue == clf_int:
                    continue
                for p in el.Parameters:
                    try:
                        if p.StorageType != StorageType.ElementId:
                            continue
                        pid = p.AsElementId()
                        if not pid or pid.IntegerValue != clf_int:
                            continue
                        try:
                            pname = p.Definition.Name
                        except Exception:
                            pname = "parametro"
                        try:
                            ename = el.Name
                        except Exception:
                            ename = el.GetType().Name
                        refs.append("{} em {} ({})".format(pname, ename, el.Id.IntegerValue))
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

    seen = set()
    unique_refs = []
    for ref in refs:
        if ref not in seen:
            seen.add(ref)
            unique_refs.append(ref)
    return unique_refs


def _replace_load_classification_references(doc, old_ids, substitute_id):
    """Troca parametros ElementId que apontam para classificacoes antigas."""
    changed = 0
    locked = []
    try:
        from Autodesk.Revit.DB import StorageType
        for el in FilteredElementCollector(doc):
            try:
                if el.Id.IntegerValue in old_ids:
                    continue
                for p in el.Parameters:
                    try:
                        if p.StorageType != StorageType.ElementId:
                            continue
                        pid = p.AsElementId()
                        if not pid or pid.IntegerValue not in old_ids:
                            continue
                        if p.IsReadOnly:
                            try:
                                pname = p.Definition.Name
                            except Exception:
                                pname = "parametro"
                            try:
                                ename = el.Name
                            except Exception:
                                ename = el.GetType().Name
                            locked.append("{} em {} ({})".format(pname, ename, el.Id.IntegerValue))
                            continue
                        p.Set(substitute_id)
                        changed += 1
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass
    return changed, locked


def find_unused_load_classifications(doc):
    """Lista classificacoes de carga para limpeza/consolidacao."""
    items = []
    try:
        from Autodesk.Revit.DB.Electrical import ElectricalLoadClassification

        all_clf = list(FilteredElementCollector(doc).OfClass(ElectricalLoadClassification))
        if not all_clf:
            return items

        for clf in all_clf:
            try:
                name = clf.Name
            except Exception:
                name = "Classificacao {}".format(clf.Id.IntegerValue)
            refs = _find_load_classification_references(doc, clf.Id, name)
            if refs:
                detail = "{} referencia(s); escolha uma substituta para consolidar".format(len(refs))
            else:
                detail = "Sem referencias encontradas"
            items.append(CleanupItem(clf.Id, "Tipo de Carga", name, detail))
    except Exception:
        pass
    return items


def find_unconnected_conduits(doc):
    """Eletrodutos com ambas as pontas livres (nenhum conector conectado)."""
    items = []
    try:
        for el in (FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_Conduit)
                   .WhereElementIsNotElementType()):
            try:
                mgr = el.ConnectorManager
                if mgr is None:
                    continue
                connectors = list(mgr.Connectors)
                if not connectors:
                    continue
                # Apenas marca se AMBAS as pontas estiverem livres
                if all(not c.IsConnected for c in connectors):
                    try:
                        loc = el.Location
                        length_mm = loc.Curve.Length * MM_PER_FT if isinstance(loc, LocationCurve) else 0
                        detail = "{:.0f} mm, ambas as pontas livres".format(length_mm)
                    except Exception:
                        detail = "Ambas as pontas livres"
                    try:
                        name = el.Name
                    except Exception:
                        name = "Eletroduto"
                    items.append(CleanupItem(el.Id, "Eletroduto Solto", name, detail))
            except Exception:
                pass
    except Exception as ex:
        logger.warning("Eletrodutos não conectados: {}".format(ex))
    return items


def find_unused_view_filters(doc):
    """ParameterFilterElement definido no projeto mas não aplicado a nenhuma vista ou modelo de vista."""
    items = []
    try:
        from Autodesk.Revit.DB import ParameterFilterElement
        used_ids = set()
        for v in FilteredElementCollector(doc).OfClass(View):
            try:
                for fid in v.GetFilters():
                    used_ids.add(fid.IntegerValue)
            except Exception:
                pass
        for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
            if f.Id.IntegerValue not in used_ids:
                items.append(CleanupItem(
                    f.Id, "Filtro de Vista", f.Name,
                    "Não aplicado a nenhuma vista ou modelo"
                ))
    except Exception as ex:
        logger.warning("Filtros de vista: {}".format(ex))
    return items


def find_unloaded_cad_links(doc):
    """Links CAD (DWG/DXF) com status diferente de Carregado."""
    items = []
    try:
        from Autodesk.Revit.DB import CADLinkType
        for lt in FilteredElementCollector(doc).OfClass(CADLinkType):
            try:
                efr = lt.GetExternalFileReference()
                if efr is None:
                    continue
                status_str = efr.LinkedFileStatus.ToString()
                if status_str == "Loaded":
                    continue
                status_labels = {
                    "Unloaded":         "Descarregado manualmente",
                    "NotFound":         "Arquivo não encontrado",
                    "LocallyUnloaded":  "Descarregado localmente",
                    "InvalidLink":      "Link inválido",
                }
                detail = status_labels.get(status_str, status_str)
                items.append(CleanupItem(lt.Id, "Link CAD", lt.Name, detail))
            except Exception:
                pass
    except Exception as ex:
        logger.warning("Links CAD: {}".format(ex))
    return items


def find_orphaned_tags(doc):
    """
    IndependentTag cujo elemento hospedeiro não existe mais no documento.
    Tags de elementos em links são ignoradas (falso positivo garantido).
    """
    items = []
    try:
        from Autodesk.Revit.DB import IndependentTag
        for tag in (FilteredElementCollector(doc)
                    .OfClass(IndependentTag)
                    .WhereElementIsNotElementType()):
            try:
                host_id = tag.TaggedLocalElementId
                # host_id == -1 indica elemento em arquivo linkado — pular
                if host_id.IntegerValue <= 0:
                    continue
                if doc.GetElement(host_id) is None:
                    try:
                        cat_name = tag.Category.Name if tag.Category else "Tag"
                    except Exception:
                        cat_name = "Tag"
                    items.append(CleanupItem(
                        tag.Id, "Tag Órfã", cat_name,
                        "Host ID {} não encontrado".format(host_id.IntegerValue)
                    ))
            except Exception:
                pass
    except Exception as ex:
        logger.warning("Tags órfãs: {}".format(ex))
    return items


# ====== Helper: reatribuição de classificação de carga ======

def _reassign_load_classifications(doc, old_names, old_ids, substitute_id):
    """
    Reatribui para substitute_id todos os elementos que referenciam as
    classificações de carga em old_names / old_ids.

    Cobre dois grupos:
      1. ElectricalSystem (circuitos) — detectado pelo nome via .LoadClassifications
      2. FamilyInstance em categorias elétricas — detectado por ElementId do parâmetro

    Deve ser chamado dentro de uma Transaction já aberta.
    """
    from Autodesk.Revit.DB.Electrical import ElectricalSystem
    from Autodesk.Revit.DB import BuiltInParameter, StorageType, BuiltInCategory

    # Descobre BIPs válidos nesta versão do Revit
    bips = []
    for bip_name in ("RBS_ELEC_CIRCUIT_LOAD_CLASSIFICATION",):
        try:
            bips.append(getattr(BuiltInParameter, bip_name))
        except Exception:
            pass
    param_names = ("Load Classification", "Classificação de Carga", "LoadClassification")

    def _try_set(el):
        """Tenta setar a classificação de carga no elemento; retorna True se conseguiu."""
        # Via BIP (armazenamento ElementId)
        for bip in bips:
            try:
                p = el.get_Parameter(bip)
                if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                    if p.AsElementId().IntegerValue in old_ids:
                        p.Set(substitute_id)
                        return True
            except Exception:
                pass
        # Via nome de parâmetro (ElementId)
        for pname in param_names:
            try:
                p = el.LookupParameter(pname)
                if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                    if p.AsElementId().IntegerValue in old_ids:
                        p.Set(substitute_id)
                        return True
            except Exception:
                pass
        return False

    # ── 1. Circuitos (ElectricalSystem) ─────────────────────────────────────
    # .LoadClassifications retorna string; usamos para filtrar, depois setamos
    # via parâmetro de ElementId.
    for circ in FilteredElementCollector(doc).OfClass(ElectricalSystem):
        try:
            raw = circ.LoadClassifications
            if not raw:
                continue
            if not {s.strip() for s in raw.split(",")}.intersection(old_names):
                continue
            _try_set(circ)
        except Exception:
            pass

    # ── 2. Instâncias de família elétricas ──────────────────────────────────
    elec_cats = [
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_LightingDevices,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_DataDevices,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_NurseCallDevices,
        BuiltInCategory.OST_SecurityDevices,
    ]
    for bic in elec_cats:
        try:
            for el in (FilteredElementCollector(doc)
                       .OfCategory(bic)
                       .WhereElementIsNotElementType()):
                try:
                    _try_set(el)
                except Exception:
                    pass
        except Exception:
            pass


# ====== Janela principal ======

class OverkillWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.cleanup_items = []
        self._init_categories()
        self._bind_events()
        self._sync_options_panel()

    def _init_categories(self):
        cats = [
            ("Notas de Texto",              BuiltInCategory.OST_TextNotes,           False),
            ("Anotação Genérica",           BuiltInCategory.OST_GenericAnnotation,   False),
            ("Identificadores (Tags)",      BuiltInCategory.OST_Tags,                False),
            ("Conduítes",                   BuiltInCategory.OST_Conduit,             False),
            ("Conexões de Conduíte",        BuiltInCategory.OST_ConduitFitting,      False),
            ("Eletrocalhas",                BuiltInCategory.OST_CableTray,           False),
            ("Conexões de Eletrocalha",     BuiltInCategory.OST_CableTrayFitting,    False),
            ("Dutos",                       BuiltInCategory.OST_DuctCurves,          False),
            ("Conexões de Duto",            BuiltInCategory.OST_DuctFitting,         False),
            ("Tubulações",                  BuiltInCategory.OST_PipeCurves,          False),
            ("Conexões de Tubulação",       BuiltInCategory.OST_PipeFitting,         False),
            ("Equipamentos Mecânicos",      BuiltInCategory.OST_MechanicalEquipment, False),
            ("Luminárias",                  BuiltInCategory.OST_LightingFixtures,    False),
            ("Dispositivos de Iluminação",  BuiltInCategory.OST_LightingDevices,     False),
            ("Dispositivos Elétricos",      BuiltInCategory.OST_ElectricalFixtures,  False),
            ("Mobiliário",                  BuiltInCategory.OST_Furniture,           False),
            ("Linhas de Detalhe",           BuiltInCategory.OST_Lines,               False),
        ]
        self.categories = [CategoryOption(n, b, c) for n, b, c in cats]
        self.categories_lb.ItemsSource = self.categories

    def _bind_events(self):
        self.run_btn.Click      += self._run_analysis
        self.cancel_btn.Click   += lambda s, a: self.Close()
        self.cat_all_btn.Click  += self._select_all_cats
        self.cat_none_btn.Click += self._select_none_cats
        self.confirm_btn.Click  += self._execute_deletion
        self.select_btn.Click   += self._execute_selection
        self.res_all_btn.Click  += self._select_all_results
        self.res_none_btn.Click += self._select_none_results
        self.back_btn.Click     += self._go_back
        self.op_duplicates_cb.Checked   += lambda s, a: self._sync_options_panel()
        self.op_duplicates_cb.Unchecked += lambda s, a: self._sync_options_panel()

    def _sync_options_panel(self):
        vis     = System.Windows.Visibility
        checked = bool(self.op_duplicates_cb.IsChecked)
        self.options_border.Visibility   = vis.Visible   if checked else vis.Collapsed
        self.categories_panel.Visibility = vis.Visible   if checked else vis.Collapsed

    # --- Seleção de categorias ---

    def _select_all_cats(self, sender, args):
        for c in self.categories:
            c.is_checked = True
        self.categories_lb.Items.Refresh()

    def _select_none_cats(self, sender, args):
        for c in self.categories:
            c.is_checked = False
        self.categories_lb.Items.Refresh()

    # --- Seleção nos resultados ---

    def _select_all_results(self, sender, args):
        for item in self.cleanup_items:
            item._is_selected = True
        self.results_dg.Items.Refresh()

    def _select_none_results(self, sender, args):
        for item in self.cleanup_items:
            item._is_selected = False
        self.results_dg.Items.Refresh()

    def _go_back(self, sender, args):
        vis = System.Windows.Visibility
        self.config_grid.Visibility  = vis.Visible
        self.results_grid.Visibility = vis.Collapsed
        self.run_btn.Visibility      = vis.Visible
        self.confirm_btn.Visibility  = vis.Collapsed
        self.select_btn.Visibility   = vis.Collapsed

    # --- Análise ---

    def _run_analysis(self, sender, args):
        ops = [
            self.op_duplicates_cb.IsChecked,
            self.op_view_templates_cb.IsChecked,
            self.op_text_types_cb.IsChecked,
            self.op_line_patterns_cb.IsChecked,
            self.op_load_cb.IsChecked,
            self.op_conduits_cb.IsChecked,
            self.op_view_filters_cb.IsChecked,
            self.op_cad_links_cb.IsChecked,
            self.op_orphan_tags_cb.IsChecked,
        ]
        if not any(ops):
            forms.alert("Selecione pelo menos uma operação de limpeza.")
            return

        self.cleanup_items = []

        # Elementos sobrepostos
        if self.op_duplicates_cb.IsChecked:
            selected_bics = [c.bic for c in self.categories if c.is_checked]
            if not selected_bics:
                forms.alert("Selecione pelo menos uma categoria para detectar sobrepostos.")
                return
            try:
                tol_curve = float(self.tol_curve_tb.Text)
                tol_point = float(self.tol_point_tb.Text)
            except Exception:
                forms.alert("Tolerâncias inválidas. Use números (ex: 1.0).")
                return
            is_active_view = bool(self.scope_view_rb.IsChecked)
            self.cleanup_items.extend(
                find_duplicate_elements(doc, selected_bics, is_active_view, tol_curve, tol_point)
            )

        # Modelos de vista não utilizados
        if self.op_view_templates_cb.IsChecked:
            self.cleanup_items.extend(find_unused_view_templates(doc))

        # Tipos de texto não utilizados
        if self.op_text_types_cb.IsChecked:
            self.cleanup_items.extend(find_unused_text_types(doc))

        # Padrões de linha não utilizados
        if self.op_line_patterns_cb.IsChecked:
            self.cleanup_items.extend(find_unused_line_patterns(doc))

        # Tipos de carga (LoadClassification, estrutural)
        if self.op_load_cb.IsChecked:
            self.cleanup_items.extend(find_unused_load_classifications(doc))

        # Eletrodutos com ambas as pontas livres
        if self.op_conduits_cb.IsChecked:
            self.cleanup_items.extend(find_unconnected_conduits(doc))

        # Filtros de vista não aplicados
        if self.op_view_filters_cb.IsChecked:
            self.cleanup_items.extend(find_unused_view_filters(doc))

        # Links CAD não carregados
        if self.op_cad_links_cb.IsChecked:
            self.cleanup_items.extend(find_unloaded_cad_links(doc))

        # Tags sem elemento hospedeiro
        if self.op_orphan_tags_cb.IsChecked:
            self.cleanup_items.extend(find_orphaned_tags(doc))

        if not self.cleanup_items:
            forms.alert("Nenhum item encontrado. O projeto está limpo nessas categorias!")
            return

        # Resumo por tipo
        counts = {}
        for item in self.cleanup_items:
            counts[item.ItemType] = counts.get(item.ItemType, 0) + 1
        parts = ["{}  {}".format(t, c) for t, c in sorted(counts.items())]
        self.summary_tb.Text = "{} item(s) encontrado(s) para limpeza   —   {}".format(
            len(self.cleanup_items), "   |   ".join(parts)
        )

        self.results_dg.ItemsSource = self.cleanup_items

        vis = System.Windows.Visibility
        self.config_grid.Visibility  = vis.Collapsed
        self.results_grid.Visibility = vis.Visible
        self.run_btn.Visibility      = vis.Collapsed
        self.confirm_btn.Visibility  = vis.Visible
        # "Selecionar no Revit" só faz sentido para elementos de modelo
        has_model = any(i.ItemType == "Sobreposto" for i in self.cleanup_items)
        self.select_btn.Visibility = vis.Visible if has_model else vis.Collapsed

    # --- Ações sobre os resultados ---

    def _execute_deletion(self, sender, args):
        to_delete = [i for i in self.cleanup_items if i.IsSelected]
        if not to_delete:
            forms.alert("Nenhum item selecionado para exclusão.")
            return

        load_items = [i for i in to_delete if i.ItemType == "Tipo de Carga"]
        substitute_id = None
        substitute_name = None
        load_refs_before = {}
        if load_items:
            from Autodesk.Revit.DB.Electrical import ElectricalLoadClassification
            deleting_ids = {i.ElementId.IntegerValue for i in load_items}
            for item in load_items:
                refs = _find_load_classification_references(doc, item.ElementId, item.Name)
                if refs:
                    load_refs_before[item.ElementId.IntegerValue] = refs

            available = [
                clf for clf in FilteredElementCollector(doc).OfClass(ElectricalLoadClassification)
                if clf.Id.IntegerValue not in deleting_ids
            ]
            if load_refs_before and not available:
                forms.alert(
                    "Nao ha classificacao substituta disponivel.\n"
                    "Mantenha pelo menos uma classificacao de carga no projeto antes de excluir.",
                    title="Operacao Cancelada"
                )
                return

            if load_refs_before:
                chosen = forms.SelectFromList.show(
                    sorted(clf.Name for clf in available),
                    title="Classificacao Substituta",
                    prompt=(
                        "As classificacoes selecionadas possuem elementos dependentes.\n"
                        "Escolha para qual classificacao essas referencias serao movidas:"
                    ),
                    multiselect=False
                )
                if not chosen:
                    return

                chosen_name = chosen[0] if isinstance(chosen, (list, tuple)) else chosen
                for clf in available:
                    if clf.Name == chosen_name:
                        substitute_id = clf.Id
                        substitute_name = clf.Name
                        break

        self.Close()
        self._delete_with_substitute(to_delete, load_items, substitute_id, substitute_name)
        return

    def _delete_with_substitute(self, to_delete, load_items, substitute_id, substitute_name):
        from Autodesk.Revit.DB import SubTransaction, Transaction as RawTransaction

        original_total = len(to_delete)
        deleted = 0
        skipped = []
        reassigned = 0
        locked_refs = []

        t = RawTransaction(doc, "LF Limpeza - Excluir")
        t.Start()
        try:
            if load_items and substitute_id:
                old_ids = {i.ElementId.IntegerValue for i in load_items}
                old_names = {i.Name for i in load_items}
                _reassign_load_classifications(doc, old_names, old_ids, substitute_id)
                changed, locked_refs = _replace_load_classification_references(
                    doc, old_ids, substitute_id
                )
                reassigned += changed
                try:
                    doc.Regenerate()
                except Exception:
                    pass

            if load_items:
                still_used_ids = set()
                for item in load_items:
                    refs = _find_load_classification_references(doc, item.ElementId, item.Name)
                    if refs:
                        still_used_ids.add(item.ElementId.IntegerValue)
                        skipped.append(u"{} (Tipo de Carga - ainda possui {} referencia(s))".format(
                            item.Name, len(refs)
                        ))
                if still_used_ids:
                    to_delete = [
                        i for i in to_delete
                        if i.ItemType != "Tipo de Carga"
                        or i.ElementId.IntegerValue not in still_used_ids
                    ]

            for item in to_delete:
                st = SubTransaction(doc)
                st.Start()
                try:
                    doc.Delete(item.ElementId)
                    st.Commit()
                    deleted += 1
                except Exception:
                    st.RollBack()
                    skipped.append(u"{} ({})".format(item.Name, item.ItemType))

            t.Commit()
        except Exception:
            t.RollBack()

        msg = u"{} de {} item(s) excluido(s).".format(deleted, original_total)
        if reassigned and substitute_name:
            msg += u"\n{} referencia(s) de tipo de carga foram movidas para '{}'.".format(
                reassigned, substitute_name
            )
        if skipped:
            msg += u"\n\nNao foi possivel excluir {} item(s):\n{}".format(
                len(skipped),
                u"\n".join(u"  - " + s for s in skipped[:15])
            )
            if len(skipped) > 15:
                msg += u"\n  ... e mais {}.".format(len(skipped) - 15)
        if locked_refs:
            msg += u"\n\n{} referencia(s) estavam bloqueadas/somente leitura e foram mantidas.".format(
                len(locked_refs)
            )
        forms.alert(msg)
        return

        protected_load_ids = set()
        for item in load_items:
            refs = _find_load_classification_references(doc, item.ElementId, item.Name)
            if refs:
                protected_load_ids.add(item.ElementId.IntegerValue)

        if protected_load_ids:
            to_delete = [
                i for i in to_delete
                if i.ItemType != "Tipo de Carga"
                or i.ElementId.IntegerValue not in protected_load_ids
            ]
            if not to_delete:
                forms.alert(
                    "As classificacoes de carga selecionadas ainda possuem dependencias "
                    "em circuitos, familias ou tipos. Nada foi excluido.",
                    title="Operacao Cancelada"
                )
                return

        # A reatribuicao automatica era instavel em alguns projetos.
        # A partir daqui, tipos de carga so seguem para exclusao se a checagem
        # acima nao encontrou nenhuma referencia.
        load_items = []

        # ── Classificações de carga exigem uma substituta ──────────────────
        substitute_id = None
        if False and load_items:
            from Autodesk.Revit.DB.Electrical import ElectricalLoadClassification
            deleting_ids = {i.ElementId.IntegerValue for i in load_items}
            available = [
                clf for clf in FilteredElementCollector(doc).OfClass(ElectricalLoadClassification)
                if clf.Id.IntegerValue not in deleting_ids
            ]
            if not available:
                forms.alert(
                    "Não há classificação substituta disponível.\n"
                    "Mantenha pelo menos uma classificação de carga no projeto antes de excluir.",
                    title="Operação Cancelada"
                )
                return

            chosen = forms.SelectFromList.show(
                sorted(clf.Name for clf in available),
                title="Classificação Substituta",
                prompt=(
                    "Os circuitos vinculados às classificações selecionadas precisam de uma\n"
                    "classificação substituta antes da exclusão. Selecione:"
                ),
                multiselect=False
            )
            if not chosen:
                return  # usuário cancelou

            chosen_name = chosen[0] if isinstance(chosen, (list, tuple)) else chosen
            for clf in available:
                if clf.Name == chosen_name:
                    substitute_id = clf.Id
                    break

        self.Close()

        from Autodesk.Revit.DB import SubTransaction, Transaction as RawTransaction

        deleted  = 0
        skipped  = []

        t = RawTransaction(doc, "LF Limpeza — Excluir")
        t.Start()
        try:
            # Reatribui referências de circuitos ANTES de deletar
            if load_items and substitute_id:
                _reassign_load_classifications(
                    doc,
                    {i.Name for i in load_items},
                    {i.ElementId.IntegerValue for i in load_items},
                    substitute_id
                )

            # Deleta item por item; usa SubTransaction para isolar falhas
            # sem disparar o diálogo de erro do Revit para o usuário
            for item in to_delete:
                st = SubTransaction(doc)
                st.Start()
                try:
                    doc.Delete(item.ElementId)
                    st.Commit()
                    deleted += 1
                except Exception:
                    st.RollBack()
                    skipped.append(u"{} ({})".format(item.Name, item.ItemType))

            t.Commit()
        except Exception:
            t.RollBack()

        msg = u"{} de {} item(s) excluído(s).".format(deleted, len(to_delete))
        if skipped:
            msg += u"\n\nNão foi possível excluir {} item(s) — provavelmente referenciados\nem definições de família ou tipos de sistema:\n{}".format(
                len(skipped),
                u"\n".join(u"  • " + s for s in skipped[:15])
            )
            if len(skipped) > 15:
                msg += u"\n  ... e mais {}.".format(len(skipped) - 15)
        forms.alert(msg)

    def _execute_selection(self, sender, args):
        model_items = [i for i in self.cleanup_items
                       if i.IsSelected and i.ItemType == "Sobreposto"]
        if not model_items:
            forms.alert("Nenhum elemento de modelo marcado para seleção.")
            return
        ids = List[ElementId]([i.ElementId for i in model_items])
        self.Close()
        uidoc.Selection.SetElementIds(ids)
        forms.alert("{} elemento(s) duplicado(s) selecionado(s) no Revit.".format(len(model_items)))


if __name__ == "__main__":
    if doc:
        OverkillWindow("OverkillWindow.xaml").show(modal=True)
