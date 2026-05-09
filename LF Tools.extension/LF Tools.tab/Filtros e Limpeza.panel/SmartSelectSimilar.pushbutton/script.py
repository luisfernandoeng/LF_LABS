# -*- coding: utf-8 -*-
"""Smart Select Similar — seleciona elementos similares por critérios.
Shift+Click: abre configuração completa.
Click normal: usa último preset salvo.
"""
__title__ = "Smart Select\nSimilar"
__author__ = "Luís Fernando"

import io
import os
import json

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import Thickness, Visibility, FontWeights
from System.Windows.Controls import CheckBox, TextBlock
from System.Windows.Media import SolidColorBrush, Color
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId, OverrideGraphicSettings,
    FillPatternElement, Color as RvtColor, Transaction,
    FamilyInstanceFilter, ElementLevelFilter,
    BuiltInParameter, RevitLinkInstance, Reference,
)
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, script as _pyscript

uidoc       = __revit__.ActiveUIDocument
doc         = uidoc.Document
active_view = uidoc.ActiveView


# ── Config ────────────────────────────────────────────────────────────────

def _script_dir():
    try:
        return __commandpath__
    except NameError:
        return os.path.dirname(os.path.abspath(__file__))


_CONFIG_FILE = os.path.join(_script_dir(), 'sss_config.json')


def _load_config():
    try:
        with io.open(_CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(data):
    try:
        with io.open(_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ── Highlight ─────────────────────────────────────────────────────────────

_solid_fill_id = None


def _get_solid_fill():
    global _solid_fill_id
    if _solid_fill_id is not None:
        return _solid_fill_id
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                _solid_fill_id = fp.Id
                return _solid_fill_id
        except:
            pass
    return None


def set_highlight(view, element_ids, apply=True):
    if apply:
        ogs = OverrideGraphicSettings()
        red = RvtColor(255, 0, 0)
        ogs.SetProjectionLineColor(red)
        try:
            ogs.SetProjectionLineWeight(8)
        except:
            pass
        sfid = _get_solid_fill()
        if sfid:
            try:
                ogs.SetSurfaceForegroundPatternId(sfid)
                ogs.SetSurfaceForegroundPatternColor(red)
                ogs.SetSurfaceForegroundPatternVisible(True)
            except:
                pass
            try:
                ogs.SetSurfaceBackgroundPatternId(sfid)
                ogs.SetSurfaceBackgroundPatternColor(red)
                ogs.SetSurfaceBackgroundPatternVisible(True)
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


# ── Extração completa de parâmetros (apenas para exibição na UI) ──────────

_BLACKLIST = frozenset([u'mm\xb2', u'Fase ', u'Fase-', u'Neutro', u'Terra', u'Retorno'])


def _to_text(value):
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except:
        try:
            return str(value)
        except:
            return u''


def _elem_type(elem):
    try:
        tid = elem.GetTypeId()
        if tid and tid != ElementId.InvalidElementId:
            return elem.Document.GetElement(tid)
    except:
        pass
    return None


def _family_name(elem):
    try:
        return elem.Symbol.FamilyName
    except:
        pass
    et = _elem_type(elem)
    if et:
        try:
            return et.FamilyName
        except:
            pass
    try:
        return elem.Name
    except:
        return None


def _type_name(elem):
    try:
        return elem.Symbol.Name or elem.Name
    except:
        pass
    et = _elem_type(elem)
    if et:
        try:
            return et.Name
        except:
            pass
    try:
        return elem.Name
    except:
        return None


def get_element_parameters(element):
    params = {}
    try:
        params['Category'] = element.Category.Name if element.Category else 'Sem Categoria'
        fam = _family_name(element)
        typ = _type_name(element)
        if fam:
            params['Family'] = fam
        if typ:
            params['Type'] = typ
        try:
            lv = element.Document.GetElement(element.LevelId)
            if lv:
                params['Level'] = lv.Name
        except:
            pass
        try:
            if element.Document.IsWorkshared:
                ws = element.Document.GetWorksetTable().GetWorkset(element.WorksetId)
                if ws:
                    params['Workset'] = ws.Name
        except:
            pass
        try:
            do = element.DesignOption
            if do:
                params['Design Option'] = do.Name
        except:
            pass
        try:
            p = element.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if p and p.AsElementId() != ElementId.InvalidElementId:
                ph = element.Document.GetElement(p.AsElementId())
                if ph:
                    params['Phase Created'] = ph.Name
        except:
            pass
    except:
        pass

    def _add(param_set):
        for p in param_set:
            try:
                name = p.Definition.Name
                if name in params or not p.HasValue:
                    continue
                st = p.StorageType
                from Autodesk.Revit.DB import StorageType as ST
                if st == ST.String:
                    v = p.AsString() or u''
                elif st == ST.Integer:
                    v = p.AsValueString() or str(p.AsInteger())
                elif st == ST.Double:
                    v = p.AsValueString() or u'{:.4f}'.format(p.AsDouble())
                else:
                    v = p.AsValueString() or u''
                params[name] = v
            except:
                continue

    try:
        _add(element.Parameters)
    except:
        pass
    try:
        et = element.Document.GetElement(element.GetTypeId())
        if et:
            _add(et.Parameters)
    except:
        pass

    return params


# ── Leitura rápida de valor único (para filtragem) ────────────────────────

def _param_value(elem, name):
    """Lê um único parâmetro pelo nome sem iterar todos os parâmetros."""
    if name == 'Category':
        return elem.Category.Name if elem.Category else None
    if name == 'Family':
        return _family_name(elem)
    if name == 'Type':
        return _type_name(elem)
    if name == 'Level':
        try:
            lv = elem.Document.GetElement(elem.LevelId)
            return lv.Name if lv else None
        except:
            return None
    if name == 'Workset':
        try:
            if elem.Document.IsWorkshared:
                ws = elem.Document.GetWorksetTable().GetWorkset(elem.WorksetId)
                return ws.Name if ws else None
        except:
            return None
    if name == 'Phase Created':
        try:
            p = elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if p and p.AsElementId() != ElementId.InvalidElementId:
                ph = elem.Document.GetElement(p.AsElementId())
                return ph.Name if ph else None
        except:
            return None
    if name == 'Design Option':
        try:
            do = elem.DesignOption
            return do.Name if do else None
        except:
            return None
    # Genérico: instância primeiro, depois tipo
    try:
        p = elem.LookupParameter(name)
        if p and p.HasValue:
            return p.AsValueString() or p.AsString() or u''
    except:
        pass
    try:
        et = elem.Document.GetElement(elem.GetTypeId())
        if et:
            p = et.LookupParameter(name)
            if p and p.HasValue:
                return p.AsValueString() or p.AsString() or u''
    except:
        pass
    return None


# ── Filtragem otimizada ────────────────────────────────────────────────────

def _has_filter_value(value):
    if value is None:
        return False
    try:
        return bool(_to_text(value).strip())
    except:
        return True


def _effective_criteria(ref_elem, configured_criteria, include_standard=True):
    """Mantem o preset salvo, mas usa so o que existe no elemento clicado."""
    ordered = []
    if include_standard:
        ordered.extend(['Category', 'Family', 'Type'])
    if configured_criteria:
        ordered.extend(configured_criteria)

    result = []
    seen = set()
    for c in ordered:
        c = _to_text(c or u'').strip()
        if not c or c in seen:
            continue
        if c == 'Category' or _has_filter_value(_param_value(ref_elem, c)):
            result.append(c)
            seen.add(c)
    return result


def filter_similar(target_doc, ref_elem, criteria, scope, ref_cat_id=None):
    """
    Usa filtros nativos C++ do Revit (FamilyInstanceFilter, ElementLevelFilter)
    para critérios comuns, depois aplica os demais via LookupParameter.

    FamilyInstanceFilter cobre Família+Tipo em uma única passagem sem tocar
    em parâmetros — é ordens de magnitude mais rápido que comparação manual.
    """
    use_view = (scope == 'active_view' and target_doc == doc)
    coll = (FilteredElementCollector(target_doc, active_view.Id)
            if use_view else FilteredElementCollector(target_doc))
    coll = coll.WhereElementIsNotElementType()
    if ref_cat_id:
        coll = coll.OfCategoryId(ref_cat_id)

    soft = list(criteria)

    # Família + Tipo → FamilyInstanceFilter (quick filter, roda em C++)
    if 'Family' in soft and 'Type' in soft:
        try:
            symbol = ref_elem.Symbol
            type_id = ref_elem.GetTypeId()
            if symbol and type_id != ElementId.InvalidElementId:
                coll = coll.WherePasses(FamilyInstanceFilter(target_doc, type_id))
                soft = [c for c in soft if c not in ('Family', 'Type', 'Category')]
        except:
            pass

    # Nível → ElementLevelFilter (também quick filter)
    if 'Level' in soft:
        try:
            lvl_id = ref_elem.LevelId
            if lvl_id != ElementId.InvalidElementId:
                coll = coll.WherePasses(ElementLevelFilter(lvl_id))
                soft = [c for c in soft if c != 'Level']
        except:
            pass

    # Aplicar ElementParameterFilter nativo (C++) para acelerar parâmetros de instância
    from Autodesk.Revit.DB import (ParameterValueProvider, FilterStringRule, FilterStringEquals,
                                   FilterNumericEquals, ElementParameterFilter, FilterIntegerRule,
                                   FilterDoubleRule, FilterElementIdRule, StorageType)

    def _create_string_rule(provider, evaluator, value):
        try:
            return FilterStringRule(provider, evaluator, value) # Revit 2022+
        except:
            return FilterStringRule(provider, evaluator, value, False) # Revit 2021-

    for c in list(soft):
        try:
            p = ref_elem.LookupParameter(c)
            # Filtro nativo só funciona direto para parâmetros da instância
            if p and p.HasValue and p.Id != ElementId.InvalidElementId:
                provider = ParameterValueProvider(p.Id)
                rule = None
                st = p.StorageType
                if st == StorageType.String:
                    rule = _create_string_rule(provider, FilterStringEquals(), p.AsString() or "")
                elif st == StorageType.Integer:
                    rule = FilterIntegerRule(provider, FilterNumericEquals(), p.AsInteger())
                elif st == StorageType.Double:
                    rule = FilterDoubleRule(provider, FilterNumericEquals(), p.AsDouble(), 1e-6)
                elif st == StorageType.ElementId:
                    rule = FilterElementIdRule(provider, FilterNumericEquals(), p.AsElementId())
                
                if rule:
                    coll = coll.WherePasses(ElementParameterFilter(rule))
                    soft.remove(c)
        except:
            pass

    # Se todos os critérios foram cobertos por filtros nativos, retorna direto
    if not soft:
        return list(coll.ToElements())

    # Pré-computa cache de IDs O(1) para os critérios que não puderam usar C++ filter (ex: de Tipo)
    param_ids = []
    try:
        type_elem = target_doc.GetElement(ref_elem.GetTypeId()) if ref_elem.GetTypeId() != ElementId.InvalidElementId else None
    except:
        type_elem = None
        
    for c in soft:
        pid = None
        is_type = False
        p = ref_elem.LookupParameter(c)
        if p:
            pid = p.Id
        elif type_elem:
            pt = type_elem.LookupParameter(c)
            if pt:
                pid = pt.Id
                is_type = True
        param_ids.append((c, pid, is_type, _param_value(ref_elem, c)))

    result = []
    for elem in coll:
        try:
            ok = True
            current_type_elem = None
            for c, pid, is_type, ref_val in param_ids:
                val = None
                if pid:
                    # Busca O(1) via ElementId
                    if is_type:
                        if not current_type_elem:
                            current_type_elem = target_doc.GetElement(elem.GetTypeId())
                        if current_type_elem:
                            p = current_type_elem.get_Parameter(pid)
                            if p and p.HasValue:
                                val = p.AsValueString() or p.AsString() or u''
                    else:
                        p = elem.get_Parameter(pid)
                        if p and p.HasValue:
                            val = p.AsValueString() or p.AsString() or u''
                
                # Fallback lento para propriedades complexas (Name, Category, Level, etc)
                if val is None:
                    val = _param_value(elem, c)

                if val != ref_val:
                    ok = False
                    break
            if ok:
                result.append(elem)
        except:
            pass
    return result


# ── Janela ────────────────────────────────────────────────────────────────

class SmartSelectSimilarWindow(forms.WPFWindow):
    def __init__(self, xaml_file, ref_elem, is_linked, target_doc, link_inst):
        forms.WPFWindow.__init__(self, xaml_file)
        self._ref_elem          = ref_elem
        self._is_linked         = is_linked
        self._target_doc        = target_doc
        self._link_inst         = link_inst
        self._cat_id            = ref_elem.Category.Id if ref_elem.Category else None
        self._params            = get_element_parameters(ref_elem)
        self._criteria          = []
        self._matches           = []
        self._highlighted       = []
        self._custom_param_names = set()

        prefix = u'[VÍNCULO] ' if is_linked else u''
        self.ElementNameLabel.Text = u'Objeto: ' + prefix + getattr(ref_elem, 'Name', u'?')
        self.ElementIdLabel.Text   = u'ID: ' + str(ref_elem.Id)

        if is_linked:
            self.RadioActiveView.Content = u'Vista Ativa (Vínculo)'
            self.RadioProject.Content    = u'Vínculo Inteiro'

        self._populate_params()
        self._update_preview()

        self.SearchBox.TextChanged          += self._on_search
        self.SearchBox.GotFocus             += self._on_search_focus
        self.SearchBox.LostFocus            += self._on_search_focus
        self.RadioActiveView.Checked        += self._on_scope
        self.RadioProject.Checked           += self._on_scope
        self.BtnPresetFamilyType.Click      += self._preset_family_type
        self.BtnPresetFamilyComments.Click  += self._preset_family_comments
        self.BtnClearAll.Click              += self._clear_all
        self.PaintButton.Click              += self._on_paint
        self.ZoomButton.Click               += self._on_zoom
        self.BtnPreview.Click               += self._on_preview
        self.BtnSelect.Click                += self._on_select
        self.Closing                        += self._on_closing

    # ── Ciclo de vida ──────────────────────────────────────────────────────

    def _on_closing(self, s, a):
        self._clear_highlight()

    def _clear_highlight(self):
        if not self._highlighted:
            return
        t = Transaction(doc, u'Limpar Highlight')
        t.Start()
        try:
            set_highlight(active_view, self._highlighted, apply=False)
            t.Commit()
        except:
            try:
                t.RollBack()
            except:
                pass
        self._highlighted = []

    # ── Parâmetros ─────────────────────────────────────────────────────────

    def _populate_params(self):
        self._custom_param_names = set()
        self.ParametersPanel.Children.Clear()
        groups = {
            u'Básicas':      ['Category', 'Family', 'Type', 'Level'],
            u'Identidade':   ['Comments', 'Mark', 'Workset'],
            u'Projeto':      ['Phase Created', 'Design Option'],
            u'Customizados': [],
        }
        base = (groups[u'Básicas'] + groups[u'Identidade'] + groups[u'Projeto'])
        for k in sorted(self._params.keys()):
            if k not in base:
                groups[u'Customizados'].append(k)

        ACCENT  = SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD7))
        MUTED   = SolidColorBrush(Color.FromRgb(0xAA, 0xAA, 0xAA))

        for group_name in [u'Básicas', u'Identidade', u'Projeto', u'Customizados']:
            params = groups[group_name]
            if not params:
                continue
            is_custom = (group_name == u'Customizados')
            header_added = False
            for p_name in params:
                if any(b in p_name for b in _BLACKLIST):
                    continue
                if p_name not in self._params:
                    continue
                if not header_added:
                    hdr            = TextBlock()
                    hdr.Text       = group_name.upper()
                    hdr.Foreground = ACCENT
                    hdr.FontWeight = FontWeights.SemiBold
                    hdr.Margin     = Thickness(0, 14, 0, 4)
                    hdr.FontSize   = 10
                    if is_custom:
                        hdr.Tag        = u'custom_header'
                        hdr.Visibility = Visibility.Collapsed
                    self.ParametersPanel.Children.Add(hdr)
                    header_added = True

                val     = self._params.get(p_name, u'')
                display = val if (val and str(val).strip()) else u'<vazio>'
                cb           = CheckBox()
                cb.Content   = u'{} : {}'.format(p_name, display)
                cb.Tag       = p_name
                cb.IsChecked = p_name in ('Category', 'Family', 'Type')
                cb.Checked   += self._on_criteria
                cb.Unchecked += self._on_criteria
                if is_custom:
                    self._custom_param_names.add(p_name)
                    cb.Visibility = Visibility.Collapsed
                self.ParametersPanel.Children.Add(cb)

        # Nota sobre parâmetros customizados (só aparece quando não está buscando)
        custom_count = sum(
            1 for k in groups[u'Customizados']
            if k in self._params and not any(b in k for b in _BLACKLIST)
        )
        if custom_count > 0:
            hint = TextBlock()
            hint.Tag        = u'custom_hint'
            hint.Text       = u'⋯  {} parâmetro(s) adicional(is) — use a busca acima'.format(custom_count)
            hint.Foreground = MUTED
            hint.FontSize   = 10
            hint.FontStyle  = System.Windows.FontStyles.Italic
            hint.Margin     = Thickness(0, 10, 0, 2)
            self.ParametersPanel.Children.Add(hint)

        # Popula critérios iniciais (Category + Family já marcados)
        self._on_criteria(None, None)

    # ── Eventos ────────────────────────────────────────────────────────────

    def _on_search_focus(self, s, a):
        txt = str(self.SearchBox.Text or u'')
        focused = self.SearchBox.IsKeyboardFocused
        self.SearchHint.Visibility = Visibility.Collapsed if (txt or focused) else Visibility.Visible

    def _on_search(self, s, a):
        txt       = str(self.SearchBox.Text or u'').lower()
        searching = bool(txt)
        self.SearchHint.Visibility = Visibility.Collapsed if searching else Visibility.Visible
        for child in self.ParametersPanel.Children:
            if isinstance(child, CheckBox):
                tag       = str(child.Tag or u'')
                is_custom = tag in self._custom_param_names
                if is_custom:
                    # Parâmetros customizados: visíveis apenas ao buscar
                    visible = searching and (txt in str(child.Content or u'').lower())
                else:
                    visible = (not txt) or (txt in str(child.Content or u'').lower())
                child.Visibility = Visibility.Visible if visible else Visibility.Collapsed
            elif isinstance(child, TextBlock):
                tag = str(child.Tag or u'')
                if tag == u'custom_header':
                    child.Visibility = Visibility.Visible if searching else Visibility.Collapsed
                elif tag == u'custom_hint':
                    child.Visibility = Visibility.Collapsed if searching else Visibility.Visible

    def _on_scope(self, s, a):
        self._update_preview()

    def _on_criteria(self, s, a):
        self._criteria = [
            str(c.Tag) for c in self.ParametersPanel.Children
            if isinstance(c, CheckBox) and c.IsChecked == True
        ]
        self._update_preview()

    def _update_preview(self):
        if not self._criteria:
            self.PreviewLabel.Text = u'Selecione ao menos 1 critério'
            self._matches = []
            return

        scope = 'active_view' if self.RadioActiveView.IsChecked else 'project'
        self._matches = filter_similar(
            self._target_doc, self._ref_elem, self._criteria, scope, self._cat_id
        )
        n = len(self._matches)
        self.PreviewLabel.Text = u'{} elemento(s) encontrado(s)'.format(n)

        if not self._is_linked:
            self._clear_highlight()
            if self._matches:
                self._highlighted = [e.Id for e in self._matches]
                t = Transaction(doc, u'Highlight Temporário')
                t.Start()
                try:
                    set_highlight(active_view, self._highlighted, apply=True)
                    t.Commit()
                except:
                    try:
                        t.RollBack()
                    except:
                        pass

    def _batch_checks(self, tags):
        tags_set = set(tags)
        has_custom_checked = False
        for c in self.ParametersPanel.Children:
            if isinstance(c, CheckBox):
                tag = str(c.Tag or u'')
                c.IsChecked = tag in tags_set
                # Custom param marcado deve ficar visível
                if tag in self._custom_param_names:
                    if tag in tags_set:
                        c.Visibility = Visibility.Visible
                        has_custom_checked = True
                    else:
                        c.Visibility = Visibility.Collapsed
            elif isinstance(c, TextBlock):
                tag = str(c.Tag or u'')
                if tag == u'custom_header':
                    c.Visibility = Visibility.Visible if has_custom_checked else Visibility.Collapsed
        self._on_criteria(None, None)

    def _preset_family_type(self, s, a):
        self._batch_checks(['Category', 'Family', 'Type'])

    def _preset_family_comments(self, s, a):
        self._batch_checks(['Category', 'Family', 'Comments'])

    def _clear_all(self, s, a):
        self._batch_checks([])

    # ── Ações ──────────────────────────────────────────────────────────────

    def _on_preview(self, s, a):
        if not self._matches:
            return
        if self._is_linked:
            refs = List[Reference]()
            for e in self._matches:
                try:
                    refs.Add(Reference(e).CreateLinkReference(self._link_inst))
                except:
                    pass
            if refs.Count:
                uidoc.Selection.SetReferences(refs)
            uidoc.ShowElements(self._link_inst.Id)
        else:
            ids = List[ElementId]()
            for e in self._matches:
                ids.Add(e.Id)
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(ids)

    def _on_paint(self, s, a):
        if not self._matches:
            return
        if self._is_linked:
            forms.alert(u'Pintura em vínculos não é suportada pelo Revit.')
            return
        ids = [e.Id for e in self._matches]
        t = Transaction(doc, u'Pintar Elementos')
        t.Start()
        try:
            set_highlight(active_view, ids, apply=True)
            t.Commit()
        except:
            try:
                t.RollBack()
            except:
                pass
        if ids == self._highlighted:
            self._highlighted = []
        forms.toast(u'Seleção pintada!')

    def _on_zoom(self, s, a):
        if not self._matches:
            return
        if self._is_linked:
            refs = List[Reference]()
            for e in self._matches:
                try:
                    refs.Add(Reference(e).CreateLinkReference(self._link_inst))
                except:
                    pass
            if refs.Count:
                uidoc.Selection.SetReferences(refs)
            uidoc.ShowElements(self._link_inst.Id)
        else:
            ids = List[ElementId]()
            for e in self._matches:
                ids.Add(e.Id)
            uidoc.ShowElements(ids)

    def _on_select(self, s, a):
        if self._matches:
            try:
                _save_config({
                    'criteria': u','.join(self._criteria),
                    'scope': 'active_view' if self.RadioActiveView.IsChecked else 'project',
                })
            except:
                pass
            if self._is_linked:
                refs = List[Reference]()
                for e in self._matches:
                    try:
                        refs.Add(Reference(e).CreateLinkReference(self._link_inst))
                    except:
                        pass
                if refs.Count:
                    uidoc.Selection.SetReferences(refs)
                else:
                    forms.alert(u'Falha ao criar referências vinculadas.')
            else:
                ids = List[ElementId]()
                for e in self._matches:
                    ids.Add(e.Id)
                uidoc.Selection.SetElementIds(ids)
        self.Close()


# ── Entry Point ───────────────────────────────────────────────────────────

ref       = None
elem      = None
is_linked = False
tgt_doc   = doc
link_inst = None

# Usa elemento já selecionado, ou pede pick
try:
    sel_ids = list(uidoc.Selection.GetElementIds())
    if sel_ids:
        elem = doc.GetElement(sel_ids[0])
except:
    pass

if not elem:
    try:
        ref  = uidoc.Selection.PickObject(ObjectType.PointOnElement,
                                          u'Selecione o elemento de referência')
        elem = doc.GetElement(ref.ElementId)
    except:
        raise SystemExit()

# Detecta se é elemento de vínculo
if isinstance(elem, RevitLinkInstance):
    is_linked = True
    link_inst = elem
    tgt_doc   = link_inst.GetLinkDocument()
    if not tgt_doc:
        forms.alert(u'Vínculo descarregado.', exitscript=True)
    linked_id = ref.LinkedElementId if ref else None
    if linked_id and linked_id != ElementId.InvalidElementId:
        elem = tgt_doc.GetElement(linked_id)
    else:
        forms.alert(u'Elemento inválido no vínculo.', exitscript=True)
elif ref:
    try:
        if ref.LinkedElementId != ElementId.InvalidElementId:
            candidate = doc.GetElement(ref.ElementId)
            if isinstance(candidate, RevitLinkInstance):
                is_linked = True
                link_inst = candidate
                tgt_doc   = link_inst.GetLinkDocument()
                elem      = tgt_doc.GetElement(ref.LinkedElementId)
    except:
        pass

if not elem:
    forms.alert(u'Não foi possível acessar o elemento.', exitscript=True)

# Carrega config salva
saved_criteria = None
saved_scope    = 'active_view'
try:
    cfg = _load_config()
    c_str = cfg.get('criteria')
    if c_str:
        saved_criteria = [x for x in _to_text(c_str).split(',') if x]
    saved_scope = _to_text(cfg.get('scope', 'active_view'))
except:
    pass

# Shift → abre janela; Click normal → executa com último preset
shift_pressed = False
try:
    from System.Windows.Input import Keyboard, Key
    shift_pressed = (Keyboard.IsKeyDown(Key.LeftShift) or
                     Keyboard.IsKeyDown(Key.RightShift))
except:
    pass

xaml_path = _pyscript.get_bundle_file('ui.xaml')

if shift_pressed or not saved_criteria:
    win = SmartSelectSimilarWindow(xaml_path, elem, is_linked, tgt_doc, link_inst)
    if saved_criteria:
        win.RadioProject.IsChecked    = (saved_scope == 'project')
        win.RadioActiveView.IsChecked = (saved_scope == 'active_view')
        win._batch_checks(_effective_criteria(elem, saved_criteria, include_standard=True))
    win.ShowDialog()
else:
    cat_id  = elem.Category.Id if elem.Category else None
    effective_criteria = _effective_criteria(elem, saved_criteria, include_standard=True)
    with forms.ProgressBar(title=u"Smart Select Similar...", cancellable=False) as pb:
        pb.update_progress(0, 1)
        matches = filter_similar(tgt_doc, elem, effective_criteria, saved_scope, cat_id)
        pb.update_progress(1, 1)
    if matches:
        if is_linked:
            refs = List[Reference]()
            for e in matches:
                try:
                    refs.Add(Reference(e).CreateLinkReference(link_inst))
                except:
                    pass
            if refs.Count:
                uidoc.Selection.SetReferences(refs)
        else:
            ids = List[ElementId]()
            for e in matches:
                ids.Add(e.Id)
            uidoc.Selection.SetElementIds(ids)
    else:
        forms.alert(u'Nenhum similar encontrado.', title=u'Smart Select Similar')
