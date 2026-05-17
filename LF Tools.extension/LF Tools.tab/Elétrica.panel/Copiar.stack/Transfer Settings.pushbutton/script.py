# -*- coding: utf-8 -*-
"""
Transfer Settings — Transferência Cirúrgica de Configurações
=============================================================
LF Tools — pyRevit Extension
"""
__title__ = "Transfer\nSettings"
__author__ = "Luís Fernando"

DEBUG_MODE = False

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows.Controls import CheckBox, TextBlock, StackPanel, Expander
from System.Windows import Thickness, FontWeights, Visibility
from System.Windows.Input import Keyboard, Key
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId,
    Transaction, ParameterFilterElement,
    FillPatternElement, View, PhaseFilter,
    ElementTransformUtils, CopyPasteOptions, Transform,
    LinePatternElement,
    RevitLinkInstance, RevitLinkType,
    BuiltInCategory, FamilySymbol, StorageType,
)
from Autodesk.Revit.DB.Electrical import (
    PanelScheduleTemplate, ElectricalLoadClassification,
    ElectricalDemandFactorDefinition, DistributionSysType, VoltageType,
    WireType, ConduitType, CableTrayType,
)
from pyrevit import revit, forms, script
from lf_utils import DebugLogger, make_warning_swallower

dbg    = DebugLogger(DEBUG_MODE)
doc    = revit.doc
uiapp  = revit.HOST_APP.uiapp
app    = revit.HOST_APP.app
output = script.get_output()


# ══════════════════════════════════════════════════════════════
#  CATEGORIAS DE FAMÍLIA
# ══════════════════════════════════════════════════════════════

FAMILY_CATEGORIES = [
    (u"Mobiliário",              BuiltInCategory.OST_Furniture),
    (u"Sistemas de Mobiliário",  BuiltInCategory.OST_FurnitureSystems),
    (u"Armários",                BuiltInCategory.OST_Casework),
    (u"Portas",                  BuiltInCategory.OST_Doors),
    (u"Janelas",                 BuiltInCategory.OST_Windows),
    (u"Paredes / Revestimentos", BuiltInCategory.OST_Walls),
    (u"Pisos / Revestimentos",   BuiltInCategory.OST_Floors),
    (u"Forros",                  BuiltInCategory.OST_Ceilings),
    (u"Equipamento Mecânico",    BuiltInCategory.OST_MechanicalEquipment),
    (u"Equipamento Elétrico",    BuiltInCategory.OST_ElectricalEquipment),
    (u"Aparelhos Hidráulicos",   BuiltInCategory.OST_PlumbingFixtures),
    (u"Equipamento Especial",    BuiltInCategory.OST_SpecialityEquipment),
    (u"Luminárias",              BuiltInCategory.OST_LightingFixtures),
    (u"Modelos Genéricos",       BuiltInCategory.OST_GenericModel),
]

_FAMILY_CAT_MAP = {cat_name: bic for cat_name, bic in FAMILY_CATEGORIES}


# ══════════════════════════════════════════════════════════════
#  CLASSES DE APOIO
# ══════════════════════════════════════════════════════════════

class TransferableItem(object):
    def __init__(self, name, el_id, category):
        self.name       = name
        self.element_id = el_id
        self.category   = category
        self.checkbox   = None


class StandardCategory(object):
    def __init__(self, name, items=None):
        self.name            = name
        self.items           = items or []
        self.header_checkbox = None


class FamilyTypeItem(object):
    def __init__(self, fam_name, type_name, type_id, category, bic, type_class):
        self.fam_name   = fam_name
        self.type_name  = type_name
        self.symbol_id  = type_id
        self.category   = category
        self.bic        = bic
        self.type_class = type_class
        self.checkbox   = None


class DuplicateNameHandler(object):
    def ask_user(self, name):
        result = forms.CommandSwitchWindow.show(
            ["Sobrescrever", "Renomear (sufixo)", "Pular"],
            message=u"O elemento '{}' já existe no destino.\nO que deseja fazer?".format(name)
        )
        if result == "Sobrescrever":
            return "overwrite"
        elif result and "Renomear" in result:
            return "rename"
        return "skip"


# ══════════════════════════════════════════════════════════════
#  FONTES DE DOCUMENTOS
# ══════════════════════════════════════════════════════════════

def get_open_documents():
    docs = []
    for d in app.Documents:
        try:
            if not d.Equals(doc) and not d.IsFamilyDocument:
                docs.append(d)
        except Exception:
            pass
    return docs


def get_linked_documents():
    docs = []
    col = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    for link in col:
        try:
            link_doc = link.GetLinkDocument()
            if link_doc and not link_doc.Equals(doc):
                docs.append(link_doc)
        except Exception:
            pass
    return docs


# ══════════════════════════════════════════════════════════════
#  COLETORES — STANDARDS
# ══════════════════════════════════════════════════════════════

def _safe_name(el):
    try:
        return el.Name
    except Exception:
        return ""


def _collect_by_class(source_doc, cls, cat_name, filter_fn=None):
    items = []
    try:
        col = FilteredElementCollector(source_doc).OfClass(cls)
        for e in col:
            try:
                name = _safe_name(e)
                if not name:
                    continue
                if filter_fn and not filter_fn(e):
                    continue
                items.append(TransferableItem(name, e.Id, cat_name))
            except Exception:
                pass
    except Exception:
        pass
    return sorted(items, key=lambda x: x.name)


def collect_view_filters(source_doc):
    return _collect_by_class(source_doc, ParameterFilterElement, "Filtros de Vista")

def collect_view_templates(source_doc):
    return _collect_by_class(source_doc, View, "View Templates",
                             filter_fn=lambda v: v.IsTemplate)

def collect_fill_patterns(source_doc):
    return _collect_by_class(source_doc, FillPatternElement, u"Padrões de Preenchimento")

def collect_line_patterns(source_doc):
    return _collect_by_class(source_doc, LinePatternElement, u"Padrões de Linha")

def collect_phase_filters(source_doc):
    return _collect_by_class(source_doc, PhaseFilter, "Filtros de Fase")

def collect_panel_schedules(source_doc):
    return _collect_by_class(source_doc, PanelScheduleTemplate, "Modelos de Tabela de Carga")

def collect_load_class(source_doc):
    return _collect_by_class(source_doc, ElectricalLoadClassification, u"Classificações de Carga")

def collect_demand_factors(source_doc):
    return _collect_by_class(source_doc, ElectricalDemandFactorDefinition, "Fatores de Demanda")

def collect_wire_types(source_doc):
    return _collect_by_class(source_doc, WireType, u"Tipos de Fiação")

def collect_distribution_sys(source_doc):
    return _collect_by_class(source_doc, DistributionSysType, u"Sistemas de Distribuição")

def collect_voltage_types(source_doc):
    return _collect_by_class(source_doc, VoltageType, u"Tipos de Tensão")

def collect_conduit_types(source_doc):
    return _collect_by_class(source_doc, ConduitType, "Tipos de Eletroduto")

def collect_cable_tray_types(source_doc):
    return _collect_by_class(source_doc, CableTrayType, "Tipos de Eletrocalha")


_CAT_CLASS_MAP = {
    "Filtros de Vista":             ParameterFilterElement,
    "View Templates":               View,
    u"Padrões de Preenchimento":    FillPatternElement,
    u"Padrões de Linha":            LinePatternElement,
    "Filtros de Fase":              PhaseFilter,
    "Modelos de Tabela de Carga":   PanelScheduleTemplate,
    u"Classificações de Carga":     ElectricalLoadClassification,
    "Fatores de Demanda":           ElectricalDemandFactorDefinition,
    u"Tipos de Fiação":             WireType,
    u"Sistemas de Distribuição":    DistributionSysType,
    u"Tipos de Tensão":             VoltageType,
    "Tipos de Eletroduto":          ConduitType,
    "Tipos de Eletrocalha":         CableTrayType,
}


def _find_in_active_doc(name, category):
    cls = _CAT_CLASS_MAP.get(category)
    if cls is None:
        return None
    try:
        col = FilteredElementCollector(doc).OfClass(cls)
        for el in col:
            try:
                if category == "View Templates" and not el.IsTemplate:
                    continue
                if _safe_name(el) == name:
                    return el
            except Exception:
                pass
    except Exception:
        pass
    return None


def check_name_exists(name, category):
    return _find_in_active_doc(name, category) is not None


# ══════════════════════════════════════════════════════════════
#  COLETORES — PARÂMETROS DE FAMÍLIA
# ══════════════════════════════════════════════════════════════

def collect_family_types(source_doc):
    """Retorna tipos usados no modelo fonte, ignorando tipos apenas carregados."""
    result = []
    for cat_name, bic in FAMILY_CATEGORIES:
        try:
            items = []
            seen_type_ids = set()
            col = FilteredElementCollector(source_doc).OfCategory(bic).WhereElementIsNotElementType()
            for inst in col:
                try:
                    type_id = inst.GetTypeId()
                    if type_id == ElementId.InvalidElementId:
                        continue
                    type_int = type_id.IntegerValue
                    if type_int in seen_type_ids:
                        continue

                    type_el = source_doc.GetElement(type_id)
                    if type_el is None:
                        continue

                    fam_name = ""
                    if isinstance(type_el, FamilySymbol):
                        fam = type_el.Family
                        if fam is None:
                            continue
                        fam_name = fam.Name

                    type_name = _safe_name(type_el)
                    if type_name:
                        seen_type_ids.add(type_int)
                        items.append(FamilyTypeItem(
                            fam_name, type_name, type_id, cat_name, bic, type_el.GetType()
                        ))
                except Exception:
                    pass
            if items:
                items.sort(key=lambda x: (x.fam_name or x.category, x.type_name))
                result.append((cat_name, items))
        except Exception:
            pass
    return result


def find_matching_symbol(dest_doc, fam_name, type_name, bic, type_class=None):
    try:
        col = FilteredElementCollector(dest_doc).OfCategory(bic).WhereElementIsElementType()
        if type_class is not None:
            col = col.OfClass(type_class)
        for sym in col:
            try:
                if _safe_name(sym) != type_name:
                    continue
                if fam_name:
                    if isinstance(sym, FamilySymbol) and sym.Family and sym.Family.Name == fam_name:
                        return sym
                    continue
                return sym
            except Exception:
                pass
    except Exception:
        pass
    return None


def find_matching_used_symbol(dest_doc, item):
    """Procura no destino o mesmo tipo usado: familia+tipo ou tipo de sistema."""
    match = find_matching_symbol(
        dest_doc, item.fam_name, item.type_name, item.bic, item.type_class
    )
    if match is None:
        return None

    try:
        used_ids = set()
        col = FilteredElementCollector(dest_doc).OfCategory(item.bic).WhereElementIsNotElementType()
        for inst in col:
            try:
                type_id = inst.GetTypeId()
                if type_id != ElementId.InvalidElementId:
                    used_ids.add(type_id.IntegerValue)
            except Exception:
                pass
        if match.Id.IntegerValue in used_ids:
            return match
    except Exception:
        pass
    return None


def find_matching_referenced_element(source_ref, dest_doc):
    if source_ref is None:
        return None
    ref_name = _safe_name(source_ref)
    if not ref_name:
        return None
    try:
        col = FilteredElementCollector(dest_doc).OfClass(source_ref.GetType())
        for dest_ref in col:
            try:
                if _safe_name(dest_ref) == ref_name:
                    return dest_ref
            except Exception:
                pass
    except Exception:
        pass
    return None


def copy_type_params(source_sym, dest_sym):
    """Copia valores de parâmetros de tipo. Retorna (copiados, pulados)."""
    copied  = 0
    skipped = 0
    for param in source_sym.Parameters:
        try:
            if param.IsReadOnly:
                continue
            dest_param = dest_sym.LookupParameter(param.Definition.Name)
            if dest_param is None or dest_param.IsReadOnly:
                skipped += 1
                continue
            if param.StorageType == StorageType.String:
                val = param.AsString()
                if val is not None:
                    dest_param.Set(val)
                    copied += 1
            elif param.StorageType == StorageType.Integer:
                dest_param.Set(param.AsInteger())
                copied += 1
            elif param.StorageType == StorageType.Double:
                dest_param.Set(param.AsDouble())
                copied += 1
            elif param.StorageType == StorageType.ElementId:
                source_ref = source_sym.Document.GetElement(param.AsElementId())
                dest_ref = find_matching_referenced_element(source_ref, dest_sym.Document)
                if dest_ref is None:
                    skipped += 1
                    continue
                dest_param.Set(dest_ref.Id)
                copied += 1
        except Exception:
            skipped += 1
    return copied, skipped


# ══════════════════════════════════════════════════════════════
#  TRANSFERÊNCIA — STANDARDS
# ══════════════════════════════════════════════════════════════

def transfer_single_element(source_doc, el_id, category, handler):
    try:
        source_el = source_doc.GetElement(el_id)
        if not source_el:
            return "error"

        el_name = _safe_name(source_el)

        if el_name and check_name_exists(el_name, category):
            if category == "Filtros de Fase":
                choice = "overwrite"
            else:
                choice = handler.ask_user(el_name)

            if choice == "skip":
                return "skipped"
            elif choice == "overwrite":
                try:
                    existing = _find_in_active_doc(el_name, category)
                    if existing:
                        doc.Delete(existing.Id)
                except Exception as e:
                    dbg.warn(u"Falha ao deletar '{}': {}".format(el_name, e))

        ids_list = List[ElementId]()
        ids_list.Add(el_id)
        opts = CopyPasteOptions()
        ElementTransformUtils.CopyElements(
            source_doc, ids_list, doc, Transform.Identity, opts
        )
        return "success"
    except Exception as ex:
        dbg.error("Erro ao transferir: {}".format(ex))
        return "error"


# ══════════════════════════════════════════════════════════════
#  JANELA WPF
# ══════════════════════════════════════════════════════════════

class TransferWindow(forms.WPFWindow):

    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.source_doc       = None
        self.source_docs_list = []
        # Standards
        self.standard_categories  = []
        self.all_item_checkboxes  = []
        self._expander_map        = []
        # Family types
        self._family_type_items       = []
        self._all_family_checkboxes   = []
        self._last_family_clicked_idx = None

        self._bind_events()
        self._init_ui()
        if self.source_docs_list:
            _, self.source_doc = self.source_docs_list[0]
            self.load_standards()
            self.load_family_types()

    def _init_ui(self):
        open_docs = get_open_documents()
        for d in open_docs:
            try:
                title = "[Aberto] " + d.Title
                self.source_docs_list.append((title, d))
            except Exception:
                pass

        linked_docs = get_linked_documents()
        for d in linked_docs:
            try:
                if any(d.Equals(existing_d) for _, existing_d in self.source_docs_list):
                    continue
                title = "[Link] " + d.Title
                self.source_docs_list.append((title, d))
            except Exception:
                pass

        if not self.source_docs_list:
            self.lbl_Info.Text = u"Nenhum documento fonte disponível. Abra outro documento ou adicione um link."
            self.btn_Transfer.IsEnabled = False
            return

        for title, _ in self.source_docs_list:
            self.cb_SourceDoc.Items.Add(title)
        self.cb_SourceDoc.SelectedIndex = 0

    def _bind_events(self):
        self.btn_Transfer.Click      += self.execute_transfer
        self.btn_Cancel.Click        += lambda s, a: self.Close()
        self.btn_SelectAll.Click     += self.select_all
        self.btn_SelectNone.Click    += self.select_none
        self.btn_FamSelectAll.Click  += self.select_all_families
        self.btn_FamSelectNone.Click += self.select_none_families
        self.cb_SourceDoc.SelectionChanged += self.on_source_changed
        self.txt_Search.TextChanged  += self.on_search_changed
        self.txt_Search.GotFocus     += self._search_got_focus
        self.txt_Search.LostFocus    += self._search_lost_focus

    def on_source_changed(self, sender, args):
        idx = self.cb_SourceDoc.SelectedIndex
        if idx < 0 or idx >= len(self.source_docs_list):
            return
        _, self.source_doc = self.source_docs_list[idx]
        self.load_standards()
        self.load_family_types()

    # ── Pesquisa ──

    def _search_got_focus(self, sender, args):
        self.txt_SearchHint.Visibility = Visibility.Collapsed

    def _search_lost_focus(self, sender, args):
        if not self.txt_Search.Text:
            self.txt_SearchHint.Visibility = Visibility.Visible

    def on_search_changed(self, sender, args):
        query = self.txt_Search.Text.strip().lower()
        for expander, cat in self._expander_map:
            visible_count = 0
            items_panel = expander.Content
            if items_panel:
                for i in range(items_panel.Children.Count):
                    child = items_panel.Children[i]
                    if hasattr(child, 'Content'):
                        name = str(child.Content).lower()
                        if not query or query in name:
                            child.Visibility = Visibility.Visible
                            visible_count += 1
                        else:
                            child.Visibility = Visibility.Collapsed
            if not query:
                expander.Visibility = Visibility.Visible
            elif visible_count > 0:
                expander.Visibility = Visibility.Visible
                expander.IsExpanded = True
            else:
                expander.Visibility = Visibility.Collapsed

    # ── Carregar Standards ──

    def load_standards(self):
        self.standard_categories = []
        self.all_item_checkboxes = []
        self._expander_map       = []
        self.sp_Standards.Children.Clear()

        if not self.source_doc:
            return

        collectors = [
            ("Filtros de Vista",           collect_view_filters),
            ("Modelos de Vista",           collect_view_templates),
            ("Filtros de Fase",            collect_phase_filters),
            (u"Padrões de Preenchimento",  collect_fill_patterns),
            (u"Padrões de Linha",          collect_line_patterns),
            ("Modelos de Tabela de Carga", collect_panel_schedules),
            (u"Classificações de Carga",   collect_load_class),
            ("Fatores de Demanda",         collect_demand_factors),
            (u"Sistemas de Distribuição",  collect_distribution_sys),
            (u"Tipos de Tensão",           collect_voltage_types),
            (u"Tipos de Fiação",           collect_wire_types),
            ("Tipos de Eletroduto",        collect_conduit_types),
            ("Tipos de Eletrocalha",       collect_cable_tray_types),
        ]

        total_items = 0
        for cat_name, collector_fn in collectors:
            items = collector_fn(self.source_doc)
            if not items:
                continue

            cat = StandardCategory(u"{} ({})".format(cat_name, len(items)), items)
            self.standard_categories.append(cat)
            total_items += len(items)

            expander     = Expander()
            header_panel = StackPanel()
            header_panel.Orientation = System.Windows.Controls.Orientation.Horizontal

            header_cb = CheckBox()
            header_cb.IsChecked = True
            header_cb.Margin = Thickness(0, 0, 8, 0)
            header_cb.VerticalAlignment = System.Windows.VerticalAlignment.Center
            cat.header_checkbox = header_cb

            header_text = TextBlock()
            header_text.Text       = cat.name
            header_text.FontWeight = FontWeights.SemiBold
            header_text.FontSize   = 12
            header_text.VerticalAlignment = System.Windows.VerticalAlignment.Center

            header_panel.Children.Add(header_cb)
            header_panel.Children.Add(header_text)
            expander.Header     = header_panel
            expander.IsExpanded = False
            expander.Margin     = Thickness(0, 4, 0, 4)

            items_panel = StackPanel()
            items_panel.Margin = Thickness(20, 4, 0, 4)

            for item in items:
                item_cb = CheckBox()
                item_cb.Content   = item.name
                item_cb.IsChecked = True
                item_cb.Margin    = Thickness(0, 2, 0, 2)
                item_cb.FontSize  = 11
                item.checkbox = item_cb
                self.all_item_checkboxes.append(item_cb)
                items_panel.Children.Add(item_cb)

            expander.Content = items_panel

            def make_toggle(items_list):
                def toggle(s, a):
                    for it in items_list:
                        if it.checkbox:
                            it.checkbox.IsChecked = s.IsChecked
                return toggle

            header_cb.Checked   += make_toggle(items)
            header_cb.Unchecked += make_toggle(items)

            self._expander_map.append((expander, cat))
            self.sp_Standards.Children.Add(expander)

        if total_items == 0:
            msg = TextBlock()
            msg.Text      = u"Nenhum standard encontrado neste documento."
            msg.FontStyle = System.Windows.FontStyles.Italic
            msg.FontSize  = 11
            self.sp_Standards.Children.Add(msg)

        self.lbl_Info.Text = u"{} standards em {} categorias.".format(
            total_items, len(self.standard_categories))

    # ── Carregar Tipos de Família ──

    def load_family_types(self):
        self._family_type_items       = []
        self._all_family_checkboxes   = []
        self._last_family_clicked_idx = None
        self.sp_FamilyTypes.Children.Clear()

        if not self.source_doc:
            return

        cats_with_items = collect_family_types(self.source_doc)
        total = 0

        for cat_name, items in cats_with_items:
            expander     = Expander()
            header_panel = StackPanel()
            header_panel.Orientation = System.Windows.Controls.Orientation.Horizontal

            header_cb = CheckBox()
            header_cb.IsChecked = True
            header_cb.Margin = Thickness(0, 0, 8, 0)
            header_cb.VerticalAlignment = System.Windows.VerticalAlignment.Center

            header_text = TextBlock()
            header_text.Text       = u"{} ({})".format(cat_name, len(items))
            header_text.FontWeight = FontWeights.SemiBold
            header_text.FontSize   = 12
            header_text.VerticalAlignment = System.Windows.VerticalAlignment.Center

            header_panel.Children.Add(header_cb)
            header_panel.Children.Add(header_text)
            expander.Header     = header_panel
            expander.IsExpanded = False
            expander.Margin     = Thickness(0, 4, 0, 4)

            items_panel = StackPanel()
            items_panel.Margin = Thickness(20, 4, 0, 4)

            items_in_cat = []
            for item in items:
                cb = CheckBox()
                if item.fam_name:
                    cb.Content = u"{} — {}".format(item.fam_name, item.type_name)
                else:
                    cb.Content = item.type_name
                cb.IsChecked = True
                cb.Margin    = Thickness(0, 2, 0, 2)
                cb.FontSize  = 11

                idx = len(self._all_family_checkboxes)
                item.checkbox = cb
                self._all_family_checkboxes.append(cb)
                self._family_type_items.append(item)
                items_in_cat.append(item)
                items_panel.Children.Add(cb)
                cb.Click += self._make_fam_cb_click(idx)
                total += 1

            expander.Content = items_panel

            def make_fam_toggle(items_list):
                def toggle(s, a):
                    for it in items_list:
                        if it.checkbox:
                            it.checkbox.IsChecked = s.IsChecked
                return toggle

            header_cb.Checked   += make_fam_toggle(items_in_cat)
            header_cb.Unchecked += make_fam_toggle(items_in_cat)

            self.sp_FamilyTypes.Children.Add(expander)

        if total == 0:
            msg = TextBlock()
            msg.Text      = u"Nenhum tipo usado encontrado neste documento."
            msg.FontStyle = System.Windows.FontStyles.Italic
            msg.FontSize  = 11
            self.sp_FamilyTypes.Children.Add(msg)

    def _make_fam_cb_click(self, idx):
        def on_click(sender, args):
            if Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift):
                if self._last_family_clicked_idx is not None:
                    start = min(self._last_family_clicked_idx, idx)
                    end   = max(self._last_family_clicked_idx, idx)
                    state = sender.IsChecked
                    for i in range(start, end + 1):
                        self._all_family_checkboxes[i].IsChecked = state
            self._last_family_clicked_idx = idx
        return on_click

    # ── Seleção ──

    def select_all(self, sender, args):
        for cb in self.all_item_checkboxes:
            cb.IsChecked = True
        for cat in self.standard_categories:
            if cat.header_checkbox:
                cat.header_checkbox.IsChecked = True

    def select_none(self, sender, args):
        for cb in self.all_item_checkboxes:
            cb.IsChecked = False
        for cat in self.standard_categories:
            if cat.header_checkbox:
                cat.header_checkbox.IsChecked = False

    def select_all_families(self, sender, args):
        for cb in self._all_family_checkboxes:
            cb.IsChecked = True

    def select_none_families(self, sender, args):
        for cb in self._all_family_checkboxes:
            cb.IsChecked = False

    # ── Transferir ──

    def execute_transfer(self, sender, args):
        if self.tab_main.SelectedIndex == 1:
            self._execute_family_transfer()
        else:
            self._execute_standards_transfer()

    def _execute_standards_transfer(self):
        if not self.source_doc:
            forms.alert(u"Selecione um documento fonte.", title="Transfer Settings")
            return

        selected_items = []
        for cat in self.standard_categories:
            for item in cat.items:
                if item.checkbox and item.checkbox.IsChecked:
                    selected_items.append(item)

        if not selected_items:
            forms.alert(u"Nenhum item selecionado.", title="Transfer Settings")
            return

        self.Close()

        handler       = DuplicateNameHandler()
        success_count = 0
        skip_count    = 0
        error_count   = 0

        preprocessor = make_warning_swallower()
        with revit.Transaction("Transfer Settings", doc):
            for item in selected_items:
                result = transfer_single_element(
                    self.source_doc, item.element_id, item.category, handler
                )
                if result == "success":
                    success_count += 1
                elif result == "skipped":
                    skip_count += 1
                else:
                    error_count += 1

        msg = u"Transferência concluída!\n\nTransferidos: {}\n".format(success_count)
        if skip_count > 0:
            msg += u"Pulados: {}\n".format(skip_count)
        if error_count > 0:
            msg += u"Erros: {}\n".format(error_count)
        forms.alert(msg, title="Transfer Settings")

    def _execute_family_transfer(self):
        if not self.source_doc:
            forms.alert(u"Selecione um documento fonte.", title="Transfer Settings")
            return

        selected_items = [it for it in self._family_type_items
                          if it.checkbox and it.checkbox.IsChecked]

        if not selected_items:
            forms.alert(u"Nenhum tipo de família selecionado.", title="Transfer Settings")
            return

        self.Close()

        total_copied  = 0
        total_skipped = 0
        no_match      = 0

        with revit.Transaction("Transfer Family Type Parameters", doc):
            for item in selected_items:
                source_sym = self.source_doc.GetElement(item.symbol_id)
                if not source_sym:
                    no_match += 1
                    continue
                dest_sym = find_matching_used_symbol(doc, item)
                if dest_sym is None:
                    no_match += 1
                    dbg.warn(u"Tipo nao encontrado: {} / {}".format(item.fam_name, item.type_name))
                    continue
                copied, skipped = copy_type_params(source_sym, dest_sym)
                total_copied  += copied
                total_skipped += skipped

        msg = u"Transferência concluída!\n\nParâmetros transferidos: {}\n".format(total_copied)
        if total_skipped > 0:
            msg += u"Sem equivalente no destino: {}\n".format(total_skipped)
        if no_match > 0:
            msg += u"Tipos equivalentes não usados no projeto ativo: {}\n".format(no_match)
        forms.alert(msg, title="Transfer Settings")


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if doc:
        TransferWindow("TransferWindow.xaml").show(modal=True)
