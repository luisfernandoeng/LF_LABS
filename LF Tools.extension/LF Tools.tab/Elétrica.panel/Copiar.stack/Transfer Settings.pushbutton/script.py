# -*- coding: utf-8 -*-
"""
Transfer Settings — Transferência Cirúrgica de Configurações
=============================================================
LF Tools — pyRevit Extension

Permite selecionar INDIVIDUALMENTE quais standards (Filtros de Vista,
View Templates, Fill Patterns, Line Patterns, Phase Filters,
elementos elétricos) copiar de um documento fonte para o documento ativo.

Diferente do Transfer Project Standards nativo do Revit:
  - Granularidade por ELEMENTO (não por categoria inteira)
  - Pergunta ao usuário em caso de conflito de nomes
  - Suporta documentos abertos E modelos linkados
"""
__title__ = "Transfer\nSettings"
__author__ = "Luís Fernando"

# ╔══════════════════════════════════════════════════════════╗
# ║  DEBUG_MODE                                              ║
# ╚══════════════════════════════════════════════════════════╝
DEBUG_MODE = False

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows.Controls import CheckBox, TextBlock, StackPanel, Expander
from System.Windows import Thickness, FontWeights, Visibility
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId,
    Transaction, ParameterFilterElement,
    FillPatternElement, View, PhaseFilter,
    ElementTransformUtils, CopyPasteOptions, Transform,
    LinePatternElement,
    RevitLinkInstance, RevitLinkType
)
from Autodesk.Revit.DB.Electrical import (
    PanelScheduleTemplate, ElectricalLoadClassification,
    ElectricalDemandFactorDefinition, DistributionSysType, VoltageType,
    WireType, ConduitType, CableTrayType
)
from pyrevit import revit, forms, script
from lf_utils import DebugLogger, make_warning_swallower

# ══════════════════════════════════════════════════════════════
#  INIT
# ══════════════════════════════════════════════════════════════

dbg    = DebugLogger(DEBUG_MODE)
doc    = revit.doc
uiapp  = revit.HOST_APP.uiapp
app    = revit.HOST_APP.app
output = script.get_output()


# ══════════════════════════════════════════════════════════════
#  CLASSES DE APOIO
# ══════════════════════════════════════════════════════════════

class TransferableItem(object):
    """Um standard que pode ser transferido."""
    def __init__(self, name, el_id, category):
        self.name = name
        self.element_id = el_id
        self.category = category
        self.checkbox = None


class StandardCategory(object):
    """Agrupador de standards."""
    def __init__(self, name, items=None):
        self.name = name
        self.items = items or []
        self.header_checkbox = None


class DuplicateNameHandler(object):
    """Pergunta ao usuário sobre duplicatas."""
    def ask_user(self, name):
        result = forms.CommandSwitchWindow.show(
            ["Sobrescrever", "Renomear (sufixo)", "Pular"],
            message="O elemento '{}' já existe no destino.\nO que deseja fazer?".format(name)
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
    """Retorna documentos abertos na sessão (exceto o ativo)."""
    docs = []
    for d in app.Documents:
        try:
            if not d.Equals(doc) and not d.IsFamilyDocument:
                docs.append(d)
        except Exception:
            pass
    return docs


def get_linked_documents():
    """Retorna documentos linkados no projeto ativo."""
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
#  COLETORES
# ══════════════════════════════════════════════════════════════

def _safe_name(el):
    """Lê .Name de forma segura."""
    try:
        return el.Name
    except Exception:
        return ""


def _collect_by_class(source_doc, cls, cat_name, filter_fn=None):
    """Coleta genérica por classe, com filtro opcional."""
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
    return _collect_by_class(source_doc, FillPatternElement, "Padrões de Preenchimento")

def collect_line_patterns(source_doc):
    return _collect_by_class(source_doc, LinePatternElement, "Padrões de Linha")

def collect_phase_filters(source_doc):
    return _collect_by_class(source_doc, PhaseFilter, "Filtros de Fase")

def collect_panel_schedules(source_doc):
    return _collect_by_class(source_doc, PanelScheduleTemplate, "Modelos de Tabela de Carga")

def collect_load_class(source_doc):
    return _collect_by_class(source_doc, ElectricalLoadClassification, "Classificações de Carga")

def collect_demand_factors(source_doc):
    return _collect_by_class(source_doc, ElectricalDemandFactorDefinition, "Fatores de Demanda")

def collect_wire_types(source_doc):
    return _collect_by_class(source_doc, WireType, "Tipos de Fiação")

def collect_distribution_sys(source_doc):
    return _collect_by_class(source_doc, DistributionSysType, "Sistemas de Distribuição")

def collect_voltage_types(source_doc):
    return _collect_by_class(source_doc, VoltageType, "Tipos de Tensão")

def collect_conduit_types(source_doc):
    return _collect_by_class(source_doc, ConduitType, "Tipos de Eletroduto")

def collect_cable_tray_types(source_doc):
    return _collect_by_class(source_doc, CableTrayType, "Tipos de Eletrocalha")


# Mapeamento categoria → classe (para busca no doc ativo)
_CAT_CLASS_MAP = {
    "Filtros de Vista":           ParameterFilterElement,
    "View Templates":             View,
    "Padrões de Preenchimento":   FillPatternElement,
    "Padrões de Linha":           LinePatternElement,
    "Filtros de Fase":            PhaseFilter,
    "Modelos de Tabela de Carga": PanelScheduleTemplate,
    "Classificações de Carga":    ElectricalLoadClassification,
    "Fatores de Demanda":         ElectricalDemandFactorDefinition,
    "Tipos de Fiação":            WireType,
    "Sistemas de Distribuição":   DistributionSysType,
    "Tipos de Tensão":            VoltageType,
    "Tipos de Eletroduto":        ConduitType,
    "Tipos de Eletrocalha":       CableTrayType,
}


def _find_in_active_doc(name, category):
    """Busca elemento por nome no doc ativo. Retorna elemento ou None."""
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
    """Verifica se já existe um elemento com o mesmo nome no doc ativo."""
    return _find_in_active_doc(name, category) is not None


# ══════════════════════════════════════════════════════════════
#  TRANSFERÊNCIA
# ══════════════════════════════════════════════════════════════

def transfer_single_element(source_doc, el_id, category, handler):
    """Transfere um único elemento, tratando conflitos.
    Retorna 'success', 'skipped', ou 'error'.
    """
    try:
        source_el = source_doc.GetElement(el_id)
        if not source_el:
            return "error"

        el_name = _safe_name(source_el)

        if el_name and check_name_exists(el_name, category):
            # Filtros de Fase: sobrescreve automaticamente sem perguntar
            if category == "Filtros de Fase":
                choice = "overwrite"
                dbg.debug("Auto-sobrescrevendo filtro de fase: {}".format(el_name))
            else:
                choice = handler.ask_user(el_name)

            if choice == "skip":
                dbg.debug("Pulado: {}".format(el_name))
                return "skipped"
            elif choice == "overwrite":
                try:
                    existing = _find_in_active_doc(el_name, category)
                    if existing:
                        doc.Delete(existing.Id)
                        dbg.debug("Deletado existente: {}".format(el_name))
                except Exception as e:
                    dbg.warn("Falha ao deletar existente '{}': {}".format(el_name, e))

        ids_list = List[ElementId]()
        ids_list.Add(el_id)
        opts = CopyPasteOptions()

        ElementTransformUtils.CopyElements(
            source_doc, ids_list, doc, Transform.Identity, opts
        )
        dbg.info("Transferido: {}".format(el_name))
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
        self.source_doc = None
        self.source_docs_list = []
        self.standard_categories = []
        self.all_item_checkboxes = []
        self._expander_map = []
        self._bind_events()
        self._init_ui()
        if self.source_docs_list:
            _, self.source_doc = self.source_docs_list[0]
            self.load_standards()

    def _init_ui(self):
        open_docs = get_open_documents()
        for d in open_docs:
            try:
                title = "[Aberto] " + d.Title
            except Exception:
                title = "[Aberto] ?"
            self.source_docs_list.append((title, d))

        linked_docs = get_linked_documents()
        for d in linked_docs:
            try:
                title = "[Link] " + d.Title
            except Exception:
                title = "[Link] ?"
            if not any(t == title for t, _ in self.source_docs_list):
                self.source_docs_list.append((title, d))

        if not self.source_docs_list:
            self.lbl_Info.Text = "Nenhum documento fonte disponível. Abra outro documento ou adicione um link."
            self.btn_Transfer.IsEnabled = False
            return

        for title, _ in self.source_docs_list:
            self.cb_SourceDoc.Items.Add(title)
        self.cb_SourceDoc.SelectedIndex = 0

    def _bind_events(self):
        self.btn_Transfer.Click  += self.execute_transfer
        self.btn_Cancel.Click    += lambda s, a: self.Close()
        self.btn_SelectAll.Click += self.select_all
        self.btn_SelectNone.Click += self.select_none
        self.cb_SourceDoc.SelectionChanged += self.on_source_changed
        self.txt_Search.TextChanged += self.on_search_changed
        self.txt_Search.GotFocus    += self._search_got_focus
        self.txt_Search.LostFocus   += self._search_lost_focus

    def on_source_changed(self, sender, args):
        idx = self.cb_SourceDoc.SelectedIndex
        if idx < 0 or idx >= len(self.source_docs_list):
            return
        _, self.source_doc = self.source_docs_list[idx]
        self.load_standards()

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
        self._expander_map = []
        self.sp_Standards.Children.Clear()

        if not self.source_doc:
            return

        dbg.section("Carregando Standards")

        collectors = [
            ("Filtros de Vista",           collect_view_filters),
            ("Modelos de Vista",           collect_view_templates),
            ("Filtros de Fase",            collect_phase_filters),
            ("Padrões de Preenchimento",   collect_fill_patterns),
            ("Padrões de Linha",           collect_line_patterns),
            ("Modelos de Tabela de Carga", collect_panel_schedules),
            ("Classificações de Carga",    collect_load_class),
            ("Fatores de Demanda",         collect_demand_factors),
            ("Sistemas de Distribuição",   collect_distribution_sys),
            ("Tipos de Tensão",            collect_voltage_types),
            ("Tipos de Fiação",            collect_wire_types),
            ("Tipos de Eletroduto",        collect_conduit_types),
            ("Tipos de Eletrocalha",       collect_cable_tray_types),
        ]

        total_items = 0

        for cat_name, collector_fn in collectors:
            items = collector_fn(self.source_doc)
            if not items:
                continue

            dbg.info("{}: {} itens".format(cat_name, len(items)))
            cat = StandardCategory("{} ({})".format(cat_name, len(items)), items)
            self.standard_categories.append(cat)
            total_items += len(items)

            # Cria Expander com checkboxes
            expander = Expander()

            header_panel = StackPanel()
            header_panel.Orientation = System.Windows.Controls.Orientation.Horizontal

            header_cb = CheckBox()
            header_cb.IsChecked = True
            header_cb.Margin = Thickness(0, 0, 8, 0)
            header_cb.VerticalAlignment = System.Windows.VerticalAlignment.Center
            cat.header_checkbox = header_cb

            header_text = TextBlock()
            header_text.Text = cat.name
            header_text.FontWeight = FontWeights.SemiBold
            header_text.FontSize = 12
            header_text.VerticalAlignment = System.Windows.VerticalAlignment.Center

            header_panel.Children.Add(header_cb)
            header_panel.Children.Add(header_text)

            expander.Header = header_panel
            expander.IsExpanded = False
            expander.Margin = Thickness(0, 4, 0, 4)

            items_panel = StackPanel()
            items_panel.Margin = Thickness(20, 4, 0, 4)

            for item in items:
                item_cb = CheckBox()
                item_cb.Content = item.name
                item_cb.IsChecked = True
                item_cb.Margin = Thickness(0, 2, 0, 2)
                item_cb.FontSize = 11

                item.checkbox = item_cb
                self.all_item_checkboxes.append(item_cb)
                items_panel.Children.Add(item_cb)

            expander.Content = items_panel

            def make_toggle(items_list):
                def toggle(s, a):
                    checked = s.IsChecked
                    for it in items_list:
                        if it.checkbox:
                            it.checkbox.IsChecked = checked
                return toggle

            header_cb.Checked   += make_toggle(items)
            header_cb.Unchecked += make_toggle(items)

            self._expander_map.append((expander, cat))
            self.sp_Standards.Children.Add(expander)

        if total_items == 0:
            msg = TextBlock()
            msg.Text = "Nenhum standard encontrado neste documento."
            msg.FontStyle = System.Windows.FontStyles.Italic
            msg.FontSize = 11
            self.sp_Standards.Children.Add(msg)

        self.lbl_Info.Text = "{} standards em {} categorias.".format(
            total_items, len(self.standard_categories))

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

    # ── Transferir ──

    def execute_transfer(self, sender, args):
        if not self.source_doc:
            forms.alert("Selecione um documento fonte.", title="Transfer Settings")
            return

        selected_items = []
        for cat in self.standard_categories:
            for item in cat.items:
                if item.checkbox and item.checkbox.IsChecked:
                    selected_items.append(item)

        if not selected_items:
            forms.alert("Nenhum item selecionado.", title="Transfer Settings")
            return

        self.Close()

        dbg.section("Transferência")
        dbg.info("{} itens selecionados".format(len(selected_items)))

        handler = DuplicateNameHandler()
        success_count = 0
        skip_count = 0
        error_count = 0

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

        dbg.section("Resultado")
        dbg.info("OK: {} | Pulados: {} | Erros: {}".format(
            success_count, skip_count, error_count))

        msg = "Transferência concluída!\n\n"
        msg += "Transferidos: {}\n".format(success_count)
        if skip_count > 0:
            msg += "Pulados: {}\n".format(skip_count)
        if error_count > 0:
            msg += "Erros: {}\n".format(error_count)

        forms.alert(msg, title="Transfer Settings")


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if doc:
        TransferWindow("TransferWindow.xaml").show(modal=True)
