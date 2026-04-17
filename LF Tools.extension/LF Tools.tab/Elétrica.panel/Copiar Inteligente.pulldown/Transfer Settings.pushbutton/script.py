# -*- coding: utf-8 -*-
"""
Transfer Settings - Transferência Cirúrgica de Configurações
=============================================================
LF Tools - pyRevit Extension

Permite selecionar INDIVIDUALMENTE quais standards (Filtros de Vista,
View Templates, Fill Patterns, Line Patterns, Line Styles, Parâmetros)
copiar de um documento fonte para o documento ativo.

Diferente do Transfer Project Standards nativo do Revit:
  - Granularidade por ELEMENTO (não por categoria inteira)
  - Pergunta ao usuário em caso de conflito de nomes
  - Suporta documentos abertos E modelos linkados
"""
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Collections.Generic import List
import System
from System.Windows.Controls import CheckBox, TextBlock, StackPanel, Expander
from System.Windows import Thickness, FontWeights
from System.Windows.Media import SolidColorBrush
from System.Windows.Media import Color as WpfColor

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId,
    Transaction, ParameterFilterElement,
    FillPatternElement, View,
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

doc = revit.doc
uiapp = revit.HOST_APP.uiapp
app = revit.HOST_APP.app
logger = script.get_logger()


# ====== Classes de Apoio ======
class TransferableItem(object):
    """Um standard que pode ser transferido."""
    def __init__(self, name, el_id, category):
        self.name = name
        self.element_id = el_id
        self.category = category
        self.checkbox = None  # Referência ao CheckBox na UI


class StandardCategory(object):
    """Uma categoria de standards (agrupador)."""
    def __init__(self, name, items=None):
        self.name = name
        self.items = items or []
        self.header_checkbox = None


# ====== Handler de Duplicatas ======
class DuplicateNameHandler(object):
    """Handler que pergunta ao usuário sobre duplicatas."""
    
    def __init__(self):
        self.user_choice = None  # "overwrite", "rename", "skip"
    
    def ask_user(self, name):
        """Exibe diálogo perguntando ao usuário."""
        result = forms.CommandSwitchWindow.show(
            ["Sobrescrever", "Renomear (sufixo)", "Pular"],
            message="O elemento '{}' já existe no destino.\nO que deseja fazer?".format(name)
        )
        if result == "Sobrescrever":
            return "overwrite"
        elif result and "Renomear" in result:
            return "rename"
        else:
            return "skip"


# ====== Funções de Coleta ======
def get_open_documents():
    """Retorna documentos abertos na sessão (exceto o ativo)."""
    docs = []
    for d in app.Documents:
        try:
            if not d.Equals(doc) and not d.IsFamilyDocument:
                docs.append(d)
        except:
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
        except:
            pass
    return docs


def collect_view_filters(source_doc):
    """Coleta ParameterFilterElement."""
    items = []
    try:
        col = FilteredElementCollector(source_doc).OfClass(ParameterFilterElement)
        for f in col:
            try:
                items.append(TransferableItem(f.Name, f.Id, "Filtros de Vista"))
            except:
                pass
    except:
        pass
    return sorted(items, key=lambda x: x.name)


def collect_view_templates(source_doc):
    """Coleta View Templates."""
    items = []
    try:
        col = FilteredElementCollector(source_doc).OfClass(View)
        for v in col:
            try:
                if v.IsTemplate:
                    items.append(TransferableItem(v.Name, v.Id, "View Templates"))
            except:
                pass
    except:
        pass
    return sorted(items, key=lambda x: x.name)


def collect_fill_patterns(source_doc):
    """Coleta FillPatternElement."""
    items = []
    try:
        col = FilteredElementCollector(source_doc).OfClass(FillPatternElement)
        for p in col:
            try:
                items.append(TransferableItem(p.Name, p.Id, "Fill Patterns"))
            except:
                pass
    except:
        pass
    return sorted(items, key=lambda x: x.name)


def collect_line_patterns(source_doc):
    """Coleta LinePatternElement."""
    items = []
    try:
        col = FilteredElementCollector(source_doc).OfClass(LinePatternElement)
        for lp in col:
            try:
                items.append(TransferableItem(lp.Name, lp.Id, "Line Patterns"))
            except:
                pass
    except:
        pass
    return sorted(items, key=lambda x: x.name)


def collect_electrical_elements(source_doc, doc_class, category_name):
    """Coleta genérica para elementos elétricos."""
    items = []
    try:
        col = FilteredElementCollector(source_doc).OfClass(doc_class)
        for e in col:
            try:
                if e.Name:
                    items.append(TransferableItem(e.Name, e.Id, category_name))
            except:
                pass
    except:
        pass
    return sorted(items, key=lambda x: x.name)


def collect_panel_schedules(source_doc): return collect_electrical_elements(source_doc, PanelScheduleTemplate, "Modelos de Tabela de Carga")
def collect_load_class(source_doc): return collect_electrical_elements(source_doc, ElectricalLoadClassification, "Classificações de Carga")
def collect_demand_factors(source_doc): return collect_electrical_elements(source_doc, ElectricalDemandFactorDefinition, "Fatores de Demanda")
def collect_wire_types(source_doc): return collect_electrical_elements(source_doc, WireType, "Tipos de Fiação")
def collect_distribution_sys(source_doc): return collect_electrical_elements(source_doc, DistributionSysType, "Sistemas de Distribuição")
def collect_voltage_types(source_doc): return collect_electrical_elements(source_doc, VoltageType, "Tipos de Tensão")
def collect_conduit_types(source_doc): return collect_electrical_elements(source_doc, ConduitType, "Tipos de Eletroduto")
def collect_cable_tray_types(source_doc): return collect_electrical_elements(source_doc, CableTrayType, "Tipos de Eletrocalha")


def check_name_exists(name, category):
    """Verifica se já existe um elemento com o mesmo nome no doc ativo."""
    try:
        col = None
        if category == "Filtros de Vista":
            col = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
        elif category == "View Templates":
            col = FilteredElementCollector(doc).OfClass(View)
            for v in col:
                if v.IsTemplate and v.Name == name:
                    return True
            return False
        elif category == "Fill Patterns":
            col = FilteredElementCollector(doc).OfClass(FillPatternElement)
        elif category == "Line Patterns":
            col = FilteredElementCollector(doc).OfClass(LinePatternElement)
        elif category == "Modelos de Tabela de Carga": col = FilteredElementCollector(doc).OfClass(PanelScheduleTemplate)
        elif category == "Classificações de Carga": col = FilteredElementCollector(doc).OfClass(ElectricalLoadClassification)
        elif category == "Fatores de Demanda": col = FilteredElementCollector(doc).OfClass(ElectricalDemandFactorDefinition)
        elif category == "Tipos de Fiação": col = FilteredElementCollector(doc).OfClass(WireType)
        elif category == "Sistemas de Distribuição": col = FilteredElementCollector(doc).OfClass(DistributionSysType)
        elif category == "Tipos de Tensão": col = FilteredElementCollector(doc).OfClass(VoltageType)
        elif category == "Tipos de Eletroduto": col = FilteredElementCollector(doc).OfClass(ConduitType)
        elif category == "Tipos de Eletrocalha": col = FilteredElementCollector(doc).OfClass(CableTrayType)
        else:
            return False
        
        if col is not None:
            for el in col:
                try:
                    if el.Name == name:
                        return True
                except:
                    pass
    except:
        pass
    return False


def transfer_single_element(source_doc, el_id, handler):
    """
    Transfere um único elemento, tratando conflitos.
    
    Returns:
        "success", "skipped", ou "error"
    """
    try:
        source_el = source_doc.GetElement(el_id)
        if not source_el:
            return "error"
        
        el_name = ""
        try:
            el_name = source_el.Name
        except:
            pass
        
        # Verifica se já existe
        cat = ""
        if isinstance(source_el, ParameterFilterElement):
            cat = "Filtros de Vista"
        elif isinstance(source_el, View) and source_el.IsTemplate:
            cat = "View Templates"
        elif isinstance(source_el, FillPatternElement):
            cat = "Fill Patterns"
        elif isinstance(source_el, LinePatternElement):
            cat = "Line Patterns"
        elif isinstance(source_el, PanelScheduleTemplate): cat = "Modelos de Tabela de Carga"
        elif isinstance(source_el, ElectricalLoadClassification): cat = "Classificações de Carga"
        elif isinstance(source_el, ElectricalDemandFactorDefinition): cat = "Fatores de Demanda"
        elif isinstance(source_el, WireType): cat = "Tipos de Fiação"
        elif isinstance(source_el, DistributionSysType): cat = "Sistemas de Distribuição"
        elif isinstance(source_el, VoltageType): cat = "Tipos de Tensão"
        elif isinstance(source_el, ConduitType): cat = "Tipos de Eletroduto"
        elif isinstance(source_el, CableTrayType): cat = "Tipos de Eletrocalha"
        
        if el_name and check_name_exists(el_name, cat):
            choice = handler.ask_user(el_name)
            if choice == "skip":
                return "skipped"
            elif choice == "overwrite":
                # Deleta o existente antes de copiar
                try:
                    existing = find_element_by_name(el_name, cat)
                    if existing:
                        doc.Delete(existing.Id)
                except:
                    pass  # Se não conseguir deletar, tenta copiar mesmo assim
        
        # Copia o elemento
        ids_list = List[ElementId]()
        ids_list.Add(el_id)
        opts = CopyPasteOptions()
        
        ElementTransformUtils.CopyElements(
            source_doc, ids_list, doc, Transform.Identity, opts
        )
        return "success"
    except Exception as ex:
        print("Erro ao transferir: " + str(ex))
        return "error"


def find_element_by_name(name, category):
    """Encontra um elemento por nome no documento ativo."""
    try:
        col = None
        if category == "Filtros de Vista":
            col = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
        elif category == "View Templates":
            col = FilteredElementCollector(doc).OfClass(View)
            for v in col:
                if v.IsTemplate and v.Name == name:
                    return v
            return None
        elif category == "Fill Patterns":
            col = FilteredElementCollector(doc).OfClass(FillPatternElement)
        elif category == "Line Patterns":
            col = FilteredElementCollector(doc).OfClass(LinePatternElement)
        elif category == "Modelos de Tabela de Carga": col = FilteredElementCollector(doc).OfClass(PanelScheduleTemplate)
        elif category == "Classificações de Carga": col = FilteredElementCollector(doc).OfClass(ElectricalLoadClassification)
        elif category == "Fatores de Demanda": col = FilteredElementCollector(doc).OfClass(ElectricalDemandFactorDefinition)
        elif category == "Tipos de Fiação": col = FilteredElementCollector(doc).OfClass(WireType)
        elif category == "Sistemas de Distribuição": col = FilteredElementCollector(doc).OfClass(DistributionSysType)
        elif category == "Tipos de Tensão": col = FilteredElementCollector(doc).OfClass(VoltageType)
        elif category == "Tipos de Eletroduto": col = FilteredElementCollector(doc).OfClass(ConduitType)
        elif category == "Tipos de Eletrocalha": col = FilteredElementCollector(doc).OfClass(CableTrayType)
        else:
            return None
        
        if col is not None:
            for el in col:
                try:
                    if el.Name == name:
                        return el
                except:
                    pass
    except:
        pass
    return None


# ====== Janela WPF ======
class TransferWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.source_doc = None
        self.source_docs_list = []  # Lista de (title, doc) 
        self.standard_categories = []
        self.all_item_checkboxes = []
        self._bind_events()
        self._init_ui()
        # Carrega os standards do primeiro documento automaticamente
        if self.source_docs_list:
            _, self.source_doc = self.source_docs_list[0]
            self.load_standards()
    
    def _init_ui(self):
        """Popula ComboBox de documentos fontes."""
        # Documentos abertos
        open_docs = get_open_documents()
        for d in open_docs:
            self.source_docs_list.append(("[Aberto] " + d.Title, d))
        
        # Documentos linkados
        linked_docs = get_linked_documents()
        for d in linked_docs:
            title = "[Link] " + d.Title
            # Evita duplicatas
            if not any(t == title for t, _ in self.source_docs_list):
                self.source_docs_list.append((title, d))
        
        if not self.source_docs_list:
            self.lbl_Info.Text = "⚠ Nenhum documento fonte disponível. Abra outro documento ou adicione um link."
            self.btn_Transfer.IsEnabled = False
            return
        
        doc_names = [t for t, _ in self.source_docs_list]
        self.cb_SourceDoc.ItemsSource = doc_names
        self.cb_SourceDoc.SelectedIndex = 0
    
    def _bind_events(self):
        self.btn_Transfer.Click += self.execute_transfer
        self.btn_Cancel.Click += lambda s, a: self.Close()
        self.btn_SelectAll.Click += self.select_all
        self.btn_SelectNone.Click += self.select_none
        self.cb_SourceDoc.SelectionChanged += self.on_source_changed
        self.txt_Search.TextChanged += self.on_search_changed
        self.txt_Search.GotFocus += self._search_got_focus
        self.txt_Search.LostFocus += self._search_lost_focus
    
    def on_source_changed(self, sender, args):
        """Quando o documento fonte muda, recarrega os standards."""
        idx = self.cb_SourceDoc.SelectedIndex
        if idx < 0 or idx >= len(self.source_docs_list):
            return
        
        _, self.source_doc = self.source_docs_list[idx]
        self.load_standards()
    
    def _search_got_focus(self, sender, args):
        self.txt_SearchHint.Visibility = System.Windows.Visibility.Collapsed
    
    def _search_lost_focus(self, sender, args):
        if not self.txt_Search.Text:
            self.txt_SearchHint.Visibility = System.Windows.Visibility.Visible
    
    def on_search_changed(self, sender, args):
        """Filtra standards pelo texto de pesquisa."""
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
                            child.Visibility = System.Windows.Visibility.Visible
                            visible_count += 1
                        else:
                            child.Visibility = System.Windows.Visibility.Collapsed
            
            if not query:
                expander.Visibility = System.Windows.Visibility.Visible
            elif visible_count > 0:
                expander.Visibility = System.Windows.Visibility.Visible
                expander.IsExpanded = True
            else:
                expander.Visibility = System.Windows.Visibility.Collapsed
    
    def load_standards(self):
        """Carrega standards do documento fonte na UI."""
        self.standard_categories = []
        self.all_item_checkboxes = []
        self._expander_map = []  # [(expander, cat)]
        self.sp_Standards.Children.Clear()
        
        if not self.source_doc:
            return
        
        # Coleta cada tipo
        collectors = [
            ("Filtros de Vista", collect_view_filters),
            ("Modelos de Vista", collect_view_templates),
            ("Padrões de Preenchimento", collect_fill_patterns),
            ("Padrões de Linha", collect_line_patterns),
            ("Modelos de Tabela de Carga", collect_panel_schedules),
            ("Classificações de Carga", collect_load_class),
            ("Fatores de Demanda", collect_demand_factors),
            ("Sistemas de Distribuição", collect_distribution_sys),
            ("Tipos de Tensão", collect_voltage_types),
            ("Tipos de Fiação", collect_wire_types),
            ("Tipos de Eletroduto", collect_conduit_types),
            ("Tipos de Eletrocalha", collect_cable_tray_types),
        ]
        
        total_items = 0
        
        for cat_name, collector_fn in collectors:
            items = collector_fn(self.source_doc)
            if not items:
                continue
            
            cat = StandardCategory("{} ({})".format(cat_name, len(items)), items)
            self.standard_categories.append(cat)
            total_items += len(items)
            
            # Cria Expander com checkboxes
            expander = Expander()
            
            # Header com checkbox
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
            try:
                header_text.Foreground = SolidColorBrush(WpfColor.FromArgb(255, 78, 201, 176))  # AccentTeal
            except:
                pass
            header_text.VerticalAlignment = System.Windows.VerticalAlignment.Center
            
            header_panel.Children.Add(header_cb)
            header_panel.Children.Add(header_text)
            
            expander.Header = header_panel
            expander.IsExpanded = False
            expander.Margin = Thickness(0, 4, 0, 4)
            try:
                expander.Foreground = SolidColorBrush(WpfColor.FromArgb(255, 241, 241, 241))
            except:
                pass
            
            # Conteúdo: lista de checkboxes individuais
            items_panel = StackPanel()
            items_panel.Margin = Thickness(20, 4, 0, 4)
            
            for item in items:
                item_cb = CheckBox()
                item_cb.Content = item.name
                item_cb.IsChecked = True
                item_cb.Margin = Thickness(0, 2, 0, 2)
                item_cb.FontSize = 11
                try:
                    item_cb.Foreground = SolidColorBrush(WpfColor.FromArgb(255, 204, 204, 204))
                except:
                    pass
                
                item.checkbox = item_cb
                self.all_item_checkboxes.append(item_cb)
                items_panel.Children.Add(item_cb)
            
            expander.Content = items_panel
            
            # Conecta o header checkbox para marcar/desmarcar todos os filhos
            def make_toggle(items_list):
                def toggle(s, a):
                    checked = s.IsChecked
                    for it in items_list:
                        if it.checkbox:
                            it.checkbox.IsChecked = checked
                return toggle
            
            header_cb.Checked += make_toggle(items)
            header_cb.Unchecked += make_toggle(items)
            
            self._expander_map.append((expander, cat))
            self.sp_Standards.Children.Add(expander)
        
        if total_items == 0:
            msg = TextBlock()
            msg.Text = "Nenhum standard encontrado neste documento."
            msg.FontStyle = System.Windows.FontStyles.Italic
            msg.FontSize = 11
            try:
                msg.Foreground = SolidColorBrush(WpfColor.FromArgb(255, 204, 204, 204))
            except:
                pass
            self.sp_Standards.Children.Add(msg)
        
        self.lbl_Info.Text = "{} standards encontrados em {} categorias.".format(
            total_items, len(self.standard_categories)
        )
    
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
    
    def execute_transfer(self, sender, args):
        """Executa a transferência dos itens selecionados."""
        if not self.source_doc:
            forms.alert("Selecione um documento fonte.")
            return
        
        # Coleta itens marcados
        selected_items = []
        for cat in self.standard_categories:
            for item in cat.items:
                if item.checkbox and item.checkbox.IsChecked:
                    selected_items.append(item)
        
        if not selected_items:
            forms.alert("Nenhum item selecionado para transferência.")
            return
        
        self.Close()
        
        handler = DuplicateNameHandler()
        success_count = 0
        skip_count = 0
        error_count = 0
        
        with revit.Transaction("Transfer Settings", doc):
            for item in selected_items:
                result = transfer_single_element(self.source_doc, item.element_id, handler)
                if result == "success":
                    success_count += 1
                elif result == "skipped":
                    skip_count += 1
                else:
                    error_count += 1
        
        # Relatório
        msg = "Transferência concluída!\n\n"
        msg += "✅ Transferidos: {}\n".format(success_count)
        if skip_count > 0:
            msg += "⏭ Pulados: {}\n".format(skip_count)
        if error_count > 0:
            msg += "❌ Erros: {}\n".format(error_count)
        
        forms.alert(msg, title="Transfer Settings")


# ====== Entry Point ======
if __name__ == "__main__":
    if doc:
        TransferWindow("TransferWindow.xaml").show(modal=True)
