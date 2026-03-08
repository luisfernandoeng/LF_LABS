# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import os

# Get UI environment
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

# Import System and WPF components
import System
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows import Window, Thickness, FontWeight, FontWeights, Media, Visibility
from System.Windows.Controls import TextBlock, CheckBox, TextBox
from System.Collections.Generic import List
from System.Windows.Forms import Control, Keys

def get_element_parameters(element):
    """Extrai exaustivamente todos os parâmetros e valores do elemento."""
    params_dict = {}
    
    # 1. Propriedades Básicas e Geométricas
    try:
        params_dict['Category'] = element.Category.Name if element.Category else "Sem Categoria"
        
        # Family e Type
        if hasattr(element, 'Symbol'):
            params_dict['Family'] = element.Symbol.FamilyName
            params_dict['Type'] = element.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or element.Name
        else:
            params_dict['Family'] = element.Name
            params_dict['Type'] = element.Name

        # Level
        level_id = element.LevelId
        if level_id != DB.ElementId.InvalidElementId:
            params_dict['Level'] = doc.GetElement(level_id).Name
            
        # Workset
        workset_id = element.WorksetId
        if workset_id and doc.IsWorkshared:
            workset = doc.GetWorksetTable().GetWorkset(workset_id)
            params_dict['Workset'] = workset.Name if workset else None
            
        # Design Option
        do_id = element.DesignOption
        if do_id:
            params_dict['Design Option'] = doc.GetElement(do_id).Name

        # Phase Created
        pc_param = element.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        if pc_param and pc_param.AsElementId() != DB.ElementId.InvalidElementId:
            params_dict['Phase Created'] = doc.GetElement(pc_param.AsElementId()).Name

    except Exception as e:
        print("Erro nas propriedades básicas: {}".format(e))
    
    # 2. Todos os Parâmetros (Instância e Tipo)
    def add_params_from_set(param_set):
        for param in param_set:
            try:
                name = param.Definition.Name
                if name in params_dict: continue
                
                # Para parâmetros vazios, retornamos string vazia para permitir o filtro
                value = "" 
                if param.HasValue:
                    if param.StorageType == DB.StorageType.String:
                        value = param.AsString() or ""
                    elif param.StorageType == DB.StorageType.Integer:
                        value = param.AsValueString() or str(param.AsInteger())
                    elif param.StorageType == DB.StorageType.Double:
                        value = param.AsValueString() or "{:.4f}".format(param.AsDouble())
                    elif param.StorageType == DB.StorageType.ElementId:
                        value = param.AsValueString() or ""
                
                params_dict[name] = value
            except:
                continue

    add_params_from_set(element.Parameters)
    
    elem_type = doc.GetElement(element.GetTypeId())
    if elem_type:
        add_params_from_set(elem_type.Parameters)
    
    return params_dict

def filter_similar_elements(reference_params, selected_criteria, scope='active_view', reference_cat_id=None):
    """Filtra elementos baseado nos critérios selecionados."""
    
    collector = DB.FilteredElementCollector(doc, active_view.Id) if scope == 'active_view' else DB.FilteredElementCollector(doc)
    collector = collector.WhereElementIsNotElementType()
    
    if reference_cat_id:
        collector = collector.OfCategoryId(reference_cat_id)
    
    all_elements = list(collector.ToElements())
    matching_elements = []
    
    for elem in all_elements:
        elem_params = get_element_parameters(elem)
        match = True
        
        for criterion in selected_criteria:
            # Comparação direta funciona bem com strings normalizadas
            if reference_params.get(criterion) != elem_params.get(criterion):
                match = False
                break
        
        if match:
            matching_elements.append(elem)
    
    return matching_elements

class SmartSelectSimilarWindow(forms.WPFWindow):
    def __init__(self, xaml_file, selected_element):
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.selected_element = selected_element
        self.reference_cat_id = selected_element.Category.Id if selected_element.Category else None
        self.element_params = get_element_parameters(selected_element)
        self.selected_criteria = []
        self.matching_elements = []
        
        # Setup UI initial state
        self.ElementNameLabel.Text = "Objeto: " + (getattr(selected_element, 'Name', 'Objeto Selecionado'))
        self.ElementIdLabel.Text = "ID: " + str(selected_element.Id)
        
        self.populate_parameters()
        self.update_preview()
    
    def populate_parameters(self):
        """Preenche a lista de parâmetros agrupados."""
        self.ParametersPanel.Children.Clear()
        
        groups = {
            'Básicas': ['Category', 'Family', 'Type', 'Level'],
            'Identidade': ['Comments', 'Mark', 'Workset'],
            'Projeto': ['Phase Created', 'Design Option', 'Design Option Set'],
            'Customizados': []
        }
        
        base_list = groups['Básicas'] + groups['Identidade'] + groups['Projeto']
        for p_name in sorted(self.element_params.keys()):
            if p_name not in base_list:
                groups['Customizados'].append(p_name)
        
        # Lista de termos para filtrar "lixo" elétrico de modelos brasileiros
        blacklist = [u"mm²", u"Fase ", u"Fase-", u"Neutro", u"Terra", u"Retorno"]

        for group_name in ['Básicas', 'Identidade', 'Projeto', 'Customizados']:
            params = groups[group_name]
            if not params: continue
            
            # Cabeçalho do grupo (temporário, será adicionado se houver itens visíveis)
            group_header = TextBlock()
            group_header.Text = "━━━ {} ━━━".format(group_name.upper())
            group_header.FontWeight = FontWeights.Bold
            group_header.Foreground = Media.BrushConverter().ConvertFrom("#4fc3f7")
            group_header.Margin = Thickness(0, 15, 0, 5)
            group_header.Tag = "GroupHeader"
            
            header_added = False
            
            for param_name in params:
                # Filtragem de fiação (blacklist)
                is_trash = False
                for term in blacklist:
                    if term in param_name:
                        is_trash = True
                        break
                if is_trash:
                    continue

                if param_name not in self.element_params:
                    continue
                    
                value = self.element_params.get(param_name)
                
                # Label para valores vazios
                display_value = value if (value and str(value).strip() != "") else "<VAZIO>"
                
                # Adicionar cabeçalho se ainda não foi adicionado
                if not header_added:
                    self.ParametersPanel.Children.Add(group_header)
                    header_added = True

                cb = CheckBox()
                cb.Content = "{} : {}".format(param_name, display_value)
                cb.Tag = param_name
                cb.Checked += self.on_criteria_changed
                cb.Unchecked += self.on_criteria_changed
                
                # Default selection
                if param_name in ['Category', 'Family']:
                    cb.IsChecked = True
                
                self.ParametersPanel.Children.Add(cb)

    def on_search_changed(self, sender, args):
        """Filtra a visibilidade dos checkboxes baseado na busca."""
        search_text = self.SearchBox.Text.lower()
        
        for child in self.ParametersPanel.Children:
            if isinstance(child, CheckBox):
                param_text = child.Content.lower()
                child.Visibility = Visibility.Visible if search_text in param_text else Visibility.Collapsed
            elif hasattr(child, 'Tag') and child.Tag == "GroupHeader":
                # headers always visible or handle logic to hide empty groups
                pass

    def on_criteria_changed(self, sender, args):
        self.selected_criteria = [c.Tag for c in self.ParametersPanel.Children if isinstance(c, CheckBox) and c.IsChecked]
        self.update_preview()

    def update_preview(self):
        if not self.selected_criteria:
            self.PreviewLabel.Text = "Selecione ao menos 1 critério"
            self.matching_elements = []
            return
            
        scope = 'active_view' if self.RadioActiveView.IsChecked else 'project'
        self.matching_elements = filter_similar_elements(
            self.element_params, 
            self.selected_criteria,
            scope,
            self.reference_cat_id
        )
        self.PreviewLabel.Text = "📊 {} elementos encontrados".format(len(self.matching_elements))

    def on_preset_family_type(self, sender, args):
        self._batch_set_checks(['Category', 'Family', 'Type'])

    def on_preset_family_comments(self, sender, args):
        self._batch_set_checks(['Category', 'Family', 'Comments'])

    def on_clear_all(self, sender, args):
        self._batch_set_checks([])

    def _batch_set_checks(self, tags):
        for child in self.ParametersPanel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = child.Tag in tags
        self.on_criteria_changed(None, None)

    def on_preview_click(self, sender, args):
        if self.matching_elements:
            ids = List[DB.ElementId]([e.Id for e in self.matching_elements])
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(ids)

    def do_select(self, sender, args):
        if self.matching_elements:
            try:
                config = script.get_config()
                config.set_option('criteria', ",".join(self.selected_criteria))
                config.set_option('scope', 'active_view' if self.RadioActiveView.IsChecked else 'project')
                script.save_config()
            except: pass
            
            ids = List[DB.ElementId]([e.Id for e in self.matching_elements])
            uidoc.Selection.SetElementIds(ids)
        self.Close()

# --- Entry Point ---
selection = uidoc.Selection.GetElementIds()

# 1. Se nada selecionado, pedir para selecionar
if not selection:
    try:
        picked = uidoc.Selection.PickObject(DB.Selection.ObjectType.Element, "Selecione o elemento de referência para o Select Similar")
        if picked:
            selection = [picked.ElementId]
    except:
        script.exit()

if len(selection) >= 1:
    elem = doc.GetElement(selection[0])
    xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
    
    # Load config
    config = script.get_config()
    saved_criteria = None
    saved_scope = 'active_view'
    try:
        c_str = config.get_option('criteria', None)
        if c_str: saved_criteria = c_str.split(",")
        saved_scope = config.get_option('scope', 'active_view')
    except: pass
    
    # Shift check
    shift_pressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift
        
    if shift_pressed or not saved_criteria:
        window = SmartSelectSimilarWindow(xaml_path, elem)
        if saved_criteria:
            window.RadioProject.IsChecked = (saved_scope == 'project')
            window.RadioActiveView.IsChecked = (saved_scope == 'active_view')
            window._batch_set_checks(saved_criteria)
        window.ShowDialog()
    else:
        # Quick execution
        matching = filter_similar_elements(get_element_parameters(elem), saved_criteria, saved_scope, elem.Category.Id if elem.Category else None)
        if matching:
            uidoc.Selection.SetElementIds(List[DB.ElementId]([e.Id for e in matching]))
        else:
            forms.alert("Nenhum similar encontrado.", title="Smart Select Similar")
