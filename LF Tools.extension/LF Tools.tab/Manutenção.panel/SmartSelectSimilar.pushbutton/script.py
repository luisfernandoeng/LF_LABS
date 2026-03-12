# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import os

# Get UI environment
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

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
        if element.Category:
            params_dict['Category'] = element.Category.Name
        else:
            params_dict['Category'] = "Sem Categoria"
            
        # Family e Type
        if hasattr(element, 'Symbol') and element.Symbol:
            params_dict['Family'] = element.Symbol.FamilyName
            params_dict['Type'] = element.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or element.Name
        else:
            params_dict['Family'] = element.Name
            params_dict['Type'] = element.Name

        # Level
        level_id = element.LevelId
        if level_id != DB.ElementId.InvalidElementId:
            # Em links, o level precisa ser procurado no doc do link
            level_el = element.Document.GetElement(level_id)
            if level_el:
                params_dict['Level'] = level_el.Name
            
        # Workset (Ignorado em links geralmente, pois o workset do link é diferente)
        # Se quiser manter para links, precisa checar isWorkshared do link
        workset_id = element.WorksetId
        if workset_id and element.Document.IsWorkshared:
            ws_table = element.Document.GetWorksetTable()
            workset = ws_table.GetWorkset(workset_id)
            params_dict['Workset'] = workset.Name if workset else None
            
        # Design Option
        do_id = element.DesignOption
        if do_id:
            do_el = element.Document.GetElement(do_id)
            if do_el:
                params_dict['Design Option'] = do_el.Name

        # Phase Created
        pc_param = element.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        if pc_param and pc_param.AsElementId() != DB.ElementId.InvalidElementId:
            pc_el = element.Document.GetElement(pc_param.AsElementId())
            if pc_el:
                params_dict['Phase Created'] = pc_el.Name
    except Exception as e:
        pass
    
    # 2. Todos os Parâmetros (Instância e Tipo)
    def add_params_from_set(param_set):
        for param in param_set:
            try:
                name = param.Definition.Name
                if name in params_dict: continue
                
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
    
    elem_type = element.Document.GetElement(element.GetTypeId())
    if elem_type:
        add_params_from_set(elem_type.Parameters)
    
    return params_dict


def filter_similar_elements(target_doc, reference_params, selected_criteria, scope, reference_cat_id=None):
    """
    Filtra elementos no doc alvo (pode ser o host ou um link).
    Se scope for active_view e for o host, usa o ActiveView. Se for link, ignora view.
    """
    if scope == 'active_view' and target_doc.IsFamilyDocument == False and target_doc.Title == doc.Title:
        collector = DB.FilteredElementCollector(target_doc, active_view.Id)
    else:
        # Para vínculos, scope 'active_view' não é diretamente suportado com Id da Vista do Host.
        # Então se for vínculo, usamos Projeto Inteiro.
        collector = DB.FilteredElementCollector(target_doc)
        
    collector = collector.WhereElementIsNotElementType()
    
    if reference_cat_id:
        collector = collector.OfCategoryId(reference_cat_id)
    
    all_elements = list(collector.ToElements())
    matching_elements = []
    
    for elem in all_elements:
        try:
            elem_params = get_element_parameters(elem)
            match = True
            
            for criterion in selected_criteria:
                if reference_params.get(criterion) != elem_params.get(criterion):
                    match = False
                    break
            
            if match:
                matching_elements.append(elem)
        except:
            pass
            
    return matching_elements


class SmartSelectSimilarWindow(forms.WPFWindow):
    def __init__(self, xaml_file, selected_element, is_linked, target_doc, link_instance):
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.selected_element = selected_element
        self.is_linked = is_linked
        self.target_doc = target_doc
        self.link_instance = link_instance
        
        self.reference_cat_id = selected_element.Category.Id if selected_element.Category else None
        self.element_params = get_element_parameters(selected_element)
        self.selected_criteria = []
        self.matching_elements = []  # Elementos do target_doc
        
        # Setup UI initial state
        prefix = "[VÍNCULO] " if is_linked else ""
        self.ElementNameLabel.Text = "Objeto: " + prefix + (getattr(selected_element, 'Name', 'Objeto Selecionado'))
        self.ElementIdLabel.Text = "ID: " + str(selected_element.Id)
        
        # Adjust UI for links (Scope becomes Link Entire Document essentially, but we can leave the radio button names)
        if self.is_linked:
            self.RadioActiveView.Content = "Vista Ativa (Vínculo)"
            self.RadioProject.Content = "Vínculo Inteiro"
        
        self.populate_parameters()
        self.update_preview()
    
    def populate_parameters(self):
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
        
        blacklist = [u"mm²", u"Fase ", u"Fase-", u"Neutro", u"Terra", u"Retorno"]

        for group_name in ['Básicas', 'Identidade', 'Projeto', 'Customizados']:
            params = groups[group_name]
            if not params: continue
            
            group_header = TextBlock()
            group_header.Text = "━━━ {} ━━━".format(group_name.upper())
            group_header.FontWeight = FontWeights.Bold
            group_header.Foreground = Media.BrushConverter().ConvertFrom("#4fc3f7")
            group_header.Margin = Thickness(0, 15, 0, 5)
            group_header.Tag = "GroupHeader"
            
            header_added = False
            
            for param_name in params:
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
                display_value = value if (value and str(value).strip() != "") else "<VAZIO>"
                
                if not header_added:
                    self.ParametersPanel.Children.Add(group_header)
                    header_added = True

                cb = CheckBox()
                cb.Content = "{} : {}".format(param_name, display_value)
                cb.Tag = param_name
                cb.Checked += self.on_criteria_changed
                cb.Unchecked += self.on_criteria_changed
                
                if param_name in ['Category', 'Family']:
                    cb.IsChecked = True
                
                self.ParametersPanel.Children.Add(cb)

    def on_search_changed(self, sender, args):
        search_text = self.SearchBox.Text.lower()
        for child in self.ParametersPanel.Children:
            if isinstance(child, CheckBox):
                param_text = child.Content.lower()
                child.Visibility = Visibility.Visible if search_text in param_text else Visibility.Collapsed

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
            self.target_doc,
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
            if self.is_linked:
                forms.alert("Preview Isolate/Show não suportado nativamente para vínculos da forma tradicional.", title="Preview")
            else:
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
            
            if self.is_linked:
                # Criar Referências Cross-Document para o vínculo
                refs = List[DB.Reference]()
                for e in self.matching_elements:
                    try:
                        # Em Revit 2014+, Reference(Element) é suportado. Depois CreateLinkReference é chamado
                        ref = DB.Reference(e)
                        link_ref = ref.CreateLinkReference(self.link_instance)
                        refs.Add(link_ref)
                    except:
                        pass
                
                if refs.Count > 0:
                    uidoc.Selection.SetReferences(refs)
                else:
                    forms.alert("Falha ao criar referências vinculadas.")
            else:
                # Documento Host (padrão)
                ids = List[DB.ElementId]([e.Id for e in self.matching_elements])
                uidoc.Selection.SetElementIds(ids)
                
        self.Close()


# --- Entry Point ---
# 1. Tentar pegar seleção existente
selection_refs = uidoc.Selection.GetReferences()

if not selection_refs or len(selection_refs) == 0:
    # Se nada selecionado, pedir para selecionar um elemento (suporta vinculados e locais)
    try:
        from Autodesk.Revit.UI.Selection import ObjectType
        picked = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Selecione o elemento de referência no Projeto ou em um Vínculo")
        if picked:
            selection_refs = [picked]
    except:
        script.exit()

if selection_refs and len(selection_refs) >= 1:
    ref = selection_refs[0]
    
    elem = doc.GetElement(ref.ElementId)
    is_linked = False
    target_doc = doc
    link_instance = None
    
    # Detecção Se é um Link
    if isinstance(elem, DB.RevitLinkInstance):
        is_linked = True
        link_instance = elem
        target_doc = link_instance.GetLinkDocument()
        
        if not target_doc:
            forms.alert("Não foi possível acessar o documento do vínculo (talvez esteja descarregado).", exitscript=True)
            
        # Pega o elemento REAL dentro do vínculo (Id contido no LinkedElementId da Referência)
        linked_id = ref.LinkedElementId
        if linked_id and linked_id != DB.ElementId.InvalidElementId:
            elem = target_doc.GetElement(linked_id)
        else:
            forms.alert("Elemento inválido no vínculo.", exitscript=True)
    else:
        # Se GetElement(Reference) não for RevitLinkInstance, também testamos Reference.LinkedElementId para ter certeza
        # pois o pickObject pode retornar o RevitLinkInstance, mas o GetReferences as vezes traz host.
        try:
            val_linked = getattr(ref.LinkedElementId, "IntegerValue", -1)
            # Em API moderna, InvalidElementId é Property. Em ironpython pode causar issue.
            if val_linked > 0 or (hasattr(ref.LinkedElementId, "Value") and ref.LinkedElementId.Value > 0):
                # E definitivamente de um link, vamos pegar a instancia baseada no ElementId principal do ref (que aponta pra instancia de link q segura a geometria)
                link_instance = doc.GetElement(ref.ElementId)
                if isinstance(link_instance, DB.RevitLinkInstance):
                    is_linked = True
                    target_doc = link_instance.GetLinkDocument()
                    elem = target_doc.GetElement(ref.LinkedElementId)
        except:
            pass
            
    if not elem:
        forms.alert("Não foi possível acessar as propriedades do elemento.", exitscript=True)
    
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
    
    shift_pressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift
        
    if shift_pressed or not saved_criteria:
        window = SmartSelectSimilarWindow(xaml_path, elem, is_linked, target_doc, link_instance)
        if saved_criteria:
            window.RadioProject.IsChecked = (saved_scope == 'project')
            window.RadioActiveView.IsChecked = (saved_scope == 'active_view')
            window._batch_set_checks(saved_criteria)
        window.ShowDialog()
    else:
        # Quick execution
        matching = filter_similar_elements(target_doc, get_element_parameters(elem), saved_criteria, saved_scope, elem.Category.Id if elem.Category else None)
        if matching:
            if is_linked:
                refs = List[DB.Reference]()
                for e in matching:
                    try:
                        r = DB.Reference(e).CreateLinkReference(link_instance)
                        refs.Add(r)
                    except: pass
                if refs.Count > 0:
                    uidoc.Selection.SetReferences(refs)
            else:
                uidoc.Selection.SetElementIds(List[DB.ElementId]([e.Id for e in matching]))
        else:
            forms.alert("Nenhum similar encontrado.", title="Smart Select Similar")
