#! python3
# -*- coding: utf-8 -*-
import os
import re
import json

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit import DB

def _script_dir():
    try:
        return __commandpath__
    except NameError:
        return os.path.dirname(os.path.abspath(__file__))

_CONFIG_FILE = os.path.join(_script_dir(), "sss_config.json")

def _load_config():
    try:
        with open(_CONFIG_FILE, 'r') as _f:
            return json.load(_f)
    except Exception:
        return {}

def _save_config(data):
    try:
        with open(_CONFIG_FILE, 'w') as _f:
            json.dump(data, _f)
    except Exception:
        pass

# Get UI environment  (__revit__ is injected by pyRevit CPython runtime)
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = uidoc.ActiveView

import System
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows import Window, Thickness, FontWeight, FontWeights, Media, Visibility
from System.Windows.Controls import TextBlock, CheckBox, TextBox
from System.Collections.Generic import List
from System.Windows.Forms import Control, Keys

# ==================== CPYTHON COMPAT ====================
try:
    clr.AddReference('PresentationFramework')
except Exception:
    pass
import System.Windows.Forms as _WF
from System.Windows import MessageBox as _MB, MessageBoxButton as _MBBtn, MessageBoxResult as _MBRes

def _alert(msg, title="LF Tools", yes=False, no=False, exitscript=False, **kw):
    if yes and no:
        r = _MB.Show(str(msg), str(title), _MBBtn.YesNo)
        ans = r == _MBRes.Yes
        if exitscript and not ans:
            raise SystemExit()
        return ans
    _MB.Show(str(msg), str(title))
    if exitscript:
        raise SystemExit()

def _toast(msg, **kw):
    pass  # No-op: pyrevit script logger not available in CPython standalone

class _WPFWindowCPy:
    """CPython drop-in for pyrevit.forms.WPFWindow."""
    _XAML_EVENTS = re.compile(
        r'\s+(?:x:Class|'
        r'Click|DoubleClick|'
        r'Mouse(?:Down|Up|Move|Enter|Leave|Wheel)|'
        r'Preview(?:Mouse(?:Down|Up|Move|LeftButtonDown|LeftButtonUp)|'
        r'Key(?:Down|Up)|TextInput)|'
        r'Key(?:Down|Up)|TextInput|TextChanged|SelectionChanged|'
        r'SelectedItemChanged|ValueChanged|ScrollChanged|'
        r'Got(?:Focus|KeyboardFocus)|Lost(?:Focus|KeyboardFocus)|'
        r'Checked|Unchecked|Indeterminate|'
        r'Loaded|Unloaded|Initialized|'
        r'Clos(?:ing|ed)|Activated|Deactivated|'
        r'SizeChanged|LayoutUpdated|ContentRendered|'
        r'Drag(?:Enter|Leave|Over)|Drop|'
        r'ContextMenu(?:Opening|Closing)|'
        r'ToolTip(?:Opening|Closing)|'
        r'DataContextChanged|IsVisibleChanged|IsEnabledChanged|'
        r'RequestBringIntoView|SourceUpdated|TargetUpdated)'
        r'\s*=\s*(?:"[^"]*"|\'[^\']*\')'
    )

    def __init__(self, xaml_source, literal_string=None):
        from System.IO import StringReader
        from System.Windows.Markup import XamlReader
        import System.Xml
        stripped = str(xaml_source).strip()
        is_inline = (literal_string is True or
                     (literal_string is None and stripped.startswith('<')))
        if not is_inline:
            with open(str(xaml_source), 'r', encoding='utf-8') as _f:
                stripped = _f.read().strip()
        xaml_clean = self._XAML_EVENTS.sub('', stripped)
        rdr = System.Xml.XmlReader.Create(StringReader(xaml_clean))
        self._window = XamlReader.Load(rdr)

    def __getattr__(self, name):
        if name.startswith('_'):
            raise AttributeError(name)
        win = object.__getattribute__(self, '_window')
        el = win.FindName(name)
        if el is not None:
            return el
        return getattr(win, name)

    def ShowDialog(self):
        return self._window.ShowDialog()

    def Show(self):
        return self._window.Show()

    def Close(self):
        self._window.Close()

# ==================== FIM CPYTHON COMPAT ====================

def set_red_highlight(view, element_ids, apply=True):
    """Aplica ou remove o preenchimento vermelho sólido nos elementos."""
    doc = view.Document
    if apply:
        # Busca o padrão de preenchimento sólido no documento
        solid_fill = None
        fill_patterns = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
        for fp in fill_patterns:
            try:
                if fp.GetFillPattern().IsSolidFill:
                    solid_fill = fp
                    break
            except: continue
        
        ogs = DB.OverrideGraphicSettings()
        red = DB.Color(255, 0, 0)
        
        # Configura as linhas em vermelho e com espessura maior
        ogs.SetProjectionLineColor(red)
        try: ogs.SetProjectionLineWeight(8)
        except: pass
        
        # Se encontrou o padrão sólido, aplica preenchimento (Foreground e Background)
        if solid_fill:
            # Foreground
            try:
                ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                ogs.SetSurfaceForegroundPatternColor(red)
                ogs.SetSurfaceForegroundPatternVisible(True)
            except: pass
            
            # Background
            try:
                ogs.SetSurfaceBackgroundPatternId(solid_fill.Id)
                ogs.SetSurfaceBackgroundPatternColor(red)
                ogs.SetSurfaceBackgroundPatternVisible(True)
            except: pass
            
        for eid in element_ids:
            try: view.SetElementOverrides(eid, ogs)
            except: pass
    else:
        # Limpa todos os overrides aplicados
        blank = DB.OverrideGraphicSettings()
        for eid in element_ids:
            try: view.SetElementOverrides(eid, blank)
            except: pass


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
        try:
            do_obj = element.DesignOption
            if do_obj:
                params_dict['Design Option'] = do_obj.Name
        except:
            pass

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
    if scope == 'active_view' and target_doc.Equals(doc) and not target_doc.IsFamilyDocument:
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


class SmartSelectSimilarWindow(_WPFWindowCPy):
    def __init__(self, xaml_file, selected_element, is_linked, target_doc, link_instance):
        _WPFWindowCPy.__init__(self, xaml_file)
        
        self.selected_element = selected_element
        self.is_linked = is_linked
        self.target_doc = target_doc
        self.link_instance = link_instance
        
        self.reference_cat_id = selected_element.Category.Id if selected_element.Category else None
        self.element_params = get_element_parameters(selected_element)
        self.selected_criteria = []
        self.matching_elements = []  # Elementos do target_doc
        self.last_highlighted_ids = [] # Cache para limpar o highlight anterior
        
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

        # Reconecta eventos removidos pelo regex do _WPFWindowCPy
        self.SearchBox.TextChanged += self.on_search_changed
        self.RadioActiveView.Checked += self.on_scope_changed
        self.RadioProject.Checked += self.on_scope_changed
        self.BtnPresetFamilyType.Click += self.on_preset_family_type
        self.BtnPresetFamilyComments.Click += self.on_preset_family_comments
        self.BtnClearAll.Click += self.on_clear_all
        self.PaintButton.Click += self.on_paint_click
        self.ZoomButton.Click += self.on_zoom_click
        self.BtnPreview.Click += self.on_preview_click
        self.BtnSelect.Click += self.do_select

        # Evento para limpar highlight ao fechar
        self.Closed += self.on_window_closed
    
    def on_window_closed(self, sender, args):
        self.clear_highlight()

    def clear_highlight(self):
        if self.last_highlighted_ids:
            t = DB.Transaction(doc, "Limpar Highlight")
            t.Start()
            try:
                set_red_highlight(active_view, self.last_highlighted_ids, apply=False)
                t.Commit()
            except:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
            self.last_highlighted_ids = []
    
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

    def on_scope_changed(self, sender, args):
        self.update_preview()

    def on_search_changed(self, sender, args):
        search_text = str(self.SearchBox.Text).lower()
        for child in self.ParametersPanel.Children:
            if isinstance(child, CheckBox):
                param_text = str(child.Content or "").lower()
                child.Visibility = Visibility.Visible if search_text in param_text else Visibility.Collapsed

    def on_criteria_changed(self, sender, args):
        self.selected_criteria = []
        for c in self.ParametersPanel.Children:
            if isinstance(c, CheckBox) and c.IsChecked == True:
                self.selected_criteria.append(str(c.Tag))
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
        
        # Atualiza o destaque vermelho em tempo real
        if not self.is_linked: # Só destaca elementos no documento host (ActiveView)
            self.clear_highlight()
            if self.matching_elements:
                self.last_highlighted_ids = [e.Id for e in self.matching_elements]
                t = DB.Transaction(doc, "Highlight Temporário")
                t.Start()
                try:
                    set_red_highlight(active_view, self.last_highlighted_ids, apply=True)
                    t.Commit()
                except:
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()

    def on_preset_family_type(self, sender, args):
        self._batch_set_checks(['Category', 'Family', 'Type'])

    def on_preset_family_comments(self, sender, args):
        self._batch_set_checks(['Category', 'Family', 'Comments'])

    def on_clear_all(self, sender, args):
        self._batch_set_checks([])

    def _batch_set_checks(self, tags):
        for child in self.ParametersPanel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = (str(child.Tag) in tags)
        self.on_criteria_changed(None, None)

    def on_preview_click(self, sender, args):
        if self.matching_elements:
            if self.is_linked:
                # Para links, a melhor forma de preview é selecionar e dar zoom aproximado
                refs = List[DB.Reference]()
                for e in self.matching_elements:
                    try:
                        ref = DB.Reference(e).CreateLinkReference(self.link_instance)
                        refs.Add(ref)
                    except: pass
                uidoc.Selection.SetReferences(refs)
                uidoc.ShowElements(self.link_instance.Id)
            else:
                ids = List[DB.ElementId]()
                for e in self.matching_elements:
                    ids.Add(e.Id)
                uidoc.Selection.SetElementIds(ids)
                uidoc.ShowElements(ids)

    def on_paint_click(self, sender, args):
        """Pintura persistente (não limpa ao fechar)."""
        if self.matching_elements:
            if self.is_linked:
                _alert("Pintura individual de elementos em VÍNCULOS não é suportada diretamente pelo Revit nesta versão. Use o Zoom para identificá-los.", title="Aviso")
                return

            ids = [e.Id for e in self.matching_elements]
            t = DB.Transaction(doc, "Pintar Elementos")
            t.Start()
            try:
                set_red_highlight(active_view, ids, apply=True)
                t.Commit()
            except:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
            
            # Se forem os mesmos do preview, desvincula da limpeza automática
            if ids == self.last_highlighted_ids:
                self.last_highlighted_ids = []
            
            _toast("🎨 Seleção pintada de vermelho!")

    def on_zoom_click(self, sender, args):
        """Dá zoom nos elementos encontrados."""
        if self.matching_elements:
            if self.is_linked:
                refs = List[DB.Reference]()
                for e in self.matching_elements:
                    try:
                        ref = DB.Reference(e).CreateLinkReference(self.link_instance)
                        refs.Add(ref)
                    except: pass
                uidoc.Selection.SetReferences(refs)
                uidoc.ShowElements(self.link_instance.Id)
            else:
                ids = List[DB.ElementId]()
                for e in self.matching_elements:
                    ids.Add(e.Id)
                uidoc.ShowElements(ids)

    def do_select(self, sender, args):
        if self.matching_elements:
            try:
                _save_config({
                    'criteria': ",".join(self.selected_criteria),
                    'scope': 'active_view' if self.RadioActiveView.IsChecked == True else 'project',
                })
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
                    _alert("Falha ao criar referências vinculadas.")
            else:
                # Documento Host (padrão)
                ids = List[DB.ElementId]()
                for e in self.matching_elements:
                    ids.Add(e.Id)
                uidoc.Selection.SetElementIds(ids)
                
        self.Close()


# --- Entry Point ---
# 1. Tentar pegar seleção existente — CPython usa GetElementIds (não GetReferences)
selection_refs = None
try:
    sel_ids = uidoc.Selection.GetElementIds()
    if sel_ids and sel_ids.Count > 0:
        # Converter ElementIds para References fake para manter compatibilidade
        first_id = list(sel_ids)[0]
        first_elem = doc.GetElement(first_id)
        if first_elem:
            class _FakeRef:
                def __init__(self, eid, linked_eid=None):
                    self.ElementId = eid
                    self.LinkedElementId = linked_eid or DB.ElementId.InvalidElementId
            selection_refs = [_FakeRef(first_id)]
except:
    pass

if not selection_refs:
    # Se nada selecionado, pedir para selecionar um elemento (suporta vinculados e locais)
    try:
        from Autodesk.Revit.UI.Selection import ObjectType
        picked = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Selecione o elemento de referência no Projeto ou em um Vínculo")
        if picked:
            selection_refs = [picked]
    except:
        raise SystemExit()

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
            _alert("Não foi possível acessar o documento do vínculo (talvez esteja descarregado).", exitscript=True)

        # Pega o elemento REAL dentro do vínculo (Id contido no LinkedElementId da Referência)
        linked_id = ref.LinkedElementId
        if linked_id and linked_id != DB.ElementId.InvalidElementId:
            elem = target_doc.GetElement(linked_id)
        else:
            _alert("Elemento inválido no vínculo.", exitscript=True)
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
        _alert("Não foi possível acessar as propriedades do elemento.", exitscript=True)

    xaml_path = os.path.join(_script_dir(), "ui.xaml")

    # Load config
    saved_criteria = None
    saved_scope = 'active_view'
    try:
        _cfg = _load_config()
        c_str = _cfg.get('criteria', None)
        if c_str: saved_criteria = str(c_str).split(",")
        saved_scope = str(_cfg.get('scope', 'active_view'))
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
                _ids = List[DB.ElementId]()
                for e in matching:
                    _ids.Add(e.Id)
                uidoc.Selection.SetElementIds(_ids)
        else:
            _alert("Nenhum similar encontrado.", title="Smart Select Similar")
