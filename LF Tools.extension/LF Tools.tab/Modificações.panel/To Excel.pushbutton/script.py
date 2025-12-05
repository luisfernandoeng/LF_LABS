# -*- coding: utf-8 -*-
import os
import xlrd
import xlsxwriter
import re
import traceback
import gc # Importante para limpar a memoria e evitar lentidao
from collections import namedtuple
from pyrevit import script, forms, coreutils, revit, DB, HOST_APP

# Logger e Documento
logger = script.get_logger()
doc = revit.doc
active_view = revit.active_view
project_units = doc.GetUnits()

# Limpar output
script.get_output().close_others()

unit_postfix_pattern = re.compile(r"\s*\[.*\]$")

# Definicoes de estruturas de dados
ParamDef = namedtuple(
    "ParamDef", ["name", "istype", "definition", "isreadonly", "isunit", "storagetype"]
)
ElementDef = namedtuple(
    "ElementDef",
    ["name", "label", "category", "family", "type_name", "is_type", "element_label", "elements", "count"],
)

# ==================== HELPERS ====================
def sanitize_filename(name):
    """Remove caracteres invalidos para nome de arquivo."""
    return re.sub(r'[\\/*?:"<>|]', "", name)

def get_elementid_value(element_id):
    if isinstance(element_id, int): return element_id
    if hasattr(element_id, "Value"): return element_id.Value
    return getattr(element_id, "IntegerValue", str(element_id))

def get_elementid_from_value(value):
    return DB.ElementId(value)

def get_parameter_data_type(param_def):
    try: return param_def.GetDataType()
    except AttributeError: return None

def is_yesno_parameter(param_def):
    try:
        return param_def.GetDataType() == DB.SpecTypeId.Boolean.YesNo
    except:
        try: return param_def.ParameterType == DB.ParameterType.YesNo
        except: return False

# ==================== EXTRACAO DE DADOS ====================

def get_panel_schedule_data(schedule_view):
    """V4.1 - Extrai dados de engenharia (MEPModel) dos paineis."""
    try:
        panel_id = None
        try: panel_id = schedule_view.GetPanel()
        except: return [], []

        if not panel_id or panel_id == DB.ElementId.InvalidElementId: return [], []
        panel = doc.GetElement(panel_id)
        if not panel: return [], []

        panel_name = panel.Name
        assigned_systems = []
        try:
            if hasattr(panel, "MEPModel") and panel.MEPModel:
                assigned_systems = panel.MEPModel.GetAssignedElectricalSystems()
            else:
                assigned_systems = panel.GetAssignedElectricalSystems()
        except: return [], []

        if not assigned_systems:
            dummy = {'Tipo': 'PAINEL', 'Nome do Quadro': panel_name, 'Descricao': 'Sem circuitos'}
            return [dummy], ['Tipo', 'Nome do Quadro', 'Descricao']

        headers = ['Nome do Quadro', 'Numero do Circuito', 'Nome da Carga', 'Classificacao (A)', 'Tensao (V)', 'Polos', 'Carga Aparente (VA)', 'Fio']
        data_rows = []

        def get_param_str(elem, builtin_param):
            if not elem: return ""
            try:
                p = elem.get_Parameter(builtin_param)
                if p and p.HasValue:
                    if p.StorageType == DB.StorageType.String: return p.AsString() or ""
                    return p.AsValueString() or ""
            except: pass
            return ""

        for sys in assigned_systems:
            try:
                row_data = {}
                row_data['Tipo'] = 'CIRCUITO'
                row_data['Nome do Quadro'] = panel_name
                row_data['Numero do Circuito'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
                row_data['Nome da Carga'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
                row_data['Classificacao (A)'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
                row_data['Tensao (V)'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
                row_data['Polos'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
                row_data['Carga Aparente (VA)'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_APPARENT_LOAD)
                row_data['Fio'] = get_param_str(sys, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM)
                data_rows.append(row_data)
            except: continue

        try:
            data_rows.sort(key=lambda x: float(re.sub("[^0-9]", "", x['Numero do Circuito'])) if re.sub("[^0-9]", "", x['Numero do Circuito']) else 9999)
        except: pass

        return data_rows, headers
    except: return [], []

def get_schedule_elements_and_params(schedule):
    try: schedule_def = schedule.Definition
    except: return [], []
    
    visible_fields = []
    param_defs_dict = {}
    non_storage_type = coreutils.get_enum_none(DB.StorageType)
    
    try: field_order = schedule_def.GetFieldOrder()
    except: return [], []
    
    for field_id in field_order:
        field = schedule_def.GetField(field_id)
        if field.IsHidden or field.IsCalculatedField: continue
        param_id = field.ParameterId
        if param_id and param_id != DB.ElementId.InvalidElementId: visible_fields.append(field)
    
    collector = DB.FilteredElementCollector(doc, schedule.Id)
    elements = []
    for el in collector:
        try:
            if hasattr(el, 'Category') and el.Category and not el.Category.IsTagCategory: elements.append(el)
        except: continue
    
    if not elements: return [], []
    
    for field in visible_fields:
        param_id = field.ParameterId
        for el in elements:
            param = None
            try: param = el.get_Parameter(param_id)
            except:
                field_name = field.GetName()
                if field_name: param = el.LookupParameter(field_name)
            if param and param.StorageType != non_storage_type:
                def_name = param.Definition.Name
                if def_name not in param_defs_dict:
                    dt = get_parameter_data_type(param.Definition)
                    param_defs_dict[def_name] = ParamDef(
                        name=def_name, istype=False, definition=param.Definition, isreadonly=param.IsReadOnly,
                        isunit=DB.UnitUtils.IsMeasurableSpec(dt) if dt else False, storagetype=param.StorageType,
                    )
                break
    return elements, sorted(param_defs_dict.values(), key=lambda pd: pd.name)

def select_elements(elements):
    if not elements: return []
    type_element_map = {}
    for el in elements:
        try:
            cat = el.Category.Name if el.Category else "Outros"
            is_type = DB.ElementIsElementTypeFilter().PassesFilter(el)
            name = getattr(el, "Name", "Elemento")
            key = (cat, name, is_type)
            if key not in type_element_map: type_element_map[key] = []
            type_element_map[key].append(el)
        except: continue
        
    defs = []
    for key, lst in type_element_map.items():
        cat, name, is_type = key
        label = "[{}] {} ({})".format(cat, name, len(lst))
        defs.append(ElementDef(name=label, label=label, category=cat, family="", type_name=name, is_type=is_type, element_label="", elements=lst, count=len(lst)))
    
    sel = forms.SelectFromList.show(sorted(defs, key=lambda x: x.label), title="Selecione Elementos", width=500, multiselect=True)
    if not sel: return []
    res = []
    for d in sel: res.extend(d.elements)
    return res

def select_parameters(src_elements, is_panel_schedule=False):
    param_defs_dict = {}
    non_storage_type = coreutils.get_enum_none(DB.StorageType)
    limit = min(50, len(src_elements))
    for i in range(limit):
        el = src_elements[i]
        if not hasattr(el, 'Parameters'): continue
        for p in el.Parameters:
            if p.StorageType != non_storage_type:
                try:
                    def_name = p.Definition.Name
                    if def_name not in param_defs_dict:
                        dt = get_parameter_data_type(p.Definition)
                        param_defs_dict[def_name] = ParamDef(
                            name=def_name, istype=False, definition=p.Definition, isreadonly=p.IsReadOnly,
                            isunit=DB.UnitUtils.IsMeasurableSpec(dt) if dt else False, storagetype=p.StorageType,
                        )
                except: continue
    sel = forms.SelectFromList.show(sorted(param_defs_dict.values(), key=lambda x: x.name), width=450, multiselect=True, title="Selecione Parametros")
    return sel if sel else []

# ==================== EXPORTACAO ====================

def export_xls(src_elements, selected_params, file_path, is_panel_schedule=False):
    try:
        workbook = xlsxwriter.Workbook(file_path)
        
        if is_panel_schedule:
            ws = workbook.add_worksheet("Quadro de Cargas")
            fmt_head = workbook.add_format({"bold": True, "bg_color": "#4F81BD", "font_color": "white", "border": 1, "align": "center"})
            fmt_data = workbook.add_format({"border": 1, "align": "left"})
            
            for i, p in enumerate(selected_params):
                ws.write(0, i, p.name, fmt_head)
                ws.set_column(i, i, 20)
                
            for r, el in enumerate(src_elements, 1):
                if hasattr(el, 'row_data'):
                    for c, p in enumerate(selected_params):
                        ws.write(r, c, str(el.row_data.get(p.name, "")), fmt_data)
            
            ws.autofilter(0, 0, len(src_elements), len(selected_params)-1)
            
        else:
            ws = workbook.add_worksheet("Export")
            fmt_bold = workbook.add_format({"bold": True})
            fmt_lock_ro = workbook.add_format({"locked": True, "font_color": "#C0504D", "italic": True})
            fmt_lock_id = workbook.add_format({"locked": True, "font_color": "#95B3D7", "italic": True})
            fmt_unlock = workbook.add_format({"locked": False})
            fmt_head_id = workbook.add_format({"bold": True, "bg_color": "#DCE6F1", "font_color": "#1F4E78"})
            
            ws.freeze_panes(1, 0)
            ws.write(0, 0, "ElementId", fmt_head_id)
            
            valid_params = [p for p in selected_params if not unit_postfix_pattern.search(p.name)]
            
            for i, p in enumerate(valid_params):
                post = ""
                fmt = fmt_bold
                if p.isreadonly: fmt = workbook.add_format({"bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006"})
                dt = get_parameter_data_type(p.definition)
                if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                    try:
                        sym = project_units.GetFormatOptions(dt).GetSymbolTypeId()
                        if not sym.Empty(): post = " [" + DB.LabelUtils.GetLabelForSymbol(sym) + "]"
                    except: pass
                ws.write(0, i+1, p.name + post, fmt)
            
            widths = [len("ElementId")] + [len(p.name) for p in valid_params]
            
            for r, el in enumerate(src_elements, 1):
                try:
                    eid = get_elementid_value(el.Id)
                    ws.write(r, 0, str(eid), fmt_lock_id)
                    for c, p in enumerate(valid_params):
                        val = ""
                        pval = el.LookupParameter(p.name)
                        if pval and pval.HasValue:
                            try:
                                if pval.StorageType == DB.StorageType.Double:
                                    dt = get_parameter_data_type(p.definition)
                                    val = pval.AsDouble()
                                    if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                                        uid = pval.GetUnitTypeId()
                                        val = DB.UnitUtils.ConvertFromInternalUnits(pval.AsDouble(), uid)
                                elif pval.StorageType == DB.StorageType.String: val = pval.AsString()
                                elif pval.StorageType == DB.StorageType.Integer:
                                    if is_yesno_parameter(p.definition): val = "Yes" if pval.AsInteger() else "No"
                                    else: val = str(pval.AsInteger())
                                elif pval.StorageType == DB.StorageType.ElementId: val = pval.AsValueString()
                            except: val = "Error"
                        fmt = fmt_lock_ro if p.isreadonly else fmt_unlock
                        ws.write(r, c+1, val, fmt)
                        slen = len(str(val)) if val else 0
                        if slen > widths[c+1]: widths[c+1] = min(slen, 50)
                except: continue
            
            for i, w in enumerate(widths): ws.set_column(i, i, w+3)
            if len(src_elements) > 0: ws.autofilter(0, 0, len(src_elements), len(valid_params))

        try:
            workbook.close()
        except Exception as e:
            msg = str(e)
            if "Permission denied" in msg or "being used" in msg or "Errno 32" in msg:
                forms.alert("O arquivo Excel esta ABERTO.\nFeche-o e tente novamente.", title="Arquivo em Uso")
            else: raise e

    except Exception as e:
        logger.error("Erro Excel: {}".format(e))
        raise e

def import_xls(file_path):
    try:
        wb = xlrd.open_workbook(file_path)
        if "Quadro de Cargas" in wb.sheet_names():
            forms.alert("Arquivos de Quadro de Cargas sao somente leitura.", title="Aviso")
            return
        sh = wb.sheet_by_name("Export")
    except:
        forms.alert("Erro ao abrir Excel ou aba 'Export' nao encontrada.", title="Erro")
        return
    
    headers = sh.row_values(0)
    if not headers or str(headers[0]).strip() != "ElementId":
        forms.alert("A celula A1 deve conter 'ElementId'.", title="Erro de Formato")
        return
        
    pnames = headers[1:]
    with revit.Transaction("Import Excel"):
        for rx in range(1, sh.nrows):
            row = sh.row_values(rx)
            try:
                eid = int(float(row[0]))
                el = doc.GetElement(get_elementid_from_value(eid))
                if not el: continue
                for cx, pname in enumerate(pnames):
                    val = row[cx+1]
                    if val in (None, ""): continue
                    pname = unit_postfix_pattern.sub("", pname).strip()
                    param = el.LookupParameter(pname)
                    if not param or param.IsReadOnly: continue
                    try:
                        st = param.StorageType
                        if st == DB.StorageType.String: param.Set(str(val))
                        elif st == DB.StorageType.Integer:
                            if is_yesno_parameter(param.Definition):
                                sval = str(val).strip().lower()
                                ival = 1 if sval in ("yes", "1", "true", "sim") else 0
                                param.Set(ival)
                            else: param.Set(int(float(val)))
                        elif st == DB.StorageType.Double:
                            dt = get_parameter_data_type(param.Definition)
                            if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                                uid = project_units.GetFormatOptions(dt).GetUnitTypeId()
                                vconv = DB.UnitUtils.ConvertToInternalUnits(float(val), uid)
                                param.Set(vconv)
                            else: param.Set(float(val))
                    except: pass
            except: pass
    
    # LIMPEZA DE MEMORIA APOS IMPORTACAO
    gc.collect()

# ==================== UI ====================

class ExportImportWindow(forms.WPFWindow):
    def __init__(self, xaml_file_path):
        forms.WPFWindow.__init__(self, xaml_file_path)
        self.export_path = None
        self.import_path = None
        self.view_schedules = []
        self.panel_schedules = []
        
        self.Button_BrowseExport.Click += self.browse_export
        self.Button_BrowseImport.Click += self.browse_import
        self.Button_Export.Click += self.do_export
        self.Button_Import.Click += self.do_import
        self.Button_Close.Click += self.close_clicked
        self.ComboBox_ExportMode.SelectionChanged += self.mode_changed
        
        self.load_schedules()
        self.mode_changed(None, None)
    
    def update_status(self, msg): self.TextBlock_Status.Text = msg
    
    def load_schedules(self):
        self.view_schedules = []
        self.panel_schedules = []
        try:
            for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements():
                if s.IsTemplate: continue
                if "PanelSchedule" in s.GetType().Name: continue
                try: 
                    if s.Definition.IsInternalKeynoteSchedule or s.Definition.IsRevisionSchedule: continue
                except: pass
                self.view_schedules.append({'schedule': s, 'display_name': s.Name})
                
            for v in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements():
                if "PanelScheduleView" in v.GetType().Name and not v.IsTemplate:
                    self.panel_schedules.append({'schedule': v, 'display_name': v.Name})
            
            self.view_schedules.sort(key=lambda x: x['display_name'])
            self.panel_schedules.sort(key=lambda x: x['display_name'])
        except: pass

    def mode_changed(self, sender, args):
        idx = self.ComboBox_ExportMode.SelectedIndex
        self.ComboBox_SubSelection.Items.Clear()
        self.ComboBox_SubSelection.IsEnabled = True
        
        if idx == 0: 
            if not self.view_schedules: self.ComboBox_SubSelection.Items.Add("Nenhuma tabela")
            else: 
                for i in self.view_schedules: self.ComboBox_SubSelection.Items.Add(i['display_name'])
            self.ComboBox_SubSelection.SelectedIndex = 0
        elif idx == 1:
            if not self.panel_schedules: self.ComboBox_SubSelection.Items.Add("Nenhum quadro")
            else:
                for i in self.panel_schedules: self.ComboBox_SubSelection.Items.Add(i['display_name'])
            self.ComboBox_SubSelection.SelectedIndex = 0
        else:
            self.ComboBox_SubSelection.IsEnabled = False
            self.ComboBox_SubSelection.Items.Add("Selecao Automatica")
            self.ComboBox_SubSelection.SelectedIndex = 0

    def browse_export(self, sender, args):
        proj_name = sanitize_filename(doc.Title.replace(".rvt", ""))
        suggested_name = "Export_" + proj_name
        
        idx = self.ComboBox_ExportMode.SelectedIndex
        sub_idx = self.ComboBox_SubSelection.SelectedIndex
        
        if idx in [0, 1] and sub_idx >= 0:
            try:
                item_name = str(self.ComboBox_SubSelection.SelectedItem)
                if item_name and "Nenh" not in item_name:
                    clean_item = sanitize_filename(item_name)
                    suggested_name = "{}_{}".format(proj_name, clean_item)
            except: pass
            
        file_path = forms.save_file(file_ext="xlsx", default_name=suggested_name + ".xlsx")
        if file_path:
            self.export_path = file_path
            self.TextBox_ExportPath.Text = file_path
            self.update_status("Salvar em: " + os.path.basename(file_path))

    def browse_import(self, sender, args):
        path = forms.pick_file(file_ext="xlsx")
        if path:
            self.import_path = path
            self.TextBox_ImportPath.Text = path
            self.update_status("Ler de: " + os.path.basename(path))

    def do_export(self, sender, args):
        if not self.export_path: return forms.alert("Defina o local do arquivo.")
        
        idx = self.ComboBox_ExportMode.SelectedIndex
        src = []
        params = []
        is_panel = False
        
        try:
            self.update_status("Lendo dados...")
            
            if idx == 0: 
                if self.view_schedules:
                    sch = self.view_schedules[self.ComboBox_SubSelection.SelectedIndex]['schedule']
                    src, params = get_schedule_elements_and_params(sch)
                    if not params:
                        self.Hide()
                        params = select_parameters(src)
                        self.Show()
            
            elif idx == 1: 
                if self.panel_schedules:
                    sch = self.panel_schedules[self.ComboBox_SubSelection.SelectedIndex]['schedule']
                    is_panel = True
                    rows, headers = get_panel_schedule_data(sch)
                    class RowObj:
                        def __init__(self, d): self.row_data = d
                    src = [RowObj(r) for r in rows]
                    params = [ParamDef(name=h, istype=False, definition=None, isreadonly=True, isunit=False, storagetype=DB.StorageType.String) for h in headers]

            elif idx == 2: 
                src = select_elements(revit.query.get_all_elements_in_view(active_view))
                if src: params = select_parameters(src)
            elif idx == 3: 
                el = revit.query.get_all_elements(doc)
                src = select_elements([e for e in el if DB.ElementIsElementTypeFilter().PassesFilter(e)])
                if src: params = select_parameters(src)
            elif idx == 4: 
                el = revit.query.get_all_elements(doc)
                src = select_elements([e for e in el if DB.ElementIsElementTypeFilter(True).PassesFilter(e)])
                if src: params = select_parameters(src)
            elif idx == 5: 
                self.Hide()
                el = revit.pick_elements()
                self.Show()
                src = select_elements(el)
                if src: params = select_parameters(src)

            if not src and not is_panel: 
                self.update_status("Nada para exportar.")
                return

            self.update_status("Gerando Excel...")
            export_xls(src, params, self.export_path, is_panel)
            
            # LIMPEZA AGRESSIVA DE MEMORIA
            src = None
            params = None
            gc.collect()
            
            self.update_status("Concluido!")
            forms.alert("Exportacao finalizada com sucesso!", title="Sucesso")
            
        except Exception as e:
            logger.error(traceback.format_exc())
            forms.alert("Erro: " + str(e))

    def do_import(self, sender, args):
        if not self.import_path: return forms.alert("Selecione o arquivo.")
        self.update_status("Importando...")
        import_xls(self.import_path)
        self.update_status("Importacao Finalizada!")
        forms.alert("Importacao concluida!", title="Sucesso")

    def close_clicked(self, sender, args):
        # Limpeza final ao fechar
        self.view_schedules = []
        self.panel_schedules = []
        gc.collect()
        self.Close()

if __name__ == "__main__":
    xaml_file = script.get_bundle_file("Exportimport.xaml")
    if os.path.exists(xaml_file):
        ExportImportWindow(xaml_file).ShowDialog()