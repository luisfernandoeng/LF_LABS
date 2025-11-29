# -*- coding: utf-8 -*-
import os
import xlrd
import xlsxwriter
import re
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

# Definições de estruturas de dados
ParamDef = namedtuple(
    "ParamDef", ["name", "istype", "definition", "isreadonly", "isunit", "storagetype"]
)

ElementDef = namedtuple(
    "ElementDef",
    ["name", "label", "category", "family", "type_name", "is_type", "element_label", "elements", "count"],
)

# ==================== HELPERS DE COMPATIBILIDADE ====================

def get_elementid_value(element_id):
    """Compatibilidade para Revit <2024 e 2024+."""
    if isinstance(element_id, int):
        return element_id
    if hasattr(element_id, "Value"):
        return element_id.Value
    return getattr(element_id, "IntegerValue", str(element_id))

def get_elementid_from_value(value):
    return DB.ElementId(value)

def get_parameter_data_type(param_def):
    try:
        return param_def.GetDataType()
    except AttributeError:
        return None

def is_yesno_parameter(param_def):
    try:
        data_type = param_def.GetDataType()
        return data_type == DB.SpecTypeId.Boolean.YesNo
    except (AttributeError, NameError):
        try:
            return param_def.ParameterType == DB.ParameterType.YesNo
        except:
            return False

# ==================== FUNÇÕES DE EXPORTAÇÃO ====================

def create_element_definitions(elements):
    type_element_map = {}

    for el in elements:
        if not el.Category or el.Category.IsTagCategory:
            continue
        try:
            category = el.Category.Name
            is_type = DB.ElementIsElementTypeFilter().PassesFilter(el)
            family = "NoFamily"
            type_name = "NoType"
            element_label = "unknown"

            if is_type:
                type_name = getattr(el, "Name", "Unnamed Type")
                element_label = "type(s)"
            else:
                if isinstance(el, DB.FamilyInstance):
                    family = el.Symbol.Family.Name
                    type_name = el.Symbol.get_Parameter(
                        DB.BuiltInParameter.SYMBOL_NAME_PARAM
                    ).AsString()
                    element_label = "family instance(s)"
                else:
                    type_name = getattr(el, "Name", "Unnamed Instance")
                    element_label = "instance(s)"

            key = (category, family, type_name, is_type, element_label)
            if key not in type_element_map:
                type_element_map[key] = []
            type_element_map[key].append(el)

        except Exception as ex:
            continue

    element_defs = []
    for key, element_list in type_element_map.items():
        category, family, type_name, is_type, element_label = key
        count = len(element_list)
        label = "[{} : {}] {} ({})".format(category, family, type_name, count)

        element_def = ElementDef(
            name=label,
            label=label,
            category=category,
            family=family,
            type_name=type_name,
            is_type=is_type,
            element_label=element_label,
            elements=element_list,
            count=count,
        )
        element_defs.append(element_def)

    return element_defs

def select_elements(elements):
    element_defs = create_element_definitions(elements)
    if not element_defs:
        return []

    element_defs = sorted(element_defs, key=lambda x: x.label)
    
    # Interface simples do pyRevit para selecionar categorias
    selected_defs = forms.SelectFromList.show(
        element_defs,
        title="Selecione Elementos para Exportar",
        width=500,
        multiselect=True,
        button_name="Selecionar"
    )

    if not selected_defs:
        return []

    src_elements = []
    for elem_def in selected_defs:
        src_elements.extend(elem_def.elements)
    return src_elements

def get_schedule_elements_and_params(schedule):
    schedule_def = schedule.Definition
    visible_fields = []
    param_defs_dict = {}
    non_storage_type = coreutils.get_enum_none(DB.StorageType)

    for field_id in schedule_def.GetFieldOrder():
        field = schedule_def.GetField(field_id)
        if field.IsHidden or field.IsCalculatedField:
            continue
        param_id = field.ParameterId
        if param_id and param_id != DB.ElementId.InvalidElementId:
            visible_fields.append(field)

    collector = DB.FilteredElementCollector(doc, schedule.Id)
    elements = [el for el in collector if el.Category and not el.Category.IsTagCategory]

    if not elements:
        raise ValueError("Nenhum elemento encontrado nesta tabela.")

    for field in visible_fields:
        param_id = field.ParameterId
        for el in elements:
            param = None
            try:
                param = el.get_Parameter(param_id)
            except:
                field_name = field.GetName()
                if field_name:
                    param = el.LookupParameter(field_name)

            if param and param.StorageType != non_storage_type:
                def_name = param.Definition.Name
                if def_name not in param_defs_dict:
                    param_data_type = get_parameter_data_type(param.Definition)
                    param_defs_dict[def_name] = ParamDef(
                        name=def_name,
                        istype=False,
                        definition=param.Definition,
                        isreadonly=param.IsReadOnly,
                        isunit=DB.UnitUtils.IsMeasurableSpec(param_data_type) if param_data_type else False,
                        storagetype=param.StorageType,
                    )
                break

    param_defs = sorted(param_defs_dict.values(), key=lambda pd: pd.name)
    return elements, param_defs

def select_parameters(src_elements):
    param_defs_dict = {}
    non_storage_type = coreutils.get_enum_none(DB.StorageType)

    for el in src_elements:
        for p in el.Parameters:
            if p.StorageType != non_storage_type:
                def_name = p.Definition.Name
                if def_name not in param_defs_dict:
                    param_data_type = get_parameter_data_type(p.Definition)
                    param_defs_dict[def_name] = ParamDef(
                        name=def_name,
                        istype=False,
                        definition=p.Definition,
                        isreadonly=p.IsReadOnly,
                        isunit=DB.UnitUtils.IsMeasurableSpec(param_data_type) if param_data_type else False,
                        storagetype=p.StorageType,
                    )

    param_defs = sorted(param_defs_dict.values(), key=lambda pd: pd.name)

    selected_params = forms.SelectFromList.show(
        param_defs,
        width=450,
        multiselect=True,
        title="Selecione os Parâmetros para Exportar",
        button_name="Selecionar"
    )
    
    return selected_params if selected_params else []

def export_xls(src_elements, selected_params, file_path):
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet("Export")

    # Formatos de célula
    bold = workbook.add_format({"bold": True})
    unlocked = workbook.add_format({"locked": False})
    
    # CORREÇÃO: Formato vermelho para parâmetros read-only
    locked_readonly = workbook.add_format({
        "locked": True, 
        "font_color": "#C0504D",  # Vermelho escuro
        "italic": True
    })
    
    locked_elementid = workbook.add_format({
        "locked": True,
        "font_color": "#95B3D7",  # Azul acinzentado
        "italic": True
    })

    worksheet.freeze_panes(1, 0)
    
    # Cabeçalho ElementId com cor especial
    header_elementid = workbook.add_format({
        "bold": True, 
        "bg_color": "#DCE6F1",  # Azul claro
        "font_color": "#1F4E78"  # Azul escuro
    })
    worksheet.write(0, 0, "ElementId", header_elementid)

    valid_params = []
    for param in selected_params:
        if unit_postfix_pattern.search(param.name):
            continue
        valid_params.append(param)

    for col_idx, param in enumerate(valid_params):
        postfix = ""
        header_format = bold
        
        # Cabeçalho com cores diferentes por tipo
        if param.storagetype == DB.StorageType.ElementId:
            header_format = workbook.add_format({"bold": True, "bg_color": "#FFBD80"})
        elif param.isreadonly:
            # CORREÇÃO: Cabeçalho vermelho para read-only
            header_format = workbook.add_format({
                "bold": True, 
                "bg_color": "#FFC7CE",  # Rosa claro
                "font_color": "#9C0006"  # Vermelho escuro
            })

        forge_type_id = get_parameter_data_type(param.definition)
        if forge_type_id and DB.UnitUtils.IsMeasurableSpec(forge_type_id):
            symbol_type_id = project_units.GetFormatOptions(forge_type_id).GetSymbolTypeId()
            if not symbol_type_id.Empty():
                symbol = DB.LabelUtils.GetLabelForSymbol(symbol_type_id)
                postfix = " [" + symbol + "]"

        worksheet.write(0, col_idx + 1, param.name + postfix, header_format)

    max_widths = [len("ElementId")] + [len(p.name) for p in valid_params]

    for row_idx, el in enumerate(src_elements, start=1):
        # ElementId sempre bloqueado com formato especial
        worksheet.write(row_idx, 0, str(get_elementid_value(el.Id)), locked_elementid)

        for col_idx, param in enumerate(valid_params):
            param_name = param.name
            param_val = el.LookupParameter(param_name)
            val = ""
            cell_format = unlocked  # Padrão: editável
            
            if param_val and param_val.HasValue:
                try:
                    if param_val.StorageType == DB.StorageType.Double:
                        forge_type_id = get_parameter_data_type(param.definition)
                        val = param_val.AsDouble()
                        if forge_type_id and DB.UnitUtils.IsMeasurableSpec(forge_type_id):
                            unit_type_id = param_val.GetUnitTypeId()
                            val = DB.UnitUtils.ConvertFromInternalUnits(param_val.AsDouble(), unit_type_id)
                    elif param_val.StorageType == DB.StorageType.String:
                        val = param_val.AsString()
                    elif param_val.StorageType == DB.StorageType.Integer:
                        if is_yesno_parameter(param.definition):
                            val = "Yes" if param_val.AsInteger() else "No"
                        else:
                            val = str(param_val.AsInteger())
                    elif param_val.StorageType == DB.StorageType.ElementId:
                        val = param_val.AsValueString()
                except:
                    val = "Error"
                
                # CORREÇÃO: Aplicar formato read-only se necessário
                if param.isreadonly:
                    cell_format = locked_readonly

            worksheet.write(row_idx, col_idx + 1, val, cell_format)
            length = len(str(val)) if val else 0
            if length > max_widths[col_idx + 1]:
                max_widths[col_idx + 1] = min(length, 50)

    for col_idx, width in enumerate(max_widths):
        worksheet.set_column(col_idx, col_idx, width + 3)

    worksheet.autofilter(0, 0, len(src_elements), len(valid_params))
    
    # CORREÇÃO: Remover proteção da planilha
    # worksheet.protect() estava bloqueando tudo
    # Agora a proteção vem apenas dos formatos locked/unlocked
    
    workbook.close()

# ==================== FUNÇÃO DE IMPORTAÇÃO ====================

def import_xls(file_path):
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_name("Export")
    except Exception as e:
        logger.error("Erro ao abrir Excel: {}".format(e))
        return

    headers = sheet.row_values(0)
    if not headers or headers[0] != "ElementId":
        logger.error("Primeira coluna deve ser 'ElementId'.")
        return

    param_names = headers[1:]

    with revit.Transaction("Import Parameters from Excel"):
        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)
            try:
                el_id_val = int(row[0])
                el = doc.GetElement(get_elementid_from_value(el_id_val))
                if not el: continue

                for col_idx, param_name in enumerate(param_names):
                    new_val = row[col_idx + 1]
                    if new_val in (None, ""): continue

                    param_name = unit_postfix_pattern.sub("", param_name).strip()
                    param = el.LookupParameter(param_name)
                    if not param or param.IsReadOnly: continue

                    try:
                        storage_type = param.StorageType
                        if storage_type == DB.StorageType.String:
                            param.Set(str(new_val))
                        elif storage_type == DB.StorageType.Integer:
                            if is_yesno_parameter(param.Definition):
                                str_val = str(new_val).strip().lower()
                                int_val = 1 if str_val in ("yes", "1", "true", "sim") else 0
                                param.Set(int_val)
                            else:
                                param.Set(int(float(new_val)))
                        elif storage_type == DB.StorageType.Double:
                            forge_type_id = get_parameter_data_type(param.Definition)
                            if forge_type_id and DB.UnitUtils.IsMeasurableSpec(forge_type_id):
                                unit_type_id = project_units.GetFormatOptions(forge_type_id).GetUnitTypeId()
                                new_val = DB.UnitUtils.ConvertToInternalUnits(float(new_val), unit_type_id)
                                param.Set(new_val)
                            else:
                                param.Set(float(new_val))
                    except Exception:
                        pass
            except Exception:
                pass

# ==================== UI CLASS ====================

class ExportImportWindow(forms.WPFWindow):
    def __init__(self, xaml_file_path):
        forms.WPFWindow.__init__(self, xaml_file_path)
        
        self.export_path = None
        self.import_path = None
        self.schedules = []
        
        # Eventos
        self.Button_BrowseExport.Click += self.browse_export
        self.Button_BrowseImport.Click += self.browse_import
        self.Button_Export.Click += self.do_export
        self.Button_Import.Click += self.do_import
        self.Button_Close.Click += self.close_clicked
        self.ComboBox_ExportMode.SelectionChanged += self.mode_changed
        
        # Inicialização
        self.load_schedules()
        self.mode_changed(None, None)
        self.update_status("Selecione uma operação acima para começar.")

    def update_status(self, message):
        self.TextBlock_Status.Text = message

    def load_schedules(self):
        """Carrega todas as tabelas do projeto para uso no dropdown."""
        all_schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
        self.schedules = []
        for s in all_schedules:
            if s.IsTemplate: continue
            if hasattr(s.Definition, "IsInternalKeynoteSchedule") and s.Definition.IsInternalKeynoteSchedule: continue
            if hasattr(s.Definition, "IsRevisionSchedule") and s.Definition.IsRevisionSchedule: continue
            self.schedules.append(s)
        self.schedules.sort(key=lambda x: x.Name)

    def mode_changed(self, sender, args):
        """Atualiza a UI baseada no modo selecionado."""
        mode_idx = self.ComboBox_ExportMode.SelectedIndex
        
        self.ComboBox_SubSelection.Items.Clear()
        
        if mode_idx == 0:
            self.ComboBox_SubSelection.IsEnabled = True
            if not self.schedules:
                self.ComboBox_SubSelection.Items.Add("Nenhuma tabela encontrada")
                self.ComboBox_SubSelection.SelectedIndex = 0
            else:
                for s in self.schedules:
                    self.ComboBox_SubSelection.Items.Add(s.Name)
                self.ComboBox_SubSelection.SelectedIndex = 0
        else:
            self.ComboBox_SubSelection.IsEnabled = False
            self.ComboBox_SubSelection.Items.Add("Seleção não necessária")
            self.ComboBox_SubSelection.SelectedIndex = 0

    def browse_export(self, sender, args):
        default_name = "Export_{}.xlsx".format(doc.Title.replace(".rvt", "").replace(" ", "_"))
        file_path = forms.save_file(file_ext="xlsx", default_name=default_name)
        if file_path:
            self.export_path = file_path
            self.TextBox_ExportPath.Text = file_path
            self.update_status("OK - Caminho definido: {}".format(os.path.basename(file_path)))

    def browse_import(self, sender, args):
        file_path = forms.pick_file(file_ext="xlsx", restore_dir=True)
        if file_path and os.path.exists(file_path):
            self.import_path = file_path
            self.TextBox_ImportPath.Text = file_path
            self.update_status("OK - Arquivo selecionado: {}".format(os.path.basename(file_path)))

    def do_export(self, sender, args):
        try:
            if not self.export_path:
                self.update_status("ERRO - Selecione onde salvar o arquivo.")
                return
            
            self.update_status("Preparando exportação...")
            mode_idx = self.ComboBox_ExportMode.SelectedIndex
            src_elements = []
            selected_params = []

            if mode_idx == 0:
                if not self.schedules:
                    self.update_status("ERRO - Nenhuma tabela disponível.")
                    return
                selected_schedule_idx = self.ComboBox_SubSelection.SelectedIndex
                if selected_schedule_idx < 0:
                    self.update_status("ERRO - Selecione uma tabela.")
                    return
                
                schedule = self.schedules[selected_schedule_idx]
                src_elements, selected_params = get_schedule_elements_and_params(schedule)

            elif mode_idx == 1:
                elements = revit.query.get_all_elements_in_view(active_view)
                src_elements = select_elements(elements)
                if src_elements: selected_params = select_parameters(src_elements)

            elif mode_idx == 2:
                elements = revit.query.get_all_elements(doc)
                type_filter = DB.ElementIsElementTypeFilter()
                elements = [el for el in elements if type_filter.PassesFilter(el)]
                src_elements = select_elements(elements)
                if src_elements: selected_params = select_parameters(src_elements)

            elif mode_idx == 3:
                elements = revit.query.get_all_elements(doc)
                type_filter = DB.ElementIsElementTypeFilter(True)
                elements = [el for el in elements if type_filter.PassesFilter(el)]
                src_elements = select_elements(elements)
                if src_elements: selected_params = select_parameters(src_elements)

            elif mode_idx == 4:
                elements = revit.get_selection()
                if not elements:
                    self.Hide()
                    with forms.WarningBar(title="Selecione elementos no Revit"):
                        elements = revit.pick_elements(message="Selecione elementos")
                    self.Show()
                
                if elements:
                    src_elements = select_elements(elements)
                    if src_elements: selected_params = select_parameters(src_elements)

            if not src_elements:
                self.update_status("Aviso - Nenhum elemento selecionado ou operação cancelada.")
                return

            export_xls(src_elements, selected_params, self.export_path)
            self.update_status("✅ Exportação concluída! {} elementos exportados.".format(len(src_elements)))

        except Exception as e:
            import traceback as tb
            self.update_status("❌ Erro: {}".format(str(e)))
            logger.error(tb.format_exc())

    def do_import(self, sender, args):
        try:
            if not self.import_path:
                self.update_status("❌ Selecione um arquivo Excel.")
                return
            
            self.update_status("⏳ Importando...")
            import_xls(self.import_path)
            self.update_status("✅ Importação concluída!")
        except Exception as e:
            import traceback as tb
            self.update_status("❌ Erro: {}".format(str(e)))
            logger.error(tb.format_exc())

    def close_clicked(self, sender, args):
        self.Close()

if __name__ == "__main__":
    try:
        xaml_file = script.get_bundle_file("Exportimport.xaml")
        if os.path.exists(xaml_file):
            window = ExportImportWindow(xaml_file)
            window.ShowDialog()
        else:
            forms.alert("Arquivo XAML não encontrado.")
    except Exception as e:
        logger.error(str(e))
