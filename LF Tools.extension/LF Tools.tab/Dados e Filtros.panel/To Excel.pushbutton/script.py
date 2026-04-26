# -*- coding: utf-8 -*-
import os
import xlsxwriter
import zipfile
import xml.etree.ElementTree as ET

# ==================== LEITOR XLSX EMBUTIDO (sem openpyxl) ====================
# .xlsx é um ZIP de XMLs. Lemos com bibliotecas padrão do Python.
_XLSX_NS = {
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

def _col_letter_to_index(col_str):
    """Converte letra de coluna Excel (A, B, ..., AA, AB) para indice 0-based."""
    result = 0
    for ch in col_str:
        result = result * 26 + (ord(ch.upper()) - ord('A') + 1)
    return result - 1

def _parse_cell_ref(ref):
    """Extrai (col_index, row_number) de referencia como 'B3'."""
    col_str = ''
    row_str = ''
    for ch in ref:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    return _col_letter_to_index(col_str), int(row_str)

class _XlsxReader:
    """Leitor leve de .xlsx usando apenas zipfile + ElementTree."""
    def __init__(self, file_path):
        self._zf = zipfile.ZipFile(file_path, 'r')
        self._shared_strings = self._load_shared_strings()
        self._sheet_names, self._sheet_paths = self._load_workbook_info()

    def _load_shared_strings(self):
        strings = []
        try:
            xml_data = self._zf.read('xl/sharedStrings.xml')
            root = ET.fromstring(xml_data)
            for si in root.findall('ss:si', _XLSX_NS):
                # Texto pode estar em <t> direto ou em múltiplos <r><t>
                parts = []
                t_elem = si.find('ss:t', _XLSX_NS)
                if t_elem is not None and t_elem.text:
                    parts.append(t_elem.text)
                else:
                    for r_elem in si.findall('ss:r', _XLSX_NS):
                        rt = r_elem.find('ss:t', _XLSX_NS)
                        if rt is not None and rt.text:
                            parts.append(rt.text)
                strings.append(''.join(parts))
        except (KeyError, ET.ParseError):
            pass
        return strings

    def _load_workbook_info(self):
        # Ler nomes das abas
        wb_xml = self._zf.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)
        sheets_elem = wb_root.find('ss:sheets', _XLSX_NS)
        sheet_entries = []  # (name, rId)
        if sheets_elem is not None:
            for s in sheets_elem.findall('ss:sheet', _XLSX_NS):
                name = s.get('name', '')
                rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
                sheet_entries.append((name, rid))

        # Mapear rId -> caminho do arquivo
        rels_xml = self._zf.read('xl/_rels/workbook.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        rid_to_path = {}
        for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rid_to_path[rel.get('Id', '')] = 'xl/' + rel.get('Target', '')

        names = []
        paths = []
        for name, rid in sheet_entries:
            names.append(name)
            paths.append(rid_to_path.get(rid, ''))
        return names, paths

    @property
    def sheetnames(self):
        return list(self._sheet_names)

    def read_sheet(self, sheet_name):
        """Retorna lista de linhas (cada linha é uma lista de valores)."""
        try:
            idx = self._sheet_names.index(sheet_name)
        except ValueError:
            return []
        path = self._sheet_paths[idx]
        if not path:
            return []

        try:
            xml_data = self._zf.read(path)
        except KeyError:
            return []

        root = ET.fromstring(xml_data)
        rows_dict = {}  # row_num -> {col_idx: value}
        max_col = 0

        for row_elem in root.findall('.//ss:row', _XLSX_NS):
            row_num = int(row_elem.get('r', '0'))
            if row_num == 0:
                continue
            cells = {}
            for cell in row_elem.findall('ss:c', _XLSX_NS):
                ref = cell.get('r', '')
                if not ref:
                    continue
                col_idx, _ = _parse_cell_ref(ref)
                if col_idx > max_col:
                    max_col = col_idx

                cell_type = cell.get('t', '')
                v_elem = cell.find('ss:v', _XLSX_NS)
                value = None

                if v_elem is not None and v_elem.text is not None:
                    raw = v_elem.text
                    if cell_type == 's':  # shared string
                        try:
                            value = self._shared_strings[int(raw)]
                        except (IndexError, ValueError):
                            value = raw
                    elif cell_type == 'b':  # boolean
                        value = 'Yes' if raw == '1' else 'No'
                    elif cell_type == 'inlineStr':
                        is_elem = cell.find('ss:is/ss:t', _XLSX_NS)
                        value = is_elem.text if is_elem is not None else raw
                    else:  # number
                        try:
                            fval = float(raw)
                            value = int(fval) if fval == int(fval) else fval
                        except ValueError:
                            value = raw
                else:
                    # Inline string sem <v>
                    is_elem = cell.find('ss:is/ss:t', _XLSX_NS)
                    if is_elem is not None:
                        value = is_elem.text

                cells[col_idx] = value
            rows_dict[row_num] = cells

        if not rows_dict:
            return []

        # Montar lista de linhas densas
        max_row = max(rows_dict.keys())
        num_cols = max_col + 1
        result = []
        for r in range(1, max_row + 1):
            row_cells = rows_dict.get(r, {})
            row = [row_cells.get(c) for c in range(num_cols)]
            result.append(tuple(row))
        return result

    def close(self):
        try:
            self._zf.close()
        except:
            pass
import io
import re
import traceback
import gc

from collections import namedtuple
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("System.Data")
clr.AddReference("System.Xml")
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB
import System
from System.Windows import Visibility
from System.Windows.Input import Cursors
from System.Data import DataTable
import System.Windows.Controls as Controls

# ==================== CPYTHON COMPAT ====================
try:
    clr.AddReference('System.Windows.Forms')
except Exception:
    pass
import System.Windows.Forms as _WF
import System.Xml
from System.IO import StringReader
from System.Windows.Markup import XamlReader
import re as _re

def _alert(msg, title="LF Tools", yes=False, no=False, exitscript=False, **kw):
    if yes and no:
        r = _WF.MessageBox.Show(str(msg), str(title), _WF.MessageBoxButtons.YesNo)
        ans = r == _WF.DialogResult.Yes
        if exitscript and not ans:
            raise SystemExit()
        return ans
    _WF.MessageBox.Show(str(msg), str(title))
    if exitscript:
        raise SystemExit()

def _pick_file(file_ext="xlsx", **kw):
    dlg = _WF.OpenFileDialog()
    dlg.Filter = "{1} (*.{0})|*.{0}|All files (*.*)|*.*".format(file_ext, file_ext.upper())
    return dlg.FileName if dlg.ShowDialog() == _WF.DialogResult.OK else None

def _save_file(file_ext="xlsx", default_name="export", **kw):
    dlg = _WF.SaveFileDialog()
    dlg.Filter = "{1} (*.{0})|*.{0}".format(file_ext, file_ext.upper())
    dlg.FileName = str(default_name)
    return dlg.FileName if dlg.ShowDialog() == _WF.DialogResult.OK else None

class _WPFWindowCPy(object):
    """CPython drop-in para pyrevit.forms.WPFWindow."""
    _XAML_EVENTS = _re.compile(
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
            with io.open(str(xaml_source), 'r', encoding='utf-8') as _f:
                stripped = _f.read().strip()
        xaml_clean = self._XAML_EVENTS.sub('', stripped)
        rdr = System.Xml.XmlReader.Create(StringReader(xaml_clean))
        self._window = XamlReader.Load(rdr)
        # Vincula ao Revit para herdar ícone e ficar na taskbar
        try:
            from System.Windows.Interop import WindowInteropHelper
            from System.Diagnostics import Process
            WindowInteropHelper(self._window).Owner = Process.GetCurrentProcess().MainWindowHandle
        except Exception:
            pass

    def __getattr__(self, name):
        if name.startswith('_'):
            raise AttributeError(name)
        win = object.__getattribute__(self, '_window')
        el = win.FindName(name)
        if el is not None:
            return el
        import System.Windows
        el = System.Windows.LogicalTreeHelper.FindLogicalNode(win, name)
        if el is not None:
            return el
        return getattr(win, name)

    def ShowDialog(self): return self._window.ShowDialog()
    def Show(self): return self._window.Show()
    def Close(self): self._window.Close()

class _FormsStub: pass
forms = _FormsStub()
forms.WPFWindow  = _WPFWindowCPy
forms.alert      = _alert
forms.pick_file  = _pick_file
forms.save_file  = _save_file

class _ScriptStub:
    @staticmethod
    def get_bundle_file(name):
        try:
            return os.path.join(__commandpath__, name)
        except NameError:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
script = _ScriptStub()

class _CoreutilsStub:
    @staticmethod
    def get_enum_none(enum_type):
        return getattr(enum_type, 'None', 0)
coreutils = _CoreutilsStub()
# ==================== FIM CPYTHON COMPAT ====================

# Documento
import logging as _logging
logger = _logging.getLogger(__name__)
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
project_units = doc.GetUnits()

unit_postfix_pattern = re.compile(r"\s*\[.*\]$")

# ==================== CACHE SIMPLES E OTIMIZADO ====================
# Substituindo implementacoes complexas por dicts padrao do Python (muito mais rapidos)
_param_def_cache = {}
_element_cache = {}
_format_pool = {}

def setup_memory_optimization():
    gc.enable()

def smart_gc_collect(force=False, threshold=50000):
    if force:
        try:
            gc.collect()
        except:
            pass
    return False

def clear_all_caches():
    try:
        _param_def_cache.clear()
        _element_cache.clear()
        _format_pool.clear()
    except:
        pass

# ==================== ESTRUTURAS DE DADOS OTIMIZADAS ====================
ParamDef = namedtuple(
    "ParamDef", ["name", "istype", "definition", "isreadonly", "isunit", "storagetype"]
)

def sanitize_filename(name):
    """Remove caracteres invalidos para nome de arquivo e abas Excel."""
    return re.sub(r'[\\/*?:"<>|\[\]]', "", name)

def get_parameter_data_type(param_def):
    try: 
        return param_def.GetDataType()
    except AttributeError: 
        return None

def is_yesno_parameter(param_def):
    try:
        return param_def.GetDataType() == DB.SpecTypeId.Boolean.YesNo
    except:
        try: 
            return param_def.ParameterType == DB.ParameterType.YesNo
        except: 
            return False

# ==================== POOL DE FORMATOS EXCEL ====================
def get_excel_format(workbook, properties):
    """Obtém ou cria formato Excel do pool para reutilização."""
    # Cria chave única baseada nas propriedades
    key_parts = []
    for k, v in sorted(properties.items()):
        key_parts.append("{}:{}".format(k, v))
    key = "|".join(key_parts)
    
    if key not in _format_pool:
        _format_pool[key] = workbook.add_format(properties)
    
    return _format_pool[key]

# ==================== FUNÇÕES DE EXTRAÇÃO OTIMIZADAS ====================
def get_panel_schedule_data(schedule_view):
    """Extrai dados dos paineis com cache seguro."""
    cache_key = ("panel_schedule", schedule_view.Id.IntegerValue)
    cached = _param_def_cache.get(cache_key)
    if cached:
        return cached
    
    try:
        panel_id = None
        try: 
            panel_id = schedule_view.GetPanel()
        except: 
            return [], []

        if not panel_id or panel_id == DB.ElementId.InvalidElementId: 
            return [], []
        
        panel = doc.GetElement(panel_id)
        if not panel: 
            return [], []

        panel_name = panel.Name
        
        # Coletar sistemas
        assigned_systems = []
        try:
            if hasattr(panel, "MEPModel") and panel.MEPModel:
                assigned_systems = list(panel.MEPModel.GetAssignedElectricalSystems())
            else:
                assigned_systems = list(panel.GetAssignedElectricalSystems())
        except: 
            return [], []

        if not assigned_systems:
            dummy = {'Tipo': 'PAINEL', 'Nome do Quadro': panel_name, 'Descricao': 'Sem circuitos'}
            result = [dummy], ['Tipo', 'Nome do Quadro', 'Descricao']
            _param_def_cache[cache_key] = result
            return result

        headers = ['Nome do Quadro', 'Numero do Circuito', 'Nome da Carga', 'Classificacao (A)', 
                  'Tensao (V)', 'Polos', 'Carga Aparente (VA)', 'Fio']
        
        # Mapear BuiltInParameters
        param_map = {
            'Numero do Circuito': DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
            'Nome da Carga': DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
            'Classificacao (A)': DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
            'Tensao (V)': DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
            'Polos': DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
            'Carga Aparente (VA)': DB.BuiltInParameter.RBS_ELEC_APPARENT_LOAD,
            'Fio': DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM,
        }
        
        data_rows = []
        for sys in assigned_systems:
            try:
                row_data = {'Tipo': 'CIRCUITO', 'Nome do Quadro': panel_name}
                
                for header, builtin_param in param_map.items():
                    try:
                        p = sys.get_Parameter(builtin_param)
                        if p and p.HasValue:
                            row_data[header] = p.AsValueString() or p.AsString() or ""
                        else:
                            row_data[header] = ""
                    except:
                        row_data[header] = ""
                
                data_rows.append(row_data)
            except: 
                continue

        # Ordenar
        try:
            data_rows.sort(key=lambda x: int(re.sub(r"\D", "", x.get('Numero do Circuito', '9999')) or '9999'))
        except: 
            pass

        result = data_rows, headers
        _param_def_cache[cache_key] = result
        return result
        
    except Exception as e:
        logger.error("Erro em get_panel_schedule_data: " + str(e))
        return [], []
    finally:
        # Sem limpeza manual de loop para performance
        pass

def get_schedule_elements_and_params(schedule):
    """Extrai elementos e parametros com cache otimizado."""
    cache_key = ("schedule_elements", schedule.Id.IntegerValue)
    cached = _param_def_cache.get(cache_key)
    if cached:
        return cached
    
    try:
        schedule_def = schedule.Definition
    except: 
        return [], []
    
    # Coletar campos visíveis na ORDEM DO SCHEDULE
    visible_fields = []
    
    try: 
        field_order = schedule_def.GetFieldOrder()
    except: 
        return [], []
    
    for field_id in field_order:
        field = schedule_def.GetField(field_id)
        if not field.IsHidden and not field.IsCalculatedField:
            visible_fields.append(field)
    
    if not visible_fields:
        return [], []
    
    # Coletar elementos do schedule - método otimizado
    elements = []
    element_ids = set()
    
    # Usar filtro otimizado
    try:
        # Método 1: GetScheduleElements (mais rápido)
        element_ids_set = schedule.GetScheduleElements()
        if element_ids_set:
            for element_id in element_ids_set:
                if element_id and element_id != DB.ElementId.InvalidElementId:
                    cached_element = _element_cache.get(element_id.IntegerValue)
                    if cached_element:
                        element = cached_element
                    else:
                        element = doc.GetElement(element_id)
                        if element:
                            _element_cache[element_id.IntegerValue] = element
                    
                    if element and element.Id not in element_ids:
                        element_ids.add(element.Id)
                        elements.append(element)
    except:
        pass
    
    # Fallback: FilteredElementCollector otimizado
    # Nota: NAO usar WhereElementIsNotElementType() pois schedules de tipo
    # contem ElementTypes, e esse filtro os removeria, gerando tabelas vazias.
    if not elements:
        try:
            collector = DB.FilteredElementCollector(doc, schedule.Id)

            for el in collector:
                try:
                    if el.Id not in element_ids:
                        element_ids.add(el.Id)
                        elements.append(el)
                        _element_cache[el.Id.IntegerValue] = el
                except:
                    continue
        except:
            pass
    
    if not elements: 
        return [], []
    
    # Processar parâmetros com cache
    param_defs_list = []
    non_storage_type = coreutils.get_enum_none(DB.StorageType)
    
    schedule_id_val = schedule.Id.IntegerValue

    for field in visible_fields:
        field_name = field.GetName()
        if not field_name:
            continue

        # Chave inclui schedule ID para evitar contaminação entre schedules
        # com campos de mesmo nome mas parâmetros diferentes
        param_cache_key = ("param_def", schedule_id_val, field_name)
        cached_param_def = _param_def_cache.get(param_cache_key)

        if cached_param_def:
            param_defs_list.append(cached_param_def)
            continue

        param_id = field.ParameterId
        param_info = None

        # Tentar múltiplos elementos para encontrar a definição do parâmetro.
        # Usando mais amostras pois alguns parâmetros só existem em certos elementos.
        sample_elements = elements[:min(15, len(elements))]
        
        for sample_el in sample_elements:
            if param_info:
                break

            if param_id and param_id != DB.ElementId.InvalidElementId:
                try:
                    param_element = doc.GetElement(param_id)
                    if param_element and hasattr(param_element, 'Definition'):
                        param_definition = param_element.Definition
                        param_info = sample_el.get_Parameter(param_definition)
                except:
                    pass

            if not param_info:
                param_info = sample_el.LookupParameter(field_name)

            if not param_info and hasattr(sample_el, 'Parameters'):
                for p in getattr(sample_el, 'Parameters', []):
                    try:
                        if p.Definition.Name == field_name:
                            param_info = p
                            break
                    except:
                        continue

            # Fallback: buscar no tipo do elemento (parâmetros de tipo aparecem
            # em schedules de instância)
            if not param_info:
                try:
                    type_id = sample_el.GetTypeId()
                    if type_id and type_id != DB.ElementId.InvalidElementId:
                        el_type = doc.GetElement(type_id)
                        if el_type:
                            param_info = el_type.LookupParameter(field_name)
                except:
                    pass
        
        # Criar definição
        if param_info:
            dt = get_parameter_data_type(param_info.Definition) if hasattr(param_info, 'Definition') else None
            param_def = ParamDef(
                name=field_name, 
                istype=False, 
                definition=param_info.Definition if hasattr(param_info, 'Definition') else None, 
                isreadonly=param_info.IsReadOnly if hasattr(param_info, 'IsReadOnly') else True,
                isunit=DB.UnitUtils.IsMeasurableSpec(dt) if dt else False, 
                storagetype=param_info.StorageType if hasattr(param_info, 'StorageType') else DB.StorageType.String,
            )
        else:
            param_def = ParamDef(
                name=field_name, 
                istype=False, 
                definition=None, 
                isreadonly=True,
                isunit=False, 
                storagetype=DB.StorageType.String,
            )
        
        param_defs_list.append(param_def)
        _param_def_cache[param_cache_key] = param_def
    
    result = elements, param_defs_list
    _param_def_cache[cache_key] = result
    
    # Sem GC forçado aqui para evitar travamentos
    if len(elements) > 2000:
        pass
    
    return result

def get_element_parameter_value(el, param_def, param_cache=None):
    """Obtém valor de parâmetro com cache simples."""
    if param_cache is None:
        param_cache = {}
    
    param_name = param_def.name
    cache_key = (el.Id.IntegerValue, param_name)
    
    # Cache L1: Por elemento/parâmetro
    cached_value = param_cache.get(cache_key)
    if cached_value is not None:
        return cached_value
    
    value = ""
    
    # Cache L2: Elemento já carregado?
    element_key = el.Id.IntegerValue
    cached_element = _element_cache.get(element_key)
    if not cached_element:
        _element_cache[element_key] = el
    
    # Extrair valor — tentar múltiplas estratégias
    if param_def.definition and hasattr(el, 'get_Parameter'):
        try:
            param = el.get_Parameter(param_def.definition)
            if param and param.HasValue:
                value = get_parameter_value(param, param_def)
        except:
            pass
    
    if not value and hasattr(el, 'LookupParameter'):
        try:
            param = el.LookupParameter(param_name)
            if param and param.HasValue:
                value = get_parameter_value(param, param_def)
        except:
            pass
    
    if not value and hasattr(el, 'Parameters'):
        for param in el.Parameters:
            try:
                if param.Definition.Name == param_name:
                    if param.HasValue:
                        value = get_parameter_value(param, param_def)
                    break
            except:
                continue
    
    # Fallback: tentar parâmetro no Type
    if not value:
        try:
            type_id = el.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                el_type = doc.GetElement(type_id)
                if el_type:
                    tp = el_type.LookupParameter(param_name)
                    if tp and tp.HasValue:
                        value = get_parameter_value(tp, param_def)
        except:
            pass
    
    # Armazenar em cache
    param_cache[cache_key] = value
    
    # Limpeza periódica do cache — threshold maior para evitar perda
    if len(param_cache) > 10000:
        param_cache.clear()
    
    return value

def get_parameter_value(param, param_def=None):
    """Obtém valor formatado de parâmetro otimizado.
    Usa AsValueString como fallback primário para garantir que valores
    formatados pelo Revit sejam capturados mesmo quando a conversão manual falha.
    """
    try:
        st = param.StorageType
        if st == DB.StorageType.Double:
            # Retorna float para que o Excel receba número, não texto.
            # Isso permite ordenação, soma e cálculos na planilha.
            try:
                dt = get_parameter_data_type(param.Definition)
                if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                    uid = param.GetUnitTypeId()
                    return DB.UnitUtils.ConvertFromInternalUnits(param.AsDouble(), uid)
            except:
                pass
            # Fallback: valor bruto como float (unidades internas do Revit)
            return param.AsDouble()
        elif st == DB.StorageType.String: 
            return param.AsString() or ""
        elif st == DB.StorageType.Integer:
            if is_yesno_parameter(param.Definition): 
                return "Yes" if param.AsInteger() else "No"
            else:
                # Tentar AsValueString para enums/listas
                try:
                    vs = param.AsValueString()
                    if vs:
                        return vs
                except:
                    pass
                return str(param.AsInteger())
        elif st == DB.StorageType.ElementId: 
            return param.AsValueString() or ""
    except Exception as e:
        # Fallback universal: AsValueString
        try:
            vs = param.AsValueString()
            if vs:
                return vs
        except:
            pass
        logger.debug("get_parameter_value falhou: {}".format(e))
        return ""
    
    return ""

def get_param_display_string(el, param_def):
    """Retorna o valor de um parâmetro como string de exibição (igual ao Revit)."""
    try:
        param = None
        if param_def.definition and hasattr(el, 'get_Parameter'):
            try:
                param = el.get_Parameter(param_def.definition)
            except Exception:
                pass
        if not param and hasattr(el, 'LookupParameter'):
            try:
                param = el.LookupParameter(param_def.name)
            except Exception:
                pass
        if param and param.HasValue:
            try:
                vs = param.AsValueString()
                if vs is not None:
                    return vs
            except Exception:
                pass
            try:
                return param.AsString() or ""
            except Exception:
                pass
    except Exception:
        pass
    return ""


# ==================== EXPORTAÇÃO OTIMIZADA ====================
def export_xls(targets, file_path, formatted=False):
    """Exporta dados para Excel com múltiplas abas se necessário."""
    workbook = None
    try:
        # Configurar workbook com otimizações
        workbook_options = {
            'constant_memory': True,
            'use_zip64': True,
        }
        
        workbook = xlsxwriter.Workbook(file_path, workbook_options)
        
        # Pool de formatos globais (reutilizados entre abas)
        fmt_head_panel = get_excel_format(workbook, {
            "bold": True, 
            "bg_color": "#4F81BD", 
            "font_color": "white", 
            "border": 1, 
            "align": "center"
        })
        fmt_data_panel = get_excel_format(workbook, {"border": 1, "align": "left"})
        
        fmt_bold = get_excel_format(workbook, {"bold": True})
        fmt_lock_ro = get_excel_format(workbook, {
            "locked": True, 
            "font_color": "#C0504D", 
            "italic": True
        })
        fmt_lock_id = get_excel_format(workbook, {
            "locked": True, 
            "font_color": "#95B3D7", 
            "italic": True
        })
        fmt_unlock = get_excel_format(workbook, {"locked": False})
        fmt_head_id = get_excel_format(workbook, {
            "bold": True, 
            "bg_color": "#DCE6F1", 
            "font_color": "#1F4E78"
        })
        fmt_head_ro = get_excel_format(workbook, {
            "bold": True, 
            "bg_color": "#FFC7CE", 
            "font_color": "#9C0006"
        })

        # Formatos extras para modo formatado (visual)
        if formatted:
            fmt_head_vis = get_excel_format(workbook, {
                "bold": True, "bg_color": "#1F4E78", "font_color": "white",
                "border": 1, "align": "center", "valign": "vcenter",
            })
            fmt_data_vis = get_excel_format(workbook, {
                "border": 1, "align": "left", "valign": "vcenter",
            })
            fmt_data_alt = get_excel_format(workbook, {
                "border": 1, "align": "left", "valign": "vcenter",
                "bg_color": "#EBF3FB",
            })

        for target in targets:
            sheet_name = sanitize_filename(target['name'])[:31] # Excel limita a 31 chars
            src_elements = target['src']
            selected_params = target['params']
            is_panel_schedule = target.get('is_panel', False)

            ws = workbook.add_worksheet(sheet_name)

            if is_panel_schedule:
                # Escrever headers
                for i, p in enumerate(selected_params):
                    ws.write(0, i, p.name, fmt_head_panel)
                    ws.set_column(i, i, 20)

                # Escrever dados
                for r, el in enumerate(src_elements, 1):
                    if hasattr(el, 'row_data'):
                        for c, p in enumerate(selected_params):
                            ws.write(r, c, str(el.row_data.get(p.name, "")), fmt_data_panel)

                if src_elements:
                    ws.autofilter(0, 0, len(src_elements), len(selected_params)-1)

            elif formatted:
                # ── Modo formatado: visual idêntico ao Revit, sem ElementId ─────
                ws.set_tab_color("#1F4E78")
                ws.freeze_panes(1, 0)
                ws.set_row(0, 18)

                header_names = [p.name for p in selected_params]
                widths = [len(n) for n in header_names]

                for i, name in enumerate(header_names):
                    ws.write(0, i, name, fmt_head_vis)

                for r, el in enumerate(src_elements, 1):
                    fmt_row = fmt_data_alt if r % 2 == 0 else fmt_data_vis
                    try:
                        for c, p in enumerate(selected_params):
                            if hasattr(el, 'row_data'):
                                value = str(el.row_data.get(p.name, ""))
                            else:
                                value = get_param_display_string(el, p)
                            ws.write(r, c, value, fmt_row)
                            slen = len(value) if value else 0
                            if slen > widths[c]:
                                widths[c] = min(slen, 60)
                    except Exception:
                        pass

                for i, w in enumerate(widths):
                    ws.set_column(i, i, w + 3)

                if src_elements:
                    ws.autofilter(0, 0, len(src_elements), len(selected_params) - 1)

            else:
                # ── Modo padrão: com ElementId, editável/importável ───────────
                ws.freeze_panes(1, 1)
                ws.write(0, 0, "ElementId", fmt_head_id)

                # Processar cabeçalhos
                param_units = []
                header_names = []

                for p in selected_params:
                    post = ""
                    header_name = p.name

                    if not unit_postfix_pattern.search(header_name):
                        dt = get_parameter_data_type(p.definition) if p.definition else None
                        if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                            try:
                                sym = project_units.GetFormatOptions(dt).GetSymbolTypeId()
                                if not sym.Empty():
                                    post = " [" + DB.LabelUtils.GetLabelForSymbol(sym) + "]"
                            except:
                                pass

                    param_units.append(post)
                    header_names.append(header_name)

                # Escrever headers
                for i, p in enumerate(selected_params):
                    fmt = fmt_head_ro if p.isreadonly else fmt_bold
                    ws.write(0, i+1, header_names[i] + param_units[i], fmt)

                # Calcular larguras
                widths = [len("ElementId")] + [len(header_names[i]) + len(param_units[i]) for i in range(len(selected_params))]

                # Processamento em lotes adaptativo
                total_elements = len(src_elements)
                batch_size = 1000

                param_cache = {}

                for batch_start in range(0, total_elements, batch_size):
                    batch_end = min(batch_start + batch_size, total_elements)

                    r_offset = 0
                    for el in src_elements[batch_start:batch_end]:
                        r = batch_start + r_offset + 1
                        try:
                            eid = el.Id.IntegerValue
                            ws.write(r, 0, str(eid), fmt_lock_id)

                            for c, p in enumerate(selected_params):
                                value = get_element_parameter_value(el, p, param_cache)
                                fmt = fmt_lock_ro if p.isreadonly else fmt_unlock
                                ws.write(r, c+1, value, fmt)

                                slen = len(str(value)) if value else 0
                                if slen > widths[c+1]:
                                    widths[c+1] = min(slen, 50)
                        except:
                            pass
                        r_offset += 1

                    if len(param_cache) > 5000:
                        param_cache.clear()

                # Coluna ElementId oculta — dados preservados para importação
                ws.set_column(0, 0, None, None, {'hidden': True})
                for i, w in enumerate(widths):
                    if i == 0:
                        continue
                    ws.set_column(i, i, w + 3)
                
                if total_elements > 0: 
                    ws.autofilter(0, 0, total_elements, len(selected_params))

    except Exception as e:
        logger.error("Erro Excel: " + str(e))
        if workbook:
            try:
                workbook.close()
            except:
                pass
        raise e
    finally:
        if workbook:
            try:
                workbook.close()
            except Exception as e:
                msg = str(e)
                if "Permission denied" in msg or "being used" in msg:
                    forms.alert("O arquivo Excel esta ABERTO.\nFeche-o e tente novamente.", 
                               title="Arquivo em Uso")
                else:
                    logger.error("Erro ao fechar workbook: " + str(e))
        
        # Limpeza pós-exportação
        clear_all_caches()

# ==================== IMPORTAÇÃO OTIMIZADA ====================
def get_param_robust(element, param_name):
    """Busca parametro na instancia ou no tipo."""
    # 1. Tenta na instancia
    p = element.LookupParameter(param_name)
    if p: return p
    
    # 2. Tenta no tipo
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != DB.ElementId.InvalidElementId:
            el_type = element.Document.GetElement(type_id)
            if el_type:
                p = el_type.LookupParameter(param_name)
                if p: return p
    except:
        pass
        
    return None

def import_xls(file_path):
    """Importa dados do Excel usando leitor embutido (sem openpyxl)."""
    try:
        wb = _XlsxReader(file_path)
    except Exception as e:
        forms.alert("Erro ao abrir Excel: " + str(e), title="Erro")
        return

    try:
        # Identifica abas válidas: com "ElementId" na célula A1
        sheets_to_process = []
        for s_name in wb.sheetnames:
            if s_name.startswith('_'):
                continue
            rows = wb.read_sheet(s_name)
            if rows and str(rows[0][0] if rows[0] else "").strip() == "ElementId":
                sheets_to_process.append((s_name, rows))

        if not sheets_to_process:
            forms.alert(
                "Nenhuma aba com formato valido (ElementId na celula A1) encontrada.\n"
                "Quadros de Cargas exportados nao podem ser importados.",
                title="Aviso"
            )
            return

        stats = {
            'elements_processed': 0, 'success_cells': 0, 'skipped_cells': 0,
            'readonly_cells': 0, 'not_found_cells': 0, 'error_cells': 0
        }

        _t = DB.Transaction(doc, "Import Excel (Multi)")
        _t.Start()
        try:
            import_element_cache = {}

            for s_name, rows in sheets_to_process:
                if not rows:
                    continue

                headers = [str(h).strip() if h is not None else "" for h in rows[0]]
                pnames = [unit_postfix_pattern.sub("", h).strip() for h in headers[1:]]

                for row in rows[1:]:
                    try:
                        if row[0] is None:
                            continue
                        eid = int(float(row[0]))

                        el = import_element_cache.get(eid)
                        if not el:
                            el = doc.GetElement(DB.ElementId(eid))
                            if el:
                                import_element_cache[eid] = el

                        if not el:
                            continue
                        stats['elements_processed'] += 1

                        for cx, pname in enumerate(pnames):
                            val = row[cx + 1] if cx + 1 < len(row) else None
                            if val is None or val == "":
                                continue

                            param = get_param_robust(el, pname)
                            if not param:
                                stats['not_found_cells'] += 1
                                continue
                            if param.IsReadOnly:
                                stats['readonly_cells'] += 1
                                continue

                            try:
                                st = param.StorageType
                                success = False
                                changed = False

                                if st == DB.StorageType.String:
                                    current_val = param.AsString() or ""
                                    if isinstance(val, float) and val == int(val):
                                        new_val = str(int(val))
                                    else:
                                        new_val = str(val)
                                    if current_val != new_val:
                                        success = param.Set(new_val)
                                        changed = True

                                elif st == DB.StorageType.Integer:
                                    current_val = param.AsInteger()
                                    if is_yesno_parameter(param.Definition):
                                        sval = str(val).strip().lower()
                                        new_val = 1 if sval in ("yes", "1", "true", "sim") else 0
                                    else:
                                        new_val = int(float(val))
                                    if current_val != new_val:
                                        success = param.Set(new_val)
                                        changed = True

                                elif st == DB.StorageType.Double:
                                    current_val = param.AsDouble()
                                    dt = get_parameter_data_type(param.Definition)
                                    new_val = float(val)
                                    if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                                        uid = param.GetUnitTypeId()
                                        new_val_internal = DB.UnitUtils.ConvertToInternalUnits(new_val, uid)
                                        if abs(current_val - new_val_internal) > 0.0001:
                                            success = param.Set(new_val_internal)
                                            changed = True
                                    else:
                                        if abs(current_val - new_val) > 0.0001:
                                            success = param.Set(new_val)
                                            changed = True

                                if changed:
                                    if success:
                                        stats['success_cells'] += 1
                                    else:
                                        stats['error_cells'] += 1
                                else:
                                    stats['skipped_cells'] += 1

                            except:
                                stats['error_cells'] += 1

                    except:
                        pass

                if len(import_element_cache) > 2000:
                    import_element_cache.clear()

            import_element_cache.clear()
            _t.Commit()
        except Exception:
            if _t.HasStarted() and not _t.HasEnded():
                _t.RollBack()
            raise

    finally:
        try:
            wb.close()
        except:
            pass
        clear_all_caches()
    
    # Relatorio Final
    report = "Importacao Concluida!\n\n"
    report += "Elementos processados: {}\n".format(stats['elements_processed'])
    report += "Celulas atualizadas: {}\n".format(stats['success_cells'])
    
    if stats['skipped_cells'] > 0:
        report += "Celulas inalteradas (valor igual): {}\n".format(stats['skipped_cells'])
        
    if stats['readonly_cells'] > 0:
        report += "Celulas somente-leitura (ignoradas): {}\n".format(stats['readonly_cells'])
        
    if stats['not_found_cells'] > 0:
        report += "Parametros nao encontrados: {}\n".format(stats['not_found_cells'])
        
    if stats['error_cells'] > 0:
        report += "Erros de conversao: {}\n".format(stats['error_cells'])
        
    forms.alert(report, title="Relatorio de Importacao")

# ==================== UI OTIMIZADA ====================
class ExportImportWindow(forms.WPFWindow):
    def __init__(self, xaml_file_path):
        forms.WPFWindow.__init__(self, xaml_file_path)
        setup_memory_optimization()

        self.export_path = None
        self.import_path = None
        self.last_export_folder = None
        self.view_schedules = []
        self.panel_schedules = []
        self._schedule_checkboxes = []  # lista de (CheckBox, schedule_dict)
        self._dropdown_open = False

        # Eventos
        self.Button_BrowseExport.Click      += self.browse_export
        self.Button_BrowseImport.Click      += self.browse_import
        self.Button_Export.Click            += self.do_export
        self.Button_Import.Click            += self.do_import
        self.Button_Close.Click             += self.close_clicked
        self.ComboBox_ExportMode.SelectionChanged += self.mode_changed
        self.btn_DropdownToggle.MouseLeftButtonDown += self._toggle_dropdown
        self.chk_SelectAllSchedules.Click   += self._on_select_all_changed
        self.CheckBox_KeepFormat.Click      += self._on_keep_format_changed
        
        self.cmb_ExportPreviewSelect.SelectionChanged += lambda s, e: self._update_export_preview(from_combo=True)
        self.cmb_ImportPreviewSelect.SelectionChanged += lambda s, e: self._update_import_preview(from_combo=True)

        self.load_schedules()
        self.mode_changed(None, None)
    
    def update_status(self, msg): 
        self.TextBlock_Status.Text = msg

    def update_stats(self, msg):
        try:
            self.TextBlock_Stats.Text = msg
        except:
            pass

    def _build_suggested_name(self):
        """Gera nome sugerido baseado na seleção atual."""
        proj_name = sanitize_filename(doc.Title.replace(".rvt", ""))
        selected = self._get_selected_schedules()
        if len(selected) == 1:
            return proj_name + "_" + sanitize_filename(selected[0]['display_name'])
        elif len(selected) > 1:
            return proj_name + "_Multiplas_{}_Abas".format(len(selected))
        return "Export_" + proj_name

    def _update_export_path(self):
        """Atualiza caminho de exportacao mantendo a pasta escolhida."""
        if not self.last_export_folder:
            return
        name = self._build_suggested_name() + ".xlsx"
        self.export_path = os.path.join(self.last_export_folder, name)
        self.TextBox_ExportPath.Text = self.export_path
        self.update_status("Salvar em: " + name)
    
    def load_schedules(self):
        """Carrega schedules com verificacao de vazios."""
        self.view_schedules = []
        self.panel_schedules = []
        try:
            for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).WhereElementIsNotElementType():
                if s.IsTemplate or "PanelSchedule" in s.GetType().Name:
                    continue
                try:
                    if s.Definition.IsInternalKeynoteSchedule or s.Definition.IsRevisionSchedule:
                        continue
                except:
                    pass
                is_empty = False
                try:
                    if s.GetTableData().GetSectionData(DB.SectionType.Body).NumberOfRows <= 0:
                        is_empty = True
                except:
                    pass
                self.view_schedules.append({
                    'schedule': s, 'display_name': s.Name,
                    'is_empty': is_empty, 'is_panel': False
                })

            for v in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType():
                if "PanelScheduleView" not in v.GetType().Name or v.IsTemplate:
                    continue
                is_empty = False
                try:
                    panel_id = v.GetPanel()
                    if not panel_id or panel_id == DB.ElementId.InvalidElementId:
                        is_empty = True
                    else:
                        panel = doc.GetElement(panel_id)
                        if not panel:
                            is_empty = True
                        else:
                            systems = []
                            try:
                                if hasattr(panel, 'MEPModel') and panel.MEPModel:
                                    systems = list(panel.MEPModel.GetAssignedElectricalSystems())
                                else:
                                    systems = list(panel.GetAssignedElectricalSystems())
                            except:
                                pass
                            if not systems:
                                is_empty = True
                except:
                    pass
                self.panel_schedules.append({
                    'schedule': v, 'display_name': v.Name,
                    'is_empty': is_empty, 'is_panel': True
                })

            self.view_schedules.sort(key=lambda x: x['display_name'])
            self.panel_schedules.sort(key=lambda x: x['display_name'])
        except Exception as e:
            logger.error("Erro ao carregar schedules: " + str(e))
        finally:
            smart_gc_collect()

    # ── Checklist dropdown ─────────────────────────────────────────────────
    def mode_changed(self, sender, args):
        """Troca o modo (Schedule / Quadro) e repopula a checklist."""
        idx = self.ComboBox_ExportMode.SelectedIndex
        schedule_list = self.view_schedules if idx == 0 else self.panel_schedules
        self._populate_checklist(schedule_list)

    def _populate_checklist(self, schedule_list):
        """Cria os CheckBoxes da checklist com base na lista de schedules."""
        from System.Windows.Media import Brushes
        self.stack_ScheduleChecklist.Children.Clear()
        self._schedule_checkboxes = []

        try:
            style = self.FindResource("ModernCheckBox")
        except:
            style = None

        for sch_dict in schedule_list:
            name = sch_dict['display_name']
            is_empty = sch_dict.get('is_empty', False)

            cb = Controls.CheckBox()
            cb.Content = name + ("  (Vazio)" if is_empty else "")
            cb.IsChecked = False
            cb.Margin = System.Windows.Thickness(0, 2, 0, 2)
            cb.Cursor = System.Windows.Input.Cursors.Hand
            cb.FontSize = 12
            if is_empty:
                cb.Foreground = Brushes.Gray
                cb.IsEnabled = False
            if style:
                cb.Style = style
            cb.Click += self._on_checklist_changed
            self._schedule_checkboxes.append((cb, sch_dict))
            self.stack_ScheduleChecklist.Children.Add(cb)

        self.chk_SelectAllSchedules.IsChecked = False
        self._update_display_text()
        self._update_export_path()
        self._update_export_preview()

    def _toggle_dropdown(self, sender, args):
        """Abre ou fecha o painel de checklist."""
        self._dropdown_open = not self._dropdown_open
        self.pnl_ChecklistDropdown.Visibility = (
            Visibility.Visible if self._dropdown_open else Visibility.Collapsed
        )
        self.txt_DropdownArrow.Text = "▴" if self._dropdown_open else "▾"

    def _on_select_all_changed(self, sender, args):
        """Marca ou desmarca todos os itens habilitados."""
        state = bool(self.chk_SelectAllSchedules.IsChecked)
        for cb, _ in self._schedule_checkboxes:
            if cb.IsEnabled:
                cb.IsChecked = state
        self._update_display_text()
        self._update_export_path()
        self._update_export_preview()

    def _on_keep_format_changed(self, _sender, _args):
        """Mostra/oculta o aviso de 'não reimportável' conforme o checkbox."""
        checked = bool(self.CheckBox_KeepFormat.IsChecked)
        self.pnl_FormatWarning.Visibility = (
            Visibility.Visible if checked else Visibility.Collapsed
        )

    def _on_checklist_changed(self, sender, args):
        """Chamado quando qualquer checkbox da lista muda."""
        enabled = [(cb, d) for cb, d in self._schedule_checkboxes if cb.IsEnabled]
        all_checked = enabled and all(cb.IsChecked for cb, _ in enabled)
        self.chk_SelectAllSchedules.IsChecked = all_checked
        self._update_display_text()
        self._update_export_path()
        self._update_export_preview()

    def _get_selected_schedules(self):
        """Retorna lista de schedule_dicts marcados na checklist."""
        return [d for cb, d in self._schedule_checkboxes if cb.IsChecked]

    def _update_display_text(self):
        """Atualiza o texto do botão dropdown com a seleção atual."""
        selected = self._get_selected_schedules()
        if not selected:
            self.txt_DropdownDisplay.Text = "Selecionar..."
            self.update_stats("Nenhuma tabela selecionada")
        elif len(selected) == 1:
            self.txt_DropdownDisplay.Text = selected[0]['display_name']
            try:
                sch = selected[0]['schedule']
                count = sch.GetTableData().GetSectionData(DB.SectionType.Body).NumberOfRows
                self.update_stats("Tabela com aprox. {} linhas".format(count))
            except:
                self.update_stats("1 tabela selecionada")
        else:
            self.txt_DropdownDisplay.Text = "{} tabelas selecionadas".format(len(selected))
            self.update_stats("{} tabelas serão exportadas como abas separadas".format(len(selected)))

    def browse_export(self, sender, args):
        suggested_name = self._build_suggested_name()
        
        # Usar ultima pasta se disponivel
        default_path = suggested_name + ".xlsx"
        if self.last_export_folder and os.path.exists(self.last_export_folder):
            default_path = os.path.join(self.last_export_folder, default_path)
            
        file_path = forms.save_file(file_ext="xlsx", default_name=default_path)
        if file_path:
            # Salvar pasta — nome vai ser atualizado automaticamente conforme seleção
            self.last_export_folder = os.path.dirname(file_path)
            self.export_path = file_path
            self.TextBox_ExportPath.Text = file_path
            self.update_status("Pasta: " + self.last_export_folder)

    def _update_export_preview(self, from_combo=False):
        """Popula dg_ExportPreview com os dados reais da tabela selecionada."""
        try:
            selected = self._get_selected_schedules()
            
            if not from_combo:
                current_items = [self.cmb_ExportPreviewSelect.Items[i] for i in range(self.cmb_ExportPreviewSelect.Items.Count)]
                target_items = [s['display_name'] for s in selected]
                
                if current_items != target_items:
                    self.cmb_ExportPreviewSelect.Items.Clear()
                    for name in target_items:
                        self.cmb_ExportPreviewSelect.Items.Add(name)
                    if self.cmb_ExportPreviewSelect.Items.Count > 0:
                        self.cmb_ExportPreviewSelect.SelectedIndex = 0
            
            if self.cmb_ExportPreviewSelect.Items.Count > 1:
                self.cmb_ExportPreviewSelect.Visibility = Visibility.Visible
            else:
                self.cmb_ExportPreviewSelect.Visibility = Visibility.Collapsed
                
            self.dg_ExportPreview.ItemsSource = None
            self.dg_ExportPreview.Columns.Clear()

            if not selected or self.cmb_ExportPreviewSelect.SelectedIndex < 0:
                self.dg_ExportPreview.Visibility = Visibility.Collapsed
                self.lbl_PreviewStatus.Text = "Selecione uma tabela para ver a pré-visualização dos dados."
                return

            idx = self.cmb_ExportPreviewSelect.SelectedIndex
            if idx >= len(selected): idx = 0
            sch_dict = selected[idx]
            schedule = sch_dict['schedule']
            is_panel = sch_dict.get('is_panel', False)

            dt = DataTable()
            max_preview_rows = 50

            if is_panel:
                rows, headers = get_panel_schedule_data(schedule)
                for h in headers:
                    dt.Columns.Add(h)
                
                for i, r in enumerate(rows):
                    if i >= max_preview_rows:
                        break
                    row = dt.NewRow()
                    for j, h in enumerate(headers):
                        row[j] = str(r.get(h, ""))
                    dt.Rows.Add(row)
            else:
                elements, param_defs = get_schedule_elements_and_params(schedule)
                dt.Columns.Add("ElementId")
                for p in param_defs:
                    dt.Columns.Add(p.name)
                
                param_cache = {}
                for i, el in enumerate(elements):
                    if i >= max_preview_rows:
                        break
                    row = dt.NewRow()
                    try:
                        row[0] = str(el.Id.IntegerValue)
                    except:
                        row[0] = ""
                    for j, p in enumerate(param_defs):
                        val = get_element_parameter_value(el, p, param_cache)
                        row[j+1] = str(val) if val is not None else ""
                    dt.Rows.Add(row)

            self.dg_ExportPreview.ItemsSource = dt.DefaultView
            self.dg_ExportPreview.Visibility  = Visibility.Visible
            
            suffix = " (mostrando '{}')".format(sch_dict['display_name']) if len(selected) > 1 else ""
            msg_rows = dt.Rows.Count
            if msg_rows >= max_preview_rows:
                msg_rows = "50+"
            self.lbl_PreviewStatus.Text = "Pré-visualização: {} colunas, {} linhas{}".format(dt.Columns.Count, msg_rows, suffix)
        except Exception as ex:
            try:
                self.lbl_PreviewStatus.Text = "Preview indisponível: " + str(ex)
                self.dg_ExportPreview.Visibility = Visibility.Collapsed
            except:
                pass

    def _update_import_preview(self, from_combo=False, first_load=False):
        """Lê o Excel e mostra preview das alterações sem aplicá-las."""
        try:
            if first_load:
                self.cmb_ImportPreviewSelect.Items.Clear()
                self.cmb_ImportPreviewSelect.Visibility = Visibility.Collapsed
                
            self.dg_ImportPreview.ItemsSource = None
            self.dg_ImportPreview.Columns.Clear()
            self.dg_ImportPreview.Visibility = Visibility.Collapsed

            if not self.import_path or not os.path.exists(self.import_path):
                return

            self.lbl_ImportStatus.Text = "Lendo arquivo..."

            dt = DataTable()
            dt.Columns.Add("Parâmetro")
            dt.Columns.Add("Elemento")
            dt.Columns.Add("Valor Atual")
            dt.Columns.Add("Novo Valor")
            dt.Columns.Add("Status")

            MAX_ROWS = 200
            row_count = 0
            modif_count = 0

            wb = _XlsxReader(self.import_path)
            try:
                valid_sheets = []
                for s_name in wb.sheetnames:
                    if s_name.startswith('_'):
                        continue
                    rows = wb.read_sheet(s_name)
                    if rows and str(rows[0][0] if rows[0] else "").strip() == "ElementId":
                        valid_sheets.append((s_name, rows))
                
                if not valid_sheets:
                    self.lbl_ImportStatus.Text = "Nenhuma aba importável encontrada."
                    return
                
                if first_load:
                    for s_name, _ in valid_sheets:
                        self.cmb_ImportPreviewSelect.Items.Add(s_name)
                    if valid_sheets:
                        self.cmb_ImportPreviewSelect.SelectedIndex = 0
                        
                if self.cmb_ImportPreviewSelect.Items.Count > 1:
                    self.cmb_ImportPreviewSelect.Visibility = Visibility.Visible
                else:
                    self.cmb_ImportPreviewSelect.Visibility = Visibility.Collapsed
                    
                idx = self.cmb_ImportPreviewSelect.SelectedIndex
                if idx < 0 or idx >= len(valid_sheets):
                    self.lbl_ImportStatus.Text = "Nenhuma aba importável selecionada."
                    return
                    
                selected_sheet_name, rows = valid_sheets[idx]

                headers = [str(h).strip() if h is not None else "" for h in rows[0]]
                pnames  = [unit_postfix_pattern.sub("", h).strip() for h in headers[1:]]
                elem_cache = {}

                for row in rows[1:]:
                    if row_count >= MAX_ROWS:
                        break
                    if row[0] is None:
                        continue
                    try:
                        eid = int(float(row[0]))
                        el  = elem_cache.get(eid)
                        if not el:
                            el = doc.GetElement(DB.ElementId(eid))
                            if el:
                                elem_cache[eid] = el
                        if not el:
                            continue

                        try:
                            elem_name = el.Name or str(eid)
                        except:
                            elem_name = str(eid)

                        for cx, pname in enumerate(pnames):
                            val = row[cx + 1] if cx + 1 < len(row) else None
                            if val is None or val == "":
                                continue

                            param = get_param_robust(el, pname)
                            if not param:
                                status = "Não Encontrado"
                                current = ""
                            elif param.IsReadOnly:
                                status  = "Somente Leitura"
                                current = get_parameter_value(param)
                            else:
                                current = get_parameter_value(param)
                                new_str = str(int(val)) if isinstance(val, float) and val == int(val) else str(val)
                                if str(current) != new_str:
                                    status = "Modificar"
                                    modif_count += 1
                                else:
                                    status = "Igual"

                            row_dt = dt.NewRow()
                            row_dt[0] = pname
                            row_dt[1] = elem_name
                            row_dt[2] = str(current)
                            row_dt[3] = str(val)
                            row_dt[4] = status
                            dt.Rows.Add(row_dt)

                            row_count += 1
                            if row_count >= MAX_ROWS:
                                break
                    except:
                        pass
            finally:
                wb.close()

            self.dg_ImportPreview.ItemsSource = dt.DefaultView
            self.dg_ImportPreview.Visibility  = Visibility.Collapsed # Começa colapsado e expande se tiver dados
            if dt.Rows.Count > 0:
                self.dg_ImportPreview.Visibility = Visibility.Visible

            msg = "{} alterações detectadas".format(modif_count)
            if row_count >= MAX_ROWS:
                msg += " (exibindo primeiras {})".format(MAX_ROWS)
            self.lbl_ImportStatus.Text = msg
        except Exception as ex:
            try:
                self.lbl_ImportStatus.Text = "Erro ao ler preview: " + str(ex)
                self.dg_ImportPreview.Visibility = Visibility.Collapsed
            except:
                pass

    def browse_import(self, sender, args):
        path = forms.pick_file(file_ext="xlsx")
        if path:
            self.import_path = path
            self.TextBox_ImportPath.Text = path
            self.update_status("Ler de: " + os.path.basename(path))
            self._update_import_preview(first_load=True)

    def do_export(self, sender, args):
        if not self.export_path:
            return forms.alert("Defina o local do arquivo.\nClique em '...' para escolher a pasta.")

        selected = self._get_selected_schedules()
        if not selected:
            return forms.alert("Selecione ao menos uma tabela para exportar.")

        targets = []
        skipped_names = []

        try:
            try:
                from pyrevit import forms as _pyforms
                _ProgressBar = _pyforms.ProgressBar
            except Exception:
                _ProgressBar = None

            total_steps = len(selected) + 1
            _pb_ctx = _ProgressBar(title=u"Exportando para Excel...", cancellable=False) if _ProgressBar else None

            def _pb_update(step):
                try:
                    if _pb_ctx:
                        _pb_ctx.update_progress(step, total_steps)
                except Exception:
                    pass

            if _pb_ctx:
                _pb_ctx.__enter__()
            try:
                self.update_status("Lendo dados...")

                for i, sch_dict in enumerate(selected):
                    _pb_update(i)
                    src     = []
                    params  = []
                    is_panel = sch_dict.get('is_panel', False)
                    name    = sch_dict['display_name']
                    sch     = sch_dict['schedule']

                    if is_panel:
                        rows, headers = get_panel_schedule_data(sch)
                        class RowObj:
                            def __init__(self, d):
                                self.row_data = d
                        src    = [RowObj(r) for r in rows]
                        params = [ParamDef(name=h, istype=False, definition=None,
                                           isreadonly=True, isunit=False,
                                           storagetype=DB.StorageType.String) for h in headers]
                    else:
                        src, params = get_schedule_elements_and_params(sch)

                    if src:
                        targets.append({'name': name or "Sheet", 'src': src,
                                        'params': params, 'is_panel': is_panel})
                    else:
                        skipped_names.append(name)

                if not targets:
                    self.update_status("Nada para exportar.")
                    msg = "Nenhuma tabela com dados para exportar."
                    if skipped_names:
                        msg += "\n\nTabelas vazias ignoradas:\n- " + "\n- ".join(skipped_names)
                    return forms.alert(msg)

                formatted = bool(self.CheckBox_KeepFormat.IsChecked)

                _pb_update(len(selected))
                self.update_status("Gerando Excel ({} abas)...".format(len(targets)))
                export_xls(targets, self.export_path, formatted=formatted)
                self.update_status("Concluído!")
            finally:
                if _pb_ctx:
                    try:
                        _pb_ctx.__exit__(None, None, None)
                    except Exception:
                        pass

            report = "Exportação finalizada!\n{} abas criadas.".format(len(targets))
            if skipped_names:
                report += "\n\nTabelas vazias ignoradas:\n- " + "\n- ".join(skipped_names)
            if formatted:
                report += "\n\n⚠ Arquivo exportado no modo formatado.\nEste arquivo não pode ser reimportado ao Revit."
            forms.alert(report, title="Sucesso")

            if not formatted:
                # Modo padrão: auto-preencher importação
                self.import_path = self.export_path
                self.TextBox_ImportPath.Text = self.export_path
                self.tab_Main.SelectedIndex = 1
                self._update_import_preview(first_load=True)

            if self.CheckBox_OpenFolder.IsChecked:
                import subprocess
                subprocess.Popen('explorer /select,"{}"'.format(self.export_path))

        except Exception as e:
            logger.error(traceback.format_exc())
            forms.alert("Erro: " + str(e))
        finally:
            targets = None

    def do_import(self, sender, args):
        if not self.import_path:
            return forms.alert("Selecione o arquivo.")

        # Verificar antes de importar se há abas importáveis (sem quadros elétricos)
        try:
            wb = _XlsxReader(self.import_path)
            has_importable = False
            try:
                for s_name in wb.sheetnames:
                    if s_name.startswith('_'):
                        continue
                    rows = wb.read_sheet(s_name)
                    if rows and str(rows[0][0] if rows[0] else "").strip() == "ElementId":
                        has_importable = True
                        break
            finally:
                wb.close()

            if not has_importable:
                return forms.alert(
                    "Este arquivo não contém abas importáveis.\n\n"
                    "Quadros de Cargas (painéis elétricos) são exportados apenas para leitura "
                    "e não podem ser reimportados para o Revit.",
                    title="Importação Bloqueada"
                )
        except Exception as e:
            logger.error("Erro na verificação pré-import: " + str(e))

        self.update_status("Importando...")
        try:
            import_xls(self.import_path)
            self.update_status("Importação Finalizada!")
            self._update_import_preview(first_load=True)  # Atualiza preview após importar
        except Exception as e:
            logger.error(traceback.format_exc())
            forms.alert("Erro na importação: " + str(e))

    def close_clicked(self, sender, args):
        # Limpeza rapida apenas de dicionarios
        clear_all_caches()
        self.Close()

# ==================== MAIN ====================
if __name__ == "__main__":
    # Inicializar otimizações
    setup_memory_optimization()
    
    xaml_file = script.get_bundle_file("Exportimport.xaml")
    if os.path.exists(xaml_file):
        try:
            window = ExportImportWindow(xaml_file)
            window.ShowDialog()
        except Exception as e:
            logger.error("Erro na janela: " + str(e))
            logger.error(traceback.format_exc())
        finally:
            # Limpeza final rapida
            clear_all_caches()
            # Fim do script - SO libera memoria naturalmente
    else:
        forms.alert("Arquivo XAML nao encontrado: " + str(xaml_file), title="Erro")