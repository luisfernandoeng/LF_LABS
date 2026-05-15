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
    ref = str(ref).replace("$", "")
    col_str = ''
    row_str = ''
    for ch in ref:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    return _col_letter_to_index(col_str), int(row_str)

def _parse_excel_range(range_ref):
    """Converte range Excel A1:C5 em indices 0-based/1-based: c1,r1,c2,r2."""
    ref = str(range_ref or "").replace("$", "")
    if "!" in ref:
        ref = ref.split("!")[-1]
    ref = ref.replace("'", "")
    if "," in ref:
        ref = ref.split(",")[0]
    parts = ref.split(":")
    if len(parts) == 1:
        c1, r1 = _parse_cell_ref(parts[0])
        return c1, r1, c1, r1
    c1, r1 = _parse_cell_ref(parts[0])
    c2, r2 = _parse_cell_ref(parts[1])
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)

class _XlsxReader:
    """Leitor leve de .xlsx usando apenas zipfile + ElementTree."""
    def __init__(self, file_path):
        self._zf = zipfile.ZipFile(file_path, 'r')
        self._shared_strings = self._load_shared_strings()
        self._sheet_names, self._sheet_paths, self._print_areas = self._load_workbook_info()

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

        print_areas = {}
        try:
            defined_names = wb_root.find('ss:definedNames', _XLSX_NS)
            if defined_names is not None:
                for dn in defined_names.findall('ss:definedName', _XLSX_NS):
                    if dn.get('name') != '_xlnm.Print_Area':
                        continue
                    txt = dn.text or ""
                    local_id = dn.get('localSheetId')
                    if local_id is not None:
                        try:
                            idx = int(local_id)
                            if idx >= 0 and idx < len(names):
                                print_areas[names[idx]] = txt
                        except:
                            pass
                    else:
                        for s_name in names:
                            if txt.startswith("'" + s_name + "'!") or txt.startswith(s_name + "!"):
                                print_areas[s_name] = txt
                                break
        except:
            pass

        return names, paths, print_areas

    @property
    def sheetnames(self):
        return list(self._sheet_names)

    def get_print_area(self, sheet_name):
        return self._print_areas.get(sheet_name, "")

    def get_column_widths(self, sheet_name, col_count):
        """Le larguras de coluna do XLSX; volta None para colunas sem largura."""
        widths = [None for _ in range(col_count)]
        try:
            idx = self._sheet_names.index(sheet_name)
            path = self._sheet_paths[idx]
            xml_data = self._zf.read(path)
            root = ET.fromstring(xml_data)
            cols = root.find('ss:cols', _XLSX_NS)
            if cols is not None:
                for col in cols.findall('ss:col', _XLSX_NS):
                    min_c = int(col.get('min', '1')) - 1
                    max_c = int(col.get('max', '1')) - 1
                    width = float(col.get('width', '0') or 0)
                    if width <= 0:
                        continue
                    for c in range(max(min_c, 0), min(max_c + 1, col_count)):
                        widths[c] = width
        except:
            pass
        return widths

    def read_print_area_or_sheet(self, sheet_name):
        rows = self.read_sheet(sheet_name)
        area = self.get_print_area(sheet_name)
        if not area:
            return rows, False
        try:
            c1, r1, c2, r2 = _parse_excel_range(area)
            sliced = []
            for row in rows[r1 - 1:r2]:
                sliced.append(tuple(row[c1:c2 + 1]))
            return sliced, True
        except:
            return rows, False

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
import unicodedata

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

def _get_builtin_parameter(element, builtin_names):
    """Busca parametros BuiltInParameter sem quebrar em versoes diferentes do Revit."""
    for builtin_name in builtin_names:
        try:
            builtin_param = getattr(DB.BuiltInParameter, builtin_name)
            param = element.get_Parameter(builtin_param)
            if param and param.HasValue:
                return param
        except:
            pass
    return None

def _get_lookup_parameter(element, lookup_names):
    """Busca parametros por nome em PT/EN como fallback."""
    for lookup_name in lookup_names:
        try:
            param = element.LookupParameter(lookup_name)
            if param and param.HasValue:
                return param
        except:
            pass
    return None

def _parameter_display_value(param):
    if not param:
        return ""
    try:
        vs = param.AsValueString()
        if vs not in (None, ""):
            return vs
    except:
        pass
    try:
        s = param.AsString()
        if s not in (None, ""):
            return s
    except:
        pass
    try:
        if param.StorageType == DB.StorageType.Double:
            return param.AsDouble()
        if param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()
    except:
        pass
    return ""

def _parse_excel_number(value, default=None):
    """Converte textos do Revit como '1.250,5 W' ou '0,92' para numero."""
    if value in (None, ""):
        return default
    try:
        if isinstance(value, (int, float)):
            return float(value)
    except:
        pass
    try:
        text = str(value).strip()
        if not text:
            return default
        text = text.replace(u"\xa0", " ")
        matches = re.findall(r"[-+]?\d[\d.,]*", text)
        if not matches:
            return default
        num = matches[0]
        if "," in num and "." in num:
            num = num.replace(".", "").replace(",", ".")
        elif "," in num:
            num = num.replace(",", ".")
        elif "." in num:
            parts = num.split(".")
            if len(parts[-1]) == 3 and len(parts) > 1:
                num = "".join(parts)
        parsed = float(num)
        if "%" in text:
            parsed = parsed / 100.0
        return parsed
    except:
        return default

def _clean_number(value, default=None):
    num = _parse_excel_number(value, default)
    if num is None:
        return default
    try:
        if abs(num - int(num)) < 0.000001:
            return int(num)
    except:
        pass
    return num

def _natural_circuit_sort_key(value):
    """Ordena circuitos por prefixo e primeiro numero: 3,4,5 fica antes de 6."""
    text = str(value or "").strip()
    first_number = re.search(r"\d+", text)
    prefix = re.sub(r"\d.*$", "", text).strip().lower()
    number = int(first_number.group(0)) if first_number else 999999
    return (prefix, number, text.lower())

def _normalize_text(value):
    """Normaliza texto para busca tolerante a acentos e caixa."""
    try:
        text = unicode(value or "")
    except:
        text = str(value or "")
    try:
        text = unicodedata.normalize('NFD', text)
        text = ''.join([ch for ch in text if unicodedata.category(ch) != 'Mn'])
    except:
        pass
    return text.lower().strip()

def _get_panel_param_value(element, builtin_names, lookup_names, default=""):
    param = _get_builtin_parameter(element, builtin_names)
    if not param:
        param = _get_lookup_parameter(element, lookup_names)
    value = _parameter_display_value(param)
    if value in (None, ""):
        return default
    return value

def _find_numeric_parameter_containing(element, include_terms, exclude_terms=None, min_value=None, max_value=None):
    """Procura um parametro numerico por partes do nome."""
    if exclude_terms is None:
        exclude_terms = []
    try:
        parameters = list(element.Parameters)
    except:
        parameters = []

    for param in parameters:
        try:
            pname = _normalize_text(param.Definition.Name)
        except:
            continue

        matched = True
        for term in include_terms:
            if _normalize_text(term) not in pname:
                matched = False
                break
        if not matched:
            continue

        blocked = False
        for term in exclude_terms:
            if _normalize_text(term) in pname:
                blocked = True
                break
        if blocked:
            continue

        value = _parse_excel_number(_parameter_display_value(param), None)
        if value is None:
            continue
        if min_value is not None and value < min_value:
            continue
        if max_value is not None and value > max_value:
            continue
        return value

    return None

def _demand_factor_from_connected_and_demand(element, class_terms):
    """Calcula FD por demanda/conectado quando o fator nao existe diretamente."""
    connected = None
    demanded = None
    try:
        parameters = list(element.Parameters)
    except:
        parameters = []

    for param in parameters:
        try:
            pname = _normalize_text(param.Definition.Name)
        except:
            continue
        if not all([term in pname for term in class_terms]):
            continue

        value = _parse_excel_number(_parameter_display_value(param), None)
        if value in (None, 0):
            continue

        is_connected = (
            'connected' in pname or 'conectad' in pname
        )
        is_demand = (
            'estimated demand' in pname or 'demanda' in pname or 'demand' in pname
        )
        is_factor = 'factor' in pname or 'fator' in pname
        is_current = 'current' in pname or 'corrente' in pname or 'atual' in pname

        if is_factor or is_current:
            continue
        if is_connected:
            connected = value
        elif is_demand:
            demanded = value

    try:
        if connected and demanded is not None and connected > 0:
            factor = float(demanded) / float(connected)
            if factor > 0 and factor <= 1:
                return factor
    except:
        pass
    return None

def _classification_terms(load_classification):
    class_norm = _normalize_text(load_classification)
    terms = []
    if not class_norm:
        return terms

    if "tue" in class_norm:
        terms.append(['tue'])
    if "ilumin" in class_norm:
        terms.append(['ilumin'])
        terms.append(['lighting'])
    if "equipamentos" in class_norm or "rede" in class_norm:
        terms.append(['equipamentos'])
        terms.append(['rede'])
    if "hvac" in class_norm:
        terms.append(['hvac'])
    if "tomadas" in class_norm and "geral" in class_norm:
        terms.append(['tomadas', 'geral'])
    if "tomadas" in class_norm and "especific" in class_norm:
        terms.append(['tomadas', 'especific'])
    if "motor" in class_norm:
        terms.append(['motor'])

    words = [w for w in re.split(r"[^a-z0-9]+", class_norm) if len(w) >= 3]
    if words:
        terms.append(words[:3])
    return terms

def _panel_load_group_label(load_classification):
    """Agrupa classificacoes eletricas na ordem desejada para quadros."""
    class_norm = _normalize_text(load_classification)
    if not class_norm:
        return u"Outros"

    if "ilumin" in class_norm or class_norm in ("em", "emergencia") or "emerg" in class_norm:
        return u"Iluminação"
    if "estabil" in class_norm or "rede" in class_norm:
        return u"Tomadas estabilizadas"
    if "tomada" in class_norm or "tug" in class_norm or "tue" in class_norm:
        return u"Tomadas comuns"
    if "hvac" in class_norm or "ar condicionado" in class_norm or "climat" in class_norm:
        return "HVAC"
    if "motor" in class_norm or "bomba" in class_norm or "elevador" in class_norm:
        return u"Motores"
    return str(load_classification or u"Outros")

def _panel_load_group_order(load_classification):
    label = _panel_load_group_label(load_classification)
    norm = _normalize_text(label)
    order = {
        "iluminacao": 0,
        "tomadas comuns": 1,
        "tomadas estabilizadas": 2,
        "hvac": 3,
        "motores": 4,
    }
    return (order.get(norm, 99), norm)

def _infer_load_classification_from_connected(element, apparent_load):
    """Infere classificacao pela carga conectada que bate com a carga do circuito."""
    target = _parse_excel_number(apparent_load, None)
    if target is None or target <= 0:
        return ""
    try:
        parameters = list(element.Parameters)
    except:
        parameters = []

    for param in parameters:
        try:
            raw_name = param.Definition.Name
            pname = _normalize_text(raw_name)
        except:
            continue
        if not ('connected' in pname or 'conectad' in pname):
            continue

        value = _parse_excel_number(_parameter_display_value(param), None)
        if value is None:
            continue
        try:
            if abs(float(value) - float(target)) > max(1.0, float(target) * 0.02):
                continue
        except:
            continue

        text = str(raw_name)
        for token in [
            u'Potência aparente conectada de ', 'Potencia aparente conectada de ',
            'Connected'
        ]:
            text = text.replace(token, '')
        return text.strip()

    return ""

def _find_demand_factor_for_circuit(element, load_classification="", panel=None):
    """Prioriza o fator de demanda por tipo de carga; ignora parametros FD zerados."""
    search_sets = []
    class_terms = _classification_terms(load_classification)
    sources = [element]
    if panel is not None:
        sources.insert(0, panel)

    for base_terms in class_terms:
        search_sets.append(base_terms + ['demand', 'factor'])
        search_sets.append(base_terms + ['fator', 'demanda'])

    for source in sources:
        for terms in search_sets:
            value = _find_numeric_parameter_containing(
                source, terms, ['total'], 0.000001, 1
            )
            if value not in (None, ""):
                return value

    for source in sources:
        for terms in class_terms:
            value = _demand_factor_from_connected_and_demand(source, terms)
            if value not in (None, ""):
                return value

    for source in sources:
        value = _find_numeric_parameter_containing(
            source, ['demand', 'factor'], ['total'], 0.000001, 1
        )
        if value not in (None, ""):
            return value
        value = _find_numeric_parameter_containing(
            source, ['fator', 'demanda'], ['total'], 0.000001, 1
        )
        if value not in (None, ""):
            return value

    return 1

def _get_circuit_element_count(electrical_system):
    """Conta os elementos ligados ao circuito; volta 1 quando a API nao expuser a lista."""
    param_qty = _get_panel_param_value(
        electrical_system,
        [],
        [u'Número de elementos', 'Numero de elementos', u'Nº de elementos',
         'N de elementos', 'Number of Elements', 'Number of elements'],
        ""
    )
    qty = _parse_excel_number(param_qty, None)
    if qty is not None and qty > 0:
        return int(qty)

    qty = _find_numeric_parameter_containing(electrical_system, ['elementos'], [], 1, None)
    if qty is not None and qty > 0:
        return int(qty)

    candidates = []
    try:
        candidates = list(electrical_system.Elements)
    except:
        pass
    if not candidates:
        try:
            candidates = list(electrical_system.GetCircuitElements())
        except:
            pass
    if not candidates:
        try:
            candidates = list(electrical_system.GetElements())
        except:
            pass

    count = 0
    seen = set()
    for element in candidates:
        try:
            eid = element.Id.IntegerValue
            if eid in seen:
                continue
            seen.add(eid)
        except:
            pass
        count += 1

    return count if count > 0 else 1

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
def get_panel_schedule_data(schedule_view, use_element_quantity=True, preserve_revit_order=False):
    """Extrai dados dos paineis com cache seguro."""
    cache_key = ("panel_schedule", schedule_view.Id.IntegerValue, bool(use_element_quantity), bool(preserve_revit_order))
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
            dummy = {
                'Nome do Quadro': panel_name,
                u'Ítem': "",
                u'Descrição equipamento': 'Sem circuitos',
                'Quant.': "",
                u'Carga Unitária (W)': "",
                u'Carga Total (W)': "",
                'Fator de Demanda': "",
                u'Fator de Potência': "",
                u'Carga Dem. (VA)': "",
                u'Seção do Condutor Adotado': "",
                u'Proteção Adotada': "",
            }
            result = [dummy], [u'Ítem', u'Descrição equipamento', 'Quant.',
                               u'Carga Unitária (W)', u'Carga Total (W)',
                               'Fator de Demanda', u'Fator de Potência',
                               u'Carga Dem. (VA)', u'Seção do Condutor Adotado',
                               u'Proteção Adotada']
            _param_def_cache[cache_key] = result
            return result

        headers = [u'Ítem', u'Descrição equipamento', 'Quant.',
                   u'Carga Unitária (W)', u'Carga Total (W)',
                   'Fator de Demanda', u'Fator de Potência',
                   u'Carga Dem. (VA)', u'Seção do Condutor Adotado',
                   u'Proteção Adotada']
        
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

                item = row_data.get('Numero do Circuito', "")
                description = row_data.get('Nome da Carga', "")
                apparent_load = row_data.get('Carga Aparente (VA)', "")
                load_classification = _get_panel_param_value(
                    sys,
                    ['RBS_ELEC_CIRCUIT_LOAD_CLASSIFICATION_PARAM'],
                    [u'Classificação de carga', 'Classificacao de carga',
                     'Load Classification'],
                    ""
                )

                true_load = _get_panel_param_value(
                    sys,
                    ['RBS_ELEC_TRUE_LOAD', 'RBS_ELEC_TRUE_LOAD_PARAM',
                     'RBS_ELEC_CIRCUIT_TRUE_LOAD'],
                    ['Carga Real', 'Potência Real', 'Potencia Real',
                     'True Load', 'Real Power', 'Carga em Watts',
                     'Carga (W)', 'Potência (W)', 'Potencia (W)'],
                    ""
                )
                power_factor = _get_panel_param_value(
                    sys,
                    ['RBS_ELEC_POWER_FACTOR', 'RBS_ELEC_CIRCUIT_POWER_FACTOR',
                     'RBS_ELEC_POWER_FACTOR_PARAM'],
                    [u'Fator de Potência', 'Fator de Potencia',
                     'Power Factor', 'FP'],
                    1
                )
                apparent_num = _parse_excel_number(apparent_load, None)
                if load_classification in (None, ""):
                    load_classification = _infer_load_classification_from_connected(sys, apparent_num)

                demand_factor = _find_demand_factor_for_circuit(sys, load_classification, panel)
                if demand_factor in (None, ""):
                    demand_factor = _get_panel_param_value(
                        sys,
                        ['RBS_ELEC_DEMAND_FACTOR', 'RBS_ELEC_CIRCUIT_DEMAND_FACTOR',
                         'RBS_ELEC_DEMAND_FACTOR_PARAM'],
                        ['Fator de Demanda', 'Demand Factor', 'FD', 'FCA'],
                        1
                    )
                conductor = _get_panel_param_value(
                    sys,
                    [],
                    [u'Seção do Condutor Adotado (mm²)', u'Seção do Condutor Adotado',
                     'Secao do Condutor Adotado (mm2)', 'Secao do Condutor Adotado',
                     u'Seção do condutor adotado', 'Secao do condutor adotado',
                     u'Seção Condutor Adotado', 'Secao Condutor Adotado',
                     'Condutor Adotado', 'Fio', 'Wire Size'],
                    ""
                )
                protection = _get_panel_param_value(
                    sys,
                    [],
                    [u'Proteção do circuito', u'Proteção do Circuito',
                     'Protecao do circuito', 'Protecao do Circuito',
                     u'Proteção Adotada', 'Protecao Adotada',
                     u'Proteção adotada', 'Protecao adotada',
                     'Disjuntor', 'Disjuntor Adotado', 'Circuit Breaker',
                     'Circuit Rating', 'Rating'],
                    ""
                )
                if protection in (None, ""):
                    protection = _get_panel_param_value(
                        sys,
                        ['RBS_ELEC_CIRCUIT_RATING_PARAM', 'RBS_ELEC_CIRCUIT_FRAME_PARAM',
                         'RBS_ELEC_CIRCUIT_TRIP_PARAM'],
                        ['Disjuntor', 'Circuit Rating', 'Rating'],
                        row_data.get('Classificacao (A)', "")
                    )

                pf_num = _parse_excel_number(power_factor, 1)
                watts_num = _parse_excel_number(true_load, None)

                if watts_num is None and apparent_num is not None:
                    watts_num = apparent_num * (pf_num or 1)

                qty = _get_circuit_element_count(sys) if use_element_quantity else 1
                try:
                    if qty < 1:
                        qty = 1
                except:
                    qty = 1

                unit_watts = watts_num
                if use_element_quantity and watts_num not in (None, ""):
                    try:
                        unit_watts = float(watts_num) / float(qty)
                    except:
                        unit_watts = watts_num

                row_data[u'Ítem'] = item
                row_data[u'Descrição equipamento'] = description
                row_data['Quant.'] = qty
                row_data[u'Carga Unitária (W)'] = _clean_number(unit_watts, "")
                row_data[u'Carga Total (W)'] = ""
                row_data['Fator de Demanda'] = _clean_number(demand_factor, 1)
                row_data[u'Fator de Potência'] = _clean_number(power_factor, 1)
                row_data[u'Carga Dem. (VA)'] = ""
                row_data[u'Seção do Condutor Adotado'] = conductor
                row_data[u'Proteção Adotada'] = protection
                row_data['Classificacao Original'] = str(load_classification) if load_classification else u'Sem classificacao'
                row_data['Classificacao'] = _panel_load_group_label(load_classification)

                data_rows.append(row_data)
            except: 
                continue

        if preserve_revit_order:
            try:
                data_rows.sort(key=lambda x: _natural_circuit_sort_key(x.get('Numero do Circuito', x.get(u'Ítem', ''))))
            except:
                pass
        else:
            try:
                data_rows.sort(key=lambda x: (
                    _panel_load_group_order(x.get('Classificacao', '')),
                    _natural_circuit_sort_key(x.get('Numero do Circuito', x.get(u'Ítem', '')))
                ))
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

def _get_section_cell_text(section, row, col):
    """Le texto de celula de uma TableSectionData com fallback."""
    try:
        return section.GetCellText(row, col) or ""
    except:
        pass
    try:
        return section.GetCellText(row, col, False) or ""
    except:
        pass
    return ""

def get_schedule_visual_table(schedule):
    """Extrai a matriz visual da tabela do Revit (cabecalho + corpo)."""
    visual_rows = []
    try:
        table_data = schedule.GetTableData()
    except:
        return visual_rows

    for section_type in [DB.SectionType.Header, DB.SectionType.Body]:
        try:
            section = table_data.GetSectionData(section_type)
            row_count = section.NumberOfRows
            col_count = section.NumberOfColumns
        except:
            continue
        if row_count <= 0 or col_count <= 0:
            continue

        for r in range(row_count):
            row_values = []
            has_value = False
            for c in range(col_count):
                value = _get_section_cell_text(section, r, c)
                if value not in (None, ""):
                    has_value = True
                row_values.append(value)
            if has_value:
                visual_rows.append(row_values)

    return visual_rows


# ==================== EXPORTAÇÃO OTIMIZADA ====================
def export_xls(targets, file_path, formatted=False):
    """Exporta dados para Excel com múltiplas abas se necessário."""
    workbook = None
    try:
        # Configurar workbook com otimizações
        has_panel_schedule = any([t.get('is_panel', False) for t in targets])
        workbook_options = {
            'constant_memory': not has_panel_schedule,
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
        fmt_panel_title = get_excel_format(workbook, {
            "bold": True, "font_size": 14, "bg_color": "#D9EAD3",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        fmt_panel_info = get_excel_format(workbook, {
            "bold": True, "bg_color": "#E2F0D9", "border": 1,
            "align": "left", "valign": "vcenter",
        })
        fmt_panel_header = get_excel_format(workbook, {
            "bold": True, "bg_color": "#B7E1A1", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        fmt_panel_text = get_excel_format(workbook, {
            "border": 1, "align": "left", "valign": "vcenter",
        })
        fmt_panel_num = get_excel_format(workbook, {
            "border": 1, "align": "center", "valign": "vcenter",
            "num_format": "#,##0.00",
        })
        fmt_panel_int = get_excel_format(workbook, {
            "border": 1, "align": "center", "valign": "vcenter",
            "num_format": "0",
        })
        
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
                ws.set_tab_color("#70AD47")
                ws.freeze_panes(5, 0)
                ws.set_column(0, 0, 3)
                ws.set_column(1, 1, 8)
                ws.set_column(2, 5, 12)
                ws.set_column(6, 6, 8)
                ws.set_column(7, 7, 13)
                ws.set_column(8, 8, 12)
                ws.set_column(9, 9, 12)
                ws.set_column(10, 10, 12)
                ws.set_column(11, 11, 12)
                ws.set_column(12, 12, 22)
                ws.set_column(13, 13, 18)

                ws.merge_range(0, 1, 0, 13, "ESTUDO DE CARGAS", fmt_panel_title)
                ws.write(1, 1, "OBRA: " + sanitize_filename(os.path.splitext(os.path.basename(file_path))[0]), fmt_panel_text)
                ws.write(2, 1, u"Responsável: ", fmt_panel_text)
                ws.write(3, 1, "Representa a carga total do(a): " + sheet_name, fmt_panel_text)

                headers_panel = [
                    u"Ítem", u"Descrição equipamento", "Quant.",
                    u"Carga \nUnitária\n(W)", u"Carga\nTotal\n(W)",
                    u"Fator de \nDemanda", u"Fator de \nPotência",
                    u"Carga \nDem.\n(VA)", u"Seção do Condutor\nAdotado",
                    u"Proteção\nAdotada"
                ]
                ws.write(4, 1, headers_panel[0], fmt_panel_header)
                ws.merge_range(4, 2, 4, 5, headers_panel[1], fmt_panel_header)
                for col, label in [(6, headers_panel[2]), (7, headers_panel[3]),
                                   (8, headers_panel[4]), (9, headers_panel[5]),
                                   (10, headers_panel[6]), (11, headers_panel[7]),
                                   (12, headers_panel[8]), (13, headers_panel[9])]:
                    ws.write(4, col, label, fmt_panel_header)
                ws.set_row(4, 42)

                # Passagem 1: acumula dados por classificacao para saber a linha
                # de cada classificacao na aba de resumo antes de escrever a principal.
                summary_data = {}
                for el in src_elements:
                    if not hasattr(el, 'row_data'):
                        continue
                    data = el.row_data
                    cls_name = data.get('Classificacao', '') or u'Sem classificacao'
                    if cls_name not in summary_data:
                        summary_data[cls_name] = {'count': 0, 'installed_w': 0.0, 'fds': []}
                    summary_data[cls_name]['count'] += 1
                    try:
                        q = float(_clean_number(data.get('Quant.'), 1) or 1)
                        uw = float(_clean_number(data.get(u'Carga Unitária (W)'), 0) or 0)
                        fd = float(_clean_number(data.get('Fator de Demanda'), 1) or 1)
                        summary_data[cls_name]['installed_w'] += q * uw
                        summary_data[cls_name]['fds'].append(fd)
                    except:
                        pass

                # Linha Excel (1-indexed) de cada classificacao na aba de resumo:
                # linha 1 = cabecalho, dados comecam na linha 2.
                cls_sorted = sorted(summary_data.keys(), key=lambda cls: _panel_load_group_order(cls))
                cls_excel_row = {cls: (i + 2) for i, cls in enumerate(cls_sorted)}

                # Passagem 2: escreve aba principal com FD referenciando a aba de resumo.
                current_panel = None
                row = 5
                for el in src_elements:
                    if not hasattr(el, 'row_data'):
                        continue
                    data = el.row_data
                    panel_name = data.get('Nome do Quadro', sheet_name)
                    if panel_name != current_panel:
                        current_panel = panel_name
                        ws.merge_range(row, 2, row, 5, "Painel: " + str(panel_name), fmt_panel_info)
                        for col in [1, 6, 7, 8, 9, 10, 11, 12, 13]:
                            ws.write(row, col, "", fmt_panel_info)
                        row += 1

                    excel_row = row + 1
                    qty = _clean_number(data.get('Quant.'), 1)
                    unit_w = _clean_number(data.get(u'Carga Unitária (W)'), "")
                    demand = _clean_number(data.get('Fator de Demanda'), 1)
                    power_factor = _clean_number(data.get(u'Fator de Potência'), 1)
                    cls_name = data.get('Classificacao', '') or u'Sem classificacao'
                    fd_row = cls_excel_row.get(cls_name)

                    ws.write(row, 1, data.get(u'Ítem', ""), fmt_panel_text)
                    ws.merge_range(row, 2, row, 5, data.get(u'Descrição equipamento', ""), fmt_panel_text)
                    ws.write(row, 6, qty, fmt_panel_int)
                    ws.write(row, 7, unit_w, fmt_panel_num)
                    ws.write_formula(row, 8, "=H{0}*G{0}".format(excel_row), fmt_panel_num)
                    if fd_row:
                        ws.write_formula(row, 9,
                            u"='Classif. de Cargas'!D{}".format(fd_row),
                            fmt_panel_num, float(demand or 1))
                    else:
                        ws.write(row, 9, demand, fmt_panel_num)
                    ws.write(row, 10, power_factor, fmt_panel_num)
                    ws.write_formula(row, 11, "=IF(K{0}=0,0,(G{0}*H{0}*J{0})/K{0})".format(excel_row), fmt_panel_num)
                    ws.write(row, 12, data.get(u'Seção do Condutor Adotado', ""), fmt_panel_text)
                    ws.write(row, 13, data.get(u'Proteção Adotada', ""), fmt_panel_text)
                    row += 1

                if summary_data:
                    ws_sum = workbook.add_worksheet(u"Classif. de Cargas")
                    ws_sum.set_tab_color("#ED7D31")
                    ws_sum.freeze_panes(1, 0)
                    ws_sum.set_row(0, 36)
                    fmt_sum_total = get_excel_format(workbook, {
                        "bold": True, "bg_color": "#B7E1A1", "border": 1,
                        "align": "left", "valign": "vcenter",
                    })
                    sum_cols = [
                        u"Classificação de Carga",
                        u"Qtd. Circuitos",
                        u"Pot. Instalada (W)",
                        u"Fator de Demanda",
                        u"Pot. Demandada (W)",
                    ]
                    for c, h in enumerate(sum_cols):
                        ws_sum.write(0, c, h, fmt_panel_header)
                    widths_sum = [len(h) + 2 for h in sum_cols]
                    sum_row = 1
                    for cls_key in cls_sorted:
                        sd = summary_data[cls_key]
                        inst = round(sd['installed_w'], 2)
                        fds = sd['fds']
                        fd_disp = max(set(fds), key=fds.count) if fds else 1.0
                        excel_sum_row = sum_row + 1  # 1-indexed para formula
                        ws_sum.write(sum_row, 0, cls_key, fmt_panel_text)
                        ws_sum.write(sum_row, 1, sd['count'], fmt_panel_int)
                        ws_sum.write(sum_row, 2, inst, fmt_panel_num)
                        ws_sum.write(sum_row, 3, fd_disp, fmt_panel_num)
                        # Pot. Demandada = Pot. Instalada * FD — formula reativa
                        ws_sum.write_formula(sum_row, 4,
                            "=C{0}*D{0}".format(excel_sum_row),
                            fmt_panel_num, round(inst * fd_disp, 2))
                        sum_row += 1
                        if len(cls_key) > widths_sum[0]:
                            widths_sum[0] = min(len(cls_key), 45)
                    last_data_row = sum_row  # linha do total (1-indexed = sum_row+1)
                    ws_sum.write(sum_row, 0, "TOTAL", fmt_sum_total)
                    ws_sum.write(sum_row, 1,
                        sum(sd['count'] for sd in summary_data.values()), fmt_panel_int)
                    ws_sum.write_formula(sum_row, 2,
                        "=SUM(C2:C{})".format(last_data_row), fmt_panel_num)
                    ws_sum.write(sum_row, 3, "", fmt_sum_total)
                    ws_sum.write_formula(sum_row, 4,
                        "=SUM(E2:E{})".format(last_data_row), fmt_panel_num)
                    for c, w in enumerate(widths_sum):
                        ws_sum.set_column(c, c, w + 2)

            elif formatted:
                # ── Modo formatado: usa a matriz visual da tabela do Revit ─────
                ws.set_tab_color("#1F4E78")
                ws.freeze_panes(1, 0)
                visual_rows = get_schedule_visual_table(target.get('schedule')) if target.get('schedule') else []

                if visual_rows:
                    col_count = max([len(r) for r in visual_rows])
                    widths = [8 for _ in range(col_count)]
                    for r, row_values in enumerate(visual_rows):
                        is_header = (r == 0)
                        fmt_row = fmt_head_vis if is_header else (fmt_data_alt if r % 2 == 0 else fmt_data_vis)
                        ws.set_row(r, 22 if is_header else 18)
                        for c in range(col_count):
                            value = row_values[c] if c < len(row_values) else ""
                            ws.write(r, c, value, fmt_row)
                            slen = len(str(value)) if value else 0
                            if slen > widths[c]:
                                widths[c] = min(slen, 55)

                    for i, w in enumerate(widths):
                        ws.set_column(i, i, max(8, w + 2))
                    if len(visual_rows) > 1 and col_count > 0:
                        ws.autofilter(0, 0, len(visual_rows) - 1, col_count - 1)
                else:
                    # Fallback antigo caso a API nao entregue TableData.
                    header_names = [p.name for p in selected_params]
                    widths = [len(n) for n in header_names]
                    for i, name in enumerate(header_names):
                        ws.write(0, i, name, fmt_head_vis)
                    for r, el in enumerate(src_elements, 1):
                        fmt_row = fmt_data_alt if r % 2 == 0 else fmt_data_vis
                        for c, p in enumerate(selected_params):
                            value = get_param_display_string(el, p)
                            ws.write(r, c, value, fmt_row)
                            widths[c] = max(widths[c], min(len(value or ""), 55))
                    for i, w in enumerate(widths):
                        ws.set_column(i, i, w + 3)

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

def _first_text_note_type_id():
    try:
        tnt = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).FirstElement()
        if tnt:
            return tnt.Id
    except:
        pass
    return DB.ElementId.InvalidElementId

def _mm_to_ft(mm_value):
    return float(mm_value) / 304.8

def _get_or_create_table_text_type_id(text_mm=2.0):
    """Cria/usa um tipo de texto pequeno para tabelas importadas."""
    type_name = "LF Excel Table {:.1f}mm".format(float(text_mm))
    first_type = None
    try:
        for tnt in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType):
            try:
                if not first_type:
                    first_type = tnt
                if tnt.Name == type_name:
                    return tnt.Id
            except:
                pass
    except:
        pass

    if not first_type:
        return DB.ElementId.InvalidElementId

    try:
        new_type = first_type.Duplicate(type_name)
    except:
        return first_type.Id

    try:
        p = new_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        if p and not p.IsReadOnly:
            p.Set(_mm_to_ft(text_mm))
    except:
        pass
    try:
        p = new_type.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE)
        if p and not p.IsReadOnly:
            p.Set(1.0)
    except:
        pass
    return new_type.Id

def _create_cell_text_note(view, text_type_id, x, y, width, text):
    """Cria texto de celula com largura controlada quando a API suporta."""
    text = (text or "")[:120]
    opts = None
    try:
        opts = DB.TextNoteOptions(text_type_id)
        opts.HorizontalAlignment = DB.HorizontalTextAlignment.Left
    except:
        opts = None

    pt = DB.XYZ(x, y, 0)
    try:
        if opts:
            return DB.TextNote.Create(doc, view.Id, pt, width, text, opts)
    except:
        pass
    try:
        if opts:
            return DB.TextNote.Create(doc, view.Id, pt, text, opts)
    except:
        pass
    return DB.TextNote.Create(doc, view.Id, pt, text, text_type_id)

def _first_drafting_view_type_id():
    try:
        for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
            try:
                if vft.ViewFamily == DB.ViewFamily.Drafting:
                    return vft.Id
            except:
                pass
    except:
        pass
    return DB.ElementId.InvalidElementId

def _first_table_drafting_view_type_id():
    """Procura um tipo de vista de desenho nomeado como tabela/table no template."""
    try:
        for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
            try:
                if vft.ViewFamily != DB.ViewFamily.Drafting:
                    continue
                name = _normalize_text(vft.Name)
                if "tabela" in name or "table" in name:
                    return vft.Id
            except:
                pass
    except:
        pass
    return DB.ElementId.InvalidElementId

def _first_legend_view_type_id():
    try:
        for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
            try:
                if vft.ViewFamily == DB.ViewFamily.Legend:
                    return vft.Id
            except:
                pass
    except:
        pass
    return DB.ElementId.InvalidElementId

def _create_table_view(view_name):
    """Prioridade: Legend, tipo de desenho para tabela, depois Drafting comum."""
    legend_id = _first_legend_view_type_id()
    if legend_id != DB.ElementId.InvalidElementId:
        try:
            view = DB.View.CreateLegend(doc, legend_id)
            view.Name = _unique_view_name(view_name)
            return view, "Legend"
        except:
            pass

    table_drafting_id = _first_table_drafting_view_type_id()
    if table_drafting_id != DB.ElementId.InvalidElementId:
        try:
            view = DB.ViewDrafting.Create(doc, table_drafting_id)
            view.Name = _unique_view_name(view_name)
            return view, "Tabela"
        except:
            pass

    drafting_id = _first_drafting_view_type_id()
    if drafting_id != DB.ElementId.InvalidElementId:
        view = DB.ViewDrafting.Create(doc, drafting_id)
        view.Name = _unique_view_name(view_name)
        return view, "Drafting"

    return None, ""

def _unique_view_name(base_name):
    existing = set()
    try:
        for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
            try:
                existing.add(v.Name)
            except:
                pass
    except:
        pass
    name = base_name
    idx = 1
    while name in existing:
        idx += 1
        name = "{} ({})".format(base_name, idx)
    return name

def _cell_to_text(value):
    if value is None:
        return ""
    try:
        if isinstance(value, float) and value == int(value):
            return str(int(value))
    except:
        pass
    return str(value)

def _trim_external_rows(rows, max_rows=200, max_cols=20):
    """Mantem a ordem da planilha e remove apenas linhas/colunas finais vazias."""
    used_rows = []
    last_col = -1
    for row in rows:
        has_value = False
        for idx, value in enumerate(row):
            if value not in (None, ""):
                has_value = True
                if idx > last_col:
                    last_col = idx
        if has_value:
            used_rows.append(row)

    if last_col < 0:
        return [], False

    truncated = len(used_rows) > max_rows or (last_col + 1) > max_cols
    out_rows = []
    for row in used_rows[:max_rows]:
        out_rows.append([_cell_to_text(row[i] if i < len(row) else "") for i in range(min(last_col + 1, max_cols))])
    return out_rows, truncated

def import_external_xlsx_as_drafting_view(file_path, sheet_name=None):
    """Importa XLSX externo como tabela visual em Legend/Drafting View."""
    wb = None
    used_print_area = False
    excel_col_widths = []
    try:
        wb = _XlsxReader(file_path)
        target_sheet = sheet_name
        if not target_sheet or target_sheet not in wb.sheetnames:
            target_sheet = None
            for s_name in wb.sheetnames:
                if s_name.startswith('_'):
                    continue
                rows = wb.read_sheet(s_name)
                if rows:
                    target_sheet = s_name
                    break
        if not target_sheet:
            forms.alert("Nenhuma aba com dados encontrada no Excel.", title="Importar Planilha Externa")
            return False

        # Determinar o offset de coluna da area de impressao para alinhar as larguras
        _print_area_col_offset = 0
        area_str = wb.get_print_area(target_sheet)
        if area_str:
            try:
                _c1, _r1, _c2, _r2 = _parse_excel_range(area_str)
                _print_area_col_offset = _c1
            except:
                _print_area_col_offset = 0

        raw_rows, used_print_area = wb.read_print_area_or_sheet(target_sheet)
        rows, truncated = _trim_external_rows(raw_rows)
        if not rows:
            forms.alert("A aba selecionada esta vazia.", title="Importar Planilha Externa")
            return False

        # col_count baseado nos dados ja recortados
        _trimmed_col_count = max([len(r) for r in rows])
        # Buscar larguras absolutas cobrindo o range real da planilha
        _abs_col_count = _trimmed_col_count + _print_area_col_offset
        _all_widths = wb.get_column_widths(target_sheet, _abs_col_count)
        # Fatiar apenas as colunas que correspondem aos dados recortados
        excel_col_widths = _all_widths[_print_area_col_offset:_print_area_col_offset + _trimmed_col_count]
    except Exception as e:
        forms.alert("Erro ao ler Excel externo: " + str(e), title="Importar Planilha Externa")
        return False
    finally:
        try:
            if wb:
                wb.close()
        except:
            pass

    base_text_type_id = _first_text_note_type_id()
    if base_text_type_id == DB.ElementId.InvalidElementId:
        forms.alert("Nao encontrei tipo de Texto no projeto.", title="Importar Planilha Externa")
        return False

    col_count = max([len(r) for r in rows])
    text_mm = 2.0
    text_h = _mm_to_ft(text_mm)
    pad_x = _mm_to_ft(1.5)
    pad_y = _mm_to_ft(1.2)
    min_col_w = _mm_to_ft(18.0)
    max_col_w = _mm_to_ft(72.0)
    col_widths = []
    for c in range(col_count):
        if c < len(excel_col_widths) and excel_col_widths[c]:
            col_widths.append(max(min_col_w, min(max_col_w, float(excel_col_widths[c]) * _mm_to_ft(2.15))))
        else:
            max_len = 6
            for row in rows:
                text = row[c] if c < len(row) else ""
                if len(text) > max_len:
                    max_len = min(len(text), 42)
            col_widths.append(max(min_col_w, min(max_col_w, _mm_to_ft(max_len * 2.2 + 8.0))))

    row_h = max(_mm_to_ft(6.0), text_h * 2.4)
    x0 = 0.0
    y0 = 0.0
    base_name = "Tabela Excel - " + sanitize_filename(os.path.splitext(os.path.basename(file_path))[0])[:45]

    t = DB.Transaction(doc, "Importar Excel Externo como Tabela")
    t.Start()
    try:
        view, view_kind = _create_table_view(base_name)
        if not view:
            raise Exception("Nao encontrei tipo de Legend ou Vista de Desenho no projeto.")
        text_type_id = _get_or_create_table_text_type_id(text_mm)
        if text_type_id == DB.ElementId.InvalidElementId:
            text_type_id = base_text_type_id

        x_positions = [x0]
        for w in col_widths:
            x_positions.append(x_positions[-1] + w)
        total_w = x_positions[-1] - x0
        total_h = row_h * len(rows)

        # Grade da tabela
        for x in x_positions:
            line = DB.Line.CreateBound(DB.XYZ(x, y0, 0), DB.XYZ(x, y0 - total_h, 0))
            doc.Create.NewDetailCurve(view, line)
        for r in range(len(rows) + 1):
            y = y0 - (r * row_h)
            line = DB.Line.CreateBound(DB.XYZ(x0, y, 0), DB.XYZ(x0 + total_w, y, 0))
            doc.Create.NewDetailCurve(view, line)

        for r, row in enumerate(rows):
            y = y0 - (r * row_h) - pad_y - text_h
            for c, text in enumerate(row):
                if not text:
                    continue
                x = x_positions[c] + pad_x
                width = max(_mm_to_ft(5.0), col_widths[c] - (pad_x * 2.0))
                _create_cell_text_note(view, text_type_id, x, y, width, text)

        t.Commit()
        try:
            uidoc.ActiveView = view
        except:
            pass

        msg = "Planilha externa importada como tabela visual:\n\n{}".format(view.Name)
        msg += "\nTipo de vista: {}".format(view_kind)
        if used_print_area:
            msg += "\nArea de impressao do Excel aplicada."
        if truncated:
            msg += "\n\nObs.: limitei a importacao a 200 linhas e 20 colunas para manter a vista leve."
        forms.alert(msg, title="Importar Planilha Externa")
        return True
    except Exception as e:
        try:
            t.RollBack()
        except:
            pass
        logger.error(traceback.format_exc())
        forms.alert("Erro ao criar tabela visual: " + str(e), title="Importar Planilha Externa")
        return False

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

        # Eventos
        self.Button_BrowseExport.Click      += self.browse_export
        self.Button_BrowseImport.Click      += self.browse_import
        self.Button_Export.Click            += self.do_export
        self.Button_Import.Click            += self.do_import
        self.Button_Close.Click             += self.close_clicked
        self.ComboBox_ExportMode.SelectionChanged += self.mode_changed
        self.chk_SelectAllSchedules.Click   += self._on_select_all_changed
        self.CheckBox_KeepFormat.Click      += self._on_keep_format_changed
        self.TextBox_ScheduleSearch.TextChanged += self._on_schedule_search_changed
        self.Button_ClearScheduleSearch.Click += self._on_clear_schedule_search
        
        self.cmb_ExportPreviewSelect.SelectionChanged += lambda s, e: self._update_export_preview(from_combo=True)
        self.cmb_ImportPreviewSelect.SelectionChanged += lambda s, e: self._update_import_preview(from_combo=True)

        self.ContentRendered += self._on_window_loaded
    
    def update_status(self, msg): 
        try:
            self.TextBlock_Status.Text = u"Exporte e importe dados com eficiência"
        except:
            pass

    def update_stats(self, msg):
        try:
            self.TextBlock_Stats.Text = msg
        except:
            pass

    def _on_window_loaded(self, sender, args):
        try:
            self.load_schedules()
            self.mode_changed(None, None)
        except Exception as e:
            logger.error("Erro ao carregar schedules: " + str(e))

    def _build_suggested_name(self):
        """Gera nome sugerido baseado na seleção atual."""
        proj_name = sanitize_filename(doc.Title.replace(".rvt", ""))
        selected = self._get_selected_schedules()
        if selected and all([s.get('is_panel', False) for s in selected]):
            return proj_name + "_Quadros_de_Cargas"
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

    def _pick_export_file(self):
        """Abre janela de salvar usando a ultima pasta escolhida na sessao."""
        suggested_name = self._build_suggested_name()
        default_path = suggested_name + ".xlsx"
        if self.last_export_folder and os.path.exists(self.last_export_folder):
            default_path = os.path.join(self.last_export_folder, default_path)

        file_path = forms.save_file(file_ext="xlsx", default_name=default_path)
        if not file_path:
            return False

        self.last_export_folder = os.path.dirname(file_path)
        self.export_path = file_path
        self.TextBox_ExportPath.Text = file_path
        self.update_status("Pasta: " + self.last_export_folder)
        return True
    
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
                self.panel_schedules.append({
                    'schedule': v, 'display_name': v.Name,
                    'is_empty': False, 'is_panel': True
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
        try:
            self.TextBox_ScheduleSearch.Text = ""
        except:
            pass
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

    def _on_schedule_search_changed(self, sender, args):
        """Filtra a lista expandida de tabelas sem perder selecoes ja marcadas."""
        query = _normalize_text(self.TextBox_ScheduleSearch.Text)
        self.Button_ClearScheduleSearch.Visibility = (
            Visibility.Visible if query else Visibility.Collapsed
        )
        visible_enabled = []
        for cb, sch_dict in self._schedule_checkboxes:
            name = _normalize_text(sch_dict.get('display_name', ''))
            show = (not query) or (query in name)
            cb.Visibility = Visibility.Visible if show else Visibility.Collapsed
            if show and cb.IsEnabled:
                visible_enabled.append(cb)

        if visible_enabled:
            self.chk_SelectAllSchedules.Visibility = Visibility.Visible
            self.chk_SelectAllSchedules.IsEnabled = True
            self.chk_SelectAllSchedules.IsChecked = all([cb.IsChecked for cb in visible_enabled])
        else:
            self.chk_SelectAllSchedules.IsChecked = False
            self.chk_SelectAllSchedules.IsEnabled = False
            self.chk_SelectAllSchedules.Visibility = Visibility.Collapsed

        self._update_display_text()
        self._update_export_preview()

    def _on_clear_schedule_search(self, sender, args):
        """Limpa o filtro de busca da lista."""
        self.TextBox_ScheduleSearch.Text = ""
        try:
            self.TextBox_ScheduleSearch.Focus()
        except:
            pass

    def _on_select_all_changed(self, sender, args):
        """Marca ou desmarca todos os itens habilitados."""
        state = bool(self.chk_SelectAllSchedules.IsChecked)
        for cb, _ in self._schedule_checkboxes:
            if cb.IsEnabled and cb.Visibility == Visibility.Visible:
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
        self._update_export_preview()

    def _on_checklist_changed(self, sender, args):
        """Chamado quando qualquer checkbox da lista muda."""
        enabled = [(cb, d) for cb, d in self._schedule_checkboxes
                   if cb.IsEnabled and cb.Visibility == Visibility.Visible]
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
            self.update_stats("Nenhuma tabela selecionada")
        elif len(selected) == 1:
            is_panel_sel = selected[0].get('is_panel', False)
            try:
                sch = selected[0]['schedule']
                if is_panel_sel:
                    rows_p, _ = get_panel_schedule_data(sch)
                    self.update_stats("Quadro com {} circuitos".format(len(rows_p)))
                else:
                    count = sch.GetTableData().GetSectionData(DB.SectionType.Body).NumberOfRows
                    self.update_stats("Tabela com aprox. {} linhas".format(count))
            except:
                self.update_stats("1 tabela selecionada")
        else:
            if all([s.get('is_panel', False) for s in selected]):
                self.update_stats("{} quadros serao exportados na mesma aba".format(len(selected)))
            else:
                self.update_stats("{} tabelas serão exportadas como abas separadas".format(len(selected)))

    def browse_export(self, sender, args):
        self._pick_export_file()

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
                keep_revit_order = bool(self.CheckBox_KeepFormat.IsChecked)
                rows, headers = get_panel_schedule_data(schedule, preserve_revit_order=keep_revit_order)
                for h in headers:
                    dt.Columns.Add(h)
                
                for i, r in enumerate(rows):
                    if i >= max_preview_rows:
                        break
                    try:
                        if not r.get('Quant.') or str(r.get('Quant.')).strip() in ("", "0"):
                            r['Quant.'] = 1
                    except:
                        pass
                    row = dt.NewRow()
                    for j, h in enumerate(headers):
                        row[j] = str(r.get(h, ""))
                    dt.Rows.Add(row)
            elif bool(self.CheckBox_KeepFormat.IsChecked):
                visual_rows = get_schedule_visual_table(schedule)
                if visual_rows:
                    col_count = max([len(r) for r in visual_rows])
                    for c in range(col_count):
                        dt.Columns.Add("C{}".format(c + 1))
                    for i, values in enumerate(visual_rows):
                        if i >= max_preview_rows:
                            break
                        row = dt.NewRow()
                        for c in range(col_count):
                            row[c] = str(values[c]) if c < len(values) and values[c] is not None else ""
                        dt.Rows.Add(row)
                else:
                    self.lbl_PreviewStatus.Text = "Preview visual indisponivel para esta tabela."
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
        """Compara Excel com Revit e mostra apenas as alteracoes (Antes/Depois)."""
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

            wb = _XlsxReader(self.import_path)
            try:
                sheet_rows = []
                for s_name in wb.sheetnames:
                    if s_name.startswith('_'):
                        continue
                    rows = wb.read_sheet(s_name)
                    if rows:
                        sheet_rows.append((s_name, rows))

                if not sheet_rows:
                    self.lbl_ImportStatus.Text = "Nenhuma aba com dados encontrada."
                    return

                if first_load:
                    for s_name, _ in sheet_rows:
                        self.cmb_ImportPreviewSelect.Items.Add(s_name)
                    if sheet_rows:
                        self.cmb_ImportPreviewSelect.SelectedIndex = 0

                if self.cmb_ImportPreviewSelect.Items.Count > 1:
                    self.cmb_ImportPreviewSelect.Visibility = Visibility.Visible
                else:
                    self.cmb_ImportPreviewSelect.Visibility = Visibility.Collapsed

                idx = self.cmb_ImportPreviewSelect.SelectedIndex
                if idx < 0 or idx >= len(sheet_rows):
                    self.lbl_ImportStatus.Text = "Nenhuma aba selecionada."
                    return

                selected_sheet_name, rows = sheet_rows[idx]
                is_importable = bool(rows and rows[0] and str(rows[0][0] if rows[0] else "").strip() == "ElementId")

                if not is_importable:
                    self.lbl_ImportStatus.Text = u"Aba '{}': Excel externo. Ao importar, sera criada uma tabela visual em Vista de Desenho.".format(selected_sheet_name)
                    return

                headers = [str(h).strip() if h is not None else "" for h in rows[0]]
                pnames = [unit_postfix_pattern.sub("", h).strip() for h in headers[1:]]

                changes = []
                MAX_CHECK = 500
                checked = 0
                for row in rows[1:]:
                    if checked >= MAX_CHECK:
                        break
                    try:
                        if row[0] is None:
                            continue
                        eid = int(float(row[0]))
                        el = doc.GetElement(DB.ElementId(eid))
                        if not el:
                            continue
                        checked += 1
                        eid_str = str(eid)
                        for cx, pname in enumerate(pnames):
                            if not pname:
                                continue
                            val = row[cx + 1] if cx + 1 < len(row) else None
                            if val is None or val == "":
                                continue
                            param = get_param_robust(el, pname)
                            if not param or param.IsReadOnly:
                                continue
                            before_val = ""
                            try:
                                st = param.StorageType
                                if st == DB.StorageType.String:
                                    before_val = param.AsString() or ""
                                elif st == DB.StorageType.Integer:
                                    if is_yesno_parameter(param.Definition):
                                        before_val = "Yes" if param.AsInteger() else "No"
                                    else:
                                        vs = param.AsValueString()
                                        before_val = vs if vs else str(param.AsInteger())
                                elif st == DB.StorageType.Double:
                                    dt_t = get_parameter_data_type(param.Definition)
                                    if dt_t and DB.UnitUtils.IsMeasurableSpec(dt_t):
                                        uid = param.GetUnitTypeId()
                                        before_val = str(DB.UnitUtils.ConvertFromInternalUnits(param.AsDouble(), uid))
                                    else:
                                        before_val = str(param.AsDouble())
                            except:
                                pass
                            if isinstance(val, float) and val == int(val):
                                after_val = str(int(val))
                            else:
                                after_val = str(val)
                            if str(before_val).strip() != after_val.strip():
                                changes.append((eid_str, pname, str(before_val).strip(), after_val.strip()))
                    except:
                        pass
            finally:
                wb.close()

            total = len(changes)
            if total == 0:
                self.lbl_ImportStatus.Text = u"Sem alteracoes detectadas em '{}'.".format(selected_sheet_name)
                return

            THRESHOLD = 20
            if total >= THRESHOLD:
                self.lbl_ImportStatus.Text = u"{} alteracoes detectadas em '{}' — muitas para exibir individualmente.".format(total, selected_sheet_name)
                return

            dt = DataTable()
            dt.Columns.Add("ElementId")
            dt.Columns.Add(u"Parametro")
            dt.Columns.Add("Antes")
            dt.Columns.Add("Depois")
            for eid_str, pname, before_val, after_val in changes:
                row_dt = dt.NewRow()
                row_dt[0] = eid_str
                row_dt[1] = pname
                row_dt[2] = before_val
                row_dt[3] = after_val
                dt.Rows.Add(row_dt)

            self.dg_ImportPreview.ItemsSource = dt.DefaultView
            self.dg_ImportPreview.Visibility = Visibility.Visible
            self.lbl_ImportStatus.Text = u"{} alteracoes em '{}'".format(total, selected_sheet_name)
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
            if not self._pick_export_file():
                return

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
                formatted = bool(self.CheckBox_KeepFormat.IsChecked)

                for i, sch_dict in enumerate(selected):
                    _pb_update(i)
                    src     = []
                    params  = []
                    is_panel = sch_dict.get('is_panel', False)
                    name    = sch_dict['display_name']
                    sch     = sch_dict['schedule']

                    if is_panel:
                        rows, headers = get_panel_schedule_data(sch, preserve_revit_order=formatted)
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

                if targets and all([t.get('is_panel', False) for t in targets]):
                    combined_src = []
                    combined_params = targets[0]['params']
                    for target in targets:
                        combined_src.extend(target['src'])
                    targets = [{
                        'name': 'Quadros de Cargas',
                        'src': combined_src,
                        'params': combined_params,
                        'is_panel': True
                    }]

                if not targets:
                    self.update_status("Nada para exportar.")
                    msg = "Nenhuma tabela com dados para exportar."
                    if skipped_names:
                        msg += "\n\nTabelas vazias ignoradas:\n- " + "\n- ".join(skipped_names)
                    return forms.alert(msg)

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
                selected_sheet = None
                try:
                    if self.cmb_ImportPreviewSelect.SelectedIndex >= 0:
                        selected_sheet = str(self.cmb_ImportPreviewSelect.SelectedItem)
                except:
                    pass
                ok = forms.alert(
                    "Este arquivo nao tem ElementId, entao nao da para atualizar elementos do Revit.\n\n"
                    "Posso importar a aba selecionada como uma tabela visual em uma Vista de Desenho.\n"
                    "Deseja criar essa tabela agora?",
                    title="Importar Planilha Externa",
                    yes=True,
                    no=True
                )
                if ok:
                    import_external_xlsx_as_drafting_view(self.import_path, selected_sheet)
                    self._update_import_preview(first_load=True)
                return
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
