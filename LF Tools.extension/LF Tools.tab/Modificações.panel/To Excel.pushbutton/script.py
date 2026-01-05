# -*- coding: utf-8 -*-
import os
import xlrd
import xlsxwriter
import re
import traceback
import gc
import sys
from collections import namedtuple, OrderedDict
from pyrevit import script, forms, coreutils, revit, DB

# ==================== CONFIGURAÇÃO DE PERFORMANCE ====================
GC_THRESHOLD = (700, 10, 10)  # (limiar1, limiar2, limiar3) para GC mais agressivo
MEMORY_MONITOR_INTERVAL = 1000  # Verificar memória a cada 1000 elementos
MAX_CACHE_SIZE = 2000  # Tamanho máximo do cache LRU
MAX_ELEMENT_CACHE_SIZE = 500  # Tamanho máximo para cache de elementos

# Logger e Documento
logger = script.get_logger()
doc = revit.doc
project_units = doc.GetUnits()

# Limpar output
script.get_output().close_others()

unit_postfix_pattern = re.compile(r"\s*\[.*\]$")

# ==================== CACHE SIMPLES E OTIMIZADO ====================
# Substituindo implementacoes complexas por dicts padrao do Python (muito mais rapidos)
_param_def_cache = {}
_element_cache = {}
_format_pool = {}

# ==================== GERENCIAMENTO DE MEMÓRIA SEGURO ====================
def setup_memory_optimization():
    """Configura otimizações de memória seguras para IronPython."""
    try:
        # Habilita garbage collector (mais seguro para IronPython)
        gc.enable()
        
        # Define estratégia de coleta (se disponível)
        if hasattr(gc, 'set_threshold'):
            try:
                gc.set_threshold(*GC_THRESHOLD)
            except:
                # Fallback para valores padrão
                pass
    except:
        # Em caso de erro, mantém configuração padrão
        pass

def smart_gc_collect(force=False, threshold=50000):
    """
    Coleta de garbage apenas quando estritamente necessario.
    No IronPython, chamar GC.Collect() frequentemente causa overhead massivo.
    """
    if force:
        try:
            gc.collect()
        except:
            pass
    return False

def clear_all_caches():
    """Limpa todos os caches de forma segura."""
    try:
        _param_def_cache.clear()
        _element_cache.clear()
        _format_pool.clear()
    except:
        # Ignora erros na limpeza de cache
        pass
    
    # Coleta removida para evitar travamento no fechamento
    pass

def get_memory_status():
    """Retorna status de memória para logging."""
    try:
        if hasattr(gc, 'get_count'):
            count = gc.get_count()
            return "GC counts: G0={}, G1={}, G2={}".format(count[0], count[1], count[2])
    except:
        pass
    return "Memória: OK"

# ==================== ESTRUTURAS DE DADOS OTIMIZADAS ====================
ParamDef = namedtuple(
    "ParamDef", ["name", "istype", "definition", "isreadonly", "isunit", "storagetype"]
)

def sanitize_filename(name):
    """Remove caracteres invalidos para nome de arquivo."""
    return re.sub(r'[\\/*?:"<>|]', "", name)

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
    if not elements:
        try:
            collector = DB.FilteredElementCollector(doc, schedule.Id)
            collector = collector.WhereElementIsNotElementType()
            
            for el in collector:
                try:
                    if el.Id not in element_ids:
                        element_ids.add(el.Id)
                        elements.append(el)
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
    
    for field in visible_fields:
        field_name = field.GetName()
        if not field_name:
            continue
        
        # Verificar cache de definição
        param_cache_key = ("param_def", field_name)
        cached_param_def = _param_def_cache.get(param_cache_key)
        
        if cached_param_def:
            param_defs_list.append(cached_param_def)
            continue
        
        param_id = field.ParameterId
        param_info = None
        
        # Usar primeiro elemento
        first_element = elements[0]
        
        if param_id and param_id != DB.ElementId.InvalidElementId:
            try:
                param_element = doc.GetElement(param_id)
                if param_element and hasattr(param_element, 'Definition'):
                    param_definition = param_element.Definition
                    param_info = first_element.get_Parameter(param_definition)
            except:
                pass
        
        if not param_info:
            param_info = first_element.LookupParameter(field_name)
            
            if not param_info and hasattr(first_element, 'Parameters'):
                for p in first_element.Parameters:
                    try:
                        if p.Definition.Name == field_name:
                            param_info = p
                            break
                    except:
                        continue
        
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
    
    # Extrair valor
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
    
    # Armazenar em cache
    param_cache[cache_key] = value
    
    # Limpeza periódica do cache
    if len(param_cache) > 1000:
        param_cache.clear()
    
    return value

def get_parameter_value(param, param_def=None):
    """Obtém valor formatado de parâmetro otimizado."""
    try:
        st = param.StorageType
        if st == DB.StorageType.Double:
            dt = get_parameter_data_type(param.Definition)
            if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                uid = param.GetUnitTypeId()
                return str(DB.UnitUtils.ConvertFromInternalUnits(param.AsDouble(), uid))
            else:
                return str(param.AsDouble())
        elif st == DB.StorageType.String: 
            return param.AsString() or ""
        elif st == DB.StorageType.Integer:
            if is_yesno_parameter(param.Definition): 
                return "Yes" if param.AsInteger() else "No"
            else: 
                return str(param.AsInteger())
        elif st == DB.StorageType.ElementId: 
            return param.AsValueString() or ""
    except: 
        return ""
    
    return ""

# ==================== EXPORTAÇÃO OTIMIZADA ====================
def export_xls(src_elements, selected_params, file_path, is_panel_schedule=False):
    """Exporta dados para Excel com otimização de memória e performance."""
    workbook = None
    try:
        # Configurar workbook com otimizações
        workbook_options = {
            'constant_memory': True,
            'use_zip64': True,
        }
        
        workbook = xlsxwriter.Workbook(file_path, workbook_options)
        
        if is_panel_schedule:
            ws = workbook.add_worksheet("Quadro de Cargas")
            
            # Usar pool de formatos
            fmt_head = get_excel_format(workbook, {
                "bold": True, 
                "bg_color": "#4F81BD", 
                "font_color": "white", 
                "border": 1, 
                "align": "center"
            })
            fmt_data = get_excel_format(workbook, {"border": 1, "align": "left"})
            
            # Escrever headers
            for i, p in enumerate(selected_params):
                ws.write(0, i, p.name, fmt_head)
                ws.set_column(i, i, 20)
            
            # Escrever dados
            for r, el in enumerate(src_elements, 1):
                if hasattr(el, 'row_data'):
                    for c, p in enumerate(selected_params):
                        ws.write(r, c, str(el.row_data.get(p.name, "")), fmt_data)
            
            if src_elements:
                ws.autofilter(0, 0, len(src_elements), len(selected_params)-1)
            
        else:
            ws = workbook.add_worksheet("Export")
            
            # Pool de formatos
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
            
            ws.freeze_panes(1, 0)
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
            
            # Determinar batch size dinâmico
            if total_elements > 10000:
                batch_size = 250
            elif total_elements > 5000:
                batch_size = 500
            elif total_elements > 1000:
                batch_size = 1000
            else:
                batch_size = 2000
            
            # Cache por lote
            param_cache = {}
            processed_count = 0
            
            for batch_start in range(0, total_elements, batch_size):
                batch_end = min(batch_start + batch_size, total_elements)
                
                for r_offset, el in enumerate(src_elements[batch_start:batch_end]):
                    r = batch_start + r_offset + 1
                    try:
                        eid = el.Id.IntegerValue
                        ws.write(r, 0, str(eid), fmt_lock_id)
                        
                        # Obter valores com cache
                        for c, p in enumerate(selected_params):
                            value = get_element_parameter_value(el, p, param_cache)
                            
                            fmt = fmt_lock_ro if p.isreadonly else fmt_unlock
                            ws.write(r, c+1, value, fmt)
                            
                            # Atualizar largura
                            slen = len(str(value)) if value else 0
                            if slen > widths[c+1]: 
                                widths[c+1] = min(slen, 50)
                    except Exception as e:
                        logger.error("Erro ao processar elemento " + str(r) + ": " + str(e))
                        continue
                
                processed_count += (batch_end - batch_start)
                
                # Limpeza inteligente de cache removida (Python dict gerencia memoria melhor sem interferencia)
                if len(param_cache) > 5000:
                    param_cache.clear()
                
                # Coleta removida
                pass
                
                # Status periódico
                if processed_count % 1000 == 0:
                    logger.debug("Processados " + str(processed_count) + " de " + str(total_elements) + " elementos")
            
            # Aplicar larguras
            for i, w in enumerate(widths): 
                ws.set_column(i, i, w+3)
            
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
    """Importa dados do Excel com busca robusta de parametros."""
    wb = None
    try:
        wb = xlrd.open_workbook(file_path, on_demand=True)
        
        if "Quadro de Cargas" in wb.sheet_names():
            forms.alert("Arquivos de Quadro de Cargas sao somente leitura.", title="Aviso")
            return
        
        sh = wb.sheet_by_name("Export")
    except:
        forms.alert("Erro ao abrir Excel ou aba 'Export' nao encontrada.", title="Erro")
        return
    finally:
        if wb:
            try:
                wb.release_resources()
            except:
                pass
    
    headers = sh.row_values(0)
    if not headers or str(headers[0]).strip() != "ElementId":
        forms.alert("A celula A1 deve conter 'ElementId'.", title="Erro de Formato")
        return
    
    # Processar nomes de parâmetros
    pnames = [unit_postfix_pattern.sub("", h).strip() for h in headers[1:]]
    
    # Cache local (dict simples)
    import_element_cache = {}
    
    # Estatisticas
    stats = {
        'total_rows': sh.nrows - 1,
        'success_cells': 0,
        'skipped_cells': 0,
        'readonly_cells': 0,
        'not_found_cells': 0,
        'error_cells': 0,
        'elements_processed': 0
    }
    
    with revit.Transaction("Import Excel"):
        total_rows = sh.nrows - 1
        batch_size = 200
        
        for batch_start in range(1, sh.nrows, batch_size):
            batch_end = min(batch_start + batch_size, sh.nrows)
            
            for rx in range(batch_start, batch_end):
                row = sh.row_values(rx)
                try:
                    eid = int(float(row[0]))
                    
                    # Usar cache simples
                    el = import_element_cache.get(eid)
                    if not el:
                        el = doc.GetElement(DB.ElementId(eid))
                        if el:
                            import_element_cache[eid] = el
                    
                    if not el: 
                        continue
                    
                    stats['elements_processed'] += 1
                    
                    for cx, pname in enumerate(pnames):
                        val = row[cx+1]
                        
                        # Ignorar celulas vazias no Excel
                        if val is None or val == "": 
                            continue
                        
                        # Busca robusta (Instancia + Tipo)
                        param = get_param_robust(el, pname)
                        
                        if not param:
                            stats['not_found_cells'] += 1
                            continue
                            
                        if param.IsReadOnly: 
                            stats['readonly_cells'] += 1
                            continue
                        
                        try:
                            # Tenta definir valor
                            st = param.StorageType
                            success = False
                            changed = False
                            
                            if st == DB.StorageType.String:
                                current_val = param.AsString() or ""
                                
                                # Remover .0 de numeros inteiros
                                if isinstance(val, float) and val.is_integer():
                                    new_val = str(int(val))
                                else:
                                    new_val = str(val)
                                    
                                if current_val != new_val:
                                    success = param.Set(new_val)
                                    changed = True
                                    
                            elif st == DB.StorageType.Integer:
                                current_val = param.AsInteger()
                                new_val = None
                                
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
                                
                                # Converter apenas se necessario (logica simplificada para evitar dupla conversao errada)
                                if dt and DB.UnitUtils.IsMeasurableSpec(dt):
                                    uid = param.GetUnitTypeId()
                                    new_val_internal = DB.UnitUtils.ConvertToInternalUnits(new_val, uid)
                                    
                                    # Comparacao com tolerancia
                                    if abs(current_val - new_val_internal) > 0.0001:
                                        success = param.Set(new_val_internal)
                                        changed = True
                                else: 
                                    if abs(current_val - new_val) > 0.0001:
                                        success = param.Set(new_val)
                                        changed = True
                                        
                            elif st == DB.StorageType.ElementId:
                                pass
                                
                            if changed:
                                if success:
                                    stats['success_cells'] += 1
                                else:
                                    stats['error_cells'] += 1 # Mudou mas falhou
                            else:
                                stats['skipped_cells'] += 1 # Valor igual
                                
                        except: 
                            stats['error_cells'] += 1
                except: 
                    pass
            
            # Limpeza periódica dict
            if len(import_element_cache) > 2000:
                import_element_cache.clear()
            
            # GC check
            if (batch_end % 1000) == 0:
                smart_gc_collect()
    
    # Limpeza final
    import_element_cache.clear()
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
        
        # Configurar otimizações
        setup_memory_optimization()
        
        self.export_path = None
        self.import_path = None
        self.last_export_folder = None
        self.view_schedules = []
        self.panel_schedules = []
        
        # Registrar eventos
        self.Button_BrowseExport.Click += self.browse_export
        self.Button_BrowseImport.Click += self.browse_import
        self.Button_Export.Click += self.do_export
        self.Button_Import.Click += self.do_import
        self.Button_Close.Click += self.close_clicked
        self.ComboBox_ExportMode.SelectionChanged += self.mode_changed
        self.ComboBox_SubSelection.SelectionChanged += self.sub_selection_changed
        
        self.load_schedules()
        self.mode_changed(None, None)
    
    def update_status(self, msg): 
        self.TextBlock_Status.Text = msg

    def update_stats(self, msg):
        try:
            self.TextBlock_Stats.Text = msg
        except:
            pass
    
    def load_schedules(self):
        """Carrega schedules com cache."""
        self.view_schedules = []
        self.panel_schedules = []
        
        try:
            # Tabelas de Quantidade
            view_collector = DB.FilteredElementCollector(doc)\
                .OfClass(DB.ViewSchedule)\
                .WhereElementIsNotElementType()
            
            for s in view_collector:
                if s.IsTemplate: 
                    continue
                if "PanelSchedule" in s.GetType().Name: 
                    continue
                try: 
                    if s.Definition.IsInternalKeynoteSchedule or s.Definition.IsRevisionSchedule: 
                        continue
                except: 
                    pass
                
                self.view_schedules.append({
                    'schedule': s, 
                    'display_name': s.Name
                })
            
            # Quadros de Carga
            panel_collector = DB.FilteredElementCollector(doc)\
                .OfClass(DB.View)\
                .WhereElementIsNotElementType()
            
            for v in panel_collector:
                if "PanelScheduleView" in v.GetType().Name and not v.IsTemplate:
                    self.panel_schedules.append({
                        'schedule': v, 
                        'display_name': v.Name
                    })
            
            self.view_schedules.sort(key=lambda x: x['display_name'])
            self.panel_schedules.sort(key=lambda x: x['display_name'])
            
        except Exception as e:
            logger.error("Erro ao carregar schedules: " + str(e))
        finally:
            smart_gc_collect()

    def mode_changed(self, sender, args):
        idx = self.ComboBox_ExportMode.SelectedIndex
        self.ComboBox_SubSelection.Items.Clear()
        self.ComboBox_SubSelection.IsEnabled = True
        
        if idx == 0:  # Tabela de Quantidades
            if not self.view_schedules: 
                self.ComboBox_SubSelection.Items.Add("Nenhuma tabela")
            else: 
                for i in self.view_schedules: 
                    self.ComboBox_SubSelection.Items.Add(i['display_name'])
            self.ComboBox_SubSelection.SelectedIndex = 0
        elif idx == 1:  # Quadro de Cargas
            if not self.panel_schedules: 
                self.ComboBox_SubSelection.Items.Add("Nenhum quadro")
            else:
                for i in self.panel_schedules: 
                    self.ComboBox_SubSelection.Items.Add(i['display_name'])
            self.ComboBox_SubSelection.SelectedIndex = 0

    def sub_selection_changed(self, sender, args):
        """Atualiza estatisticas quando a selecao muda."""
        try:
            idx = self.ComboBox_ExportMode.SelectedIndex
            sub_idx = self.ComboBox_SubSelection.SelectedIndex
            
            if sub_idx < 0:
                self.update_stats("")
                return
            
            count = 0
            if idx == 0:  # Tabela
                if self.view_schedules:
                    sch = self.view_schedules[sub_idx]['schedule']
                    # Metodo rapido de contagem
                    try:
                        # Tenta pegar contagem sem abrir a tabela toda
                        data = sch.GetTableData().GetSectionData(DB.SectionType.Body)
                        count = data.NumberOfRows
                    except:
                        # Fallback simples
                        count = "?"
                        
                    self.update_stats("Tabela com aprox. {} linhas".format(count))
            
            elif idx == 1:  # Quadro
                if self.panel_schedules:
                    # Quadros sao tipicamente pequenos
                    self.update_stats("Quadro de Cargas (Tamanho Estimado)")
                    
        except:
            self.update_stats("")

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
                    suggested_name = proj_name + "_" + clean_item
            except: 
                pass
            
        # Usar ultima pasta se disponivel
        default_path = suggested_name + ".xlsx"
        if self.last_export_folder and os.path.exists(self.last_export_folder):
            default_path = os.path.join(self.last_export_folder, default_path)
            
        file_path = forms.save_file(file_ext="xlsx", default_name=default_path)
        if file_path:
            # Salvar pasta para proxima vez
            self.last_export_folder = os.path.dirname(file_path)
            
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
        if not self.export_path: 
            return forms.alert("Defina o local do arquivo.")
        
        idx = self.ComboBox_ExportMode.SelectedIndex
        src = []
        params = []
        is_panel = False
        
        try:
            self.update_status("Lendo dados...")
            
            if idx == 0:  # Tabela de Quantidades
                if self.view_schedules:
                    sch = self.view_schedules[self.ComboBox_SubSelection.SelectedIndex]['schedule']
                    src, params = get_schedule_elements_and_params(sch)
            
            elif idx == 1:  # Quadro de Cargas
                if self.panel_schedules:
                    sch = self.panel_schedules[self.ComboBox_SubSelection.SelectedIndex]['schedule']
                    is_panel = True
                    rows, headers = get_panel_schedule_data(sch)
                    
                    class RowObj:
                        def __init__(self, d): 
                            self.row_data = d
                    
                    src = [RowObj(r) for r in rows]
                    params = [ParamDef(
                        name=h, 
                        istype=False, 
                        definition=None, 
                        isreadonly=True, 
                        isunit=False, 
                        storagetype=DB.StorageType.String
                    ) for h in headers]

            if not src: 
                self.update_status("Nada para exportar.")
                return

            self.update_status("Gerando Excel...")
            export_xls(src, params, self.export_path, is_panel)
            
            self.update_status("Concluido!")
            forms.alert("Exportacao finalizada com sucesso!", title="Sucesso")
            
        except Exception as e:
            logger.error(traceback.format_exc())
            forms.alert("Erro: " + str(e))
        finally:
            src = None
            params = None
            # GC automatico do Python/Revit cuidara do resto
            pass

    def do_import(self, sender, args):
        if not self.import_path: 
            return forms.alert("Selecione o arquivo.")
        
        self.update_status("Importando...")
        try:
            import_xls(self.import_path)
            self.update_status("Importacao Finalizada!")
            forms.alert("Importacao concluida!", title="Sucesso")
        except Exception as e:
            logger.error(traceback.format_exc())
            forms.alert("Erro na importacao: " + str(e))
        finally:
            # GC automatico
            pass

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