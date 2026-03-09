# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ElementId, StorageType,
    BuiltInParameter, Element, UnitUtils, UnitTypeId
)
from System.Collections.Generic import List
import traceback
import os
import json
import io
import System
import sys
from System.Windows.Forms import Application as WinFormsApp
import re
import datetime

# ==================== LOGICA DE HISTORICO ====================
class SelectionHistory:
    def __init__(self):
        # Pasta de dados do usuario
        appdata = os.getenv('APPDATA')
        self.history_dir = os.path.join(appdata, 'pyRevit', 'Extensions', 'LFTools')
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir)
            
        self.history_file = os.path.join(self.history_dir, 'filtro_avancado_history.json')
        self.history = self.load_history()
        self.max_size = 15
    
    def load_history(self):
        if os.path.exists(self.history_file):
            try:
                with io.open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return []
    
    def save_history(self):
        try:
            with io.open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, indent=4, ensure_ascii=False)
        except:
            pass
    
    def add(self, element_ids, category, criteria_desc, count, state):
        entry = {
            'element_ids': [eid.IntegerValue for eid in element_ids],
            'category': category,
            'criteria': criteria_desc,
            'count': count,
            'timestamp': datetime.datetime.now().strftime("%H:%M:%S"),
            'state': state
        }
        # Evitar duplicados consecutivos idênticos
        if self.history and self.history[0]['criteria'] == criteria_desc and self.history[0]['category'] == category:
            self.history.pop(0)
            
        self.history.insert(0, entry)
        if len(self.history) > self.max_size:
            self.history = self.history[:self.max_size]
        self.save_history()

    def clear(self):
        self.history = []
        self.save_history()



# --- HELPER UNICODE PARA IRONPYTHON 2.7 (CODEC-FREE) ---
def to_unicode(val):
    """
    Converte para unicode SEM usar o sistema de codecs do Python (decode/encode),
    pois o registro de encodings parece estar corrompido no ambiente do usuario.
    """
    if val is None: return u""
    if isinstance(val, unicode): return val
    
    # Se ja for tipo basico
    if isinstance(val, (int, float, bool, System.Guid)):
        return unicode(val)

    # Se for string .NET ou byte string python (str)
    # No IronPython 2.7, str eh basicamente um array de bytes
    if isinstance(val, (str, bytes)):
        try:
            # Mapeamento Manual (0-255 Latin-1) - Nao faz lookup de codec
            return u"".join([unichr(ord(c)) for c in val])
        except:
            pass
    
    # Fallback para objetos .NET (como Element.Name)
    try:
        res = val.ToString() if hasattr(val, "ToString") else unicode(val)
        # Se o ToString retornar uma byte string com acento, limpamos de novo
        if isinstance(res, str):
            return u"".join([unichr(ord(c)) for c in res])
        return unicode(res)
    except:
        return u""

def force_unicode(data):
    """Garante recursivamente que TUDO eh unicode purista para o json.dumps"""
    if isinstance(data, dict):
        new_dict = {}
        for k, v in data.items():
            k_u = to_unicode(k)
            v_u = force_unicode(v)
            new_dict[k_u] = v_u
        return new_dict
    elif isinstance(data, (list, tuple)):
        return [force_unicode(v) for v in data]
    else:
        return to_unicode(data)

def safe_unicode_inspect(obj, path="root"):
    """Inspeciona recursivamente um objeto para achar problemas de encode"""
    issues = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            issues.extend(safe_unicode_inspect(k, "{}.KEY".format(path)))
            issues.extend(safe_unicode_inspect(v, "{}.{}".format(path, k)))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            issues.extend(safe_unicode_inspect(v, "{}[{}]".format(path, i)))
    elif isinstance(obj, (str, unicode)):
        try:
            json.dumps(obj)
        except Exception as e:
            hex_val = " ".join([hex(ord(c)) for c in str(obj)]) if isinstance(obj, str) else "unicode_str"
            issues.append(u"ERRO em '{}': {} | Tipo: {} | Hex: {}".format(path, unicode(e), type(obj), hex_val))
    return issues

# ARQUIVO DE PRESETS (Pasta segura) - Garantir Unicode no IronPython usando .NET
def get_config_path():
    try:
        appdata = os.getenv('APPDATA')
        path = os.path.join(appdata, 'pyRevit', 'Extensions', 'LFTools')
        if not os.path.exists(path):
            os.makedirs(path)
        return path
    except Exception as e:
        logger.error("Erro ao criar pasta config: " + str(e))
        return os.path.dirname(__file__)

CONFIG_DIR = get_config_path()
PRESETS_FILE = os.path.join(CONFIG_DIR, "filtro_avancado_presets.json")

def load_presets():
    default_data = {"LastUsed": "", "Presets": {}}
    try:
        if os.path.exists(PRESETS_FILE):
            with io.open(PRESETS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        logger.error("Erro ao carregar presets: " + str(e))
    return default_data

def save_presets(data):
    try:
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        with io.open(PRESETS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        logger.error("Erro ao salvar presets: " + str(e))
        forms.alert("Erro ao salvar configuração: " + str(e))
        return False

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# ==================== HELPERS DE ELEVAÇÃO ====================
def get_element_elevation(el):
    """Obtém a elevação Z do elemento convertida de pés para metros."""
    z_feet = None
    
    # 1. Tentar Location.Point (elementos pontuais: devices, equipment, etc.)
    try:
        loc = el.Location
        if loc and hasattr(loc, 'Point') and loc.Point:
            z_feet = loc.Point.Z
    except:
        pass
    
    # 2. Tentar Location.Curve midpoint (elementos lineares: conduit, pipe, duct)
    if z_feet is None:
        try:
            loc = el.Location
            if loc and hasattr(loc, 'Curve') and loc.Curve:
                mid = loc.Curve.Evaluate(0.5, True)
                z_feet = mid.Z
        except:
            pass
    
    # 3. Fallback: BoundingBox center
    if z_feet is None:
        try:
            bb = el.get_BoundingBox(None)
            if bb:
                z_feet = (bb.Min.Z + bb.Max.Z) / 2.0
        except:
            pass
    
    if z_feet is None:
        return None
    
    # Converter de pés internos para metros
    try:
        z_meters = UnitUtils.ConvertFromInternalUnits(z_feet, UnitTypeId.Meters)
    except:
        # Fallback manual: 1 pé = 0.3048 metros
        z_meters = z_feet * 0.3048
    
    return z_meters

def check_elevation_condition(z_meters, condition, val_min, val_max=None, tolerance=0.01):
    """Verifica condição de elevação. Valores em metros."""
    if z_meters is None:
        return False
    
    if condition == u"Igual a (\u00b11cm)":
        return abs(z_meters - val_min) <= tolerance
    elif condition == u"Maior que":
        return z_meters > val_min
    elif condition == u"Menor que":
        return z_meters < val_min
    elif condition == u"Maior ou igual":
        return z_meters >= val_min
    elif condition == u"Menor ou igual":
        return z_meters <= val_min
    return False

def smart_parameter_compare(param, condition, user_text, text_compare_func):
    """
    Compara o parâmetro com o input do usuário de forma inteligente.
    Se o parâmetro for numérico (Double), tenta converter e comparar como número com regras de unidade.
    Se falhar ou não for numérico, usa a comparação de texto original.
    """
    if condition == u"Em branco":
        return not param.HasValue
    if condition == u"Não em branco":
        return param.HasValue
        
    try:
        # Tenta lógica numérica APENAS se for StorageType.Double
        if param.StorageType == StorageType.Double:
            internal_val = param.AsDouble()
            
            # Tentar ler o input do usuário (esperando metros se for comprimento)
            user_val_str = user_text.replace(',', '.').strip()
            if user_val_str:
                user_val = float(user_val_str)
                
                # Obter o tipo de unidade do parâmetro
                try:
                    # Tenta converter o valor interno (pés) para unidades de exibição (ex: metros)
                    # UnitUtils.ConvertFromInternalUnits foi introduzido nas versões mais recentes
                    val_in_meters = UnitUtils.ConvertFromInternalUnits(internal_val, param.GetUnitTypeId())
                except:
                    # Fallback manual assumindo pés -> metros se GetUnitTypeId falhar
                    val_in_meters = internal_val * 0.3048
                    
                tolerance = 0.01 # Tolerância numérica (ex: 1cm se em metros)
                
                if condition == u"Igual a":
                    return abs(val_in_meters - user_val) <= tolerance
                elif condition == u"Maior que":
                    return val_in_meters > user_val
                elif condition == u"Menor que":
                    return val_in_meters < user_val
                elif condition == u"Maior ou igual":
                    return val_in_meters >= (user_val - tolerance)
                elif condition == u"Menor ou igual":
                    return val_in_meters <= (user_val + tolerance)
                elif condition == u"Diferente de":
                    return abs(val_in_meters - user_val) > tolerance
    except Exception as e:
        # Se falhar a extração ou parse float, segue pro fallback textual
        pass
        
    # --- Fallback: Texto ---
    if not param.HasValue:
        return False
        
    param_txt = param.AsValueString()
    if not param_txt:
        # Alguns parâmetros (ex: texto simples) usam AsString() em vez de AsValueString()
        if param.StorageType == StorageType.String:
            param_txt = param.AsString()
        else:
            param_txt = ""
            
    return text_compare_func(param_txt, condition, user_text)


class FiltroAvancadoWindow(forms.WPFWindow):
    def __init__(self):
        try:
            xaml_file = script.get_bundle_file('FiltroAvancado.xaml')
            if xaml_file and os.path.exists(xaml_file):
                forms.WPFWindow.__init__(self, xaml_file)
            else:
                forms.WPFWindow.__init__(self, self.get_fallback_xaml())
            
            # Cache de parâmetros para evitar recalcular
            self._parametros_cache = {}
            
            # Configurações iniciais
            self.categoria_opcoes = {
                u"Conduítes e Curvas": {
                    "categorias": [BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting],
                    "classes": ["Conduit", "FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None,
                    "exclude_if_not": "angle_or_conduit"
                },
                u"Conexões de conduíte": {
                    "categorias": [BuiltInCategory.OST_ConduitFitting],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Isolamento de Tubo": {
                    "categorias": [BuiltInCategory.OST_PipeInsulations],
                    "classes": ["PipeInsulation"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Dutos": {
                    "categorias": [BuiltInCategory.OST_DuctCurves],
                    "classes": ["Duct"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Bandeja de cabos": {
                    "categorias": [BuiltInCategory.OST_CableTray],
                    "classes": ["CableTray"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Tubulações": {
                    "categorias": [BuiltInCategory.OST_PipeCurves],
                    "classes": ["Pipe"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Conexões de Conduíte (sem RN)": {
                    "categorias": [BuiltInCategory.OST_ConduitFitting],
                    "classes": [],
                    "requires_param": None,
                    "requires_param_absent": "RN"
                },
                u"Dispositivos Elétricos": {
                    "categorias": [BuiltInCategory.OST_ElectricalFixtures],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Equipamentos Elétricos": {
                    "categorias": [BuiltInCategory.OST_ElectricalEquipment],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Anotações Genéricas": {
                    "categorias": [BuiltInCategory.OST_GenericAnnotation],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Aparelhos de Iluminação": {
                    "categorias": [BuiltInCategory.OST_LightingFixtures],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Equipamentos Mecânicos": {
                    "categorias": [BuiltInCategory.OST_MechanicalEquipment],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Acessórios de Tubulação": {
                    "categorias": [BuiltInCategory.OST_PipeFitting],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Acessórios de Dutos": {
                    "categorias": [BuiltInCategory.OST_DuctFitting],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Válvulas de Tubulação": {
                    "categorias": [BuiltInCategory.OST_PipeAccessory],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                u"Acessórios de Dutos (Registros, etc.)": {
                    "categorias": [BuiltInCategory.OST_DuctAccessory],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                }
            }
            
            # Inicializar Histórico
            self.selection_history = SelectionHistory()
            
            # Set de IDs já mapeados para evitar duplicidade na injeção
            self.ids_mapeados = set()
            for info in self.categoria_opcoes.values():
                for cat in info["categorias"]:
                    try:
                        if isinstance(cat, BuiltInCategory):
                            self.ids_mapeados.add(int(cat))
                        elif hasattr(cat, "IntegerValue"):
                            self.ids_mapeados.add(cat.IntegerValue)
                        elif hasattr(cat, "Value"):
                            self.ids_mapeados.add(int(cat.Value))
                    except:
                        pass

            self.inicializar_controles()
            
        except Exception as e:
            logger.error("Erro no init: {}".format(traceback.format_exc()))
            forms.alert("Erro ao carregar interface. Verifique o arquivo XAML.")
            raise
    
    def get_fallback_xaml(self):
        """XAML de fallback caso o arquivo não seja encontrado"""
        return '''
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Filtro Avançado - Fallback" Height="400" Width="500">
            <Grid Margin="20">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Filtro Avançado (Modo Simplificado)" 
                          FontSize="16" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,20"/>
                
                <TextBlock Grid.Row="1" Text="Arquivo de interface não encontrado. Use o modo padrão do pyRevit." 
                          TextWrapping="Wrap" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                
                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="Button_AplicarFiltro" Content="Aplicar Filtro" Background="#27AE60" 
                           Foreground="White" Width="100" Height="30" Margin="0,0,10,0"/>
                    <Button x:Name="Button_Fechar" Content="Fechar" Background="#95A5A6" 
                           Foreground="White" Width="80" Height="30"/>
                </StackPanel>
            </Grid>
        </Window>
        '''
    
    def inicializar_controles(self):
        """Inicializa os controles da interface - OTIMIZADO"""
        try:
            # Preencher categorias (operação leve)
            self.ComboBox_Categoria.Items.Clear()
            for categoria in sorted(self.categoria_opcoes.keys()):
                self.ComboBox_Categoria.Items.Add(categoria)
            
            # ⚡ OTIMIZAÇÃO: Conectar eventos APÓS preencher
            self.ComboBox_Categoria.SelectionChanged += self.categoria_selecionada
            self.CheckBox_UsarSegundoFiltro.Checked += self.segundo_filtro_alterado
            self.CheckBox_UsarSegundoFiltro.Unchecked += self.segundo_filtro_alterado
            self.Button_AplicarFiltro.Click += self.aplicar_filtro_click
            self.Button_Fechar.Click += self.fechar_click
            
            # --- Sets de Filtros ---
            self.ComboBox_FiltrosSalvos.SelectionChanged += self.preset_selecionado
            self.Button_SalvarFiltro.Click += self.salvar_preset_click
            self.Button_DeletarFiltro.Click += self.deletar_preset_click
            self.Button_CapturarSelecao.Click += self.capturar_selecao_click
            
            # --- Eventos de Histórico ---
            try:
                self.Button_HistoricoAnterior.Click += self.historico_anterior_click
                self.Button_LimparHistorico.Click += self.limpar_historico_click
            except: pass

            self.carregar_presets_ui()
            
            # Estado inicial do segundo filtro
            self.segundo_filtro_alterado(None, None)
            
            # ✨ NOVO: Detectar seleção ativa e habilitar radio / injetar categorias
            try:
                selecao_ids = list(uidoc.Selection.GetElementIds())
                self._selecao_inicial = selecao_ids
                if selecao_ids:
                    n = len(selecao_ids)
                    self.Radio_SelecaoAtual.Content = u"Seleção Atual ({} elemento{})".format(n, u"s" if n != 1 else u"")
                    self.Radio_SelecaoAtual.Visibility = System.Windows.Visibility.Visible
                    self.Radio_SelecaoAtual.IsChecked = True
                    
                    # INJEÇÃO INTELIGENTE DE CATEGORIAS DA SELEÇÃO
                    used_categories = set()
                    for eid in selecao_ids:
                        el = doc.GetElement(eid)
                        if el and el.Category:
                            # Se a categoria já está mapeada nos botões/opções padrão, não injeta como "*"
                            cat_id_val = getattr(el.Category.Id, "Value", None)
                            if cat_id_val is None:
                                cat_id_val = el.Category.Id.IntegerValue
                            
                            if int(cat_id_val) in self.ids_mapeados:
                                continue

                            c_name = el.Category.Name
                            # Insere marcador visual
                            used_categories.add(u"* [SELEÇÃO] " + c_name)
                    
                    if used_categories:
                        # Adiciona no topo das opções do ComboBox
                        for c_sel in sorted(used_categories):
                            self.ComboBox_Categoria.Items.Insert(0, c_sel)
                            
                else:
                    self._selecao_inicial = []
                    self.Radio_SelecaoAtual.Visibility = System.Windows.Visibility.Collapsed
            except:
                self._selecao_inicial = []
            
            self.atualizar_status("Selecione uma categoria para começar...")
            
        except Exception as e:
            logger.error("Erro ao inicializar controles: {}".format(traceback.format_exc()))
            forms.alert("Erro ao inicializar interface: {}".format(str(e)))
    
    def expand_nested_families(self, elements):
        """✨ NOVO: Expande recursivamente sub-elementos de FamilyInstances aninhadas."""
        expanded = []
        visited = set()
        
        def _expand(el):
            try:
                eid = el.Id.IntegerValue
                if eid in visited:
                    return
                visited.add(eid)
                expanded.append(el)
                # Tentar expandir sub-componentes (famílias aninhadas)
                try:
                    sub_ids = el.GetSubComponentIds()
                    if sub_ids:
                        for sub_id in sub_ids:
                            sub = doc.GetElement(sub_id)
                            if sub:
                                _expand(sub)
                except:
                    pass  # Elemento não suporta GetSubComponentIds
            except:
                pass
        
        for el in elements:
            _expand(el)
        
        return expanded
    
    def get_available_parameters(self, elements):
        """Coleta parâmetros de instância E de tipo - OTIMIZADO"""
        param_set = set()
        
        # ⚡ OTIMIZAÇÃO: Reduzir amostra para 20 elementos (mais rápido)
        amostra = elements[:20] if len(elements) > 20 else elements
        
        for el in amostra:
            # Parâmetros de Instância
            for param in el.Parameters:
                if param.Definition:
                    param_set.add(param.Definition.Name)
            
            # Parâmetros de Tipo
            try:
                type_id = el.GetTypeId()
                if type_id != ElementId.InvalidElementId:
                    type_elem = doc.GetElement(type_id)
                    if type_elem:
                        for param in type_elem.Parameters:
                            if param.Definition:
                                param_set.add(param.Definition.Name)
            except:
                pass

        # Injetar Parâmetros Virtuais no topo da lista (ordenados via colchetes)
        param_set.add(u"[VIRTUAL] Elevação Z (m)")
        param_set.add(u"[VIRTUAL] Nível Hospedeiro")
        param_set.add(u"[VIRTUAL] Nome do Workset")

        return sorted(list(param_set))
    
    def get_parameter_value(self, param):
        """Obtém o valor do parâmetro priorizando o formato visual (String)"""
        if not param:
            return None
        
        try:
            # Tentar AsValueString() primeiro
            val_string = param.AsValueString()
            if val_string:
                return val_string

            # Fallback para StorageType
            storage_type = param.StorageType
            if storage_type == StorageType.String:
                return param.AsString()
            elif storage_type == StorageType.Integer:
                return str(param.AsInteger())
            elif storage_type == StorageType.Double:
                return "{:.2f}".format(param.AsDouble())
            elif storage_type == StorageType.ElementId:
                elem_id = param.AsElementId()
                id_val = getattr(elem_id, "Value", None) 
                if id_val is None: 
                    id_val = elem_id.IntegerValue
                
                if id_val > 0:
                    elem = doc.GetElement(elem_id)
                    if elem and hasattr(elem, 'Name'):
                        return elem.Name
                    else:
                        return str(id_val)
                return None
        except:
            return None
        
        return None
    
    def find_parameter(self, element, param_name):
        """Encontra parâmetro na instância OU no tipo, ou resolve Parâmetros Virtuais"""
        # Checar se é virtual
        if param_name == u"[VIRTUAL] Elevação Z (m)":
            z_meters = get_element_elevation(element)
            if z_meters is not None:
                return "{:.3f}".format(z_meters)
            return ""
            
        if param_name == u"[VIRTUAL] Nível Hospedeiro":
            try:
                lvl_id = element.LevelId
                if lvl_id and lvl_id != ElementId.InvalidElementId:
                    return doc.GetElement(lvl_id).Name
            except: pass
            return ""
            
        if param_name == u"[VIRTUAL] Nome do Workset":
            try:
                if element.WorksetId and element.WorksetId != DB.WorksetId.InvalidWorksetId:
                    ws_table = doc.GetWorksetTable()
                    ws = ws_table.GetWorkset(element.WorksetId)
                    if ws: return ws.Name
            except: pass
            return ""

        # 1. Tentar na Instância
        param = element.LookupParameter(param_name)
        if param:
            return param
        
        # 2. Tentar no Tipo
        try:
            type_id = element.GetTypeId()
            if type_id != ElementId.InvalidElementId:
                type_elem = doc.GetElement(type_id)
                if type_elem:
                    param = type_elem.LookupParameter(param_name)
                    if param:
                        return param
        except:
            pass

        return None
    
    def check_condition(self, param_value, condition, value):
        """Verifica a condição"""
        if param_value is None:
            param_value = ""
        else:
            param_value = str(param_value).strip()
        
        value = value.strip() if value else ""
        
        p_val_lower = param_value.lower()
        val_lower = value.lower()

        if condition == "Igual a":
            return p_val_lower == val_lower
        elif condition == "Contém":
            return val_lower in p_val_lower
        elif condition == "Não Contém":
            return val_lower not in p_val_lower
        elif condition == "Diferente de":
            return p_val_lower != val_lower
        elif condition == "Começa com":
            return p_val_lower.startswith(val_lower)
        elif condition == "Termina com":
            return p_val_lower.endswith(val_lower)
        elif condition == "Corresponde ao Padrão":
            try:
                if not val_lower: return False
                # Converter Smart Patern para Regex
                pattern = ""
                for char in val_lower:
                    if char == '#': pattern += r'\d'
                    elif char == '@': pattern += r'[a-z]'
                    elif char == '?': pattern += r'.'
                    elif char in '()|.*+^{}\\$[]':
                        if char in '()': pattern += char # Manter grupos de captura
                        else: pattern += "\\" + char # Escapar metas do regex
                    else:
                        pattern += char
                # Forçar correspondência exata ou find parcial? Vamos fazer find parcial para ser flexível, amarrando metacaracteres de borda
                match = re.search(pattern, p_val_lower)
                return bool(match)
            except:
                return False
        elif condition == "Em branco":
            return param_value == ""
        elif condition == "Não em branco":
            return param_value != ""
        return False
    
    def categoria_selecionada(self, sender, args):
        """⚡ OTIMIZADO: Carregamento lazy de parâmetros"""
        try:
            if not self.ComboBox_Categoria.SelectedItem:
                return
                
            categoria_nome = self.ComboBox_Categoria.SelectedItem
            usar_selecao = getattr(self.Radio_SelecaoAtual, 'IsChecked', False)
            
            # ⚡ Verificar cache primeiro
            cache_key = "{}_{}_sel{}".format(categoria_nome, self.Radio_VistaAtual.IsChecked, usar_selecao)
            if cache_key in self._parametros_cache:
                self._preencher_combos_parametros(self._parametros_cache[cache_key])
                return
            
            elementos_amostra = []
            max_elementos = 20
            
            # ✨ NOVO: Escopo de Seleção Ativa — amostrar dos elementos já selecionados
            if usar_selecao and self._selecao_inicial:
                for eid in self._selecao_inicial[:max_elementos]:
                    try:
                        el = doc.GetElement(eid)
                        if el:
                            elementos_amostra.append(el)
                    except:
                        pass
            else:
                raw_cat = str(categoria_nome)
                # Modo Selecao Injetado Especial
                if raw_cat.startswith("* [SELEÇÃO] "):
                    clean_name = raw_cat.replace("* [SELEÇÃO] ", "")
                    # Pegar Ids da categoria na seleção ativa
                    if self._selecao_inicial:
                        for eid in self._selecao_inicial:
                            try:
                                el = doc.GetElement(eid)
                                if el and el.Category and el.Category.Name == clean_name:
                                    elementos_amostra.append(el)
                            except: pass
                else:
                    # Categoria Base do Hardcoded Dict
                    config = self.categoria_opcoes[categoria_nome]
                    categorias = config["categorias"]
                    usar_vista_atual = self.Radio_VistaAtual.IsChecked
                    elementos_coletados = 0
                    
                    for cat in categorias:
                        if elementos_coletados >= max_elementos:
                            break
                            
                        col = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
                        col = col.OfCategory(cat).WhereElementIsNotElementType()
                        
                        iterator = col.GetElementIterator()
                        while iterator.MoveNext() and elementos_coletados < max_elementos:
                            elementos_amostra.append(iterator.Current)
                            elementos_coletados += 1
            
            if elementos_amostra:
                parametros_disponiveis = self.get_available_parameters(elementos_amostra)
                self._parametros_cache[cache_key] = parametros_disponiveis
                self._preencher_combos_parametros(parametros_disponiveis)
                self.atualizar_status(u"Pronto. {} parâmetros encontrados.".format(len(parametros_disponiveis)))
            else:
                self.atualizar_status(u"Nenhum elemento encontrado nesta categoria.")
                
        except Exception as e:
            logger.error("Erro em categoria_selecionada: {}".format(traceback.format_exc()))
            self.atualizar_status(u"Erro ao carregar parâmetros")
    
    def _preencher_combos_parametros(self, parametros):
        """⚡ Método auxiliar para preencher ComboBoxes"""
        self.ComboBox_Parametro1.Items.Clear()
        self.ComboBox_Parametro2.Items.Clear()
        for param in parametros:
            self.ComboBox_Parametro1.Items.Add(param)
            self.ComboBox_Parametro2.Items.Add(param)

    def segundo_filtro_alterado(self, sender, args):
        try:
            usar = self.CheckBox_UsarSegundoFiltro.IsChecked
            self.ComboBox_Parametro2.IsEnabled = usar
            self.ComboBox_Condicao2.IsEnabled = usar
            self.TextBox_Valor2.IsEnabled = usar
            self.ComboBox_Operador.IsEnabled = usar
        except Exception as e:
            logger.error("Erro em segundo_filtro_alterado: {}".format(traceback.format_exc()))
    
    def atualizar_status(self, mensagem):
        try:
            self.TextBlock_Status.Text = mensagem
        except:
            pass
    
    def validar_campos(self):
        try:
            if not self.ComboBox_Categoria.SelectedItem:
                forms.alert("Selecione uma categoria.")
                return False
            if not self.ComboBox_Parametro1.SelectedItem:
                forms.alert("Selecione o parâmetro 1.")
                return False
            if not self.ComboBox_Condicao1.SelectedItem:
                forms.alert("Selecione a condição 1.")
                return False
            
            cond1 = self.ComboBox_Condicao1.SelectedItem.Content
                
            if self.CheckBox_UsarSegundoFiltro.IsChecked:
                if not self.ComboBox_Parametro2.SelectedItem:
                    forms.alert("Selecione o parâmetro 2.")
                    return False
                if not self.ComboBox_Condicao2.SelectedItem:
                    forms.alert("Selecione a condição 2.")
                    return False
                cond2 = self.ComboBox_Condicao2.SelectedItem.Content
            return True
        except Exception as e:
            logger.error("Erro em validar_campos: {}".format(traceback.format_exc()))
            return False

    def aplicar_filtro_click(self, sender, args):
        """OTIMIZADO: Filtragem com 3 escopos (Selecao/Vista/Projeto) + Familias Aninhadas"""
        try:
            if not self.validar_campos():
                return
            
            self.Button_AplicarFiltro.IsEnabled = False
            self.Button_AplicarFiltro.Content = "PROCESSANDO..."
            
            categoria_nome = self.ComboBox_Categoria.SelectedItem
            
            usar_selecao      = bool(getattr(self.Radio_SelecaoAtual,       'IsChecked', False))
            usar_vista_atual  = bool(self.Radio_VistaAtual.IsChecked)
            incluir_aninhadas = bool(getattr(self.CheckBox_FamiliasAninhadas, 'IsChecked', False))
            
            p1_nome    = self.ComboBox_Parametro1.SelectedItem
            c1         = self.ComboBox_Condicao1.SelectedItem.Content
            v1         = self.TextBox_Valor1.Text
            usar_f2    = self.CheckBox_UsarSegundoFiltro.IsChecked
            p2_nome    = self.ComboBox_Parametro2.SelectedItem if usar_f2 else None
            c2         = self.ComboBox_Condicao2.SelectedItem.Content if usar_f2 else None
            v2         = self.TextBox_Valor2.Text if usar_f2 else ""
            operador_e = self.Radio_And.IsChecked
            
            # Recupera as configuracoes da categoria se for originaria do dicionario padrao
            config = self.categoria_opcoes.get(categoria_nome, None)
            
            ids_selecionados = []
            self.atualizar_status(u"Processando elementos...")
            
            def avaliar(el):
                try:
                    
                    # Logica Avancada de Exclusao (Parametros Reais)
                    # Verifica se o config existe (pois categorias INJETADAS nao terao config customizado)
                    if config:
                        exc_rule = config.get("exclude_if_not")
                        if exc_rule == "angle_or_conduit":
                            tipo = el.GetType().Name
                            # Se for um Conduit (segmento reto), ele ta seguro, deixa de fora as exclusoes das Familias
                            if tipo == "FamilyInstance":
                                # Tenta pegar pelo PartType (Parametro NATIVO de pecas MEP)
                                # Se for uma caixa (Equipamento/Panel), o PartType costuma ser diferente de Elbow, Tee, Cross, Transition, Union, etc.
                                try:
                                    # O parametro PartType fica no FamilySymbol ou na Family
                                    part_type_param = el.Symbol.Family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
                                    if part_type_param and part_type_param.HasValue:
                                        part_type_val = part_type_param.AsInteger()
                                        # Valores corriqueiros para caixas (Equipment = 6, Panelboard = ... Varias outras coisas que não sao 1 ao 5)
                                        # Conduletes/Curvas sao geralmente Elbow(1), Tee(2), Cross(3), Transition(4), Union(5)
                                        # Se a pessoa criadora fez a caixa com PartType de curva (Elbow), então é uma falha na criação da familia.
                                        # Na dúvida: Filtramos pelo nome também para reforçar a segurança (Abordagem híbrida)
                                        pass
                                except:
                                    pass
                                    
                                # Abordagem Hibrida: Se o Revit deixou passar, usamos o nome por segurança
                                nome = ""
                                try:
                                    if hasattr(el, 'Name') and el.Name:
                                        nome += el.Name.lower()
                                    if hasattr(el, 'Symbol') and el.Symbol and hasattr(el.Symbol, 'FamilyName'):
                                        nome += " " + el.Symbol.FamilyName.lower()
                                except:
                                    pass
                                
                                if "caixa" in nome or "condulete" in nome:
                                    return False

                    res1 = False
                    param1 = self.find_parameter(el, p1_nome)
                    # Se foi Virtual Parameter, ele volta como STRING do find_parameter()
                    if param1 is not None:
                        if isinstance(param1, (str, unicode)):
                             # Rotear para Text Compare diretamente sem passar pelo tratador de Double do Parameter
                             res1 = self.check_condition(param1, c1, v1)
                        else:
                             res1 = smart_parameter_compare(param1, c1, v1, self.check_condition)
                    
                    res2 = False
                    if usar_f2:
                        param2 = self.find_parameter(el, p2_nome)
                        if param2 is not None:
                            if isinstance(param2, (str, unicode)):
                                 res2 = self.check_condition(param2, c2, v2)
                            else:
                                 res2 = smart_parameter_compare(param2, c2, v2, self.check_condition)
                    
                    return res1 if not usar_f2 else ((res1 and res2) if operador_e else (res1 or res2))
                except Exception as eval_err:
                    return False
            
            # ESCOPO 1: Selecao Ativa
            if usar_selecao and self._selecao_inicial:
                pool = []
                for eid in self._selecao_inicial:
                    try:
                        el = doc.GetElement(eid)
                        if el:
                            pool.append(el)
                    except:
                        pass
                if incluir_aninhadas:
                    pool = self.expand_nested_families(pool)
                for el in pool:
                    if avaliar(el):
                        ids_selecionados.append(el.Id)
            
            # ESCOPOS 2 e 3: Vista Atual / Projeto Inteiro
            else:
                view_atual = revit.active_view if usar_vista_atual else None
                
                raw_cat = str(categoria_nome)
                todas_categorias_analisar = []
                filtros_class = []
                requires_param = None
                requires_param_absent = None
                
                # Se for Injetada "* [SELEÇÃO]"
                if raw_cat.startswith("* [SELEÇÃO] "):
                    clean_name = raw_cat.replace("* [SELEÇÃO] ", "")
                    # O Revit nao possui filtro limpo apenas por String para Collectors globais
                    # Temos que varrer ou usar um filtro mais devagar mas vamos fazer Collector sem Classe
                    # Mas como é Global, precisamos tentar achar o BuiltInCategory via BuiltInCategories
                    # Ou varrer filtrando pela classe "clássica" global... Nao e ideal por string.
                    # Mas como isso soh ocorre se a pessoa NAO clicou no Radio "Selecao Atual", vamos alertar ou pegar da vista
                    
                    # Hack super eficiente para achar todos da mesma categoria sem ter o BuiltInCategory enum
                    # O pyRevit permite passar ids de categorias se acharmos um...
                    if self._selecao_inicial:
                        for eid in self._selecao_inicial:
                            el_dummy = doc.GetElement(eid)
                            if el_dummy and el_dummy.Category and el_dummy.Category.Name == clean_name:
                                todas_categorias_analisar.append(el_dummy.Category.Id)
                                break
                    
                else: 
                    # Categorias hardcoded originais
                    todas_categorias_analisar = config["categorias"] if config else []
                    filtros_class = config["classes"] if config else []
                    requires_param = config.get("requires_param") if config else None
                    requires_param_absent = config.get("requires_param_absent") if config else None
                
                # Para cada id_cat descoberto hardcoded OU pela injeção da seleçao
                for cat in todas_categorias_analisar:
                    col = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
                    
                    if isinstance(cat, BuiltInCategory): # Categoria BuiltIn Hardcoded
                        col = col.OfCategory(cat).WhereElementIsNotElementType()
                    else: # ElementId Category vindo do Injetor do Selecionado!
                        # Garante que é ElementId (pode vir como BuiltInCategory se o config for gerado dinamicamente)
                        if isinstance(cat, int):
                            cat = ElementId(cat)
                        col = col.OfCategoryId(cat).WhereElementIsNotElementType()
                    
                    pool_cat = []
                    for el in col:
                        try:
                            if usar_vista_atual and el.IsHidden(view_atual):
                                continue
                                
                            # Se não possui config, é uma categoria injetada (custom) e não tem exclusão fina
                            if config:
                                if filtros_class and el.GetType().Name not in filtros_class:
                                    if requires_param == "RN_optional" and not el.LookupParameter("RN"):
                                        continue
                                    if requires_param != "RN_optional":
                                        continue
                                if requires_param and requires_param != "RN_optional":
                                    if not el.LookupParameter(requires_param):
                                        continue
                                if requires_param_absent:
                                    if el.LookupParameter(requires_param_absent):
                                        continue
                                        
                            pool_cat.append(el)
                        except:
                            continue
                    
                    if incluir_aninhadas:
                        pool_cat = self.expand_nested_families(pool_cat)
                    
                    for el in pool_cat:
                        if avaliar(el):
                            ids_selecionados.append(el.Id)
            
            if ids_selecionados:
                sel_ids_list = List[ElementId](ids_selecionados)
                uidoc.Selection.SetElementIds(sel_ids_list)
                
                # ADICIONAR AO HISTÓRICO
                cat_name_history = to_unicode(categoria_nome) if categoria_nome else to_unicode(self.ComboBox_Categoria.Text)
                p1_name_history = to_unicode(p1_nome) if p1_nome else to_unicode(self.ComboBox_Parametro1.Text)
                p2_name_history = to_unicode(p2_nome) if p2_nome else (to_unicode(self.ComboBox_Parametro2.Text) if usar_f2 else "")
                
                desc = u"{} {} {}".format(p1_name_history, c1, v1)
                if usar_f2:
                    desc += u" ({} {} {} {})".format("E" if operador_e else "OU", p2_name_history, c2, v2)
                
                state = {
                    'p1': p1_name_history, 'c1': to_unicode(c1), 'v1': to_unicode(v1),
                    'usar_f2': usar_f2, 'p2': p2_name_history, 'c2': to_unicode(c2) if c2 else "", 'v2': to_unicode(v2),
                    'op_e': operador_e
                }
                
                self.selection_history.add(sel_ids_list, cat_name_history, desc, len(ids_selecionados), state)
                
                self.atualizar_status(u"OK {} elementos selecionados!".format(len(ids_selecionados)))
                forms.alert(u"OK {} elementos selecionados com sucesso!".format(len(ids_selecionados)))
            else:
                forms.alert(u"Nenhum elemento atende aos criterios.")
                
            self.Button_AplicarFiltro.IsEnabled = True
            self.Button_AplicarFiltro.Content = "APLICAR FILTRO"
                
        except Exception as e:
            logger.error(u"Erro em aplicar_filtro_click: {}".format(to_unicode(e)))
            forms.alert(u"Erro durante a filtragem:\n{}".format(to_unicode(e)))
            self.Button_AplicarFiltro.IsEnabled = True
            self.Button_AplicarFiltro.Content = "APLICAR FILTRO"



    def fechar_click(self, sender, args):
        """⚡ OTIMIZADO: Limpar cache ao fechar"""
        try:
            self._parametros_cache.clear()
            self.Close()
        except: pass
        
    # --- LOGICA DE PRESETS (SETS) ---
    def carregar_presets_ui(self):
        try:
            data = load_presets()
            self.preset_data = data
            
            self.ComboBox_FiltrosSalvos.Items.Clear()
            self.ComboBox_FiltrosSalvos.Items.Add("(Selecione um Filtro)")
            
            presets = sorted(data.get("Presets", {}).keys())
            for p in presets:
                self.ComboBox_FiltrosSalvos.Items.Add(p)
            
            # Sempre iniciar vazio (sem carregar LastUsed)
            self.ComboBox_FiltrosSalvos.SelectedIndex = 0
        except: pass

    def salvar_preset_click(self, sender, args):
        try:
            name_raw = forms.ask_for_string(prompt="Nome do Filtro (Set):", title="Salvar Filtro")
            if not name_raw: return
            name = str(name_raw)
            
            # Função auxiliar melhorada para pegar valor de ComboBox
            def get_cb_val(cb):
                try:
                    if cb.SelectedItem:
                        val = str(cb.SelectedItem)
                    elif cb.Text:
                        val = str(cb.Text)
                    else:
                        val = ""
                    return val
                except:
                    return ""

            # Captura estado atual
            cat_val = get_cb_val(self.ComboBox_Categoria)
            param1_val = get_cb_val(self.ComboBox_Parametro1)
            param2_val = get_cb_val(self.ComboBox_Parametro2)
            
            # Log para debug
            logger.debug("Salvando - Categoria: '{}', Param1: '{}', Param2: '{}'".format(cat_val, param1_val, param2_val))
            
            state = {
                "Categoria": cat_val,
                "Escopo": "Selecao Atual" if bool(getattr(self.Radio_SelecaoAtual, 'IsChecked', False)) else ("Vista Atual" if self.Radio_VistaAtual.IsChecked else "Projeto Inteiro"),
                "UseNested": bool(getattr(self.CheckBox_FamiliasAninhadas, 'IsChecked', False)),
                
                "Param1": param1_val,
                "Cond1": str(self.ComboBox_Condicao1.SelectedItem.Content) if self.ComboBox_Condicao1.SelectedItem else "",
                "Val1": str(self.TextBox_Valor1.Text) if getattr(self.TextBox_Valor1, "Text", None) else "",
                
                "UseF2": bool(self.CheckBox_UsarSegundoFiltro.IsChecked),
                "Param2": param2_val,
                "Cond2": str(self.ComboBox_Condicao2.SelectedItem.Content) if self.ComboBox_Condicao2.SelectedItem else "",
                "Val2": str(self.TextBox_Valor2.Text) if getattr(self.TextBox_Valor2, "Text", None) else "",
                
                "Logic": "AND" if self.Radio_And.IsChecked else "OR"
            }
            
            # Salva
            data = load_presets()
            if "Presets" not in data: 
                data["Presets"] = {}
            data["Presets"][name] = state
            
            if save_presets(data):
                self.carregar_presets_ui()
                forms.alert("Filtro '{}' salvo!\n\nCategoria: {}\nParametro: {}".format(name, cat_val, param1_val))
        except Exception as e:
            logger.error("Erro ao salvar preset: " + str(e))
            forms.alert("Erro ao salvar: " + str(e))

    def deletar_preset_click(self, sender, args):
        sel = self.ComboBox_FiltrosSalvos.SelectedItem
        if not sel or sel == "(Selecione um Filtro)": return
        
        if forms.alert(u"Deletar filtro '{}'?".format(sel), yes=True, no=True):
            data = load_presets()
            if "Presets" in data and sel in data["Presets"]:
                del data["Presets"][sel]
                save_presets(data)
                self.carregar_presets_ui()

    def preset_selecionado(self, sender, args):
        sel = self.ComboBox_FiltrosSalvos.SelectedItem
        if not sel or sel == "(Selecione um Filtro)": return
        
        data = load_presets()
        preset = data.get("Presets", {}).get(sel)
        if not preset: return
        
        # Aplicar Preset
        try:
            # 1. Categoria
            cat = preset.get("Categoria")
            if cat:
                for item in self.ComboBox_Categoria.Items:
                    if str(item) == cat:
                        self.ComboBox_Categoria.SelectedItem = item
                        WinFormsApp.DoEvents()
                        self.categoria_selecionada(None, None)  # Força atualizar os parâmetros
                        break
            
            # 2. Parametros Filtro 1
            p1 = preset.get("Param1")
            if p1:
                for item in self.ComboBox_Parametro1.Items:
                    if str(item) == p1:
                        self.ComboBox_Parametro1.SelectedItem = item
                        break
            
            c1 = preset.get("Cond1")
            for item in self.ComboBox_Condicao1.Items:
                if item.Content == c1:
                    self.ComboBox_Condicao1.SelectedItem = item
                    break
            
            self.TextBox_Valor1.Text = preset.get("Val1", "")
            
            # 3. Filtro 2
            use_f2 = preset.get("UseF2", False)
            self.CheckBox_UsarSegundoFiltro.IsChecked = use_f2
            WinFormsApp.DoEvents()
            
            if use_f2:
                p2 = preset.get("Param2")
                if p2:
                    for item in self.ComboBox_Parametro2.Items:
                        if str(item) == p2:
                            self.ComboBox_Parametro2.SelectedItem = item
                            break
                            
                c2 = preset.get("Cond2")
                if c2:
                    for item in self.ComboBox_Condicao2.Items:
                        if getattr(item, 'Content', str(item)) == c2:
                            self.ComboBox_Condicao2.SelectedItem = item
                            break
                            
                self.TextBox_Valor2.Text = preset.get("Val2", "")
            
            # 4. Escopo + Opções adicionais
            escopo = preset.get("Escopo", "Vista Atual")
            # "Selecao Atual" não pode ser restaurado (seleção mudou) — usa Vista Atual
            if escopo == "Projeto Inteiro":
                self.Radio_ProjetoInteiro.IsChecked = True
            else:
                self.Radio_VistaAtual.IsChecked = True
            
            # Restaurar CheckBox de famílias aninhadas
            try:
                self.CheckBox_FamiliasAninhadas.IsChecked = bool(preset.get("UseNested", False))
            except:
                pass
            
            if preset.get("Logic") == "AND": self.Radio_And.IsChecked = True
            else: self.Radio_Or.IsChecked = True
                
        except Exception as e:
            logger.error(u"Erro ao aplicar preset: " + to_unicode(e))

    def capturar_selecao_click(self, sender, args):
        """✨ NOVO: Preenche a Categoria automaticamente com base no elemento selecionado no Revit"""
        try:
            selection = uidoc.Selection.GetElementIds()
            if not selection:
                forms.alert(u"Selecione um elemento no Revit primeiro para capturar sua categoria.")
                return
                
            el = doc.GetElement(selection[0])
            cat = el.Category
            if not cat: return
            
            target_cat_id = cat.Id.IntegerValue
            
            found_cat_name = None
            for name, info in self.categoria_opcoes.items():
                cat_list = []
                for c in info["categorias"]:
                    try:
                        if isinstance(c, BuiltInCategory):
                            cat_list.append(int(c))
                        elif hasattr(c, "IntegerValue"):
                            cat_list.append(c.IntegerValue)
                        elif hasattr(c, "Value"):
                            cat_list.append(int(c.Value))
                    except:
                        pass
                
                if int(target_cat_id) in cat_list:
                    found_cat_name = name
                    break
            
            if found_cat_name:
                self.ComboBox_Categoria.SelectedItem = found_cat_name
                self.categoria_selecionada(None, None)
                self.atualizar_status(u"Categoria '{}' capturada!".format(found_cat_name))
            else:
                injected_name = u"* [SELEÇÃO] " + cat.Name
                # Se não foi pre-injetada ao abrir o script, a gente injeta agora para permitir o uso pleno
                if injected_name not in self.ComboBox_Categoria.Items:
                    self.ComboBox_Categoria.Items.Insert(0, injected_name)
                    
                self.ComboBox_Categoria.SelectedItem = injected_name
                self.categoria_selecionada(None, None)
                self.atualizar_status(u"Categoria '{}' injetada da seleção!".format(cat.Name))
                
        except Exception as e:
            logger.error(u"Erro ao capturar seleção: " + to_unicode(e))

    def historico_anterior_click(self, sender, args):
        """Recupera a última seleção realizada na sessão."""
        try:
            history = self.selection_history.history
            if not history:
                forms.alert("Nenhum histórico disponível.")
                return
            
            # Se houver histórico, mostrar lista para escolher
            options = []
            for h in history:
                options.append("{} | {} ({} elem) - {}".format(
                    h['timestamp'], h['category'], h['count'], h['criteria']
                ))
            
            selected = forms.SelectFromList.show(
                options, 
                title="Histórico de Seleções",
                multiselect=False,
                button_name="Selecionar"
            )
            
            if selected:
                # Extrai o índice do item do histórico original pela string de opção
                idx = options.index(selected)
                h_entry = history[idx]
                
                # 1. restaurar IDs no Revit
                eids = [ElementId(int(eid)) for eid in h_entry['element_ids']]
                valid_ids = [eid for eid in eids if doc.GetElement(eid)]
                if valid_ids:
                    uidoc.Selection.SetElementIds(List[ElementId](valid_ids))
                
                # 2. Restaurar UI
                state = h_entry.get('state')
                if state:
                    self.atualizar_status("Restaurando configuração...")
                    
                    # Categoria
                    cat_name = h_entry['category']
                    found_cat = False
                    for item in self.ComboBox_Categoria.Items:
                        if to_unicode(item) == to_unicode(cat_name):
                            self.ComboBox_Categoria.SelectedItem = item
                            found_cat = True
                            break
                    
                    if not found_cat:
                        self.ComboBox_Categoria.Text = cat_name
                    
                    # Força carregamento dos parâmetros (Processamento Síncrono)
                    WinFormsApp.DoEvents() 
                    self.categoria_selecionada(None, None)
                    WinFormsApp.DoEvents()

                    # Filtro 1
                    p1_val = state.get('p1')
                    found_p1 = False
                    for item in self.ComboBox_Parametro1.Items:
                        if to_unicode(item) == to_unicode(p1_val):
                            self.ComboBox_Parametro1.SelectedItem = item
                            found_p1 = True
                            break
                    if not found_p1: self.ComboBox_Parametro1.Text = p1_val
                    
                    c1_val = state.get('c1')
                    for item in self.ComboBox_Condicao1.Items:
                        if item.Content == c1_val:
                            self.ComboBox_Condicao1.SelectedItem = item
                            break
                    self.TextBox_Valor1.Text = state.get('v1', '')
                    
                    # Filtro 2
                    use_f2 = state.get('usar_f2', False)
                    self.CheckBox_UsarSegundoFiltro.IsChecked = use_f2
                    WinFormsApp.DoEvents()
                    
                    if use_f2:
                        p2_val = state.get('p2')
                        found_p2 = False
                        for item in self.ComboBox_Parametro2.Items:
                            if to_unicode(item) == to_unicode(p2_val):
                                self.ComboBox_Parametro2.SelectedItem = item
                                found_p2 = True
                                break
                        if not found_p2: self.ComboBox_Parametro2.Text = p2_val
                            
                        c2_val = state.get('c2')
                        for item in self.ComboBox_Condicao2.Items:
                            if item.Content == c2_val:
                                self.ComboBox_Condicao2.SelectedItem = item
                                break
                        self.TextBox_Valor2.Text = state.get('v2', '')
                    
                    # Operador
                    if state.get('op_e'): self.Radio_And.IsChecked = True
                    else: self.Radio_Or.IsChecked = True

                    self.atualizar_status(u"Histórico carregado: {} elementos.".format(len(valid_ids)))
                else:
                    self.atualizar_status(u"Histórico antigo: Seleção restaurada.")
        except Exception as e:
            logger.error("Erro ao recuperar historico: " + str(e))

    def limpar_historico_click(self, sender, args):
        """Limpa o arquivo de histórico."""
        self.selection_history.clear()
        self.atualizar_status("Histórico limpo.")

# Execução principal
try:
    script_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(script_dir, 'FiltroAvancado.xaml')
    
    if not os.path.exists(xaml_path):
        logger.warning("Arquivo XAML não encontrado em: {}".format(xaml_path))
        forms.alert("⚠️ Arquivo de interface não encontrado. Usando modo simplificado.")
    
    window = FiltroAvancadoWindow()
    window.ShowDialog()
    
except Exception as e:
    logger.error("Erro fatal: {}".format(traceback.format_exc()))
    forms.alert("❌ Erro crítico ao abrir o filtro:\n{}".format(str(e)))
