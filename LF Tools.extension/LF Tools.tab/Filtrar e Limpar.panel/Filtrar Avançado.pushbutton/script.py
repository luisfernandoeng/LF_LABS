# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ElementId, StorageType,
    BuiltInParameter, Element
)
from System.Collections.Generic import List
import traceback
import os
import json
import io
import System
import sys
from System.Windows.Forms import Application as WinFormsApp


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
                u"Eletrodutos (segmentos + curvas reais)": {
                    "categorias": [BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting],
                    "classes": ["Conduit"],
                    "requires_param": "RN_optional",
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
                u"Eletrocalhas": {
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
            self.carregar_presets_ui()
            
            # Estado inicial do segundo filtro
            self.segundo_filtro_alterado(None, None)
            
            self.atualizar_status("Selecione uma categoria para começar...")
            
        except Exception as e:
            logger.error("Erro ao inicializar controles: {}".format(traceback.format_exc()))
            forms.alert("Erro ao inicializar interface: {}".format(str(e)))
    
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
        """Encontra parâmetro na instância OU no tipo"""
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
        elif condition == "Diferente de":
            return p_val_lower != val_lower
        elif condition == "Começa com":
            return p_val_lower.startswith(val_lower)
        elif condition == "Termina com":
            return p_val_lower.endswith(val_lower)
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
            
            # ⚡ Verificar cache primeiro
            cache_key = "{}_{}".format(categoria_nome, self.Radio_VistaAtual.IsChecked)
            if cache_key in self._parametros_cache:
                self._preencher_combos_parametros(self._parametros_cache[cache_key])
                return
            
            config = self.categoria_opcoes[categoria_nome]
            categorias = config["categorias"]
            usar_vista_atual = self.Radio_VistaAtual.IsChecked
            
            elementos_amostra = []
            
            # ⚡ OTIMIZAÇÃO: Coletar apenas 20 elementos no total (não 10 por categoria)
            max_elementos = 20
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
                
                # Armazenar no cache
                self._parametros_cache[cache_key] = parametros_disponiveis
                
                self._preencher_combos_parametros(parametros_disponiveis)
                self.atualizar_status("Pronto. {} parâmetros encontrados.".format(len(parametros_disponiveis)))
            else:
                self.atualizar_status("Nenhum elemento encontrado nesta categoria.")
                
        except Exception as e:
            logger.error("Erro em categoria_selecionada: {}".format(traceback.format_exc()))
            self.atualizar_status("Erro ao carregar parâmetros")
    
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
        """⚡ OTIMIZADO: Lógica de filtragem com precisão de vista e segurança de encode"""
        try:
            if not self.validar_campos():
                return
            
            self.Button_AplicarFiltro.IsEnabled = False
            self.Button_AplicarFiltro.Content = "PROCESSANDO..."
            
            categoria_nome = self.ComboBox_Categoria.SelectedItem
            usar_vista_atual = self.Radio_VistaAtual.IsChecked
            
            # Filtro 1
            p1_nome = self.ComboBox_Parametro1.SelectedItem
            c1 = self.ComboBox_Condicao1.SelectedItem.Content
            v1 = self.TextBox_Valor1.Text
            
            # Filtro 2
            usar_f2 = self.CheckBox_UsarSegundoFiltro.IsChecked
            p2_nome = self.ComboBox_Parametro2.SelectedItem if usar_f2 else None
            c2 = self.ComboBox_Condicao2.SelectedItem.Content if usar_f2 else None
            v2 = self.TextBox_Valor2.Text if usar_f2 else ""
            
            operador_e = self.Radio_And.IsChecked
            
            config = self.categoria_opcoes[categoria_nome]
            ids_selecionados = []
            
            self.atualizar_status(u"Processando elementos...")
            
            for cat in config["categorias"]:
                col = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
                col = col.OfCategory(cat).WhereElementIsNotElementType()
                
                view_atual = revit.active_view if usar_vista_atual else None
                
                for el in col:
                    # MELHORIA: Filtrar apenas elementos visíveis (Respeita filtros/HH/VV)
                    if usar_vista_atual and el.IsHidden(view_atual):
                        continue

                    # Lógica de Classes e RN
                    tipo_el = el.GetType().Name
                    if config["classes"] and tipo_el not in config["classes"]:
                        if config.get("requires_param") == "RN_optional" and not el.LookupParameter("RN"):
                            continue
                        if config.get("requires_param") != "RN_optional":
                             continue

                    if config.get("requires_param") and config["requires_param"] != "RN_optional":
                        if not el.LookupParameter(config["requires_param"]): 
                            continue
                        
                    if config.get("requires_param_absent"):
                        if el.LookupParameter(config["requires_param_absent"]): 
                            continue
                    
                    # Verificação dos Filtros
                    res1 = False
                    param1 = self.find_parameter(el, p1_nome)
                    if param1:
                        val1 = self.get_parameter_value(param1)
                        res1 = self.check_condition(val1, c1, v1)
                    elif c1 == "Em branco":
                        res1 = True
                        
                    res2 = False
                    if usar_f2:
                        param2 = self.find_parameter(el, p2_nome)
                        if param2:
                            val2 = self.get_parameter_value(param2)
                            res2 = self.check_condition(val2, c2, v2)
                        elif c2 == "Em branco":
                            res2 = True
                    
                    match = False
                    if not usar_f2:
                        match = res1
                    else:
                        match = (res1 and res2) if operador_e else (res1 or res2)
                    
                    if match:
                        ids_selecionados.append(el.Id)
            
            # Resultado
            if ids_selecionados:
                uidoc.Selection.SetElementIds(List[ElementId](ids_selecionados))
                self.atualizar_status(u"✅ {} elementos selecionados!".format(len(ids_selecionados)))
                forms.alert(u"✅ {} elementos selecionados com sucesso!".format(len(ids_selecionados)))
            else:
                forms.alert(u"❌ Nenhum elemento atende aos critérios.")
                
            self.Button_AplicarFiltro.IsEnabled = True
            self.Button_AplicarFiltro.Content = "APLICAR FILTRO"
                
        except Exception as e:
            logger.error(u"Erro em aplicar_filtro_click: {}".format(to_unicode(e)))
            forms.alert(u"❌ Erro durante a filtragem:\n{}".format(to_unicode(e)))
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
                "Escopo": "Vista Atual" if self.Radio_VistaAtual.IsChecked else "Projeto Inteiro",
                
                "Param1": param1_val,
                "Cond1": str(self.ComboBox_Condicao1.SelectedItem.Content) if self.ComboBox_Condicao1.SelectedItem else "",
                "Val1": str(self.TextBox_Valor1.Text) if self.TextBox_Valor1.Text else "",
                
                "UseF2": bool(self.CheckBox_UsarSegundoFiltro.IsChecked),
                "Param2": param2_val,
                "Cond2": str(self.ComboBox_Condicao2.SelectedItem.Content) if self.ComboBox_Condicao2.SelectedItem else "",
                "Val2": str(self.TextBox_Valor2.Text) if self.TextBox_Valor2.Text else "",
                
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
            
            # 4. Outros
            if preset.get("Escopo") == "Vista Atual": self.Radio_VistaAtual.IsChecked = True
            else: self.Radio_ProjetoInteiro.IsChecked = True
            
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
                cat_list = [c.value if hasattr(c, 'value') else int(c) for c in info["categorias"]]
                if target_cat_id in cat_list:
                    found_cat_name = name
                    break
            
            if found_cat_name:
                self.ComboBox_Categoria.SelectedItem = found_cat_name
                self.categoria_selecionada(None, None)
                self.atualizar_status(u"Categoria '{}' capturada!".format(found_cat_name))
            else:
                forms.alert(u"A categoria '{}' não está mapeada.".format(cat.Name))
                
        except Exception as e:
            logger.error(u"Erro ao capturar seleção: " + to_unicode(e))

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
