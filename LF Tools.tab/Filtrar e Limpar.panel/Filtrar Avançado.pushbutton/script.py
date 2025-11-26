# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ElementId, StorageType,
    BuiltInParameter, Element
)
from System.Collections.Generic import List
import traceback
import os

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

class FiltroAvancadoWindow(forms.WPFWindow):
    def __init__(self):
        try:
            # Tenta carregar o XAML da bundle
            xaml_file = script.get_bundle_file('FiltroAvancado.xaml')
            if xaml_file and os.path.exists(xaml_file):
                forms.WPFWindow.__init__(self, xaml_file)
            else:
                # Fallback: usa XAML embutido se arquivo não for encontrado
                forms.WPFWindow.__init__(self, self.get_fallback_xaml())
            
            # Configurações iniciais
            self.categoria_opcoes = {
                "Eletrodutos (segmentos + curvas reais)": {
                    "categorias": [BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting],
                    "classes": ["Conduit"],
                    "requires_param": "RN_optional",
                    "requires_param_absent": None
                },
                "Isolamento de Tubo": {
                    "categorias": [BuiltInCategory.OST_PipeInsulations],
                    "classes": ["PipeInsulation"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Dutos": {
                    "categorias": [BuiltInCategory.OST_DuctCurves],
                    "classes": ["Duct"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Eletrocalhas": {
                    "categorias": [BuiltInCategory.OST_CableTray],
                    "classes": ["CableTray"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Tubulações": {
                    "categorias": [BuiltInCategory.OST_PipeCurves],
                    "classes": ["Pipe"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Conexões de Conduíte (sem RN)": {
                    "categorias": [BuiltInCategory.OST_ConduitFitting],
                    "classes": [],
                    "requires_param": None,
                    "requires_param_absent": "RN"
                },
                "Dispositivos Elétricos": {
                    "categorias": [BuiltInCategory.OST_ElectricalFixtures],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Equipamentos Elétricos": {
                    "categorias": [BuiltInCategory.OST_ElectricalEquipment],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Aparelhos de Iluminação": {
                    "categorias": [BuiltInCategory.OST_LightingFixtures],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Equipamentos Mecânicos": {
                    "categorias": [BuiltInCategory.OST_MechanicalEquipment],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Acessórios de Tubulação": {
                    "categorias": [BuiltInCategory.OST_PipeFitting],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Acessórios de Dutos": {
                    "categorias": [BuiltInCategory.OST_DuctFitting],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Válvulas de Tubulação": {
                    "categorias": [BuiltInCategory.OST_PipeAccessory],
                    "classes": ["FamilyInstance"],
                    "requires_param": None,
                    "requires_param_absent": None
                },
                "Acessórios de Dutos (Registros, etc.)": {
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
        """Inicializa os controles da interface"""
        try:
            # Preencher categorias
            self.ComboBox_Categoria.Items.Clear()
            for categoria in sorted(self.categoria_opcoes.keys()):
                self.ComboBox_Categoria.Items.Add(categoria)
            
            # Configurar eventos
            self.ComboBox_Categoria.SelectionChanged += self.categoria_selecionada
            self.CheckBox_UsarSegundoFiltro.Checked += self.segundo_filtro_alterado
            self.CheckBox_UsarSegundoFiltro.Unchecked += self.segundo_filtro_alterado
            self.Button_AplicarFiltro.Click += self.aplicar_filtro_click
            self.Button_Fechar.Click += self.fechar_click
            
            # Estado inicial do segundo filtro
            self.segundo_filtro_alterado(None, None)
            
            self.atualizar_status("Selecione uma categoria para começar...")
            
        except Exception as e:
            logger.error("Erro ao inicializar controles: {}".format(traceback.format_exc()))
            forms.alert("Erro ao inicializar interface: {}".format(str(e)))
    
    def get_available_parameters(self, elements):
        """Coleta parâmetros de instância E de tipo"""
        param_set = set()
        
        # Limitar a 50 elementos para performance
        amostra = elements[:50] if len(elements) > 50 else elements
        
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
        except Exception as e:
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
        try:
            if self.ComboBox_Categoria.SelectedItem:
                categoria_nome = self.ComboBox_Categoria.SelectedItem
                config = self.categoria_opcoes[categoria_nome]
                categorias = config["categorias"]
                usar_vista_atual = self.ComboBox_Escopo.SelectedItem.Content == "Somente na Vista Atual"
                
                elementos_amostra = []
                for cat in categorias:
                    collector = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
                    collector = collector.OfCategory(cat).WhereElementIsNotElementType()
                    
                    # Pega apenas os primeiros 50 para ser rápido
                    iterator = collector.GetElementIterator()
                    count = 0
                    while iterator.MoveNext() and count < 50:
                        elementos_amostra.append(iterator.Current)
                        count += 1
                
                if elementos_amostra:
                    parametros_disponiveis = self.get_available_parameters(elementos_amostra)
                    
                    self.ComboBox_Parametro1.Items.Clear()
                    self.ComboBox_Parametro2.Items.Clear()
                    for param in parametros_disponiveis:
                        self.ComboBox_Parametro1.Items.Add(param)
                        self.ComboBox_Parametro2.Items.Add(param)
                    
                    self.atualizar_status("Pronto. {} parâmetros encontrados. Configure os filtros e clique em 'APLICAR FILTRO'.".format(len(parametros_disponiveis)))
                else:
                    self.atualizar_status("Nenhum elemento encontrado nesta categoria.")
            else:
                self.atualizar_status("Selecione uma categoria...")
        except Exception as e:
            logger.error("Erro em categoria_selecionada: {}".format(traceback.format_exc()))
            self.atualizar_status("Erro ao carregar parâmetros")

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
            if "branco" not in cond1 and not self.TextBox_Valor1.Text.strip():
                forms.alert("Informe o valor para o filtro 1.")
                return False
                
            if self.CheckBox_UsarSegundoFiltro.IsChecked:
                if not self.ComboBox_Parametro2.SelectedItem:
                    forms.alert("Selecione o parâmetro 2.")
                    return False
                if not self.ComboBox_Condicao2.SelectedItem:
                    forms.alert("Selecione a condição 2.")
                    return False
                cond2 = self.ComboBox_Condicao2.SelectedItem.Content
                if "branco" not in cond2 and not self.TextBox_Valor2.Text.strip():
                    forms.alert("Informe o valor para o filtro 2.")
                    return False
            return True
        except Exception as e:
            logger.error("Erro em validar_campos: {}".format(traceback.format_exc()))
            return False

    def aplicar_filtro_click(self, sender, args):
        """Evento do botão APLICAR FILTRO - PRINCIPAL"""
        try:
            # Validar campos primeiro
            if not self.validar_campos():
                return
            
            # Desabilitar botão durante o processamento
            self.Button_AplicarFiltro.IsEnabled = False
            self.Button_AplicarFiltro.Content = "PROCESSANDO..."
            
            categoria_nome = self.ComboBox_Categoria.SelectedItem
            escopo = self.ComboBox_Escopo.SelectedItem.Content
            usar_vista_atual = escopo == "Somente na Vista Atual"
            
            # Filtro 1
            p1_nome = self.ComboBox_Parametro1.SelectedItem
            c1 = self.ComboBox_Condicao1.SelectedItem.Content
            v1 = self.TextBox_Valor1.Text
            
            # Filtro 2
            usar_f2 = self.CheckBox_UsarSegundoFiltro.IsChecked
            p2_nome = self.ComboBox_Parametro2.SelectedItem if usar_f2 else None
            c2 = self.ComboBox_Condicao2.SelectedItem.Content if usar_f2 else None
            v2 = self.TextBox_Valor2.Text if usar_f2 else ""
            
            operador_e = "E" in self.ComboBox_Operador.SelectedItem.Content
            
            config = self.categoria_opcoes[categoria_nome]
            ids_selecionados = []
            
            # Contador de progresso
            total_elementos = 0
            elementos_processados = 0
            
            # Primeiro contar elementos totais
            for cat in config["categorias"]:
                col = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
                col = col.OfCategory(cat).WhereElementIsNotElementType()
                total_elementos += len(list(col))
            
            self.atualizar_status("Processando {} elementos...".format(total_elementos))
            
            # Aplicar filtros
            for cat in config["categorias"]:
                col = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
                col = col.OfCategory(cat).WhereElementIsNotElementType()
                
                for el in col:
                    elementos_processados += 1
                    
                    # Atualizar status a cada 100 elementos
                    if elementos_processados % 100 == 0:
                        self.atualizar_status("Processando... {}/{} elementos".format(elementos_processados, total_elementos))
                    
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
                self.atualizar_status("✅ {} elementos selecionados com sucesso!".format(len(ids_selecionados)))
                forms.alert("✅ {} elementos selecionados com sucesso!".format(len(ids_selecionados)))
                
                # Manter a janela aberta para permitir novos filtros
                self.Button_AplicarFiltro.IsEnabled = True
                self.Button_AplicarFiltro.Content = "APLICAR FILTRO"
                
            else:
                forms.alert("❌ Nenhum elemento atende aos critérios.")
                self.Button_AplicarFiltro.IsEnabled = True
                self.Button_AplicarFiltro.Content = "APLICAR FILTRO"
                
        except Exception as e:
            logger.error("Erro em aplicar_filtro_click: {}".format(traceback.format_exc()))
            forms.alert("❌ Erro durante a filtragem:\n{}".format(str(e)))
            self.Button_AplicarFiltro.IsEnabled = True
            self.Button_AplicarFiltro.Content = "APLICAR FILTRO"

    def fechar_click(self, sender, args):
        """Evento do botão Fechar"""
        self.Close()

# Execução principal
try:
    # Verificar se os arquivos existem antes de executar
    script_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(script_dir, 'FiltroAvancado.xaml')
    
    if not os.path.exists(xaml_path):
        logger.warning("Arquivo XAML não encontrado em: {}".format(xaml_path))
        forms.alert("⚠️ Arquivo de interface não encontrado. Usando modo simplificado.")
    
    # Criar e mostrar a janela
    window = FiltroAvancadoWindow()
    window.ShowDialog()
    
except Exception as e:
    logger.error("Erro fatal: {}".format(traceback.format_exc()))
    forms.alert("❌ Erro crítico ao abrir o filtro:\n{}".format(str(e)))