# -*- coding: utf-8 -*-
"""
Renumerador PRO - Interface Moderna
Versao otimizada com identidade visual profissional
Autor: Luis Fernando
"""
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import StorageType, Transaction, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
import clr
import re
import os

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, ColorConverter
from System.IO import StreamReader

__title__ = "Renumerador\nPRO"
__author__ = "Luis Fernando"
__doc__ = "Interface moderna para renumeracao inteligente"

uidoc = revit.uidoc
doc = revit.doc

SCRIPT_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(SCRIPT_DIR, "RenumerarElementos.xaml")


class RenumeradorModerno:
    """Classe principal da interface moderna"""
    
    def __init__(self, xaml_file, elementos_selecionados, metodo_selecao):
        self.selected_elements = elementos_selecionados
        self.metodo_original = metodo_selecao
        self.window = self._load_xaml(xaml_file)
        self.color_converter = ColorConverter()
        self.should_retry_selection = False # Flag para indicar se deve reabrir com nova selecao
        
        if self.window:
            self._setup_controls()
            self._setup_events()
            self._configurar_metodo_inicial()
            self._carregar_parametros()
            self._update_preview()
        
    def _load_xaml(self, xaml_file):
        """Carrega a interface XAML"""
        if not os.path.exists(xaml_file):
            forms.alert("Arquivo XAML nao encontrado:\n{}".format(xaml_file))
            return None
            
        try:
            stream = StreamReader(xaml_file)
            window = XamlReader.Load(stream.BaseStream)
            stream.Close()
            return window
        except Exception as e:
            forms.alert("Erro ao carregar interface:\n{}".format(str(e)))
            return None
    
    def _setup_controls(self):
        """Mapeia os controles da interface"""
        # RadioButtons
        self.radio_manual = self.window.FindName("Radio_ManterOrdem")
        self.radio_reordenar = self.window.FindName("Radio_Reordenar")
        
        # TextBoxes
        self.txt_prefixo = self.window.FindName("TextBox_Prefixo")
        self.txt_numero_inicial = self.window.FindName("TextBox_NumeroInicial")
        self.txt_incremento = self.window.FindName("TextBox_Incremento")
        self.txt_digitos = self.window.FindName("TextBox_Digitos")
        
        # ComboBoxes
        self.combo_parametro = self.window.FindName("ComboBox_Parametro")
        self.combo_criterio = self.window.FindName("ComboBox_Criterio")
        
        # TextBlocks de Info
        self.txt_status = self.window.FindName("TextBlock_Status")
        self.txt_previa = self.window.FindName("TextBlock_Previa")
        self.txt_preview_example = self.window.FindName("TextBlock_PreviewExample")
        self.txt_info_selecao = self.window.FindName("TextBlock_InfoSelecao")
        
        # Buttons
        self.btn_aplicar = self.window.FindName("Button_Aplicar")
        self.btn_fechar = self.window.FindName("Button_Fechar")
        self.btn_selecionar = self.window.FindName("Button_Selecionar")
    
    def _setup_events(self):
        """Configura os eventos dos controles"""
        self.btn_aplicar.Click += self.on_aplicar_renumeracao
        self.btn_fechar.Click += self.on_fechar
        if self.btn_selecionar:
            self.btn_selecionar.Click += self.on_selecionar
        
        # Preview em tempo real
        self.txt_prefixo.TextChanged += self.on_preview_update
        self.txt_numero_inicial.TextChanged += self.on_preview_update
        self.txt_incremento.TextChanged += self.on_preview_update
        self.txt_digitos.TextChanged += self.on_preview_update
        
        # Mudanca de metodo
        self.radio_manual.Checked += self.on_metodo_changed
        self.radio_reordenar.Checked += self.on_metodo_changed
    
    def _configurar_metodo_inicial(self):
        """Configura o metodo baseado na selecao inicial"""
        total = len(self.selected_elements)
        
        if self.metodo_original == "manual":
            self.radio_manual.IsChecked = True
            self.combo_criterio.IsEnabled = False
            info_text = "{} elementos selecionados | Ordem PRESERVADA (sequencia de cliques)".format(total)
        else:
            self.radio_manual.IsChecked = True
            self.combo_criterio.IsEnabled = False
            info_text = "{} elementos selecionados | Ordem atual da selecao".format(total)
        
        self.txt_info_selecao.Text = info_text
        self._update_status("Pronto para configurar", "#FF4EC99B")
    
    def on_metodo_changed(self, sender, args):
        """Handler quando muda o metodo de ordenacao"""
        total = len(self.selected_elements)
        
        if self.radio_reordenar.IsChecked:
            self.combo_criterio.IsEnabled = True
            criterio = str(self.combo_criterio.SelectedItem.Content) if self.combo_criterio.SelectedItem else "criterio"
            self.txt_info_selecao.Text = "{} elementos | Serao REORDENADOS por {}".format(total, criterio)
            self._update_status("Elementos serao reordenados", "#FFD7BA7D")
        else:
            self.combo_criterio.IsEnabled = False
            self.txt_info_selecao.Text = "{} elementos | Ordem PRESERVADA".format(total)
            self._update_status("Ordem sera preservada", "#FF4EC99B")
    
    def _update_status(self, message, color_hex):
        """Atualiza o texto de status com cor"""
        self.txt_status.Text = message
        self.txt_status.Foreground = SolidColorBrush(
            self.color_converter.ConvertFromString(color_hex)
        )

    def on_selecionar(self, sender, args):
        """Handler para o botao de selecionar novos elementos"""
        self.should_retry_selection = True
        self.window.Close()
        
    def _update_preview(self, sender=None, args=None):
        """Atualiza a previa da numeracao em tempo real"""
        try:
            prefixo = self.txt_prefixo.Text or ""
            num_inicial = int(self.txt_numero_inicial.Text or "1")
            incremento = int(self.txt_incremento.Text or "1")
            digitos = int(self.txt_digitos.Text or "2")
            
            # Gerar exemplos
            exemplos = []
            num = num_inicial
            for i in range(min(5, len(self.selected_elements))):
                num_formatado = str(num).zfill(digitos)
                exemplos.append("{}{}".format(prefixo, num_formatado))
                num += incremento
            
            # Atualizar exemplo
            if len(exemplos) > 3:
                preview_text = "{}, {}, {} ... {}".format(
                    exemplos[0], 
                    exemplos[1], 
                    exemplos[2],
                    exemplos[-1]
                )
            else:
                preview_text = ", ".join(exemplos)
                if len(self.selected_elements) > len(exemplos):
                    preview_text += " ..."
            
            self.txt_preview_example.Text = preview_text
            
            # Calcular ultimo numero
            total = len(self.selected_elements)
            if total > 0:
                ultimo_num = num_inicial + (incremento * (total - 1))
                ultimo_formatado = str(ultimo_num).zfill(digitos)
                
                self.txt_previa.Text = "{} elementos | De {} ate {}{}".format(
                    total,
                    exemplos[0] if exemplos else "?",
                    prefixo,
                    ultimo_formatado
                )
            else:
                self.txt_previa.Text = "Nenhum elemento selecionado"
                
        except ValueError:
            self.txt_preview_example.Text = "⚠ Valores invalidos - use apenas numeros"
            self.txt_previa.Text = "Erro nos parametros"
    
    def on_preview_update(self, sender, args):
        """Handler para atualizacao de previa"""
        self._update_preview()
    
    def _carregar_parametros(self):
        """Carrega parametros comuns editaveis"""
        if not self.selected_elements:
            self._update_status("Nenhum elemento selecionado!", "#FFD25555")
            return
        
        first_elem = self.selected_elements[0]
        params_dict = {}
        
        # Coletar parametros do primeiro elemento
        for p in first_elem.Parameters:
            if p.StorageType in (StorageType.String, StorageType.Integer) and not p.IsReadOnly:
                params_dict[p.Definition.Name] = p.StorageType
        
        # Filtrar apenas parametros comuns a todos
        for elem in self.selected_elements[1:]:
            elem_params = set()
            for p in elem.Parameters:
                if p.Definition.Name in params_dict:
                    if p.StorageType == params_dict[p.Definition.Name] and not p.IsReadOnly:
                        elem_params.add(p.Definition.Name)
            
            # Manter apenas parametros comuns
            params_dict = {k: v for k, v in params_dict.items() if k in elem_params}
        
        # Preencher ComboBox
        self.combo_parametro.Items.Clear()
        for param_name in sorted(params_dict.keys()):
            self.combo_parametro.Items.Add(param_name)
        
        if self.combo_parametro.Items.Count > 0:
            # Tentar selecionar "Mark" por padrao
            mark_index = -1
            for i in range(self.combo_parametro.Items.Count):
                if str(self.combo_parametro.Items[i]).lower() == "mark":
                    mark_index = i
                    break
            
            if mark_index >= 0:
                self.combo_parametro.SelectedIndex = mark_index
            else:
                self.combo_parametro.SelectedIndex = 0
            
            self._update_status("Pronto! Configure e clique em RENUMERAR", "#FF4EC99B")
        else:
            self._update_status("AVISO: Nenhum parametro editavel comum!", "#FFD7BA7D")
            self.btn_aplicar.IsEnabled = False
    
    def _limpar_texto(self, texto):
        """Remove caracteres invalidos"""
        return re.sub(r'[<>:"/\\|?*]', '', texto)
    
    def _get_parameter(self, elem, param_name):
        """Busca parametro no elemento ou tipo"""
        param = elem.LookupParameter(param_name)
        if not param:
            type_id = elem.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type:
                    param = elem_type.LookupParameter(param_name)
        return param
    
    def _ordenar_elementos(self):
        """Retorna elementos na ordem correta"""
        if self.radio_manual.IsChecked:
            return self.selected_elements
        else:
            criterio = str(self.combo_criterio.SelectedItem.Content) if self.combo_criterio.SelectedItem else None
            
            if not criterio or "ID do Elemento" in criterio:
                return sorted(self.selected_elements, key=lambda e: e.Id.IntegerValue)
            elif "Nome" in criterio:
                return sorted(self.selected_elements, key=lambda e: e.Name or "")
            elif "vel" in criterio:  # Nivel
                return sorted(
                    self.selected_elements, 
                    key=lambda e: e.LevelId.IntegerValue if hasattr(e, 'LevelId') else 0
                )
            else:
                return self.selected_elements
    
    def on_aplicar_renumeracao(self, sender, args):
        """Handler para aplicar a renumeracao"""
        if not self.selected_elements:
            forms.alert("Nenhum elemento selecionado!")
            return
        
        if not self.combo_parametro.SelectedItem:
            forms.alert("Selecione um parametro de destino!")
            return
        
        # Validar entradas
        try:
            prefixo = self._limpar_texto(self.txt_prefixo.Text or "")
            num_inicial = int(self.txt_numero_inicial.Text)
            incremento = int(self.txt_incremento.Text)
            digitos = max(1, int(self.txt_digitos.Text))
            param_name = str(self.combo_parametro.SelectedItem)
            
        except ValueError:
            forms.alert("Valores numericos invalidos!\n\nVerifique os campos numericos.", title="Erro")
            return
        
        # Info sobre ordenacao
        if self.radio_manual.IsChecked:
            ordem_info = "Ordem PRESERVADA (atual)"
        else:
            criterio = str(self.combo_criterio.SelectedItem.Content) if self.combo_criterio.SelectedItem else "criterio"
            ordem_info = "REORDENADOS por: {}".format(criterio)
        
        # Calcular ultimo numero
        ultimo_num = num_inicial + (incremento * (len(self.selected_elements) - 1))
        
        # Calcular ultimo numero
        ultimo_num = num_inicial + (incremento * (len(self.selected_elements) - 1))
        
        # Executa direto sem confirmacao
        self._executar_renumeracao(param_name, prefixo, num_inicial, incremento, digitos)
    
    def _executar_renumeracao(self, param_name, prefixo, num_inicial, incremento, digitos):
        """Executa a renumeracao em uma transacao"""
        erros = []
        sucesso = 0
        
        # Ordenar elementos
        elementos_ordenados = self._ordenar_elementos()
        
        # Atualizar status
        self._update_status("Processando renumeracao...", "#FF569CD6")
        self.btn_aplicar.IsEnabled = False
        
        try:
            with Transaction(doc, "Renumerar Elementos") as t:
                t.Start()
                
                contador = num_inicial
                for idx, elem in enumerate(elementos_ordenados, 1):
                    try:
                        param = self._get_parameter(elem, param_name)
                        
                        if param and not param.IsReadOnly:
                            numero_formatado = str(contador).zfill(digitos)
                            
                            if param.StorageType == StorageType.String:
                                valor_final = self._limpar_texto(prefixo + numero_formatado)
                                param.Set(valor_final)
                            elif param.StorageType == StorageType.Integer:
                                param.Set(contador)
                            
                            sucesso += 1
                            contador += incremento
                        else:
                            erros.append("#{} - ID {}: Parametro nao editavel".format(
                                idx, elem.Id.IntegerValue
                            ))
                            
                    except Exception as e:
                        erros.append("#{} - ID {}: {}".format(
                            idx, elem.Id.IntegerValue, str(e)
                        ))
                
                t.Commit()
            
            self._mostrar_resultado(sucesso, erros)
            
        except Exception as e:
            self._update_status("Erro na transacao!", "#FFD25555")
            forms.alert("Erro ao executar renumeracao:\n\n{}".format(str(e)), title="Erro")
            self.btn_aplicar.IsEnabled = True
    
    def _mostrar_resultado(self, sucesso, erros):
        """Mostra o resultado da operacao"""
        # Se tiver erros, avisa o usuario
        if erros:
            mensagem = "Concluido com {} erros.\n\n".format(len(erros))
            for erro in erros[:5]:
                mensagem += erro + "\n"
            if len(erros) > 5: mensagem += "..."
            
            forms.alert(mensagem, title="Avisos")
            # Reabilita interface para tentar novamente
            self.btn_aplicar.IsEnabled = True
        else:
            # Sucesso silencioso - Mantem aberto e avisa no rodape
            self.btn_aplicar.IsEnabled = True
            self._update_status("Concluido com sucesso!", "#FF4EC99B")
            # Opcional: Atualizar preview ou limpar campos se necessario
    
    def on_fechar(self, sender, args):
        """Handler para fechar a janela"""
        self.window.Close()
    
    def show(self):
        """Exibe a janela"""
        if self.window:
            self.window.ShowDialog()


# ═══════════════════════════════════════
#   EXECUCAO PRINCIPAL
# ═══════════════════════════════════════

if __name__ == '__main__':
    try:
        # Tenta pegar selecao atual
        selection = uidoc.Selection.GetElementIds()
        elementos = []
        if selection:
            elementos = [doc.GetElement(eid) for eid in selection]
        
        metodo = "atual"
        if not elementos: metodo = "manual"
            
        while True:
            # Passo 1: Abrir interface moderna
            app = RenumeradorModerno(XAML_FILE, elementos, metodo)
            app.show()
            
            # Passo 2: Verifica se usuario quer selecionar novos elementos
            if app.should_retry_selection:
                try:
                    # Selecao manual UM POR UM para garantir a ordem
                    forms.alert(
                        "Clique nos elementos NA ORDEM DESEJADA\n\n"
                        "Use ESC quando terminar.", 
                        title="Seleção Sequencial"
                    )
                    
                    novos_elementos = []
                    while True:
                        try:
                            ref = uidoc.Selection.PickObject(
                                ObjectType.Element, 
                                "Selecione o elemento {} (ESC para terminar)".format(len(novos_elementos) + 1)
                            )
                            novos_elementos.append(doc.GetElement(ref))
                        except OperationCanceledException:
                            break # Sai do loop de selecao ao cancelar
                    
                    if novos_elementos:
                        elementos = novos_elementos
                        metodo = "manual" # Se selecionou manualmente, respeita a ordem
                    # Se nao selecionou nada ou cancelou sem selecionar, mantem elementos anteriores
                    
                except Exception as ex_sel:
                    pass # Ignora erro geral de selecao
            else:
                # Se fechou ou aplicou, sai do loop
                break
                
    except Exception as e:
        forms.alert(
            "Erro fatal:\n\n{}".format(str(e)), 
            title="Renumerador PRO - Erro"
        )