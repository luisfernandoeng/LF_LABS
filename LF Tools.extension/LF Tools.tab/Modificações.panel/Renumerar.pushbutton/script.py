# -*- coding: utf-8 -*-
"""
Renumeracao Inteligente de Elementos
Versao profissional com interface WPF
"""
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import StorageType, Transaction, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
import clr
import re
import os
import sys

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

__title__ = "Renumerar\nElementos"
__author__ = "Luis Fernando - Eng. Eletricista"
__doc__ = "Interface profissional para renumeracao inteligente de elementos"

uidoc = revit.uidoc
doc = revit.doc

SCRIPT_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(SCRIPT_DIR, "RenumerarElementos.xaml")


class RenumerarWindow:
    """Classe principal da interface de renumeracao"""
    
    def __init__(self, xaml_file, elementos_selecionados, metodo_selecao):
        """
        Inicializa a janela WPF
        
        Args:
            xaml_file: Caminho do arquivo XAML
            elementos_selecionados: Lista de elementos ja selecionados
            metodo_selecao: "manual" ou "atual"
        """
        self.selected_elements = elementos_selecionados
        self.metodo_original = metodo_selecao
        self.window = self._load_xaml(xaml_file)
        if self.window:
            self._setup_controls()
            self._setup_events()
            self._configurar_metodo_inicial()
            self._carregar_parametros()
            self._update_preview()
        
    def _load_xaml(self, xaml_file):
        """Carrega a interface XAML usando StreamReader"""
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
        
        # ComboBox
        self.combo_parametro = self.window.FindName("ComboBox_Parametro")
        self.combo_criterio = self.window.FindName("ComboBox_Criterio")
        
        # TextBlocks
        self.txt_status = self.window.FindName("TextBlock_Status")
        self.txt_previa = self.window.FindName("TextBlock_Previa")
        self.txt_preview_example = self.window.FindName("TextBlock_PreviewExample")
        self.txt_info_selecao = self.window.FindName("TextBlock_InfoSelecao")
        
        # Buttons
        self.btn_aplicar = self.window.FindName("Button_Aplicar")
        self.btn_fechar = self.window.FindName("Button_Fechar")
    
    def _setup_events(self):
        """Configura os eventos dos controles"""
        self.btn_aplicar.Click += self.on_aplicar_renumeracao
        self.btn_fechar.Click += self.on_fechar
        
        self.txt_prefixo.TextChanged += self.on_preview_update
        self.txt_numero_inicial.TextChanged += self.on_preview_update
        self.txt_incremento.TextChanged += self.on_preview_update
        self.txt_digitos.TextChanged += self.on_preview_update
        self.combo_parametro.SelectionChanged += self.on_preview_update
        
        self.radio_manual.Checked += self.on_metodo_changed
        self.radio_reordenar.Checked += self.on_metodo_changed
    
    def _configurar_metodo_inicial(self):
        """Configura o metodo baseado na selecao inicial"""
        if self.metodo_original == "manual":
            self.radio_manual.IsChecked = True
            self.combo_criterio.IsEnabled = False
            info_text = "Ordem PRESERVADA conforme cliques (1º clique = 1º numero)"
        else:
            self.radio_manual.IsChecked = True  # Default para manter ordem atual
            self.combo_criterio.IsEnabled = False
            info_text = "Ordem atual da selecao do Revit"
        
        self.txt_info_selecao.Text = "{} elementos selecionados | {}".format(
            len(self.selected_elements),
            info_text
        )
    
    def on_metodo_changed(self, sender, args):
        """Handler quando muda o metodo de ordenacao"""
        if self.radio_reordenar.IsChecked:
            self.combo_criterio.IsEnabled = True
            self.txt_info_selecao.Text = "{} elementos selecionados | Serao REORDENADOS".format(
                len(self.selected_elements)
            )
        else:
            self.combo_criterio.IsEnabled = False
            self.txt_info_selecao.Text = "{} elementos selecionados | Ordem PRESERVADA".format(
                len(self.selected_elements)
            )
    
    def _update_status(self, message, color="#2C3E50"):
        """Atualiza o texto de status"""
        self.txt_status.Text = message
        from System.Windows.Media import SolidColorBrush, ColorConverter
        converter = ColorConverter()
        self.txt_status.Foreground = SolidColorBrush(converter.ConvertFromString(color))
    
    def _update_preview(self, sender=None, args=None):
        """Atualiza a previa da numeracao"""
        try:
            prefixo = self.txt_prefixo.Text or ""
            num_inicial = int(self.txt_numero_inicial.Text or "1")
            incremento = int(self.txt_incremento.Text or "1")
            digitos = int(self.txt_digitos.Text or "2")
            
            exemplos = []
            num = num_inicial
            for i in range(3):
                num_formatado = str(num).zfill(digitos)
                exemplos.append("{}{}".format(prefixo, num_formatado))
                num += incremento
            
            preview_text = ", ".join(exemplos) + "..."
            self.txt_preview_example.Text = "Exemplo: " + preview_text
            
            total = len(self.selected_elements)
            ultimo_num = num_inicial + (incremento * (total - 1))
            ultimo_formatado = str(ultimo_num).zfill(digitos)
            
            self.txt_previa.Text = "{} elementos | De {} ate {}{}".format(
                total,
                exemplos[0],
                prefixo,
                ultimo_formatado
            )
                
        except ValueError:
            self.txt_preview_example.Text = "Valores invalidos - use apenas numeros"
    
    def on_preview_update(self, sender, args):
        """Handler para atualizacao de previa"""
        self._update_preview()
    
    def _carregar_parametros(self):
        """Carrega parametros comuns aos elementos selecionados"""
        if not self.selected_elements:
            return
        
        first_elem = self.selected_elements[0]
        params_dict = {}
        
        for p in first_elem.Parameters:
            if p.StorageType in (StorageType.String, StorageType.Integer) and not p.IsReadOnly:
                params_dict[p.Definition.Name] = p.StorageType
        
        for elem in self.selected_elements[1:]:
            elem_params = set()
            for p in elem.Parameters:
                if p.Definition.Name in params_dict:
                    if p.StorageType == params_dict[p.Definition.Name] and not p.IsReadOnly:
                        elem_params.add(p.Definition.Name)
            
            params_dict = {k: v for k, v in params_dict.items() if k in elem_params}
        
        self.combo_parametro.Items.Clear()
        for param_name in sorted(params_dict.keys()):
            self.combo_parametro.Items.Add(param_name)
        
        if self.combo_parametro.Items.Count > 0:
            self.combo_parametro.SelectedIndex = 0
            self._update_status(
                "Pronto! Selecione o parametro e clique em RENUMERAR",
                "#27AE60"
            )
        else:
            self._update_status(
                "AVISO: Nenhum parametro editavel comum encontrado!",
                "#E67E22"
            )
    
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
            # Mantem ordem original
            return self.selected_elements
        else:
            # Reordena por criterio
            criterio = str(self.combo_criterio.SelectedItem) if self.combo_criterio.SelectedItem else None
            
            if not criterio or criterio == "ID do Elemento":
                return sorted(self.selected_elements, key=lambda e: e.Id.IntegerValue)
            elif criterio == "Nome (A-Z)":
                return sorted(self.selected_elements, key=lambda e: e.Name or "")
            elif criterio == "Nivel":
                return sorted(self.selected_elements, key=lambda e: e.LevelId.IntegerValue if hasattr(e, 'LevelId') else 0)
            else:
                return self.selected_elements
    
    def on_aplicar_renumeracao(self, sender, args):
        """Handler para aplicar a renumeracao"""
        if not self.selected_elements:
            forms.alert("Nenhum elemento selecionado!")
            return
        
        if not self.combo_parametro.SelectedItem:
            forms.alert("Selecione um parametro!")
            return
        
        try:
            prefixo = self._limpar_texto(self.txt_prefixo.Text or "")
            num_inicial = int(self.txt_numero_inicial.Text)
            incremento = int(self.txt_incremento.Text)
            digitos = max(1, int(self.txt_digitos.Text))
            param_name = str(self.combo_parametro.SelectedItem)
            
        except ValueError:
            forms.alert("Valores numericos invalidos!")
            return
        
        # Info sobre ordenacao
        ordem_info = "Ordem PRESERVADA" if self.radio_manual.IsChecked else "REORDENADOS por {}".format(
            self.combo_criterio.SelectedItem if self.combo_criterio.SelectedItem else "criterio"
        )
        
        confirmacao = forms.alert(
            "Confirma a renumeracao de {} elementos?\n\n"
            "Parametro: {}\n"
            "Inicio: {}{}\n"
            "Incremento: {}\n"
            "Ultimo: {}{}\n\n"
            "Ordenacao: {}".format(
                len(self.selected_elements),
                param_name,
                prefixo,
                str(num_inicial).zfill(digitos),
                incremento,
                prefixo,
                str(num_inicial + incremento * (len(self.selected_elements) - 1)).zfill(digitos),
                ordem_info
            ),
            yes=True,
            no=True
        )
        
        if not confirmacao:
            return
        
        self._executar_renumeracao(param_name, prefixo, num_inicial, incremento, digitos)
    
    def _executar_renumeracao(self, param_name, prefixo, num_inicial, incremento, digitos):
        """Executa a renumeracao em uma transacao"""
        erros = []
        sucesso = 0
        
        # Ordena elementos
        elementos_ordenados = self._ordenar_elementos()
        
        with Transaction(doc, "Renumerar Elementos") as t:
            t.Start()
            
            contador = num_inicial
            for elem in elementos_ordenados:
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
                        erros.append("ID {}: Parametro nao editavel".format(elem.Id.IntegerValue))
                        
                except Exception as e:
                    erros.append("ID {}: {}".format(elem.Id.IntegerValue, str(e)))
            
            t.Commit()
        
        self._mostrar_resultado(sucesso, erros)
    
    def _mostrar_resultado(self, sucesso, erros):
        """Mostra o resultado da operacao"""
        if erros:
            mensagem = "{} elementos renumerados com sucesso\n\n{} erros:\n\n{}".format(
                sucesso,
                len(erros),
                "\n".join(erros[:10])
            )
            if len(erros) > 10:
                mensagem += "\n\n... e mais {} erros".format(len(erros) - 10)
            
            forms.alert(mensagem, title="Renumeracao Concluida")
        else:
            forms.alert(
                "Renumeracao concluida com sucesso!\n\n"
                "{} elementos foram renumerados.".format(sucesso),
                title="Sucesso"
            )
            self.window.Close()
    
    def on_fechar(self, sender, args):
        """Handler para fechar a janela"""
        self.window.Close()
    
    def show(self):
        """Exibe a janela"""
        if self.window:
            self.window.ShowDialog()


def selecionar_elementos():
    """
    Fluxo inicial: pergunta metodo e seleciona elementos
    Retorna: (lista_elementos, metodo)
    """
    # Pergunta o metodo de selecao
    opcoes = [
        "Selecao Manual (clique um por um - preserva ordem exata)",
        "Usar Selecao Atual do Revit"
    ]
    
    escolha = forms.CommandSwitchWindow.show(
        opcoes,
        message="Como deseja selecionar os elementos?"
    )
    
    if not escolha:
        script.exit()
    
    elementos = []
    metodo = ""
    
    try:
        if "Manual" in escolha:
            # Selecao manual
            metodo = "manual"
            selected_refs = []
            forms.alert(
                "Selecione os elementos NA ORDEM DESEJADA.\n\n"
                "Cada clique = ordem de numeracao\n\n"
                "Pressione ESC quando terminar.",
                title="Selecao Manual"
            )
            
            while True:
                try:
                    ref = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        "Selecione o proximo elemento (ESC para finalizar)"
                    )
                    selected_refs.append(ref)
                except OperationCanceledException:
                    break
            
            elementos = [doc.GetElement(ref) for ref in selected_refs]
            
        else:
            # Selecao atual
            metodo = "atual"
            selection = uidoc.Selection.GetElementIds()
            
            if not selection:
                forms.alert("Nenhum elemento selecionado no Revit!")
                script.exit()
            
            elementos = [doc.GetElement(eid) for eid in selection]
        
        if not elementos:
            forms.alert("Nenhum elemento selecionado!")
            script.exit()
        
        return elementos, metodo
        
    except Exception as e:
        forms.alert("Erro ao selecionar elementos:\n{}".format(str(e)))
        script.exit()


# Execucao principal
if __name__ == '__main__':
    try:
        # Passo 1: Seleciona elementos primeiro
        elementos, metodo = selecionar_elementos()
        
        # Passo 2: Abre interface com elementos ja selecionados
        app = RenumerarWindow(XAML_FILE, elementos, metodo)
        app.show()
        
    except Exception as e:
        forms.alert("Erro fatal: {}".format(str(e)), title="Erro")