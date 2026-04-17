# -*- coding: utf-8 -*-
import re
import clr
from pyrevit import revit, DB, forms

# Importações necessárias para ObservableCollection
clr.AddReference('System')
from System.Collections.ObjectModel import ObservableCollection

# ---------------------------
# MODELO DE ITEM 
# ---------------------------
class PreviewItem(object):
    """Representa um item na tabela de preview (pode ser Tipo ou Instância)"""
    
    def __init__(self, main_element, instances=None, is_itemized=False):
        """
        Args:
            main_element: ElementType (agrupado) ou Instance (itemizado)
            instances: Lista de instâncias relacionadas
            is_itemized: True = cada instância é uma linha | False = agrupa por tipo
        """
        self.element = main_element
        self.instances = instances if instances else [main_element]
        self.is_itemized = is_itemized
        
        self.elem_id = main_element.Id.IntegerValue
        self.category = main_element.Category.Name if main_element.Category else "Outros"
        self.elem_name = self._get_display_name(main_element)
        
        self.original = ""
        self.new = ""
        self.IsSelected = True  # IMPORTANTE: IsSelected com maiúscula para o binding WPF

    def _get_display_name(self, el):
        """Gera nome de exibição baseado no modo (itemizado ou agrupado)"""
        try:
            if self.is_itemized:
                # Modo Itemizado: Família : Tipo [ID]
                el_type = revit.doc.GetElement(el.GetTypeId())
                fam = el_type.FamilyName if el_type else ""
                typ = revit.query.get_name(el_type) if el_type else ""
                return "{} : {} [{}]".format(fam, typ, el.Id)
            else:
                # Modo Agrupado: Família : Tipo
                fam_name = el.FamilyName if hasattr(el, 'FamilyName') else ""
                type_name = revit.query.get_name(el)
                if fam_name and fam_name != type_name:
                    return "{} : {}".format(fam_name, type_name)
                return type_name
        except:
            return str(el.Id)

    def apply_rules(self, name_pattern="{Original}", regex_find="", regex_replace="", 
                    find_simple="", replace_simple="", use_numbering=False, numbering_config=None, case_converter=None, use_pure_regex=False):
        """Aplica todas as regras de transformação ao texto"""
        original_txt = self.original or ""
        
        # 1. Composição de Nome e Parâmetros Dinâmicos {ParamNome}
        txt = name_pattern.replace("{Original}", original_txt)
        
        def replace_dynamic(match):
            param_name = match.group(1)
            if param_name == "Original": return original_txt
            # Tenta pegar o valor do parâmetro do elemento
            p = self.element.LookupParameter(param_name)
            if p:
                return str(p.AsString() or p.AsValueString() or "")
            # Fallback para as instâncias se não achar no tipo
            if not self.is_itemized and self.instances:
                p_inst = self.instances[0].LookupParameter(param_name)
                if p_inst:
                     return str(p_inst.AsString() or p_inst.AsValueString() or "")
            return match.group(0) # Retorna a tag original se não achar
            
        txt = re.sub(r'\{(.*?)\}', replace_dynamic, txt)
        
        # 2. Padrões Inteligentes (ou Regex Puro)
        if regex_find:
            try:
                if use_pure_regex:
                    txt = re.sub(regex_find, regex_replace or "", txt)
                else:
                    # Smart Pattern Converter
                    pattern = ""
                    for char in regex_find:
                        if char == '#': pattern += r'\d'
                        elif char == '@': pattern += r'[a-zA-Z]'
                        elif char == '?': pattern += r'.'
                        elif char in '()|.*+^{}\\$[]':
                            if char in '()': pattern += char # Keep capturing groups intact
                            else: pattern += "\\" + char # Escape other regex special chars
                        else:
                            pattern += char

                    # Convert $1, $2 substitution logic to Python \g<1> syntax
                    repl = regex_replace or ""
                    repl = re.sub(r'\$(\d+)', r'\\g<\1>', repl)

                    txt = re.sub(pattern, repl, txt)
            except:
                pass
        
        # 3. Substituição Simples (AGORA CASE-INSENSITIVE)
        if find_simple:
            try:
                pattern = re.escape(find_simple)
                txt = re.sub(pattern, replace_simple or "", txt, flags=re.IGNORECASE)
            except:
                pass
        
        # 4. Numeração Sequencial
        if use_numbering and numbering_config:
            sep = numbering_config.get('separator', '-')
            val = numbering_config.get('current_val', 1)
            padding = numbering_config.get('padding', 2)
            pos = numbering_config.get('position', 0)
            
            num_str = str(val).zfill(padding)
            if pos == 1: # Início
                txt = "{}{}{}".format(num_str, sep, txt)
            else: # Fim
                txt = "{}{}{}".format(txt, sep, num_str)
        
        # 5. Conversão de Caixa (Case Converter)
        if case_converter == 'upper':
            txt = txt.upper()
        elif case_converter == 'lower':
            txt = txt.lower()
        elif case_converter == 'title':
            txt = txt.title()

        self.new = txt


# ---------------------------
# JANELA PRINCIPAL
# ---------------------------
class RenameWindow(forms.WPFWindow):
    """Interface principal do plugin Renomear+"""
    
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        
        # IMPORTANTE: Usar ObservableCollection como no exportador de folhas
        self.preview_items = ObservableCollection[PreviewItem]()
        self.preview_dg.ItemsSource = self.preview_items
        self.case_mode = None  # ['upper', 'lower', 'title']
        
        self._bind_events()
        self._load_elements_and_params()
        self.update_preview(None, None)

    def _bind_events(self):
        """Conecta todos os eventos da interface"""
        # Botões principais
        if hasattr(self, 'update_btn'):
            self.update_btn.Click += self.update_preview
        if hasattr(self, 'apply_btn'):
            self.apply_btn.Click += self.apply_changes
        if hasattr(self, 'cancel_btn'):
            self.cancel_btn.Click += lambda s, a: self.Close()
        if hasattr(self, 'select_all_btn'):
            self.select_all_btn.Click += self.select_all
        if hasattr(self, 'select_none_btn'):
            self.select_none_btn.Click += self.select_none
        
        # Campos de texto (atualização automática)
        text_controls = [
            'name_pattern_tb', 'old_format_tb', 'new_format_tb', 
            'find_simple_tb', 'replace_simple_tb', 'filter_tb', 
            'num_separator_tb', 'start_num_tb', 'num_digits_tb'
        ]
        for tb in text_controls:
            if hasattr(self, tb):
                getattr(self, tb).TextChanged += self.update_preview
        
        # Checkboxes
        checkboxes = ['numbering_cb', 'use_regex_cb']
        for cb in checkboxes:
            if hasattr(self, cb):
                checkbox = getattr(self, cb)
                checkbox.Checked += self.on_config_changed
                checkbox.Unchecked += self.on_config_changed

        # ComboBox de parâmetros
        if hasattr(self, 'param_cb'):
            self.param_cb.SelectionChanged += self.update_preview

        # Case Converters
        case_buttons = ['upper_btn', 'lower_btn', 'title_btn']
        for btn in case_buttons:
            if hasattr(self, btn):
                getattr(self, btn).Click += self.on_case_convert_clicked

        # Filtro para a Tabela
        if hasattr(self, 'filter_tb'):
            self.filter_tb.TextChanged += self.update_preview

        # Botão de adicionar parâmetro ao composer
        if hasattr(self, 'btn_AddParam'):
            self.btn_AddParam.Click += self.on_add_param_clicked

    def on_add_param_clicked(self, sender, args):
        """Adiciona a tag selecionada ao TextBox de Composição"""
        if hasattr(self, 'p_AvailableParams') and self.p_AvailableParams.Text:
            param = self.p_AvailableParams.Text
            if param:
                tag = "{%s}" % param
                current_text = self.name_pattern_tb.Text
                # Se o texto atual for vazio, só adiciona. Se não, adiciona com espaço.
                if current_text:
                    self.name_pattern_tb.Text = current_text + " " + tag
                else:
                    self.name_pattern_tb.Text = tag
                
                # Joga o foco pro textbox
                self.name_pattern_tb.Focus()
                self.name_pattern_tb.CaretIndex = len(self.name_pattern_tb.Text)

    def on_case_convert_clicked(self, sender, args):
        """Define o modo de conversão e atualiza a prévia"""
        if sender.Name == 'upper_btn':
            self.case_mode = 'upper'
        elif sender.Name == 'lower_btn':
            self.case_mode = 'lower'
        elif sender.Name == 'title_btn':
            self.case_mode = 'title'
        self.update_preview(None, None)

    def on_config_changed(self, sender, args):
        """Recarrega elementos quando modo de exibição muda"""
        self.update_preview(None, None)

    def _load_elements_and_params(self):
        """Carrega elementos selecionados e parâmetros disponíveis"""
        doc = revit.doc
        
        raw_elements = list(revit.get_selection().elements)
        
        # IMPORTANTE: Limpar a ObservableCollection
        self.preview_items.Clear()

        # MODO ITEMIZADO: Cada instância = uma linha
        for el in raw_elements:
            self.preview_items.Add(PreviewItem(el, instances=[el], is_itemized=True))

        # Carrega parâmetros disponíveis
        self._load_available_parameters()

    def _load_available_parameters(self):
        """Detecta parâmetros disponíveis nos elementos selecionados"""
        params = set([
            "Type Name", "Nome do Tipo",
            "Comments", "Comentários",
            "Mark", "Marca"
        ])

        # Amostra até 20 itens para performance
        sample_size = min(20, self.preview_items.Count)
        has_text_notes = False

        for i in range(sample_size):
            item = self.preview_items[i]

            if isinstance(item.element, DB.TextNote):
                has_text_notes = True
                params.add("Text")
                continue

            elements_to_check = [item.element]
            if not item.is_itemized and item.instances:
                elements_to_check.append(item.instances[0])

            for el in elements_to_check:
                for p in el.Parameters:
                    if not p.IsReadOnly and p.StorageType == DB.StorageType.String:
                        params.add(p.Definition.Name)

        # Atualiza ComboBox de Alvo
        if hasattr(self, 'param_cb'):
            current = self.param_cb.SelectedItem
            sorted_params = sorted(list(params))
            self.param_cb.ItemsSource = sorted_params

            # Mantém seleção ou escolhe padrão
            if current in sorted_params:
                self.param_cb.SelectedItem = current
            else:
                defaults = (["Text"] if has_text_notes else []) + ["Mark", "Marca", "Type Mark", "Marca de tipo"]
                for d in defaults:
                    if d in sorted_params:
                        self.param_cb.SelectedItem = d
                        break
                else:
                    if sorted_params:
                        self.param_cb.SelectedIndex = 0

        # Atualiza ComboBox do Composer
        if hasattr(self, 'p_AvailableParams'):
            self.p_AvailableParams.ItemsSource = sorted_params
            if sorted_params and not self.p_AvailableParams.SelectedItem:
                self.p_AvailableParams.SelectedIndex = 0

    def update_preview(self, sender, args):
        """Atualiza preview com as regras aplicadas"""
        # Coleta valores dos controles
        name_pattern = self.name_pattern_tb.Text if hasattr(self, 'name_pattern_tb') else "{Original}"
        regex_find = self.old_format_tb.Text if hasattr(self, 'old_format_tb') else ""
        regex_replace = self.new_format_tb.Text if hasattr(self, 'new_format_tb') else ""
        find_simple = self.find_simple_tb.Text if hasattr(self, 'find_simple_tb') else ""
        replace_simple = self.replace_simple_tb.Text if hasattr(self, 'replace_simple_tb') else ""
        filter_text = (self.filter_tb.Text or "").lower() if hasattr(self, 'filter_tb') else ""
        selected_param = self.param_cb.SelectedItem if hasattr(self, 'param_cb') else None

        # Configurações de numeração
        use_numbering = self.numbering_cb.IsChecked if hasattr(self, 'numbering_cb') else False
        num_separator = self.num_separator_tb.Text if hasattr(self, 'num_separator_tb') else "-"
        try:
            start_num = int(self.start_num_tb.Text) if hasattr(self, 'start_num_tb') and self.start_num_tb.Text else 1
        except:
            start_num = 1

        try:
            num_padding = int(self.num_digits_tb.Text) if hasattr(self, 'num_digits_tb') and self.num_digits_tb.Text else 2
            if num_padding < 1: num_padding = 1
        except:
            num_padding = 2

        num_position = self.num_pos_cb.SelectedIndex if hasattr(self, 'num_pos_cb') else 0
        use_pure_regex = self.use_regex_cb.IsChecked if hasattr(self, 'use_regex_cb') else False

        # Carrega valores originais dos parâmetros
        self._load_original_values(selected_param)

        # Aplica regras de transformação
        current_counter = start_num
        for item in self.preview_items:
            if item.IsSelected:
                numbering_config = {
                    'separator': num_separator, 
                    'current_val': current_counter,
                    'padding': num_padding,
                    'position': num_position
                }
                item.apply_rules(
                    name_pattern, regex_find, regex_replace,
                    find_simple, replace_simple, use_numbering, numbering_config,
                    case_converter=self.case_mode, use_pure_regex=use_pure_regex
                )
                current_counter += 1
            else:
                item.new = item.original

        # --- LÓGICA DE FILTRAGEM (AJUSTADA PARA PARÂMETRO) ---
        filter_txt = ""
        if hasattr(self, 'filter_tb'):
            filter_txt = self.filter_tb.Text.lower()
        elif hasattr(self, 'find_simple_tb'):
            filter_txt = self.find_simple_tb.Text.lower()

        visible_items = []
        for item in self.preview_items:
            # 1. Pega o valor do PARÂMETRO ALVO (Valor Atual)
            val_param = item.original.lower() if item.original else ""
            
            # 2. Pega o NOME DO ELEMENTO (Opcional, mas útil manter)
            val_name = item.elem_name.lower() if item.elem_name else ""
            
            # A Mágica: Verifica se o texto está no Parâmetro OU no Nome
            if not filter_txt or (filter_txt in val_param) or (filter_txt in val_name):
                visible_items.append(item)
        
        # ATUALIZAÇÃO DO GRID
        if hasattr(self, 'preview_dg'):
            from System.Collections.ObjectModel import ObservableCollection
            self.preview_dg.ItemsSource = ObservableCollection[object](visible_items)

    def _load_original_values(self, selected_param):
        """Carrega valores atuais do parâmetro selecionado"""
        if not selected_param:
            return

        for item in self.preview_items:
            val = ""

            if selected_param == "Text" and isinstance(item.element, DB.TextNote):
                val = item.element.Text or ""
            elif selected_param in ["Type Name", "Nome do Tipo"]:
                if item.is_itemized:
                    el_type = revit.doc.GetElement(item.element.GetTypeId())
                    val = revit.query.get_name(el_type) if el_type else ""
                else:
                    val = revit.query.get_name(item.element)
            else:
                p = item.element.LookupParameter(selected_param)
                if p:
                    val = p.AsString() or ""
                elif not item.is_itemized and item.instances:
                    p_inst = item.instances[0].LookupParameter(selected_param)
                    if p_inst:
                        val = p_inst.AsString() or ""

            item.original = val

    def select_all(self, sender, args):
        """Marca todas as checkboxes"""
        for item in self.preview_items:
            item.IsSelected = True
        self.preview_dg.Items.Refresh()

    def select_none(self, sender, args):
        """Desmarca todas as checkboxes"""
        for item in self.preview_items:
            item.IsSelected = False
        self.preview_dg.Items.Refresh()

    def apply_changes(self, sender, args):
        """Aplica as alterações no modelo do Revit"""
        selected_param = self.param_cb.SelectedItem if hasattr(self, 'param_cb') else None
        if not selected_param:
            forms.alert("Selecione um parâmetro.", title="Aviso")
            return

        items_to_apply = [it for it in self.preview_items if it.IsSelected and it.new != it.original]
        
        if not items_to_apply:
            forms.alert("Nenhuma alteração pendente.", title="Aviso")
            return

        if not forms.alert(
            "Aplicar alterações em {} itens?".format(len(items_to_apply)), 
            yes=True, no=True
        ):
            return

        # Executa transação
        with revit.Transaction("Renomear+"):
            count = 0
            errors = []
            
            for item in items_to_apply:
                try:
                    # Texto de Nota de Texto
                    if selected_param == "Text" and isinstance(item.element, DB.TextNote):
                        item.element.Text = item.new
                        count += 1
                        continue

                    # Alteração de nome do tipo
                    if selected_param in ["Type Name", "Nome do Tipo"]:
                        if item.is_itemized:
                            el_type = revit.doc.GetElement(item.element.GetTypeId())
                            if el_type:
                                el_type.Name = item.new
                                count += 1
                        else:
                            item.element.Name = item.new
                            count += 1
                        continue
                    
                    # Alteração de parâmetro
                    p = item.element.LookupParameter(selected_param)
                    if p and not p.IsReadOnly:
                        p.Set(item.new)
                        count += 1
                    elif not item.is_itemized and item.instances:
                        # Tenta nas instâncias se for modo agrupado
                        sub_count = 0
                        for inst in item.instances:
                            p_sub = inst.LookupParameter(selected_param)
                            if p_sub and not p_sub.IsReadOnly:
                                p_sub.Set(item.new)
                                sub_count += 1
                        
                        if sub_count > 0:
                            count += sub_count
                        else:
                            errors.append(item.elem_name)
                    else:
                        errors.append(item.elem_name)

                except Exception as e:
                    errors.append("{} ({})".format(item.elem_name, str(e)))

            # Mensagem final
            msg = "Processados: {} operações.".format(count)
            if errors:
                msg += "\n\nErros em:\n" + "\n".join(errors[:5])
                if len(errors) > 5:
                    msg += "\n... e mais {} itens".format(len(errors) - 5)
            
            forms.alert(msg, title="Concluído")
        
        # Recarrega interface
        self._load_elements_and_params()
        self.update_preview(None, None)


# ---------------------------
# PONTO DE ENTRADA
# ---------------------------
if __name__ == "__main__":
    RenameWindow("RenameWindow.xaml").show(modal=True)