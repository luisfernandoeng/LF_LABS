# -*- coding: utf-8 -*-
import re
from pyrevit import revit, DB, forms

# ---------------------------
# MODELO DE ITEM 
# ---------------------------
class PreviewItem(object):
    def __init__(self, main_element, instances=None, is_itemized=False):
        """
        main_element: Pode ser um ElementType (se agrupado) ou uma Instance (se itemizado)
        instances: Lista de instancias (se agrupado). Se itemizado, é None ou lista unitária.
        is_itemized: Flag para saber como tratar o nome
        """
        self.element = main_element
        self.instances = instances if instances else [main_element]
        self.is_itemized = is_itemized
        
        self.elem_id = main_element.Id.IntegerValue
        self.category = main_element.Category.Name if main_element.Category else "Outros"
        
        self.elem_name = self._get_display_name(main_element)
        
        self.original = ""
        self.new = ""
        self.is_selected = True

    def _get_display_name(self, el):
        try:
            if self.is_itemized:
                # Se itemizado, mostra: Familia : Tipo [ID]
                el_type = revit.doc.GetElement(el.GetTypeId())
                fam = el_type.FamilyName if el_type else ""
                typ = revit.query.get_name(el_type) if el_type else ""
                return "{} : {} [{}]".format(fam, typ, el.Id)
            else:
                # Se agrupado (Tipo): Familia : Tipo
                fam_name = el.FamilyName if hasattr(el, 'FamilyName') else ""
                type_name = revit.query.get_name(el)
                if fam_name and fam_name != type_name:
                    return "{} : {}".format(fam_name, type_name)
                return type_name
        except:
            return str(el.Id)

    def apply_rules(self, prefix="", suffix="", regex_find="", regex_replace="", find_simple="", replace_simple="", use_numbering=False, numbering_config=None):
        txt = self.original or ""
        
        # Regex
        if regex_find:
            try:
                txt = re.sub(regex_find, regex_replace or "", txt)
            except: pass 
        
        # Simples
        if find_simple:
            try:
                txt = txt.replace(find_simple, replace_simple or "")
            except: pass
                
        # Prefixo/Sufixo
        if prefix: txt = prefix + txt
        if suffix: txt = txt + suffix
        
        # Numeração Customizada
        if use_numbering and numbering_config:
            # numbering_config = {'separator': '-', 'current_val': 1}
            sep = numbering_config.get('separator', '-')
            val = numbering_config.get('current_val', 1)
            
            # Formato: Texto + Separador + Numero (02 digitos)
            txt = "{}{:02d}".format(txt + sep, val)
            
        self.new = txt

# ---------------------------
# JANELA PRINCIPAL
# ---------------------------
class RenameWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.preview_items = []
        self._bind_events()
        self._load_elements_and_params()
        self.update_preview(None, None)

    def _bind_events(self):
        if hasattr(self, 'update_btn'): self.update_btn.Click += self.update_preview
        if hasattr(self, 'apply_btn'): self.apply_btn.Click += self.apply_changes
        if hasattr(self, 'cancel_btn'): self.cancel_btn.Click += lambda s, a: self.Close()
        if hasattr(self, 'select_all_btn'): self.select_all_btn.Click += self.select_all
        if hasattr(self, 'select_none_btn'): self.select_none_btn.Click += self.select_none
        
        # Inputs de texto
        text_controls = ['prefix_tb', 'suffix_tb', 'old_format_tb', 'new_format_tb', 
                        'find_simple_tb', 'replace_simple_tb', 'filter_tb', 
                        'num_separator_tb', 'start_num_tb']
        for tb in text_controls:
            if hasattr(self, tb): getattr(self, tb).TextChanged += self.update_preview
        
        # Checkboxes
        cbs = ['numbering_cb', 'use_selection_cb', 'itemize_cb']
        for cb in cbs:
            if hasattr(self, cb):
                getattr(self, cb).Checked += self.on_config_changed
                getattr(self, cb).Unchecked += self.on_config_changed

        if hasattr(self, 'param_cb'):
            self.param_cb.SelectionChanged += self.update_preview

    def on_config_changed(self, sender, args):
        # Se mudou "itemize" ou "selection", precisa recarregar os elementos base
        if sender.Name in ['use_selection_cb', 'itemize_cb']:
            self._load_elements_and_params()
        self.update_preview(None, None)

    def _load_elements_and_params(self):
        doc = revit.doc
        use_selection = self.use_selection_cb.IsChecked if hasattr(self, 'use_selection_cb') else True
        is_itemized = self.itemize_cb.IsChecked if hasattr(self, 'itemize_cb') else False
        
        raw_elements = list(revit.get_selection().elements)
        
        self.preview_items = []

        # MODO 1: ITEMIZADO (Cada instância é uma linha)
        if is_itemized:
            for el in raw_elements:
                self.preview_items.append(PreviewItem(el, instances=[el], is_itemized=True))

        # MODO 2: AGRUPADO POR TIPO (Padrão)
        else:
            type_map = {} 
            type_obj_map = {}

            if raw_elements:
                for el in raw_elements:
                    try:
                        type_id = el.GetTypeId()
                        if type_id and type_id != DB.ElementId.InvalidElementId:
                            tid_val = type_id.IntegerValue
                            if tid_val not in type_obj_map:
                                el_type = doc.GetElement(type_id)
                                if el_type:
                                    type_obj_map[tid_val] = el_type
                                    type_map[tid_val] = []
                            if tid_val in type_map:
                                type_map[tid_val].append(el)
                    except: pass

            for tid, el_type in type_obj_map.items():
                instances = type_map.get(tid, [])
                self.preview_items.append(PreviewItem(el_type, instances, is_itemized=False))

        # Atualiza Grid
        if hasattr(self, 'preview_dg'):
            self.preview_dg.ItemsSource = self.preview_items

        # Carregar Parâmetros (Mesma lógica híbrida)
        params = set(["Type Name", "Nome do Tipo", "Comments", "Comentários", "Mark", "Marca"]) 
        
        # Amostragem para performance
        sample_items = self.preview_items[:20] if len(self.preview_items) > 20 else self.preview_items
        
        for item in sample_items:
            # Se itemizado, olha a instância. Se agrupado, olha o tipo e a 1a instância
            els_to_check = [item.element]
            if not item.is_itemized and item.instances:
                els_to_check.append(item.instances[0])
            
            for el in els_to_check:
                for p in el.Parameters:
                    if not p.IsReadOnly and p.StorageType == DB.StorageType.String:
                        params.add(p.Definition.Name)
        
        if hasattr(self, 'param_cb'):
            current = self.param_cb.SelectedItem
            sorted_params = sorted(list(params))
            self.param_cb.ItemsSource = sorted_params
            
            if current in sorted_params:
                self.param_cb.SelectedItem = current
            else:
                defaults = ["Mark", "Marca", "Type Mark", "Marca de tipo"]
                for d in defaults:
                    if d in sorted_params:
                        self.param_cb.SelectedItem = d
                        break
                else:
                    if sorted_params: self.param_cb.SelectedIndex = 0

    def update_preview(self, sender, args):
        prefix = self.prefix_tb.Text if hasattr(self, 'prefix_tb') else ""
        suffix = self.suffix_tb.Text if hasattr(self, 'suffix_tb') else ""
        regex_find = self.old_format_tb.Text if hasattr(self, 'old_format_tb') else ""
        regex_replace = self.new_format_tb.Text if hasattr(self, 'new_format_tb') else ""
        find_simple = self.find_simple_tb.Text if hasattr(self, 'find_simple_tb') else ""
        replace_simple = self.replace_simple_tb.Text if hasattr(self, 'replace_simple_tb') else ""
        filter_text = (self.filter_tb.Text or "").lower() if hasattr(self, 'filter_tb') else ""
        selected_param = self.param_cb.SelectedItem if hasattr(self, 'param_cb') else None

        # Configurações de Numeração
        use_numbering = self.numbering_cb.IsChecked if hasattr(self, 'numbering_cb') else False
        num_separator = self.num_separator_tb.Text if hasattr(self, 'num_separator_tb') else "-"
        try:
            start_num = int(self.start_num_tb.Text) if hasattr(self, 'start_num_tb') and self.start_num_tb.Text else 1
        except:
            start_num = 1

        # 1. Carregar valores originais
        for item in self.preview_items:
            val = ""
            el_to_read = item.element # Pode ser Tipo ou Instancia
            
            if selected_param == "Type Name" or selected_param == "Nome do Tipo":
                # Se for nome do tipo, precisa garantir que lemos o Tipo
                if item.is_itemized:
                    # Se itemizado, item.element é Instance. Pegamos o tipo dele.
                    el_type = revit.doc.GetElement(item.element.GetTypeId())
                    val = revit.query.get_name(el_type)
                else:
                    val = revit.query.get_name(item.element)
            else:
                # Tenta ler do elemento principal
                p = el_to_read.LookupParameter(selected_param)
                if p:
                    val = p.AsString()
                else:
                    # Se não achou e for agrupado, tenta na instância filha
                    if not item.is_itemized and item.instances:
                        p_inst = item.instances[0].LookupParameter(selected_param)
                        if p_inst: val = p_inst.AsString()
            
            item.original = val or ""

        # 2. Aplicar regras
        current_counter = start_num
        for item in self.preview_items:
            if item.is_selected:
                # Prepara config de numeracao
                n_conf = {'separator': num_separator, 'current_val': current_counter}
                
                item.apply_rules(prefix, suffix, regex_find, regex_replace, 
                               find_simple, replace_simple, use_numbering, n_conf)
                
                current_counter += 1
            else:
                item.new = item.original

        # 3. Filtro
        if hasattr(self, 'preview_dg'):
            if filter_text:
                filtered = [it for it in self.preview_items if 
                           filter_text in it.elem_name.lower() or 
                           filter_text in it.original.lower()]
                self.preview_dg.ItemsSource = filtered
            else:
                self.preview_dg.ItemsSource = self.preview_items
            try:
                self.preview_dg.Items.Refresh()
            except: pass

    def select_all(self, sender, args):
        for it in self.preview_items: it.is_selected = True
        self.update_preview(None, None)

    def select_none(self, sender, args):
        for it in self.preview_items: it.is_selected = False
        self.update_preview(None, None)

    def apply_changes(self, sender, args):
        selected_param = self.param_cb.SelectedItem if hasattr(self, 'param_cb') else None
        if not selected_param: return

        items_to_apply = [it for it in (self.preview_dg.ItemsSource or []) if it.is_selected and it.new != it.original]
        if not items_to_apply:
            forms.alert("Nenhuma alteração pendente.", title="Aviso")
            return

        if not forms.alert("Aplicar alterações em {} itens?".format(len(items_to_apply)), yes=True, no=True):
            return

        with revit.Transaction("Renomear+"):
            count = 0
            errors = []
            
            for it in items_to_apply:
                try:
                    # Define qual elemento será editado
                    # Se itemizado: editamos item.element (Instance)
                    # Se agrupado: editamos item.element (Type) OU suas instâncias
                    
                    targets = []
                    
                    # Se for parametro de TIPO explicito
                    if selected_param == "Type Name" or selected_param == "Nome do Tipo":
                        # Sempre edita o Tipo
                        if it.is_itemized:
                            el_type = revit.doc.GetElement(it.element.GetTypeId())
                            el_type.Name = it.new
                            count += 1
                            continue
                        else:
                            it.element.Name = it.new
                            count += 1
                            continue
                    
                    # Parametros normais
                    # Estrategia: Tenta no elemento principal. Se falhar e for agrupado, tenta nas instancias.
                    p = it.element.LookupParameter(selected_param)
                    if p and not p.IsReadOnly:
                        p.Set(it.new)
                        count += 1
                    else:
                        # Se não deu no principal, verifica se temos sub-instancias (modo agrupado)
                        if not it.is_itemized and it.instances:
                            sub_ok = False
                            for inst in it.instances:
                                p_sub = inst.LookupParameter(selected_param)
                                if p_sub and not p_sub.IsReadOnly:
                                    p_sub.Set(it.new)
                                    sub_ok = True
                            if sub_ok: count += len(it.instances)
                            else:
                                errors.append(it.elem_name)
                        else:
                            errors.append(it.elem_name)

                except Exception as e:
                    errors.append("{} ({})".format(it.elem_name, e))

            res_msg = "Processados: {} operações.".format(count)
            if errors:
                res_msg += "\nErros em: " + ", ".join(errors[:3])
            forms.alert(res_msg, title="Concluído")
        
        self._load_elements_and_params()
        self.update_preview(None, None)

if __name__ == "__main__":
    RenameWindow("RenameWindow.xaml").show(modal=True)