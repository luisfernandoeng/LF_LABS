# -*- coding: utf-8 -*-
import re
from pyrevit import revit, DB, forms

# ---------------------------
# MODELO DE ITEM PARA PREVIEW
# ---------------------------
class PreviewItem(object):
    def __init__(self, element):
        self.element = element
        self.elem_id = element.Id.IntegerValue
        self.category = element.Category.Name if element.Category else ""
        # nome apresentável (instância ou tipo)
        try:
            self.elem_name = revit.query.get_name(element)
        except Exception:
            self.elem_name = str(self.elem_id)
        self.original = ""   # valor atual do parâmetro escolhido (preenchido depois)
        self.new = ""        # valor após aplicar regras
        self.is_selected = True

    def apply_rules(self, prefix="", suffix="", regex_find="", regex_replace="", find_simple="", replace_simple="", use_numbering=False, counter=0):
        txt = self.original or ""
        # regex substitution first (if provided)
        if regex_find:
            try:
                txt = re.sub(regex_find, regex_replace or "", txt)
            except Exception as e:
                txt = "ERRO_REGEX: {}".format(e)
        # simple find/replace (literal)
        if find_simple:
            try:
                txt = txt.replace(find_simple, replace_simple or "")
            except Exception:
                pass
        # prefix/suffix
        if prefix:
            txt = prefix + txt
        if suffix:
            txt = txt + suffix
        # numbering
        if use_numbering:
            txt = "{}_{:02d}".format(txt, counter)
        self.new = txt

# ---------------------------
# MAIN WINDOW
# ---------------------------
class RenameWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)

        # lista de PreviewItem (padrão: elementos selecionados)
        self.preview_items = []

        # Conectar eventos
        self._bind_events()

        # Popular parâmetros e elementos iniciais
        self._load_elements_and_params()

        # atualizar preview inicial
        self.update_preview(None, None)

    # --------------------------------
    # BIND DE EVENTOS
    # --------------------------------
    def _bind_events(self):
        # safe attach (control pode existir no XAML)
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
        
        # text changed live preview
        text_controls = [
            'prefix_tb', 'suffix_tb', 'old_format_tb', 'new_format_tb', 
            'find_simple_tb', 'replace_simple_tb', 'filter_tb'
        ]
        for tb in text_controls:
            if hasattr(self, tb):
                getattr(self, tb).TextChanged += self.update_preview
        
        if hasattr(self, 'numbering_cb'):
            self.numbering_cb.Checked += self.update_preview
            self.numbering_cb.Unchecked += self.update_preview
        if hasattr(self, 'param_cb'):
            # when param changes, reload preview values
            self.param_cb.SelectionChanged += self.update_preview
        if hasattr(self, 'use_selection_cb'):
            self.use_selection_cb.Checked += self.on_selection_changed
            self.use_selection_cb.Unchecked += self.on_selection_changed

    def on_selection_changed(self, sender, args):
        """Recarrega elementos quando a opção de seleção muda"""
        self._load_elements_and_params()
        self.update_preview(None, None)

    # --------------------------------
    # LOAD ELEMENTS AND PARAMETERS
    # --------------------------------
    def _load_elements_and_params(self):
        # carregar elementos: usa seleção se marcada, caso contrário pede seleção
        use_selection = hasattr(self, 'use_selection_cb') and self.use_selection_cb.IsChecked
        sel = list(revit.get_selection().elements) if use_selection else []
        
        # If nothing selected and use_selection is checked, show warning
        if use_selection and not sel:
            forms.alert("Nenhum elemento selecionado. Selecione elementos na vista antes de usar esta opção.", title="Aviso")
            self.preview_items = []
            if hasattr(self, 'preview_dg'):
                self.preview_dg.ItemsSource = self.preview_items
            return

        # If no selection and not using selection, get all elements of certain categories
        if not sel and not use_selection:
            collector = DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType()
            sel = list(collector.ToElements())
            # Filter to common editable categories
            editable_categories = [
                DB.BuiltInCategory.OST_Walls,
                DB.BuiltInCategory.OST_Floors,
                DB.BuiltInCategory.OST_Doors,
                DB.BuiltInCategory.OST_Windows,
                DB.BuiltInCategory.OST_Rooms,
                DB.BuiltInCategory.OST_GenericModel
            ]
            sel = [el for el in sel if el.Category and el.Category.Id.IntegerValue in 
                  [cat.value__ for cat in editable_categories]]

        # criar PreviewItem por elemento
        self.preview_items = [PreviewItem(el) for el in sel if el]
        if hasattr(self, 'preview_dg'):
            self.preview_dg.ItemsSource = self.preview_items

        # carregar parâmetros string editáveis (a partir da seleção)
        params = set(["Name", "Family: Name"])
        for el in sel:
            if el:
                for p in el.Parameters:
                    if not p.IsReadOnly and p.StorageType == DB.StorageType.String:
                        params.add(p.Definition.Name)
        # ordenar e popular combobox
        if hasattr(self, 'param_cb'):
            self.param_cb.ItemsSource = sorted(list(params))
            if len(params) > 0:
                self.param_cb.SelectedIndex = 0

    # --------------------------------
    # UPDATE PREVIEW
    # --------------------------------
    def update_preview(self, sender, args):
        # pegar valores dos controles com segurança
        prefix = self.prefix_tb.Text if hasattr(self, 'prefix_tb') else ""
        suffix = self.suffix_tb.Text if hasattr(self, 'suffix_tb') else ""
        regex_find = self.old_format_tb.Text if hasattr(self, 'old_format_tb') else ""
        regex_replace = self.new_format_tb.Text if hasattr(self, 'new_format_tb') else ""
        find_simple = self.find_simple_tb.Text if hasattr(self, 'find_simple_tb') else ""
        replace_simple = self.replace_simple_tb.Text if hasattr(self, 'replace_simple_tb') else ""
        use_numbering = self.numbering_cb.IsChecked if hasattr(self, 'numbering_cb') else False
        filter_text = (self.filter_tb.Text or "").lower() if hasattr(self, 'filter_tb') else ""
        selected_param = self.param_cb.SelectedItem if hasattr(self, 'param_cb') else None

        # atualizar original values based on selected_param
        for item in self.preview_items:
            el = item.element
            # determine original depending on param choice
            if selected_param == "Name":
                try:
                    item.original = revit.query.get_name(el) or ""
                except Exception:
                    item.original = ""
            elif selected_param == "Family: Name":
                try:
                    item.original = revit.query.get_name(el.Family) if hasattr(el, 'Family') and el.Family else ""
                except Exception:
                    item.original = ""
            else:
                try:
                    p = el.LookupParameter(selected_param) if selected_param else None
                    item.original = p.AsString() if p else ""
                except Exception:
                    item.original = ""

        # apply rules only to selected items and apply counter for numbering
        counter = 1
        for item in self.preview_items:
            if item.is_selected:
                item.apply_rules(prefix=prefix, suffix=suffix,
                                 regex_find=regex_find, regex_replace=regex_replace,
                                 find_simple=find_simple, replace_simple=replace_simple,
                                 use_numbering=use_numbering, counter=counter)
                counter += 1
            else:
                # if not selected, keep new empty or same as original
                item.new = item.original

        # filter datasource if filter provided (search in elem_name or original)
        if hasattr(self, 'preview_dg'):
            if filter_text:
                filtered = [it for it in self.preview_items if 
                           (it.elem_name and filter_text in it.elem_name.lower()) or 
                           (it.original and filter_text in it.original.lower())]
                self.preview_dg.ItemsSource = filtered
            else:
                self.preview_dg.ItemsSource = self.preview_items

            # refresh grid
            try:
                self.preview_dg.Items.Refresh()
            except Exception:
                pass

    # --------------------------------
    # Selecionar / Deselecionar
    # --------------------------------
    def select_all(self, sender, args):
        for it in self.preview_items:
            it.is_selected = True
        try:
            self.preview_dg.Items.Refresh()
        except Exception:
            pass

    def select_none(self, sender, args):
        for it in self.preview_items:
            it.is_selected = False
        try:
            self.preview_dg.Items.Refresh()
        except Exception:
            pass

    # --------------------------------
    # APLICAR ALTERAÇÕES
    # --------------------------------
    def apply_changes(self, sender, args):
        selected_param = self.param_cb.SelectedItem if hasattr(self, 'param_cb') else None
        if not selected_param:
            forms.alert("Selecione um parâmetro para alterar.", title="Aviso")
            return

        items_to_apply = [it for it in (self.preview_dg.ItemsSource or []) if it.is_selected and it.new != it.original]
        if not items_to_apply:
            forms.alert("Nenhuma alteração detectada nos itens selecionados.", title="Aviso")
            return

        # confirmar
        if not forms.alert("Aplicar alterações em {} itens?".format(len(items_to_apply)), yes=True, no=True):
            return

        try:
            with revit.Transaction("Alteração massiva de parâmetros"):
                success_count = 0
                error_items = []
                
                for it in items_to_apply:
                    el = it.element
                    try:
                        if selected_param == "Name":
                            # element name change (works for certain element types)
                            try:
                                el.Name = it.new
                                success_count += 1
                            except Exception:
                                # some elements not allowed to set Name property: fallback to param "Name" if exists
                                p = el.LookupParameter("Name")
                                if p and not p.IsReadOnly:
                                    p.Set(it.new)
                                    success_count += 1
                                else:
                                    error_items.append("{} (não é possível renomear)".format(it.elem_name))
                        elif selected_param == "Family: Name":
                            if hasattr(el, 'Family') and el.Family:
                                el.Family.Name = it.new
                                success_count += 1
                            else:
                                error_items.append("{} (sem família)".format(it.elem_name))
                        else:
                            p = el.LookupParameter(selected_param)
                            if p and not p.IsReadOnly:
                                p.Set(it.new)
                                success_count += 1
                            else:
                                error_items.append("{} (parâmetro não encontrado ou somente leitura)".format(it.elem_name))
                    except Exception as e:
                        error_items.append("{} ({})".format(it.elem_name, str(e)))

            # Mostrar resultado
            result_msg = "{} de {} itens alterados com sucesso!".format(success_count, len(items_to_apply))
            if error_items:
                result_msg += "\n\nErros em {} itens:\n- ".format(len(error_items)) + "\n- ".join(error_items)
                forms.alert(result_msg, title="Resultado com Erros")
            else:
                forms.alert(result_msg, title="Sucesso")
                
        except Exception as e:
            forms.alert("Erro ao aplicar alterações:\n{}".format(str(e)), title="Erro")

        # recarregar preview (puxa novos valores)
        self._load_elements_and_params()
        self.update_preview(None, None)

# ---------------------------
# RUN
# ---------------------------
if __name__ == "__main__":
    try:
        if revit.doc:
            RenameWindow("RenameWindow.xaml").show(modal=True)
        else:
            forms.alert("Documento do Revit não acessível.", title="Erro")
    except Exception as e:
        forms.alert("Erro ao iniciar a ferramenta:\n{}".format(str(e)), title="Erro")
    