# -*- coding: utf-8 -*-
"""
Script: Smart Crop Turbo - OTIMIZADO
Descricao: Recorte ultra-rapido (2 cliques), analise de elementos perdidos e copia de crop region.
Otimizacoes: Snaps minimos, caching, colecoes rapidas
"""

from pyrevit import forms, revit, DB, UI, script
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import Visibility
import os
import re

# Variaveis Globais
doc = revit.doc
uidoc = revit.uidoc
view = revit.active_view

# ============================================================================
# CLASSE PARA BINDING DE VISTAS
# ============================================================================

class ViewItem(INotifyPropertyChanged):
    """Item de vista com binding para WPF"""
    def __init__(self, view_obj, level_name, discipline):
        self._view = view_obj
        self._name = view_obj.Name
        self._is_selected = False
        self._level_name = level_name
        self._discipline = discipline
        self._property_changed_handlers = []

    @property
    def View(self):
        return self._view

    @property
    def Name(self):
        return self._name

    @property
    def LevelName(self):
        return self._level_name

    @property
    def Discipline(self):
        return self._discipline

    @property
    def DisplaySubtitle(self):
        parts = []
        if self._level_name:
            parts.append(self._level_name)
        if self._discipline and self._discipline != u'geral':
            parts.append(self._discipline.capitalize())
        return u'  ·  '.join(parts)

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")

    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def OnPropertyChanged(self, prop_name):
        for handler in self._property_changed_handlers:
            handler(self, PropertyChangedEventArgs(prop_name))

# ============================================================================
# JANELA WPF UNIFICADA - SMART CROP
# ============================================================================

class SmartCropWindow(forms.WPFWindow):
    def __init__(self, current_view, crop_info, has_selection, has_active_crop):
        try:
            xaml_file = script.get_bundle_file('SmartCrop.xaml')
            forms.WPFWindow.__init__(self, xaml_file)
            
            self.current_view = current_view
            self.crop_info = crop_info
            self.has_selection = has_selection
            self.has_active_crop = has_active_crop
            self._all_view_items = []
            self.view_items = ObservableCollection[object]()

            self.TextBlock_CurrentViewName.Text = current_view.Name
            self.TextBlock_CropBoxInfo.Text = crop_info if crop_info else u"Não ativo"

            if has_selection:
                self.Panel_AdjustMode.Visibility = Visibility.Visible
                self.Button_AdjustFromSelection.Click += self.adjust_from_selection_click

            if has_active_crop:
                self.Panel_CopyMode.Visibility = Visibility.Visible
                self.Button_ApplyCopy.Visibility = Visibility.Visible
                self._collect_views()
                self.ListBox_Views.ItemsSource = self.view_items
                self.Button_FilterSameLevel.Click      += self.filter_same_level
                self.Button_FilterSameDiscipline.Click += self.filter_same_discipline
                self.Button_FilterAll.Click            += self.filter_all
                self.Button_ApplyCopy.Click            += self.apply_copy_click
                self.TextBox_SearchViews.TextChanged   += self.on_search_changed
                for item in self.view_items:
                    item.add_PropertyChanged(self.on_selection_changed)
                self.update_counter()

            self.Button_Cancel.Click += self.cancel_click
            
        except Exception as e:
            forms.alert("Erro ao carregar interface: {}".format(str(e)))
            raise
    
    def _collect_views(self):
        """Coleta vistas de planta e corte/elevação que suportam crop."""
        def _add(v, level_name):
            if v.Id == self.current_view.Id or v.IsTemplate:
                return
            if not hasattr(v, 'CropBox'):
                return
            discipline = self._detect_discipline(v.Name)
            item = ViewItem(v, level_name, discipline)
            self._all_view_items.append(item)
            self.view_items.Add(item)

        for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan):
            level = v.GenLevel
            _add(v, level.Name if level else u"")

        for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewSection):
            _add(v, u"")
    
    def _detect_discipline(self, view_name):
        """Detecta a disciplina pelo nome da vista - MELHORADO"""
        view_name_lower = view_name.lower()
        
        # Remover acentos para melhor matching
        import unicodedata
        view_name_normalized = unicodedata.normalize('NFD', view_name_lower)
        view_name_normalized = ''.join(c for c in view_name_normalized if unicodedata.category(c) != 'Mn')
        
        disciplines = {
            u'elétrica': [
                u'eletrica', u'eletrica', u'eletrico', u'eletrico',
                u'elet', u'ele', u'electrical', u'power'
            ],
            u'hidráulica': [
                u'hidraulica', u'hidraulica', u'hidro', u'agua', u'agua', 
                u'esgoto', u'plumbing', u'water', u'sanitary', u'hidr'
            ],
            u'avac': [
                u'avac', u'hvac', u'ar condicionado', u'ventilacao', u'ventilacao',
                u'climatizacao', u'climatizacao', u'mechanical', u'mech', u'mecanica', u'mecanica'
            ],
            u'estrutura': [
                u'estrutura', u'estrutural', u'concreto', u'fundacao', u'fundacao',
                u'structural', u'struct', u'est', u'pilar', u'viga', u'laje'
            ],
            u'arquitetura': [
                u'arquitetura', u'arquitetonica', u'arquitetonica', u'arq',
                u'architectural', u'arch', u'layout', u'planta', u'piso'
            ]
        }
        
        # Tentar match em ambas as versões (original e normalizada)
        for discipline, keywords in disciplines.items():
            for keyword in keywords:
                if keyword in view_name_lower or keyword in view_name_normalized:
                    return discipline
        
        return u'geral'
    
    # ========================================================================
    # MODO 1: AJUSTAR DA SELEÇÃO
    # ========================================================================
    
    def adjust_from_selection_click(self, sender, args):
        """Ajusta o crop box para a seleção atual"""
        selection = revit.get_selection()

        if not selection:
            forms.alert("Nenhum elemento selecionado.", warn_icon=True)
            return

        # Ler margem do campo (metros → pés internos do Revit)
        try:
            margin_m = float((self.TextBox_Margin.Text or u"0.60").replace(u",", u"."))
            margin_m = max(0.0, margin_m)
        except (ValueError, AttributeError):
            margin_m = 0.60
        margin = margin_m / 0.3048

        min_x, min_y = float('inf'), float('inf')
        max_x, max_y = float('-inf'), float('-inf')

        for el in selection:
            bbox = el.get_BoundingBox(self.current_view)
            if bbox:
                min_x = min(min_x, bbox.Min.X)
                min_y = min(min_y, bbox.Min.Y)
                max_x = max(max_x, bbox.Max.X)
                max_y = max(max_y, bbox.Max.Y)

        if min_x == float('inf'):
            forms.alert("Nenhum elemento com bounding box válido.", warn_icon=True)
            return
        
        # Aplicar com TransactionGroup
        tg = DB.TransactionGroup(doc, "Ajustar Crop da Seleção")
        tg.Start()
        
        try:
            # Ativar crop se necessário
            if not self.current_view.CropBoxActive:
                t1 = DB.Transaction(doc, "Ativar Crop")
                t1.Start()
                self.current_view.CropBoxActive = True
                self.current_view.CropBoxVisible = True
                t1.Commit()
            
            current = self.current_view.CropBox
            
            # Ajustar crop
            t2 = DB.Transaction(doc, "Ajustar Crop")
            t2.Start()
            
            new_crop = DB.BoundingBoxXYZ()
            new_crop.Min = DB.XYZ(min_x - margin, min_y - margin, current.Min.Z)
            new_crop.Max = DB.XYZ(max_x + margin, max_y + margin, current.Max.Z)
            new_crop.Transform = current.Transform
            
            self.current_view.CropBox = new_crop
            
            t2.Commit()
            tg.Assimilate()
            
            forms.alert(u"✅ Crop ajustado para a seleção!", title="Sucesso")
            
            # Atualizar info do crop na interface
            crop = self.current_view.CropBox
            self.TextBlock_CropBoxInfo.Text = u"({:.2f}, {:.2f}) até ({:.2f}, {:.2f})".format(
                crop.Min.X, crop.Min.Y, crop.Max.X, crop.Max.Y
            )
            
            # Mostrar o painel de cópia se estava escondido
            if self.Panel_CopyMode.Visibility != Visibility.Visible:
                self.Panel_CopyMode.Visibility = Visibility.Visible
                self.Button_ApplyCopy.Visibility = Visibility.Visible
                self._collect_views()
                self.ListBox_Views.ItemsSource = self.view_items
                self.Button_FilterSameLevel.Click      += self.filter_same_level
                self.Button_FilterSameDiscipline.Click += self.filter_same_discipline
                self.Button_FilterAll.Click            += self.filter_all
                self.Button_ApplyCopy.Click            += self.apply_copy_click
                self.TextBox_SearchViews.TextChanged   += self.on_search_changed
                for item in self.view_items:
                    item.add_PropertyChanged(self.on_selection_changed)
                self.update_counter()
            
        except:
            tg.RollBack()
            raise
    
    # ========================================================================
    # MODO 2: COPIAR PARA OUTRAS VISTAS
    # ========================================================================
    
    def filter_same_level(self, sender, args):
        current_level = self.current_view.GenLevel
        current_level_name = current_level.Name if current_level else u""
        for item in self._all_view_items:
            item.IsSelected = (item.LevelName == current_level_name)
        self.update_counter()

    def filter_same_discipline(self, sender, args):
        current_discipline = self._detect_discipline(self.current_view.Name)
        for item in self._all_view_items:
            item.IsSelected = (item.Discipline == current_discipline)
        self.update_counter()

    def filter_all(self, sender, args):
        for item in self._all_view_items:
            item.IsSelected = True
        self.update_counter()
    
    def on_search_changed(self, sender, args):
        """Filtra a lista de vistas pelo texto digitado."""
        term = (self.TextBox_SearchViews.Text or u'').strip().lower()
        if not term:
            self.ListBox_Views.ItemsSource = self.view_items
        else:
            filtered = [i for i in self._all_view_items
                        if term in i.Name.lower()
                        or term in i.LevelName.lower()
                        or term in i.Discipline.lower()]
            self.ListBox_Views.ItemsSource = filtered

    def on_selection_changed(self, sender, args):
        self.update_counter()

    def update_counter(self):
        count = sum(1 for item in self._all_view_items if item.IsSelected)
        self.Button_ApplyCopy.Content = u"Copiar para {} Vista{}".format(
            count, u"s" if count != 1 else u""
        )
    
    def apply_copy_click(self, sender, args):
        """Aplica o crop region às vistas selecionadas"""
        selected_views = [item.View for item in self._all_view_items if item.IsSelected]
        
        if not selected_views:
            forms.alert(u"Selecione pelo menos uma vista.", warn_icon=True)
            return
        
        copy_annotation = self.CheckBox_CopyAnnotationCrop.IsChecked
        enable_crop = self.CheckBox_EnableCrop.IsChecked
        show_crop_box = self.CheckBox_ShowCropBox.IsChecked
        show_annotation_crop = self.CheckBox_ShowAnnotationCrop.IsChecked
        
        # Obter crop box da vista atual
        source_crop = self.current_view.CropBox
        
        # Obter annotation crop se necessário
        source_annotation_crop = None
        if copy_annotation:
            try:
                source_annotation_crop = self.current_view.GetAnnotationCrop()
            except:
                pass
        
        # Aplicar às vistas selecionadas
        tg = DB.TransactionGroup(doc, "Copiar Crop Region")
        tg.Start()
        
        success_count = 0
        error_count = 0
        
        try:
            for target_view in selected_views:
                try:
                    t = DB.Transaction(doc, "Aplicar Crop")
                    t.Start()
                    
                    # Ativar crop se necessário
                    if enable_crop and not target_view.CropBoxActive:
                        target_view.CropBoxActive = True
                    
                    # Controlar visibilidade do crop box
                    target_view.CropBoxVisible = show_crop_box
                    
                    # Copiar crop box
                    new_crop = DB.BoundingBoxXYZ()
                    new_crop.Min = DB.XYZ(source_crop.Min.X, source_crop.Min.Y, source_crop.Min.Z)
                    new_crop.Max = DB.XYZ(source_crop.Max.X, source_crop.Max.Y, source_crop.Max.Z)
                    new_crop.Transform = source_crop.Transform
                    
                    target_view.CropBox = new_crop
                    
                    # Copiar annotation crop se solicitado
                    if copy_annotation and source_annotation_crop:
                        try:
                            target_view.SetAnnotationCrop(source_annotation_crop)
                        except:
                            pass
                    
                    # Controlar visibilidade do annotation crop
                    try:
                        target_view.AnnotationCropActive = show_annotation_crop
                    except:
                        pass
                    
                    t.Commit()
                    success_count += 1
                    
                except Exception as e:
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()
                    error_count += 1
            
            tg.Assimilate()
            
        except:
            tg.RollBack()
            raise
        
        # Mensagem de resultado
        msg = u"✅ Crop region copiado com sucesso!\n\n"
        msg += u"Vistas atualizadas: {}\n".format(success_count)
        if error_count > 0:
            msg += u"Erros: {}".format(error_count)
        
        forms.alert(msg, title="Concluído")
        self.Close()
    
    def cancel_click(self, sender, args):
        """Fecha a janela"""
        self.Close()

# ============================================================================
# MAIN - INTERFACE UNIFICADA
# ============================================================================

def main():
    """Ponto de entrada principal com validação"""
    
    # Verificar se a vista atual suporta crop
    if not hasattr(view, 'CropBox'):
        forms.alert(u"A vista atual não suporta Crop Region.", warn_icon=True)
        return
    
    # Verificar se há elementos selecionados OU se o crop está ativo
    selection = revit.get_selection()
    has_selection = len(selection) > 0
    has_active_crop = view.CropBoxActive
    
    if not has_selection and not has_active_crop:
        forms.alert(
            u"Para usar o Smart Crop, você precisa:\n\n"
            u"• Selecionar elementos (para ajustar crop), OU\n"
            u"• Ter o Crop Box ativado (para copiar crop)\n\n"
            u"Faça uma das duas opções e tente novamente.",
            title="Smart Crop",
            warn_icon=True
        )
        return
    
    # Obter informações do crop box (se ativo)
    crop_info = None
    if has_active_crop:
        crop = view.CropBox
        crop_info = u"({:.2f}, {:.2f}) até ({:.2f}, {:.2f})".format(
            crop.Min.X, crop.Min.Y,
            crop.Max.X, crop.Max.Y
        )
    
    # Abrir interface unificada
    window = SmartCropWindow(view, crop_info, has_selection, has_active_crop)
    window.ShowDialog()

# Executar
main()