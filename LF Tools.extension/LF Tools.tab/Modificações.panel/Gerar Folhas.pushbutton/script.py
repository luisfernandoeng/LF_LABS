# coding: utf-8
"""
EXPORTADOR PRO | LUIS FERNANDO
Versao Final: Selecionar Tudo + Interface Corrigida
"""
import clr
import System
from pyrevit import revit, forms, script

# --- CARREGAMENTO DE DLLs ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms') 

import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.Windows.Forms import Application as WinFormsApp
from System.Windows import Input
from System.Windows.Data import CollectionViewSource
import os
import re
import time
import json

doc = revit.doc
uidoc = revit.uidoc

# --- DETECAO DE VERSAO ---
REVIT_YEAR = int(doc.Application.VersionNumber)
HAS_PDF_SUPPORT = REVIT_YEAR >= 2022

# --- CAMINHO DE CONFIGURACAO ---
CONFIG_DIR = os.path.join(os.getenv('APPDATA'), 'pyRevit', 'Extensions', 'LF Tools')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'gerar_folhas_config.json')

# --- CLASSE DE DADOS PARA FOLHAS ---
class SheetItem(object):
    def __init__(self, sheet, name_pattern="{Nome da folha}"):
        self.Element = sheet
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name if sheet.Name else "" 
        self.IsSelected = False
        self.name_pattern = name_pattern
        self._file_name = "" 
        self.PdfFileName = ""
        
        # Status Properties
        self.StatusIcon = ""
        self.StatusColor = "Transparent"
        self.StatusToolTip = ""
        self.HasViews = True # Padrao True, verificado dps
        
        # Inicializa via setter
        self.FileName = self._generate_filename(sheet, name_pattern)

    @property
    def FileName(self):
        return self._file_name
    
    @FileName.setter
    def FileName(self, value):
        self._file_name = value
        self.PdfFileName = value # Sincroniza PDF com o nome editado 

    def _get_param_value(self, sheet, p_name):
        val = ""
        if p_name == "Sheet Number": return sheet.SheetNumber
        if p_name == "Sheet Name": return sheet.Name
        if p_name == "Current Revision": return sheet.GetCurrentRevision()
            
        try:
            param = sheet.LookupParameter(p_name)
            if not param:
                p_list = sheet.GetParameters(p_name)
                if p_list: param = p_list[0]
            
            if param:
                if param.StorageType == DB.StorageType.String:
                    val = param.AsString()
                else:
                    val = param.AsValueString()
        except:
            val = ""
        return val if val else ""

    def _generate_filename(self, sheet, pattern):
        final_name = pattern
        tags = re.findall(r'\{(.*?)\}', pattern)
        for tag in tags:
            val = self._get_param_value(sheet, tag)
            final_name = final_name.replace("{" + tag + "}", val)
        clean_val = re.sub(r'[<>:"/\\|?*]', '', final_name)
        return clean_val.strip()
    
    def update_filename(self, new_pattern):
        self.name_pattern = new_pattern
        self.FileName = self._generate_filename(self.Element, new_pattern)
        self.PdfFileName = self.FileName

    def update_status(self, folder_path):
        """Atualiza icone e cor baseado na existencia do arquivo e conteudo"""
        self.StatusIcon = ""
        self.StatusColor = "Transparent"
        self.StatusToolTip = ""

        # 1. Verifica se folha esta vazia (sem views)
        if not self.HasViews:
            self.StatusIcon = "📄"
            self.StatusColor = "#FF888888" # Cinza
            self.StatusToolTip = "Folha vazia (sem vistas)"
        
        # 2. Verifica se arquivo ja existe (sobrescreve o aviso de vazia se for o caso, ou soma?)
        # Vamos dar prioridade ao aviso de sobrescrita pois eh destrutivo
        if folder_path and os.path.exists(folder_path):
            # Verifica PDF
            pdf_path = os.path.join(folder_path, "{}.pdf".format(self.PdfFileName))
            # Verifica DWG
            dwg_path = os.path.join(folder_path, "{}.dwg".format(self.FileName))
            
            exists_pdf = os.path.exists(pdf_path)
            exists_dwg = os.path.exists(dwg_path)
            
            if exists_pdf or exists_dwg:
                self.StatusIcon = "⚠️"
                self.StatusColor = "#FFD7BA7D" # Amarelo Warning
                found = []
                if exists_pdf: found.append("PDF")
                if exists_dwg: found.append("DWG")
                self.StatusToolTip = "Arquivo(s) existente(s): {}".format(", ".join(found))

# --- FUNCAO AUXILIAR PARA DWG SETTINGS ---
def create_default_dwg_settings():
    """Cria configuracao padrao de DWG Export se nao existir nenhuma"""
    try:
        # Verifica se ja existe algum ExportDWGSettings
        existing = DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings).ToElements()
        if len(list(existing)) > 0:
            return  # Ja existe, nao precisa criar
        
        # Cria as opcoes de exportacao DWG
        dwg_opts = DB.DWGExportOptions()
        
        # 1. Padrao AIA para camadas
        dwg_opts.LayerMapping = "AIA"
        
        # 2. True Colors da vista do Revit
        dwg_opts.Colors = DB.ExportColorMode.TrueColorPerView
        
        # 3. Coordenadas compartilhadas
        dwg_opts.SharedCoords = True
        
        # 4. Unidades em metros
        dwg_opts.TargetUnit = DB.ExportUnit.Meter
        
        # Cria o ExportDWGSettings no documento
        with DB.Transaction(doc, "Criar DWG Export Settings") as t:
            t.Start()
            DB.ExportDWGSettings.Create(doc, "Padrão Luís Fernando", dwg_opts)
            t.Commit()
            
    except Exception as ex:
        print("Erro ao criar DWG settings: " + str(ex))

# --- FUNCOES DE CONFIGURACAO PERSISTENTE ---
def load_config():
    """Carrega configuracoes salvas do arquivo JSON"""
    default_config = {
        "last_folder": "",
        "open_folder_after": True
    }
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
    except:
        pass
    return default_config

def save_config(config):
    """Salva configuracoes no arquivo JSON"""
    try:
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    except Exception as ex:
        print("Erro ao salvar config: " + str(ex))


# --- JANELA PRINCIPAL ---
class LuisExporterWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = script.get_bundle_file('folhas_window.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.sheet_items = ObservableCollection[SheetItem]()
        
        # Configurar CollectionView para filtragem
        self.view_source = CollectionViewSource.GetDefaultView(self.sheet_items)
        self.view_source.Filter = System.Predicate[object](self.filter_sheets)
        
        self.dg_Sheets.ItemsSource = self.view_source
        
        # Eventos Principais
        self.btn_SelectFolder.Click += self.pick_folder
        self.btn_Start.Click += self.run_export
        self.btn_Cancel.Click += self.cancel_export
        self.btn_AddParam.Click += self.add_param_to_pattern
        self.txt_NamePattern.TextChanged += self.on_pattern_changed
        
        # NOVO: Evento Selecionar Tudo
        self.chk_SelectAll.Click += self.toggle_all_sheets
        self.btn_InvertSelection.Click += self.invert_selection
        
        # NOVO: Busca e Filtros
        self.txt_Search.TextChanged += self.on_search_text_changed
        self.btn_SelectVisible.Click += self.select_visible
        
        # NOVO: Sets do Revit
        # NOVO: Sets do Revit
        self.cb_RevitSets.SelectionChanged += self.apply_current_set # Auto-apply
        self.btn_SaveSet.Click += self.save_new_set
        self.btn_DeleteSet.Click += self.delete_current_set

        # Configurações Iniciais
        self.config = load_config()
        self.txt_OutputFolder.Text = self.config.get("last_folder", "")
        self.chk_OpenFolderAfter.IsChecked = self.config.get("open_folder_after", True)


        # Evento de clique na linha (RESTAURADO)
        self.dg_Sheets.SelectionChanged += self.on_row_clicked
        self.is_handling_click = False 
        

        
        self.is_cancelled = False
        self.rb_ZoomFit.IsChecked = True
        self.chk_ExportDWG.IsChecked = True
        
        self.load_sheet_parameters() 
        self.load_sheet_parameters() 
        self.load_sheets()
        
        # Checar arquivos existentes na inicializacao
        self.check_existing_files()
        
        # Carregar Sets do Revit
        self.load_revit_sets()
        
        create_default_dwg_settings()  # Cria settings padrao se nao existir
        self.load_dwg_setups()
        
        if HAS_PDF_SUPPORT:
            self.chk_ExportPDF.IsEnabled = True
            self.pnl_PDFSettings.IsEnabled = True
            self.chk_ExportPDF.IsChecked = True
        else:
            self.chk_ExportPDF.IsChecked = False
            self.chk_ExportPDF.IsEnabled = False
            self.chk_ExportPDF.Content = "PDF Requer Revit 2022+"
            self.pnl_PDFSettings.IsEnabled = False

    # --- LOGICA SELECIONAR TUDO ---
    def toggle_all_sheets(self, sender, args):
        """Marca ou desmarca todas as folhas VISIVEIS"""
        is_checked = bool(self.chk_SelectAll.IsChecked)
        for item in self.sheet_items:
            # Bugfix: Respeitar o filtro
            if self.filter_sheets(item):
                item.IsSelected = is_checked
        self.dg_Sheets.Items.Refresh()

    def invert_selection(self, sender, args):
        """Inverte a seleção atual"""
        for item in self.sheet_items:
            item.IsSelected = not item.IsSelected
        self.dg_Sheets.Items.Refresh()

    def Window_PreviewKeyDown(self, sender, args):
        """Atalhos de teclado: Ctrl+A (Tudo), Ctrl+Shift+A (Inverter)"""
        try:
            if args.Key == Input.Key.A and (Input.Keyboard.Modifiers & Input.ModifierKeys.Control) == Input.ModifierKeys.Control:
                if (Input.Keyboard.Modifiers & Input.ModifierKeys.Shift) == Input.ModifierKeys.Shift:
                    # Ctrl + Shift + A -> Inverter
                    self.invert_selection(None, None)
                else:
                    # Ctrl + A -> Selecionar Tudo (respeitando filtro)
                    visible_items = [i for i in self.sheet_items if self.filter_sheets(i)]
                    
                    if visible_items:
                        # Se TODOS os visiveis estao selecionados -> Desseleciona eles
                        # Senao -> Seleciona todos os visiveis
                        all_visible_selected = all(i.IsSelected for i in visible_items)
                        new_state = not all_visible_selected
                        
                        self.chk_SelectAll.IsChecked = new_state
                        for item in visible_items:
                            item.IsSelected = new_state
                        
                        self.dg_Sheets.Items.Refresh()
                args.Handled = True
        except Exception as ex:
            print("Erro KeyDown: " + str(ex))

    # --- LOGICA DE BUSCA E FILTRO ---
    def filter_sheets(self, item):
        if not self.txt_Search.Text: return True
        s_text = self.txt_Search.Text.lower()
        if s_text in item.Number.lower(): return True
        if s_text in item.Name.lower(): return True
        if s_text in item.FileName.lower(): return True
        return False

    def on_search_text_changed(self, sender, args):
        self.view_source.Refresh()

    def select_visible(self, sender, args):
        """Seleciona apenas os itens visiveis no filtro atual"""
        for item in self.sheet_items:
            # Verifica se item passa no filtro
            if self.filter_sheets(item):
                item.IsSelected = True
            else:
                 # Opcional: Desselecionar os que nao estao visiveis?
                 # Melhor nao mexer nos invisiveis para nao perder selecao anterior
                 pass
        self.dg_Sheets.Items.Refresh()

    # --- LOGICA DE VERIFICACAO DE ARQUIVOS ---
    def check_existing_files(self, folder=None):
        """Verifica quais arquivos ja existem na pasta de destino"""
        try:
            if not folder:
                folder = self.txt_OutputFolder.Text
            
            if not folder or not os.path.exists(folder):
                return

            for item in self.sheet_items:
                item.update_status(folder)
            
            self.dg_Sheets.Items.Refresh()
        except:
            pass

    # --- LOGICA DE SETS DO REVIT (FASE 3) ---
    # --- LOGICA DE SETS DO REVIT (FASE 3 & 4) ---
    def load_revit_sets(self):
        try:
            self.cb_RevitSets.Items.Clear()
            
            # 1. Opcao Padrao "Todas as Folhas"
            self.cb_RevitSets.Items.Add("(Todas as Folhas)")
            
            # 2. Sets do Revit
            all_sets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet).ToElements()
            sorted_sets = sorted(all_sets, key=lambda x: x.Name)
            
            self.revit_sets_map = {}
            for vset in sorted_sets:
                self.cb_RevitSets.Items.Add(vset.Name)
                self.revit_sets_map[vset.Name] = vset
            
            # Auto-selecionar o ultimo usado ou o primeiro
            last_set = self.config.get("last_set", "(Todas as Folhas)")
            
            if last_set in self.cb_RevitSets.Items:
                self.cb_RevitSets.SelectedItem = last_set
            elif self.cb_RevitSets.Items.Count > 0:
                self.cb_RevitSets.SelectedIndex = 0
                
        except Exception as ex:
            self.log_message("Erro load_sets: " + str(ex))

    def apply_current_set(self, sender, args):
        """Aplica o set assim que mudado no combo"""
        sel_name = self.cb_RevitSets.SelectedItem
        if not sel_name: return
        
        # Persistir a escolha
        try:
            if self.config.get("last_set") != sel_name:
                self.config["last_set"] = sel_name
                save_config(self.config)
        except: pass

        try:
            # Caso Especial: Todas as Folhas
            if sel_name == "(Todas as Folhas)":
                for item in self.sheet_items:
                    item.IsSelected = True # Seleciona Tudo (padrão clean)
                self.dg_Sheets.Items.Refresh()
                return

            # Caso Sets do Revit
            if sel_name in self.revit_sets_map:
                vset = self.revit_sets_map[sel_name]
                view_ids = [v.Id for v in vset.Views]
                
                count = 0
                for item in self.sheet_items:
                    if item.Element.Id in view_ids:
                        item.IsSelected = True
                        count += 1
                    else:
                        item.IsSelected = False # Exclusivo
                
                self.dg_Sheets.Items.Refresh()
                
        except Exception as ex:
            forms.alert("Erro ao aplicar set: " + str(ex))

    def save_new_set(self, sender, args):
        selected_items = [i for i in self.sheet_items if i.IsSelected]
        if not selected_items:
            forms.alert("Selecione folhas para salvar.")
            return

        name = forms.ask_for_string(prompt="Nome do Set:", title="Novo Set")
        if not name: return
        
        # Validacao nome
        if name == "(Todas as Folhas)":
             forms.alert("Nome reservado.")
             return

        try:
            with DB.Transaction(doc, "Criar Set") as t:
                t.Start()
                view_set = DB.ViewSet()
                for item in selected_items: view_set.Insert(item.Element)
                
                print_mgr = doc.PrintManager
                print_mgr.PrintRange = DB.PrintRange.Select
                settings = print_mgr.ViewSheetSetting
                settings.CurrentViewSheetSet.Views = view_set
                settings.SaveAs(name)
                t.Commit()
            
            self.load_revit_sets()
            self.cb_RevitSets.SelectedItem = name # Ja aplica? Sim, o SelectionChanged dispara
            forms.alert("Set salvo!")
        except Exception as ex:
            forms.alert("Erro ao salvar: " + str(ex))

    def delete_current_set(self, sender, args):
        sel_name = self.cb_RevitSets.SelectedItem
        if not sel_name: return
        
        if sel_name == "(Todas as Folhas)":
            forms.alert("Não é possível deletar o conjunto padrão.")
            return
            
        if not forms.alert("Deletar o set '{}'? (Isso remove do Revit)".format(sel_name), yes=True, no=True):
            return

        try:
            if sel_name in self.revit_sets_map:
                vset = self.revit_sets_map[sel_name]
                with DB.Transaction(doc, "Deletar Set") as t:
                    t.Start()
                    doc.Delete(vset.Id)
                    t.Commit()
                
                self.load_revit_sets() # Recarrega
                self.cb_RevitSets.SelectedIndex = 0 # Volta para Default
                forms.alert("Set deletado.")
        except Exception as ex:
            forms.alert("Erro ao deletar: " + str(ex))

    # --- LOGICA DE CLIQUE NA LINHA (RESTAURADA) ---
    def on_row_clicked(self, sender, args):
        if self.is_handling_click: return
        self.is_handling_click = True
        try:
            if args.AddedItems and args.AddedItems.Count > 0:
                for item in args.AddedItems:
                    item.IsSelected = not item.IsSelected
                self.dg_Sheets.UnselectAll()
                self.dg_Sheets.Items.Refresh()
        except: pass
        self.is_handling_click = False

    # --- FUNCOES DE UI E PROGRESSO ---

    def log_message(self, message):
        """Escreve no log detalhado"""
        try:
            timestamp = time.strftime("%H:%M:%S")
            if "===" in message:
                full_msg = "{}\n".format(message)
            else:
                full_msg = "[{}] {}\n".format(timestamp, message)
            
            self.txt_Log.AppendText(full_msg)
            self.txt_Log.ScrollToEnd()
            WinFormsApp.DoEvents()
        except:
            pass
    
    def reset_progress_ui(self):
        """Reseta a interface de progresso para estado inicial"""
        try:
            self.lbl_ProgressTitle.Text = "Aguardando início..."
            self.lbl_ProgressDetail.Text = ""
            self.lbl_ProgressPercent.Text = "0%"
            self.lbl_SuccessCount.Text = "0"
            self.lbl_ErrorCount.Text = "0"
            self.lbl_TimeElapsed.Text = "0s"
            self.progressBar.Width = 0
            self.pnl_ExportItems.Children.Clear()
            self.txt_Log.Text = ""  # Limpa log detalhado
            WinFormsApp.DoEvents()
        except:
            pass
    
    def update_progress(self, current, total, current_item_name=""):
        """Atualiza a barra de progresso e informacoes"""
        try:
            percent = int((current / float(total)) * 100)
            self.lbl_ProgressTitle.Text = "Exportando {} de {}".format(current, total)
            self.lbl_ProgressDetail.Text = current_item_name
            self.lbl_ProgressPercent.Text = "{}%".format(percent)
            
            # Atualiza largura da barra (assume container ~500px)
            parent_width = self.progressBar.Parent.ActualWidth if self.progressBar.Parent else 500
            self.progressBar.Width = (percent / 100.0) * parent_width
            
            WinFormsApp.DoEvents()
        except:
            pass
    
    def update_counters(self, success, errors, elapsed_seconds):
        """Atualiza contadores de sucesso/erro/tempo"""
        try:
            self.lbl_SuccessCount.Text = str(success)
            self.lbl_ErrorCount.Text = str(errors)
            self.lbl_TimeElapsed.Text = self.format_time(elapsed_seconds)
            WinFormsApp.DoEvents()
        except:
            pass
    
    def add_export_item(self, sheet_number, file_name, status, file_type=""):
        """Adiciona um item na lista de exportados com status visual"""
        try:
            import System.Windows.Controls as Controls
            import System.Windows.Media as Media
            
            # Container do item
            border = Controls.Border()
            border.CornerRadius = System.Windows.CornerRadius(4)
            border.Padding = System.Windows.Thickness(10, 8, 10, 8)
            border.Margin = System.Windows.Thickness(0, 0, 0, 4)
            
            if status == "success":
                border.Background = Media.SolidColorBrush(Media.Color.FromArgb(26, 78, 201, 155))  # Verde translucido
                icon = "✓"
                icon_color = Media.Color.FromArgb(255, 78, 201, 155)
            elif status == "error":
                border.Background = Media.SolidColorBrush(Media.Color.FromArgb(26, 210, 85, 85))  # Vermelho translucido
                icon = "✕"
                icon_color = Media.Color.FromArgb(255, 210, 85, 85)
            else:  # processing
                border.Background = Media.SolidColorBrush(Media.Color.FromArgb(26, 86, 156, 214))  # Azul translucido
                icon = "◐"
                icon_color = Media.Color.FromArgb(255, 86, 156, 214)
            
            # Grid interno
            grid = Controls.Grid()
            col1 = Controls.ColumnDefinition()
            col1.Width = System.Windows.GridLength(25)
            col2 = Controls.ColumnDefinition()
            col2.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            col3 = Controls.ColumnDefinition()
            col3.Width = System.Windows.GridLength.Auto
            grid.ColumnDefinitions.Add(col1)
            grid.ColumnDefinitions.Add(col2)
            grid.ColumnDefinitions.Add(col3)
            
            # Icone
            icon_txt = Controls.TextBlock()
            icon_txt.Text = icon
            icon_txt.FontSize = 14
            icon_txt.FontWeight = System.Windows.FontWeights.Bold
            icon_txt.Foreground = Media.SolidColorBrush(icon_color)
            icon_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
            Controls.Grid.SetColumn(icon_txt, 0)
            grid.Children.Add(icon_txt)
            
            # Texto (numero + nome)
            stack = Controls.StackPanel()
            
            num_txt = Controls.TextBlock()
            num_txt.Text = sheet_number
            num_txt.FontWeight = System.Windows.FontWeights.SemiBold
            num_txt.Foreground = Media.SolidColorBrush(Media.Color.FromArgb(255, 241, 241, 241))
            num_txt.FontSize = 12
            stack.Children.Add(num_txt)
            
            name_txt = Controls.TextBlock()
            name_txt.Text = file_name
            name_txt.Foreground = Media.SolidColorBrush(Media.Color.FromArgb(255, 204, 204, 204))
            name_txt.FontSize = 10
            stack.Children.Add(name_txt)
            
            Controls.Grid.SetColumn(stack, 1)
            grid.Children.Add(stack)
            
            # Tipo de arquivo
            if file_type:
                type_txt = Controls.TextBlock()
                type_txt.Text = file_type
                type_txt.FontSize = 10
                type_txt.Foreground = Media.SolidColorBrush(Media.Color.FromArgb(255, 150, 150, 150))
                type_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
                Controls.Grid.SetColumn(type_txt, 2)
                grid.Children.Add(type_txt)
            
            border.Child = grid
            self.pnl_ExportItems.Children.Add(border)
            
            # Scroll para o final
            parent = self.pnl_ExportItems.Parent
            if hasattr(parent, 'ScrollToEnd'):
                parent.ScrollToEnd()
            
            WinFormsApp.DoEvents()
        except Exception as ex:
            print("Erro add_export_item: " + str(ex))

    def load_sheet_parameters(self):
        try:
            self.cb_AvailableParams.Items.Clear()
            defaults = ["Sheet Number", "Sheet Name", "Current Revision", "Approved By", "Designed By", "Nome da folha"]
            for d in defaults: self.cb_AvailableParams.Items.Add(d)
            
            sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType().ToElements()
            if sheets:
                sample = sheets[0]
                for p in sample.Parameters:
                    if p.Definition.Name not in defaults and p.StorageType == DB.StorageType.String:
                        self.cb_AvailableParams.Items.Add(p.Definition.Name)
            if self.cb_AvailableParams.Items.Count > 0: self.cb_AvailableParams.SelectedIndex = 0
        except Exception as ex:
            self.log_message("Erro ao carregar parametros: " + str(ex))

    def add_param_to_pattern(self, sender, args):
        if self.cb_AvailableParams.SelectedItem:
            param = self.cb_AvailableParams.SelectedItem
            current_text = self.txt_NamePattern.Text
            to_insert = "{{{}}}".format(param) 
            if current_text and not current_text.strip().endswith("-") and not current_text.endswith(" "):
                to_insert = " - " + to_insert
            self.txt_NamePattern.Text += to_insert

    def on_pattern_changed(self, sender, args):
        new_pattern = self.txt_NamePattern.Text
        if not new_pattern: return
        for item in self.sheet_items: item.update_filename(new_pattern)
        # Re-checar existencia pois nomes mudaram
        self.check_existing_files() 
        self.dg_Sheets.Items.Refresh()

    def load_sheets(self):
        try:
            sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType().ToElements()
            sorted_sheets = sorted(sheets, key=lambda x: x.SheetNumber)
            
            current_pattern = self.txt_NamePattern.Text
            if not current_pattern:
                current_pattern = "{Nome da folha}"
                self.txt_NamePattern.Text = current_pattern
            
            self.sheet_items.Clear()
            for s in sorted_sheets:
                if not s.IsPlaceholder: 
                    item = SheetItem(s, current_pattern)
                    # Checar se tem views (para status icon)
                    try:
                         if not s.GetAllPlacedViews():
                             item.HasViews = False
                    except: pass
                    self.sheet_items.Add(item)
            self.lbl_Count.Text = str(len(self.sheet_items))
        except Exception as e:
            forms.alert("Erro ao carregar folhas: " + str(e))

    def load_dwg_setups(self):
        try:
            setups = DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings).ToElements()
            self.cb_DWGSetups.Items.Clear()
            for s in setups: self.cb_DWGSetups.Items.Add(s.Name)
            if self.cb_DWGSetups.Items.Count > 0: self.cb_DWGSetups.SelectedIndex = 0
        except Exception as ex:
            print("Erro setups: " + str(ex))

    def pick_folder(self, sender, args):
        folder = forms.pick_folder()
        if folder: 
            self.txt_OutputFolder.Text = folder
            # Salvar config
            self.config["last_folder"] = folder
            save_config(self.config)
            
            # Re-checar existencia na nova pasta
            self.check_existing_files(folder)
    
    def cancel_export(self, sender, args):
        # Se NAO esta exportando, fecha o programa
        if self.btn_Start.IsEnabled:
            self.Close()
            return
        
        # Se ESTA exportando, cancela a exportacao
        if forms.alert("Deseja cancelar a exportacao?", yes=True, no=True):
            self.is_cancelled = True
            self.log_message("!!! CANCELAMENTO SOLICITADO !!!")

    def create_pdf_options(self):
        pdf_opts = DB.PDFExportOptions()
        pdf_opts.Combine = False
        try:
            if self.rb_ZoomFit.IsChecked:
                pdf_opts.ZoomType = DB.ZoomType.FitToPage
                pdf_opts.PaperPlacement = DB.PaperPlacementType.Center
            else:
                pdf_opts.ZoomType = DB.ZoomType.Zoom
                pdf_opts.ZoomPercentage = 100
                pdf_opts.PaperPlacement = DB.PaperPlacementType.Margins
        except: pass
        
        # Configurar ocultação de Crop Region (Região de Recorte)
        try:
            hide_crop = bool(self.chk_HideCrop.IsChecked)
            
            if hasattr(pdf_opts, 'HideCropBoundaries'):
                pdf_opts.HideCropBoundaries = hide_crop
                self.log_message("  > HideCropBoundaries = {}".format(hide_crop))
            else:
                self.log_message("  > AVISO: HideCropBoundaries nao disponivel nesta versao")
        except Exception as ex:
            self.log_message("  > ERRO ao configurar HideCropBoundaries: " + str(ex))
        
        return pdf_opts

    def create_dwg_options_compatible(self):
        dwg_opts = DB.DWGExportOptions()
        dwg_opts.MergedViews = True 
        try:
            if hasattr(dwg_opts, 'ExportScope'): dwg_opts.ExportScope = DB.DWGExportScope.View
        except: pass
        return dwg_opts

    def safe_rename_file(self, old_path, new_path, max_attempts=5):
        for attempt in range(max_attempts):
            try:
                if os.path.exists(new_path): os.remove(new_path)
                os.rename(old_path, new_path)
                return True
            except:
                if attempt < max_attempts - 1: time.sleep(0.5)
        return False
    
    def format_time(self, seconds):
        if seconds < 60: return "{:.0f}s".format(seconds)
        elif seconds < 3600: return "{:.1f}min".format(seconds/60)
        else: return "{:.1f}h".format(seconds/3600)

    def run_export(self, sender, args):
        folder = self.txt_OutputFolder.Text
        if not folder or not os.path.exists(folder):
            folder = forms.pick_folder()
            if folder:
                self.txt_OutputFolder.Text = folder
            else:
                return

        selected_items = [i for i in self.sheet_items if i.IsSelected]
        if not selected_items:
            forms.alert("Selecione pelo menos uma folha.")
            return

        do_pdf = self.chk_ExportPDF.IsChecked
        do_dwg = self.chk_ExportDWG.IsChecked
        
        if not do_pdf and not do_dwg:
            forms.alert("Selecione PDF ou DWG.")
            return

        # VERIFICACAO DE DUPLICADOS
        name_counts = {}
        for item in selected_items:
            name = item.FileName
            if name in name_counts: name_counts[name] += 1
            else: name_counts[name] = 1
        
        duplicates = [n for n, c in name_counts.items() if c > 1]
        if duplicates:
            msg = "Existem {} nomes de arquivos duplicados:\n\n".format(len(duplicates))
            msg += "\n".join(duplicates[:10])
            if len(duplicates) > 10: msg += "\n...e mais {}.".format(len(duplicates)-10)
            msg += "\n\nDeseja continuar mesmo assim? (Arquivos podem ser sobrescritos)"
            if not forms.alert(msg, yes=True, no=True):
                return
        
        # UI Setup
        self.btn_Start.IsEnabled = False
        self.btn_Cancel.Visibility = System.Windows.Visibility.Visible
        self.tab_Main.SelectedIndex = 1 
        self.reset_progress_ui()
        
        self.is_cancelled = False
        
        success_count = 0
        error_count = 0
        start_time = time.time()
        total_items = len(selected_items)
        
        self.log_message("INICIANDO EXPORTAÇÃO...")
        self.log_message("Total de folhas: {}".format(total_items))
        self.log_message("Pasta: {}".format(folder))
        
        # Merge PDF Info
        do_merge_pdf = do_pdf and self.chk_CombinePDF.IsChecked
        pdf_merge_list = []
        if do_merge_pdf:
            self.log_message("Modo PDF Combinado (Merge) ATIVADO")
            
        self.log_message("="*50)

        # Estimativa de tempo
        last_times = []
        
        
        # Pre-configurar DWG options fora do loop
        dwg_opts = None
        if do_dwg:
            if self.cb_DWGSetups.SelectedItem:
                s_name = self.cb_DWGSetups.SelectedItem
                for s in DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings):
                    if s.Name == s_name:
                        dwg_opts = s.GetDWGExportOptions()
                        break
            if not dwg_opts: 
                dwg_opts = self.create_dwg_options_compatible()
            dwg_opts.MergedViews = True

        for idx, item in enumerate(selected_items):
            loop_start = time.time()
            if self.is_cancelled: 
                break
            
            current_num = idx + 1
            
            # Calculo de estimativa
            remaining_txt = ""
            if len(last_times) > 0:
                avg_time = sum(last_times) / len(last_times)
                remaining_items = total_items - idx
                est_seconds = remaining_items * avg_time
                remaining_txt = " | Restante: ~{}".format(self.format_time(est_seconds))

            self.update_progress(current_num, total_items, "Exportando: {} - {}{}".format(item.Number, item.Name, remaining_txt))
            self.update_counters(success_count, error_count, time.time() - start_time)
            
            try:
                self.log_message("PROCESSANDO: {} - {}".format(item.Number, item.Name))
                view_ids = List[DB.ElementId]([item.Element.Id])

                # --- DWG ---
                if do_dwg:
                    try:
                        doc.Export(folder, item.FileName, view_ids, dwg_opts)
                        self.log_message("  [DWG] OK: {}".format(item.FileName))

                        # Verificar Renomear DWG (Revit as vezes poe prefixo)
                        dwg_path = os.path.join(folder, "{}.dwg".format(item.FileName))
                        for _ in range(15):
                            if self.is_cancelled or os.path.exists(dwg_path): break
                            time.sleep(0.3)
                            WinFormsApp.DoEvents()
                        
                        if not os.path.exists(dwg_path):
                            # Tenta achar arquivo similar com sufixo -Sheet ou _Sheet
                            alt = os.path.join(folder, "{}_Sheet.dwg".format(item.FileName))
                            if os.path.exists(alt):
                                self.safe_rename_file(alt, dwg_path)
                                self.log_message("  > DWG Renomeado de _Sheet")
                        
                        self.add_export_item(item.Number, item.FileName, "success", "DWG")
                    except Exception as e:
                        error_count += 1
                        self.add_export_item(item.Number, "Erro DWG", "error", "")
                        self.log_message("  > ERRO DWG: " + str(e))

                # --- PDF ---
                if do_pdf:
                    if do_merge_pdf:
                        pdf_merge_list.append(item.Element.Id)
                        self.log_message("  [PDF] Adicionado ao Merge")
                        if not do_dwg:
                             self.add_export_item(item.Number, "Na Fila Merge", "processing", "PDF")
                    else:
                        try:
                            pdf_opts = self.create_pdf_options()
                            pdf_opts.FileName = item.PdfFileName

                            # --- ESTRATEGIA UNIVERSAL SNAPSHOT ---
                            # 1. Tenta limpar o arquivo esperado final (se existir)
                            expected_path = os.path.join(folder, "{}.pdf".format(item.PdfFileName))
                            try:
                                if os.path.exists(expected_path): os.remove(expected_path)
                            except: pass

                            # 2. Tira foto da pasta ANTES
                            before_files = set(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
                            
                            # 3. Exporta
                            doc.Export(folder, view_ids, pdf_opts)
                            
                            # 4. Procura QUALQUER arquivo novo
                            final_path = None
                            found_new_file = None
                            
                            for _ in range(30): # 10s
                                if self.is_cancelled: break
                                
                                try:
                                    current_files = set(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
                                    new_files = current_files - before_files
                                    
                                    if new_files:
                                        # Achou arquivo(s) novo(s)!
                                        found_new_file = list(new_files)[0] # Pega o primeiro
                                        
                                        # Log do que foi achado
                                        self.log_message("  > Arquivo gerado: '{}'".format(found_new_file))
                                        
                                        generated_path = os.path.join(folder, found_new_file)
                                        
                                        # Renomeia para o esperado
                                        if self.safe_rename_file(generated_path, expected_path):
                                            final_path = expected_path
                                            self.log_message("  > Renomeado com sucesso!")
                                        else:
                                            # Se falhar renomear (ex: mesmo nome), assume que eh ele
                                            if found_new_file.lower() == "{}.pdf".format(item.PdfFileName).lower():
                                                 final_path = generated_path
                                            else:
                                                 self.log_message("  > Falha ao renomear.")
                                        break
                                except: pass
                                
                                time.sleep(0.33)
                                WinFormsApp.DoEvents()
                            
                            if final_path:
                                self.log_message("  [PDF] OK: {}".format(item.PdfFileName))
                                self.add_export_item(item.Number, item.FileName, "success", "PDF")
                            else:
                                self.log_message("  [PDF] ERRO: Nenhum arquivo novo detectado.")
                                self.add_export_item(item.Number, "PDF?", "error", "PDF")

                        except Exception as e:
                            error_count += 1
                            self.log_message("  > ERRO PDF: " + str(e))
                            self.add_export_item(item.Number, "Erro PDF", "error", "")
            
            except Exception as ex:
                error_count += 1
                self.add_export_item(item.Number, "Erro: {}".format(str(ex)[:30]), "error", "")
                self.log_message("  > ERRO GERAL LOOP: {}".format(str(ex)))
            
            # Registra tempo do loop
            loop_time = time.time() - loop_start
            last_times.append(loop_time)
            if len(last_times) > 5: last_times.pop(0)

        # --- POS-LOOP MERGE PDF ---
        if do_merge_pdf and pdf_merge_list and not self.is_cancelled:
            self.update_progress(total_items, total_items, "Gerando PDF Único (Merge)... Aguarde.")
            try:
                # Nome do arquivo = Nome do RVT + Data
                base_name = doc.Title
                # Remove extensao .rvt se existir (doc.Title as vezes traz)
                if base_name.lower().endswith(".rvt"):
                    base_name = base_name[:-4]
                
                cleaned_name = re.sub(r'[\\/*?:"<>|]', "", base_name)
                merge_name = "{}_{}".format(cleaned_name, time.strftime("%Y%m%d"))
                
                pdf_opts = self.create_pdf_options() 
                pdf_opts.FileName = merge_name
                pdf_opts.Combine = True
                
                merge_ids = List[DB.ElementId](pdf_merge_list)
                doc.Export(folder, merge_ids, pdf_opts)
                
                self.log_message("="*30)
                self.log_message("[PDF MERGE] ARQUIVO GERADO: {}.pdf".format(merge_name))
                self.add_export_item("MERGE", "{}.pdf".format(merge_name), "success", "PDF COMBINADO")
                success_count += 1 
            except Exception as ex:
                error_count += 1
                self.log_message("[PDF MERGE] ERRO: " + str(ex))
                self.add_export_item("MERGE", "Falha ao combinar", "error", "")

        # --- FINALIZACAO ---
        total_time = time.time() - start_time
        self.btn_Start.IsEnabled = True
        self.btn_Cancel.Visibility = System.Windows.Visibility.Collapsed
        self.update_counters(success_count, error_count, total_time)
        
        self.config["open_folder_after"] = self.chk_OpenFolderAfter.IsChecked
        save_config(self.config)

        self.log_message("="*50)
        
        if self.is_cancelled:
            self.lbl_ProgressTitle.Text = "Cancelado"
        elif error_count == 0:
            self.lbl_ProgressTitle.Text = "✓ Concluído com sucesso!"
        else:
            self.lbl_ProgressTitle.Text = "Concluído com avisos"

        try:
             if self.chk_OpenFolderAfter.IsChecked and not self.is_cancelled:
                 os.startfile(folder)
        except: pass

try:
    LuisExporterWindow().ShowDialog()
except Exception as e:
    forms.alert("Erro fatal: " + str(e))