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
from System.Windows.Forms import Application as WinFormsApp
import os
import re
import time

doc = revit.doc
uidoc = revit.uidoc

# --- DETECAO DE VERSAO ---
REVIT_YEAR = int(doc.Application.VersionNumber)
HAS_PDF_SUPPORT = REVIT_YEAR >= 2022

# --- CLASSE DE DADOS PARA FOLHAS ---
class SheetItem(object):
    def __init__(self, sheet, name_pattern="{Nome da folha}"):
        self.Element = sheet
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name if sheet.Name else "" 
        self.IsSelected = False
        self.name_pattern = name_pattern
        self._file_name = "" # Backing field
        self.PdfFileName = "" 
        
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

# --- JANELA PRINCIPAL ---
class LuisExporterWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = script.get_bundle_file('folhas_window.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.sheet_items = ObservableCollection[SheetItem]()
        self.dg_Sheets.ItemsSource = self.sheet_items
        
        # Eventos Principais
        self.btn_SelectFolder.Click += self.pick_folder
        self.btn_Start.Click += self.run_export
        self.btn_Cancel.Click += self.cancel_export
        self.btn_AddParam.Click += self.add_param_to_pattern
        self.txt_NamePattern.TextChanged += self.on_pattern_changed
        
        # NOVO: Evento Selecionar Tudo
        self.chk_SelectAll.Click += self.toggle_all_sheets

        # Evento de clique na linha (RESTAURADO)
        self.dg_Sheets.SelectionChanged += self.on_row_clicked
        self.is_handling_click = False 
        
        # Configurar Grid de Edicao (mesma fonte)
        if hasattr(self, 'dg_Edits'):
            self.dg_Edits.ItemsSource = self.sheet_items
        
        self.is_cancelled = False
        self.rb_ZoomFit.IsChecked = True
        self.chk_ExportDWG.IsChecked = True
        
        self.load_sheet_parameters() 
        self.load_sheets()
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
        """Marca ou desmarca todas as folhas"""
        is_checked = self.chk_SelectAll.IsChecked
        for item in self.sheet_items:
            item.IsSelected = is_checked
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
    def save_changes_revit(self, sender, args):
        """Aplica os nomes editados as folhas no Revit"""
        updated_count = 0
        errors = []
        
        try:
            # Tenta atualizar
            with revit.Transaction("Atualizar Nomes de Folhas"):
                for item in self.sheet_items:
                    # Verifica se mudou (e se nao eh vazio)
                    if item.FileName and item.FileName != item.Name:
                        try:
                            item.Element.Name = item.FileName
                            # Atualiza item.Name para refletir a mudanca
                            item.Name = item.FileName
                            updated_count += 1
                        except Exception as e:
                            errors.append("{} ({}): {}".format(item.Number, item.FileName, str(e)))
            
            if errors:
                forms.alert("Concluído com {} erros:\n{}".format(len(errors), "\n".join(errors[:5])))
            else:
                forms.alert("Sucesso! {} folhas atualizadas no Revit.".format(updated_count))
            
            # Atualiza Grids
            self.dg_Edits.Items.Refresh()
            self.dg_Sheets.Items.Refresh()
                
        except Exception as e:
            forms.alert("Erro ao aplicar mudanças: " + str(e))

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
                    self.sheet_items.Add(SheetItem(s, current_pattern))
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
        if folder: self.txt_OutputFolder.Text = folder
    
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
        
        # UI Setup
        self.btn_Start.IsEnabled = False
        self.btn_Cancel.Visibility = System.Windows.Visibility.Visible
        self.tab_Main.SelectedIndex = 2 # Aba Progresso agora eh indice 2
        self.reset_progress_ui()
        
        self.is_cancelled = False
        success_count = 0
        error_count = 0
        start_time = time.time()
        total_items = len(selected_items)
        
        self.log_message("INICIANDO EXPORTAÇÃO...")
        self.log_message("Total de folhas: {}".format(total_items))
        self.log_message("Pasta: {}".format(folder))
        self.log_message("="*50)
        
        for idx, item in enumerate(selected_items):
            if self.is_cancelled: 
                break
            
            current_num = idx + 1
            self.update_progress(current_num, total_items, "Folha {} - {}".format(item.Number, item.Name))
            self.update_counters(success_count, error_count, time.time() - start_time)
            
            try:
                self.log_message("PROCESSANDO: {} - {}".format(item.Number, item.Name))
                view_ids = System.Collections.Generic.List[DB.ElementId]([item.Id])
                
                # --- PDF ---
                if do_pdf and HAS_PDF_SUPPORT and not self.is_cancelled:
                    try:
                        pdf_opts = self.create_pdf_options()
                        pdf_files_before = set([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                        
                        doc.Export(folder, view_ids, pdf_opts)
                        
                        max_wait = 15
                        elapsed = 0
                        pdf_ok = False
                        
                        while elapsed < max_wait and not self.is_cancelled:
                            WinFormsApp.DoEvents()
                            time.sleep(0.5)
                            elapsed += 0.5
                            self.update_counters(success_count, error_count, time.time() - start_time)
                            
                            pdf_files_after = set([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                            new_files = pdf_files_after - pdf_files_before
                            
                            if new_files:
                                latest = list(new_files)[0]
                                old_p = os.path.join(folder, latest)
                                new_p = os.path.join(folder, "{}.pdf".format(item.PdfFileName))
                                if self.safe_rename_file(old_p, new_p):
                                    success_count += 1
                                    pdf_ok = True
                                    self.add_export_item(item.Number, "{}.pdf".format(item.PdfFileName), "success", "PDF")
                                    self.log_message("  > PDF OK: {}.pdf".format(item.PdfFileName))
                                else:
                                    error_count += 1
                                    self.add_export_item(item.Number, "Erro ao renomear PDF", "error", "PDF")
                                    self.log_message("  > ERRO PDF: Falha ao renomear arquivo")
                                break
                        
                        if not pdf_ok and not self.is_cancelled:
                            if os.path.exists(os.path.join(folder, "{}.pdf".format(item.PdfFileName))):
                                success_count += 1
                                self.add_export_item(item.Number, "{}.pdf".format(item.PdfFileName), "success", "PDF")
                                self.log_message("  > PDF OK: {}.pdf".format(item.PdfFileName))
                            else:
                                error_count += 1
                                self.add_export_item(item.Number, "PDF não encontrado", "error", "PDF")
                                self.log_message("  > ERRO PDF: Arquivo nao encontrado apos exportacao")

                    except Exception as ex:
                        error_count += 1
                        self.add_export_item(item.Number, str(ex)[:40], "error", "PDF")
                        self.log_message("  > ERRO CRITICO PDF: {}".format(str(ex)))
                
                # --- DWG ---
                if do_dwg and not self.is_cancelled:
                    try:
                        dwg_opts = None
                        if self.cb_DWGSetups.SelectedItem:
                            s_name = self.cb_DWGSetups.SelectedItem
                            for s in DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings):
                                if s.Name == s_name:
                                    dwg_opts = s.GetDWGExportOptions()
                                    break
                        if not dwg_opts: 
                            dwg_opts = self.create_dwg_options_compatible()
                        dwg_opts.MergedViews = True
                        
                        doc.Export(folder, item.FileName, view_ids, dwg_opts)
                        
                        dwg_path = os.path.join(folder, "{}.dwg".format(item.FileName))
                        found_dwg = False
                        
                        for _ in range(15):
                            if self.is_cancelled:
                                break
                            if os.path.exists(dwg_path):
                                found_dwg = True
                                break
                            time.sleep(0.3)
                            WinFormsApp.DoEvents()
                            self.update_counters(success_count, error_count, time.time() - start_time)

                        if found_dwg:
                            success_count += 1
                            self.add_export_item(item.Number, "{}.dwg".format(item.FileName), "success", "DWG")
                            self.log_message("  > DWG OK: {}.dwg".format(item.FileName))
                        else:
                            alt = os.path.join(folder, "{}_Sheet.dwg".format(item.FileName))
                            if os.path.exists(alt):
                                self.safe_rename_file(alt, dwg_path)
                                success_count += 1
                                self.add_export_item(item.Number, "{}.dwg".format(item.FileName), "success", "DWG")
                                self.log_message("  > DWG OK: {}.dwg (renomeado)".format(item.FileName))
                            else:
                                error_count += 1
                                self.add_export_item(item.Number, "DWG não gerado", "error", "DWG")
                                self.log_message("  > ERRO DWG: Arquivo nao foi gerado")

                    except Exception as ex:
                        error_count += 1
                        self.add_export_item(item.Number, str(ex)[:40], "error", "DWG")
                        self.log_message("  > ERRO CRITICO DWG: {}".format(str(ex)))
                
            except Exception as ex:
                error_count += 1
                self.add_export_item(item.Number, "Erro: {}".format(str(ex)[:30]), "error", "")
                self.log_message("  > ERRO GERAL: {}".format(str(ex)))
        
        # --- FINALIZACAO ---
        total_time = time.time() - start_time
        self.btn_Start.IsEnabled = True
        self.btn_Cancel.Visibility = System.Windows.Visibility.Collapsed
        
        # Atualiza UI final
        self.update_counters(success_count, error_count, total_time)
        
        # Log de finalizacao
        self.log_message("="*50)
        if self.is_cancelled:
            self.lbl_ProgressTitle.Text = "Cancelado pelo usuário"
            self.lbl_ProgressDetail.Text = ""
            self.log_message("CANCELADO PELO USUARIO")
        elif error_count == 0:
            self.lbl_ProgressTitle.Text = "✓ Concluído com sucesso!"
            self.lbl_ProgressDetail.Text = "{} arquivos exportados em {}".format(success_count, self.format_time(total_time))
            self.update_progress(total_items, total_items, "")
            self.log_message("CONCLUIDO COM SUCESSO!")
            self.log_message("Arquivos: {} | Tempo: {}".format(success_count, self.format_time(total_time)))
        else:
            self.lbl_ProgressTitle.Text = "Concluído com avisos"
            self.lbl_ProgressDetail.Text = "{} sucessos, {} erros".format(success_count, error_count)
            self.update_progress(total_items, total_items, "")
            self.log_message("CONCLUIDO COM AVISOS")
            self.log_message("Sucessos: {} | Erros: {} | Tempo: {}".format(success_count, error_count, self.format_time(total_time)))
        
        self.log_message("="*50)
        
        self.log_message("="*50)
        
        try:
             # Pergunta se quer abrir a pasta com TaskDialog padrao do pyrevit
             if forms.alert("Processo finalizado. Deseja abrir a pasta de destino?", yes=True, no=True):
                 os.startfile(folder)
        except:
             pass

try:
    LuisExporterWindow().ShowDialog()
except Exception as e:
    forms.alert("Erro fatal: " + str(e))