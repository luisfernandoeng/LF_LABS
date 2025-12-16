# -*- coding: utf-8 -*-
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
    def __init__(self, sheet, name_pattern="{C-NOME-FOLHA}"):
        self.Element = sheet
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name if sheet.Name else "" 
        self.IsSelected = False
        self.name_pattern = name_pattern
        self.FileName = self._generate_filename(sheet, name_pattern)
        self.PdfFileName = self.FileName 

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

        # Evento de clique na linha
        self.dg_Sheets.SelectionChanged += self.on_row_clicked
        self.is_handling_click = False 
        
        self.is_cancelled = False
        self.rb_ZoomFit.IsChecked = True
        self.chk_ExportDWG.IsChecked = True
        
        self.load_sheet_parameters() 
        self.load_sheets()           
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
        self.dg_Sheets.Items.Refresh()

    # --- LOGICA DE CLIQUE NA LINHA ---
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

    # --- FUNCOES DE UI E LOG ---
    def log_message(self, message):
        """Escreve na aba de Log e atualiza a tela"""
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

    def load_sheet_parameters(self):
        try:
            self.cb_AvailableParams.Items.Clear()
            defaults = ["Sheet Number", "Sheet Name", "Current Revision", "Approved By", "Designed By", "C-NOME-FOLHA"]
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
                current_pattern = "{C-NOME-FOLHA}"
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
        if forms.alert("Deseja cancelar?", yes=True, no=True):
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
            forms.alert("Selecione uma pasta valida.")
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
        self.txt_Log.Text = "" 
        
        self.tab_Main.SelectedIndex = 1 
        self.log_message("INICIANDO EXPORTAÇÃO...")
        self.log_message("Total de folhas: {}".format(len(selected_items)))
        self.log_message("="*40)
        
        self.is_cancelled = False
        success_count = 0
        error_count = 0
        start_time = time.time()
        
        for item in selected_items:
            if self.is_cancelled: break
            
            try:
                self.log_message("PROCESSANDO: {}".format(item.Number))
                view_ids = System.Collections.Generic.List[DB.ElementId]([item.Id])
                
                # --- PDF ---
                if do_pdf and HAS_PDF_SUPPORT and not self.is_cancelled:
                    try:
                        self.log_message("  > Gerando PDF...")
                        pdf_opts = self.create_pdf_options()
                        pdf_files_before = set([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                        
                        doc.Export(folder, view_ids, pdf_opts)
                        
                        max_wait = 15
                        elapsed = 0
                        found = False
                        while elapsed < max_wait and not self.is_cancelled:
                            WinFormsApp.DoEvents() 
                            time.sleep(0.5)
                            elapsed += 0.5
                            
                            pdf_files_after = set([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                            new_files = pdf_files_after - pdf_files_before
                            
                            if new_files:
                                latest = list(new_files)[0]
                                old_p = os.path.join(folder, latest)
                                new_p = os.path.join(folder, "{}.pdf".format(item.PdfFileName))
                                if self.safe_rename_file(old_p, new_p):
                                    success_count += 1
                                    found = True
                                    self.log_message("  > PDF OK: " + item.PdfFileName)
                                else:
                                    error_count += 1
                                    self.log_message("  > ERRO ao renomear PDF.")
                                break
                        
                        if not found and not self.is_cancelled:
                            if os.path.exists(os.path.join(folder, "{}.pdf".format(item.PdfFileName))):
                                success_count += 1
                                self.log_message("  > PDF OK (Direto).")
                            else:
                                error_count += 1
                                self.log_message("  > ERRO: PDF nao encontrado.")

                    except Exception as ex:
                        error_count += 1
                        self.log_message("  > ERRO CRITICO PDF: " + str(ex))
                
                # --- DWG ---
                if do_dwg and not self.is_cancelled:
                    try:
                        self.log_message("  > Gerando DWG...")
                        dwg_opts = None
                        if self.cb_DWGSetups.SelectedItem:
                            s_name = self.cb_DWGSetups.SelectedItem
                            for s in DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings):
                                if s.Name == s_name:
                                    dwg_opts = s.GetDWGExportOptions()
                                    break
                        if not dwg_opts: dwg_opts = self.create_dwg_options_compatible()
                        dwg_opts.MergedViews = True 
                        
                        doc.Export(folder, item.FileName, view_ids, dwg_opts)
                        
                        dwg_path = os.path.join(folder, "{}.dwg".format(item.FileName))
                        found_dwg = False
                        
                        for _ in range(15): 
                            if os.path.exists(dwg_path): 
                                found_dwg = True
                                break
                            time.sleep(0.3)
                            WinFormsApp.DoEvents()

                        if found_dwg:
                            success_count += 1
                            self.log_message("  > DWG OK.")
                        else:
                            alt = os.path.join(folder, "{}_Sheet.dwg".format(item.FileName))
                            if os.path.exists(alt):
                                self.safe_rename_file(alt, dwg_path)
                                success_count += 1
                                self.log_message("  > DWG OK (Renomeado).")
                            else:
                                error_count += 1
                                self.log_message("  > ERRO: DWG nao gerado.")

                    except Exception as ex:
                        error_count += 1
                        self.log_message("  > ERRO CRITICO DWG: " + str(ex))
                
            except Exception as ex:
                self.log_message("ERRO GERAL Folha {}: {}".format(item.Number, str(ex)))
                error_count += 1
        
        # --- FINALIZACAO ---
        total_time = time.time() - start_time
        self.btn_Start.IsEnabled = True
        self.btn_Cancel.Visibility = System.Windows.Visibility.Collapsed
        
        self.log_message("="*40)
        
        if self.is_cancelled:
            self.log_message("CANCELADO PELO USUÁRIO.")
        elif error_count == 0:
            self.log_message("CONCLUÍDO COM SUCESSO!")
            self.log_message("Tempo: {}".format(self.format_time(total_time)))
            self.log_message("Arquivos: {}".format(success_count))
        else:
            self.log_message("CONCLUÍDO COM AVISOS.")
            self.log_message("Sucessos: {} | Erros: {}".format(success_count, error_count))
            
        self.log_message("="*40)
        self.log_message("Abrindo pasta de destino...")
        
        try:
            os.startfile(folder)
        except:
            self.log_message("Nao foi possivel abrir a pasta automaticamente.")

try:
    LuisExporterWindow().ShowDialog()
except Exception as e:
    forms.alert("Erro fatal: " + str(e))