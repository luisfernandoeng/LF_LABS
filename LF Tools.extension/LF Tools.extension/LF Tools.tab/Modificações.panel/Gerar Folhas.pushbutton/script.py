# -*- coding: utf-8 -*-
"""
EXPORTADOR PRO | LUIS FERNANDO
Versao sem Progresso Individual
"""
import clr
import System
from pyrevit import revit, forms, script

# --- CARREGAMENTO DE DLLs ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
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
    def __init__(self, sheet, param_name="C-NOME-FOLHA"):
        self.Element = sheet
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name if sheet.Name else "" 
        self.IsSelected = False
        self.param_name = param_name
        self.FileName = self._generate_filename(sheet, param_name)
        self.PdfFileName = self._generate_pdf_filename(sheet, param_name)

    def _generate_filename(self, sheet, param_name):
        """Gera nome para exportacao"""
        val = ""
        try:
            param = sheet.LookupParameter(param_name)
            if not param or not param.AsString():
                param_list = sheet.GetParameters(param_name)
                if param_list and len(param_list) > 0:
                    param = param_list[0]
            
            if param and param.AsString() and param.AsString().strip():
                val = param.AsString().strip()
            else:
                val = "{} - {}".format(self.Number, self.Name)
                
        except Exception as ex:
            print("Erro ao gerar filename: {}".format(str(ex)))
            val = "{} - {}".format(self.Number, self.Name)
        
        clean_val = re.sub(r'[<>:"/\\|?*]', '', val)
        return clean_val

    def _generate_pdf_filename(self, sheet, param_name):
        """Gera nome para PDF"""
        return self._generate_filename(sheet, param_name)
    
    def update_filename(self, param_name):
        """Atualiza o nome do arquivo baseado em novo parametro"""
        self.param_name = param_name
        self.FileName = self._generate_filename(self.Element, param_name)
        self.PdfFileName = self._generate_pdf_filename(self.Element, param_name)

# --- JANELA PRINCIPAL ---
class LuisExporterWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = script.get_bundle_file('exp_luis_tool.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.sheet_items = ObservableCollection[SheetItem]()
        self.dg_Sheets.ItemsSource = self.sheet_items
        
        self.btn_SelectFolder.Click += self.pick_folder
        self.btn_Start.Click += self.run_export
        self.btn_Cancel.Click += self.cancel_export
        self.cb_SheetParameters.SelectionChanged += self.on_parameter_changed
        
        # Flag de cancelamento
        self.is_cancelled = False
        
        # Configuracoes iniciais
        self.rb_ZoomFit.IsChecked = True
        self.chk_ExportDWG.IsChecked = True
        
        # Carregar dados
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

    def load_sheet_parameters(self):
        """Carrega parametros disponiveis nas folhas"""
        try:
            self.cb_SheetParameters.Items.Clear()
            
            sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType().ToElements()
            if sheets and len(sheets) > 0:
                sample_sheet = sheets[0]
                
                common_params = ["C-NOME-FOLHA", "Sheet Number", "Sheet Name"]
                
                for param in sample_sheet.Parameters:
                    param_name = param.Definition.Name
                    if param.StorageType == DB.StorageType.String and param_name not in common_params:
                        common_params.append(param_name)
                
                for param_name in common_params:
                    self.cb_SheetParameters.Items.Add(param_name)
                
                if "C-NOME-FOLHA" in common_params:
                    self.cb_SheetParameters.SelectedItem = "C-NOME-FOLHA"
                elif self.cb_SheetParameters.Items.Count > 0:
                    self.cb_SheetParameters.SelectedIndex = 0
                    
        except Exception as ex:
            print("Erro ao carregar parametros: {}".format(str(ex)))
            self.cb_SheetParameters.Items.Add("C-NOME-FOLHA")
            self.cb_SheetParameters.SelectedIndex = 0

    def on_parameter_changed(self, sender, args):
        """Atualiza nomes dos arquivos quando o parametro muda"""
        if not self.cb_SheetParameters.SelectedItem:
            return
            
        selected_param = str(self.cb_SheetParameters.SelectedItem)
        
        for item in self.sheet_items:
            item.update_filename(selected_param)
        
        self.dg_Sheets.Items.Refresh()

    def load_sheets(self):
        """Carrega folhas"""
        try:
            sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType().ToElements()
            sorted_sheets = sorted(sheets, key=lambda x: x.SheetNumber)
            
            selected_param = "C-NOME-FOLHA"
            if self.cb_SheetParameters.SelectedItem:
                selected_param = str(self.cb_SheetParameters.SelectedItem)
            
            self.sheet_items.Clear()
            for s in sorted_sheets:
                if not s.IsPlaceholder: 
                    self.sheet_items.Add(SheetItem(s, selected_param))
            self.lbl_Count.Text = str(len(self.sheet_items))
            
        except Exception as e:
            forms.alert("Erro ao carregar folhas: " + str(e))

    def load_dwg_setups(self):
        """Carrega configuracoes DWG"""
        try:
            setups = DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings).ToElements()
            self.cb_DWGSetups.Items.Clear()
            for s in setups:
                self.cb_DWGSetups.Items.Add(s.Name)
            if self.cb_DWGSetups.Items.Count > 0:
                self.cb_DWGSetups.SelectedIndex = 0
        except Exception as ex:
            print("Erro ao carregar DWG setups: {}".format(str(ex)))

    def pick_folder(self, sender, args):
        folder = forms.pick_folder()
        if folder:
            self.txt_OutputFolder.Text = folder
    
    def cancel_export(self, sender, args):
        """Cancela a exportacao em andamento"""
        if forms.alert("Deseja realmente cancelar a exportacao?", yes=True, no=True):
            self.is_cancelled = True

    def create_pdf_options(self):
        """Cria opcoes de PDF"""
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
        except Exception as ex:
            print("Erro ao configurar zoom PDF: {}".format(str(ex)))
            
        return pdf_opts

    def create_dwg_options_compatible(self):
        """Cria opcoes de DWG compativeis"""
        dwg_opts = DB.DWGExportOptions()
        dwg_opts.MergedViews = True
        
        try:
            if hasattr(dwg_opts, 'ExportScope'):
                dwg_opts.ExportScope = DB.DWGExportScope.View
        except Exception as ex:
            print("Erro ao configurar DWG options: {}".format(str(ex)))
        
        return dwg_opts

    def safe_rename_file(self, old_path, new_path, max_attempts=5):
        """Renomeia arquivo com retry em caso de bloqueio"""
        for attempt in range(max_attempts):
            try:
                if os.path.exists(new_path):
                    os.remove(new_path)
                os.rename(old_path, new_path)
                return True
            except Exception as ex:
                if attempt < max_attempts - 1:
                    time.sleep(0.5)
                else:
                    print("Erro ao renomear {}: {}".format(old_path, str(ex)))
                    return False
        return False
    
    def format_time(self, seconds):
        """Formata tempo em segundos para string legivel"""
        if seconds < 60:
            return "{:.0f}s".format(seconds)
        elif seconds < 3600:
            minutes = seconds / 60
            return "{:.1f}min".format(minutes)
        else:
            hours = seconds / 3600
            return "{:.1f}h".format(hours)

    def run_export(self, sender, args):
        """Executa a exportacao"""
        folder = self.txt_OutputFolder.Text
        if not folder or not os.path.exists(folder):
            forms.alert("Selecione uma pasta de destino valida.")
            return

        selected_items = [i for i in self.sheet_items if i.IsSelected]
        if not selected_items:
            forms.alert("Selecione pelo menos uma folha.")
            return

        do_pdf = self.chk_ExportPDF.IsChecked
        do_dwg = self.chk_ExportDWG.IsChecked
        
        if not do_pdf and not do_dwg:
            forms.alert("Ative pelo menos uma opcao de exportacao (PDF ou DWG).")
            return
        
        # Prepara a contagem total para o resumo final
        total_sheets = len(selected_items)
        total_exports = total_sheets * (1 if do_pdf else 0) + total_sheets * (1 if do_dwg else 0)
        
        # Desativa/Ativa botoes e UI
        self.btn_Start.IsEnabled = False
        self.btn_Cancel.Visibility = System.Windows.Visibility.Visible
        self.tab_Main.SelectedIndex = 1
        self.UpdateLayout()
        time.sleep(0.2)
        
        # Reset do estado
        self.is_cancelled = False
        success_count = 0
        error_count = 0
        start_time = time.time()
        
        
        for item in selected_items:
            if self.is_cancelled:
                break
            
            try:
                view_ids = System.Collections.Generic.List[DB.ElementId]([item.Id])
                
                # Exportar PDF
                if do_pdf and HAS_PDF_SUPPORT and not self.is_cancelled:
                    try:
                        print("Exporting PDF: {}".format(item.FileName))
                        
                        pdf_opts = self.create_pdf_options()
                        pdf_files_before = set([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                        
                        doc.Export(folder, view_ids, pdf_opts)
                        
                        # Aguarda arquivo e renomeia
                        max_wait = 10
                        wait_interval = 0.3
                        elapsed = 0
                        found = False
                        
                        while elapsed < max_wait and not self.is_cancelled:
                            time.sleep(wait_interval)
                            elapsed += wait_interval
                            
                            pdf_files_after = set([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                            new_files = pdf_files_after - pdf_files_before
                            
                            if new_files:
                                latest_file = list(new_files)[0]
                                exported_file = os.path.join(folder, latest_file)
                                desired_name = "{}.pdf".format(item.PdfFileName)
                                new_file = os.path.join(folder, desired_name)
                                
                                if self.safe_rename_file(exported_file, new_file):
                                    success_count += 1
                                    found = True
                                else:
                                    error_count += 1
                                break
                        
                        if not found and not self.is_cancelled:
                            error_count += 1
                        
                            
                    except Exception as ex:
                        if not self.is_cancelled:
                            error_count += 1
                            print("Erro PDF {}: {}".format(item.FileName, str(ex)))
                
                # Exportar DWG
                if do_dwg and not self.is_cancelled:
                    try:
                        print("Exporting DWG: {}".format(item.FileName))
                        
                        dwg_opts = None
                        if self.cb_DWGSetups.SelectedItem:
                            setup_name = self.cb_DWGSetups.SelectedItem
                            setups = DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings)
                            for setup in setups:
                                if setup.Name == setup_name:
                                    dwg_opts = setup.GetDWGExportOptions()
                                    break
                        
                        if dwg_opts is None:
                            dwg_opts = self.create_dwg_options_compatible()
                        
                        dwg_opts.MergedViews = True
                        doc.Export(folder, item.FileName, view_ids, dwg_opts)
                        
                        time.sleep(0.5)
                        
                        if self.is_cancelled:
                            break
                        
                        dwg_file = os.path.join(folder, "{}.dwg".format(item.FileName))
                        found = False
                        
                        if os.path.exists(dwg_file):
                            found = True
                        else:
                            for ext in ['_Sheet.dwg', '_Model.dwg']:
                                alt_file = os.path.join(folder, "{}{}".format(item.FileName, ext))
                                if os.path.exists(alt_file):
                                    if self.safe_rename_file(alt_file, dwg_file):
                                        found = True
                                        break
                        
                        if found:
                            success_count += 1
                        else:
                            error_count += 1
                            
                    except Exception as ex:
                        if not self.is_cancelled:
                            error_count += 1
                            print("Erro DWG {}: {}".format(item.FileName, str(ex)))
                
                time.sleep(0.05) 
                
            except Exception as ex:
                if not self.is_cancelled:
                    print("Erro geral {}: {}".format(item.Number, str(ex)))
                    error_count += 1
        
        # Finalizacao
        total_time = time.time() - start_time
        self.btn_Start.IsEnabled = True
        self.btn_Cancel.Visibility = System.Windows.Visibility.Collapsed
        
        # Mensagens de resultado final
        if self.is_cancelled:
            forms.alert("Exportacao cancelada!\n{} arquivos foram exportados.".format(success_count))
        elif error_count == 0:
            forms.alert("Exportacao concluida!\n{} arquivos em {}.".format(success_count, self.format_time(total_time)))
        else:
            forms.alert("Finalizado com {} erros.\n{} arquivos exportados.".format(error_count, success_count))

# --- EXECUCAO PRINCIPAL ---
try:
    LuisExporterWindow().ShowDialog()
except Exception as e:
    forms.alert("Erro ao abrir o exportador: " + str(e))
    print("Stack trace: {}".format(str(e)))