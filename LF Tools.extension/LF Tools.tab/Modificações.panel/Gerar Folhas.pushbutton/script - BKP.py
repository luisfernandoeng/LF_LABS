# -*- coding: utf-8 -*-
"""
EXPORTADOR PRO | LUÍS FERNANDO
Versão Final - Produção
"""
import clr
import System
from pyrevit import revit, forms, script

# --- CARREGAMENTO DE DLLs ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Threading import Dispatcher, DispatcherPriority
import os
import re
import time

doc = revit.doc
uidoc = revit.uidoc

# --- DETECÇÃO DE VERSÃO ---
REVIT_YEAR = int(doc.Application.VersionNumber)
HAS_PDF_SUPPORT = REVIT_YEAR >= 2022

# --- CLASSE DE DADOS ---
class SheetItem(object):
    def __init__(self, sheet):
        self.Element = sheet
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name if sheet.Name else "" 
        self.IsSelected = False
        self.FileName = self._generate_filename(sheet)
        self.PdfFileName = self._generate_pdf_filename(sheet)

    def _generate_filename(self, sheet):
        """Gera nome para exportação"""
        val = ""
        try:
            param = sheet.LookupParameter("C-NOME-FOLHA")
            if not param or not param.AsString():
                param_list = sheet.GetParameters("C-NOME-FOLHA")
                if param_list and len(param_list) > 0:
                    param = param_list[0]
            
            if param and param.AsString() and param.AsString().strip():
                val = param.AsString().strip()
            else:
                val = "{} - {}".format(self.Number, self.Name)
                
        except:
            val = "{} - {}".format(self.Number, self.Name)
        
        clean_val = re.sub(r'[<>:"/\\|?*]', '', val)
        return clean_val

    def _generate_pdf_filename(self, sheet):
        """Gera nome para PDF"""
        return self._generate_filename(sheet)

# --- JANELA PRINCIPAL ---
class LuisExporterWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = script.get_bundle_file('exp_luis_tool.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.sheet_items = ObservableCollection[SheetItem]()
        self.dg_Sheets.ItemsSource = self.sheet_items
        
        self.btn_SelectFolder.Click += self.pick_folder
        self.btn_Start.Click += self.run_export
        
        # Configurações iniciais
        self.cb_Colors.SelectedIndex = 0
        self.rb_ZoomFit.IsChecked = True
        self.chk_ExportDWG.IsChecked = True
        
        # Carregar dados
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

    def load_sheets(self):
        """Carrega folhas"""
        try:
            sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType().ToElements()
            sorted_sheets = sorted(sheets, key=lambda x: x.SheetNumber)
            
            self.sheet_items.Clear()
            for s in sorted_sheets:
                if not s.IsPlaceholder: 
                    self.sheet_items.Add(SheetItem(s))
            self.lbl_Count.Text = str(len(self.sheet_items))
            
        except Exception as e:
            forms.alert("Erro ao carregar folhas: " + str(e))

    def load_dwg_setups(self):
        """Carrega configurações DWG"""
        try:
            setups = DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings).ToElements()
            self.cb_DWGSetups.Items.Clear()
            for s in setups:
                self.cb_DWGSetups.Items.Add(s.Name)
            if self.cb_DWGSetups.Items.Count > 0:
                self.cb_DWGSetups.SelectedIndex = 0
        except:
            pass

    def pick_folder(self, sender, args):
        folder = forms.pick_folder()
        if folder:
            self.txt_OutputFolder.Text = folder

    def create_pdf_options(self):
        """Cria opções de PDF"""
        pdf_opts = DB.PDFExportOptions()
        pdf_opts.Combine = False
        
        color_index = self.cb_Colors.SelectedIndex
        try:
            if color_index == 0:
                if hasattr(pdf_opts, 'ColorDepth'):
                    pdf_opts.ColorDepth = 2
                elif hasattr(pdf_opts, 'ColorMode'):
                    pdf_opts.ColorMode = 2
            elif color_index == 1:
                if hasattr(pdf_opts, 'ColorDepth'):
                    pdf_opts.ColorDepth = 0
                elif hasattr(pdf_opts, 'ColorMode'):
                    pdf_opts.ColorMode = 0
            else:
                if hasattr(pdf_opts, 'ColorDepth'):
                    pdf_opts.ColorDepth = 1
                elif hasattr(pdf_opts, 'ColorMode'):
                    pdf_opts.ColorMode = 1
        except:
            pass
        
        try:
            if self.rb_ZoomFit.IsChecked:
                pdf_opts.ZoomType = DB.ZoomType.FitToPage
                pdf_opts.PaperPlacement = DB.PaperPlacementType.Center
            else:
                pdf_opts.ZoomType = DB.ZoomType.Zoom
                pdf_opts.ZoomPercentage = 100
                pdf_opts.PaperPlacement = DB.PaperPlacementType.Margins
        except:
            pass
            
        return pdf_opts

    def create_dwg_options_compatible(self):
        """Cria opções de DWG compatíveis"""
        dwg_opts = DB.DWGExportOptions()
        dwg_opts.MergedViews = True
        
        try:
            if hasattr(dwg_opts, 'ExportScope'):
                dwg_opts.ExportScope = DB.DWGExportScope.View
        except:
            pass
        
        return dwg_opts

    def run_export(self, sender, args):
        """Executa a exportação"""
        folder = self.txt_OutputFolder.Text
        if not folder or not os.path.exists(folder):
            forms.alert("Selecione uma pasta de destino válida.")
            return

        selected_items = [i for i in self.sheet_items if i.IsSelected]
        if not selected_items:
            forms.alert("Selecione pelo menos uma folha.")
            return

        do_pdf = self.chk_ExportPDF.IsChecked
        do_dwg = self.chk_ExportDWG.IsChecked
        
        success_count = 0
        error_count = 0
        
        self.pb_Progress.Maximum = len(selected_items)
        self.pb_Progress.Value = 0
        self.btn_Start.IsEnabled = False
        
        for i, item in enumerate(selected_items):
            try:
                current_status = "Exportando: {} ({}/{})".format(item.FileName, i+1, len(selected_items))
                self.txt_Status.Text = current_status
                self.refresh_ui()
                
                view_ids = System.Collections.Generic.List[DB.ElementId]([item.Id])
                
                # Exportar PDF
                if do_pdf and HAS_PDF_SUPPORT:
                    try:
                        pdf_opts = self.create_pdf_options()
                        doc.Export(folder, view_ids, pdf_opts)
                        time.sleep(1)
                        
                        # Renomear arquivo PDF
                        pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
                        if pdf_files:
                            latest_file = max(pdf_files, key=lambda f: os.path.getmtime(os.path.join(folder, f)))
                            exported_file = os.path.join(folder, latest_file)
                            
                            if os.path.exists(exported_file):
                                desired_name = "{}.pdf".format(item.PdfFileName)
                                new_file = os.path.join(folder, desired_name)
                                
                                if os.path.exists(new_file):
                                    os.remove(new_file)
                                
                                os.rename(exported_file, new_file)
                                success_count += 1
                    except:
                        error_count += 1
                
                # Exportar DWG
                if do_dwg:
                    try:
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
                        
                        dwg_file = os.path.join(folder, "{}.dwg".format(item.FileName))
                        if os.path.exists(dwg_file):
                            success_count += 1
                        else:
                            # Tentar encontrar arquivo com nome alternativo
                            for ext in ['_Sheet.dwg', '_Model.dwg', '.dwg']:
                                alt_file = os.path.join(folder, "{}{}".format(item.FileName, ext))
                                if os.path.exists(alt_file):
                                    if ext != '.dwg':
                                        os.rename(alt_file, dwg_file)
                                    success_count += 1
                                    break
                            
                    except:
                        error_count += 1
                
                self.pb_Progress.Value = i + 1
                
            except:
                error_count += 1
                if not forms.alert("Erro em {}.\nContinuar exportação?".format(item.Number), yes=True, no=True):
                    break
        
        self.btn_Start.IsEnabled = True
        
        if error_count == 0:
            final_msg = "✅ Exportação concluída com sucesso!\n{} arquivos exportados.\n\nCriado pelo Luís Fernando".format(success_count)
            self.txt_Status.Text = "Concluído! {} arquivos exportados.".format(success_count)
            forms.alert(final_msg)
        else:
            final_msg = "⚠ Exportação finalizada com {} erros.\n{} arquivos exportados com sucesso.".format(error_count, success_count)
            self.txt_Status.Text = "Concluído com {} erros.".format(error_count)
            forms.alert(final_msg)

    def refresh_ui(self):
        """Atualiza a interface"""
        try:
            Dispatcher.CurrentDispatcher.Invoke(
                DispatcherPriority.Background, 
                System.Action(lambda: None)
            )
        except:
            pass

# --- EXECUÇÃO PRINCIPAL ---
try:
    LuisExporterWindow().ShowDialog()
except Exception as e:
    forms.alert("Erro ao abrir o exportador: " + str(e))