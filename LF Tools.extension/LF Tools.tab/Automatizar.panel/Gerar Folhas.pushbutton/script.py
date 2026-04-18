# -*- coding: utf-8 -*-
"""
GERAR FOLHAS | LUIS FERNANDO
"""
from pyrevit import forms, script
import clr
import System
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB
clr.AddReference('System.Windows.Forms')
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.Windows.Forms import Application as WinFormsApp
from System.Windows import Input
from System.Windows.Data import CollectionViewSource
import System.Windows.Controls as Controls
import System.Windows.Media as Media
import os
import re
import time
import json
import codecs

# --- BIBLIOTECA LF TOOLS ---
try:
    from lf_utils import DebugLogger
except:
    class DebugLogger(object):
        def __init__(self, *args, **kwargs): pass
        def info(self, *args): pass
        def section(self, *args): pass
        def warn(self, *args): pass
        def error(self, *args): pass

# --- DETECAO DE VERSAO ---
try:
    REVIT_YEAR = int(__revit__.Application.VersionNumber)
except:
    REVIT_YEAR = 2024

# --- INICIALIZAÇÃO DO DOCUMENTO ---
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None

if not doc:
    forms.alert("Erro Crítico de Inicialização:\nNão foi possível localizar nenhum projeto aberto no Revit.\n\nPor favor, feche e abra o Revit ou verifique se há um arquivo (.rvt) ativo.", exitscript=True)

HAS_PDF_SUPPORT = REVIT_YEAR >= 2022

# --- CAMINHO DE CONFIGURACAO ---
CONFIG_DIR = os.path.join(os.getenv('APPDATA'), 'LFTools')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'config.json')

# --- CLASSE DE DADOS PARA FOLHAS ---
import System.ComponentModel as ComponentModel

class SheetItem(object):
    def __init__(self, sheet, name_pattern="{Nome da folha}"):
        self.Element = sheet
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.Name = sheet.Name if sheet.Name else ""
        self._is_selected = False
        self.name_pattern = name_pattern
        self._file_name = ""
        self.PdfFileName = ""

        # Status
        self._status_icon = ""
        self._status_color = "Transparent"
        self._status_tooltip = ""
        self.HasViews = True

        self._file_name = self._generate_filename(sheet, name_pattern)
        self.PdfFileName = self._file_name

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value

    @property
    def StatusIcon(self):
        return self._status_icon

    @StatusIcon.setter
    def StatusIcon(self, value):
        self._status_icon = value

    @property
    def StatusColor(self):
        return self._status_color

    @StatusColor.setter
    def StatusColor(self, value):
        self._status_color = value

    @property
    def StatusToolTip(self):
        return self._status_tooltip

    @StatusToolTip.setter
    def StatusToolTip(self, value):
        self._status_tooltip = value

    @property
    def FileName(self):
        return self._file_name

    @FileName.setter
    def FileName(self, value):
        self._file_name = value
        self.PdfFileName = value

    def _get_param_value(self, sheet, p_name):
        if p_name == "Sheet Number": return sheet.SheetNumber
        if p_name == "Sheet Name": return sheet.Name
        if p_name == "Current Revision": return str(sheet.GetCurrentRevision())
        try:
            param = sheet.LookupParameter(p_name)
            if not param:
                p_list = sheet.GetParameters(p_name)
                if p_list: param = p_list[0]
            if param:
                if param.StorageType == DB.StorageType.String:
                    return param.AsString() or ""
                else:
                    return param.AsValueString() or ""
        except:
            pass
        return ""

    def _generate_filename(self, sheet, pattern):
        final_name = pattern
        for tag in re.findall(r'\{(.*?)\}', pattern):
            val = self._get_param_value(sheet, tag)
            final_name = final_name.replace("{" + tag + "}", str(val))
        return re.sub(r'[<>:"/\\|?*]', '', final_name).strip()

    def update_filename(self, new_pattern):
        self.name_pattern = new_pattern
        self.FileName = self._generate_filename(self.Element, new_pattern)

    def update_status(self, folder_path):
        self.StatusIcon = ""
        self.StatusColor = "Transparent"
        self.StatusToolTip = ""

        if not self.HasViews:
            self.StatusIcon = "\U0001f4c4"
            self.StatusColor = "#FF888888"
            self.StatusToolTip = "Folha vazia (sem vistas)"

        if folder_path and os.path.exists(folder_path):
            exists_pdf = os.path.exists(os.path.join(folder_path, "{}.pdf".format(self.PdfFileName)))
            exists_dwg = os.path.exists(os.path.join(folder_path, "{}.dwg".format(self.FileName)))
            if exists_pdf or exists_dwg:
                # Usa sinal de warning padrao texto sem "Variacao Emoji" (\uFE0F) para permitir customizacao de cor 
                self.StatusIcon = "\u26A0"
                self.StatusColor = "#FFF7C453"
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
        "open_folder_after": True,
        "profiles": {},
        "last_profile": "",
        "debug_mode": False
    }
    try:
        if os.path.exists(CONFIG_FILE):
            with codecs.open(CONFIG_FILE, 'r', 'utf-8') as f:
                data = json.load(f)
                # Garante que as chaves de profiles existam caso o arquivo seja antigo
                if "profiles" not in data:
                    data["profiles"] = {}
                return data
    except:
        pass
    return default_config

def save_config(config):
    """Salva configuracoes no arquivo JSON"""
    try:
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        with codecs.open(CONFIG_FILE, 'w', 'utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as ex:
        forms.alert("Erro ao salvar config: " + str(ex))


# --- JANELA PRINCIPAL ---
class LuisExporterWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = script.get_bundle_file('folhas_window.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.sheet_items = ObservableCollection[object]()

        # Configurar CollectionView para filtragem
        self.view_source = CollectionViewSource.GetDefaultView(self.sheet_items)
        self.view_source.Filter = self.filter_sheets
        
        self.lst_Sheets.ItemsSource = self.view_source
        
        # Eventos Principais
        self.btn_SelectFolder.Click += self.pick_folder
        self.btn_Start.Click += self.run_export
        self.btn_Cancel.Click += self.cancel_export
        self.btn_AddParam.Click += self.add_param_to_pattern
        self.txt_NamePattern.TextChanged += self.on_pattern_changed
        
        # NOVO: Evento Selecionar Tudo
        self.chk_SelectAll.Click += self.toggle_all_sheets
        self.lst_Sheets.SelectionChanged += self.on_list_selection_changed
        
        # NOVO: Busca e Filtros
        self.txt_Search.TextChanged += self.on_search_text_changed

        # NOVO: Perfis de Exportação (Presets)
        try:
            self.cb_Profiles.SelectionChanged += self.apply_profile
            self.btn_AddProfile.Click += self.add_profile
            self.btn_SaveProfile.Click += self.save_profile
            self.btn_DeleteProfile.Click += self.delete_profile
        except Exception:
            pass

        # NOVO: Persistencia imediata do Debug
        self.chk_DebugMode.Click += self.toggle_debug_persistence

        # Configurações Iniciais
        self.config = load_config()
        self.txt_OutputFolder.Text = self.config.get("last_folder", "")
        self.chk_OpenFolderAfter.IsChecked = self.config.get("open_folder_after", True)
        self.chk_DebugMode.IsChecked = self.config.get("debug_mode", False)
        
        # Inicializa o Logger de Debug (integrado com o log da UI)
        self.dbg = DebugLogger(self.chk_DebugMode.IsChecked, log_func=None)
        
        self.is_handling_click = False 
        

        
        self.is_cancelled = False
        self.rb_ZoomFit.IsChecked = True
        self.chk_ExportDWG.IsChecked = True
        
        self.load_sheet_parameters() 
        self.load_sheets()
        
        # Checar arquivos existentes na inicializacao
        self.check_existing_files()
        create_default_dwg_settings()  # Cria settings padrao se nao existir
        self.load_dwg_setups()
        self.load_profiles()  # Carrega os perfis do JSON
        
        if HAS_PDF_SUPPORT:
            self.chk_ExportPDF.IsEnabled = True
            self.pnl_PDFSettings.IsEnabled = True
            self.chk_ExportPDF.IsChecked = True
        else:
            self.chk_ExportPDF.IsChecked = False
            self.chk_ExportPDF.IsEnabled = False
            self.chk_ExportPDF.Content = "PDF Requer Revit 2022+"
            self.pnl_PDFSettings.IsEnabled = False

        self.dbg.info("Interface inicializada com sucesso.")

    def toggle_debug_persistence(self, sender, args):
        """Salva a escolha do modo debug imediatamente no config global"""
        is_on = bool(self.chk_DebugMode.IsChecked)
        self.dbg.enabled = is_on
        self.config["debug_mode"] = is_on
        save_config(self.config)
        if is_on:
            self.dbg.section("MODO DEBUG ATIVADO")
            self.dbg.info("As configurações de debug agora serão persistentes.")

    def force_refresh_sheets(self):
        """Atualiza a lista visualmente"""
        try:
            self.lst_Sheets.Items.Refresh()
        except:
            pass

    # --- LOGICA SELECIONAR TUDO ---
    def toggle_all_sheets(self, sender, args):
        """Marca ou desmarca todas as folhas VISIVEIS"""
        is_checked = bool(self.chk_SelectAll.IsChecked)
        for item in self.sheet_items:
            # Bugfix: Respeitar o filtro
            if self.filter_sheets(item):
                item.IsSelected = is_checked
        self.force_refresh_sheets()


    def Window_PreviewKeyDown(self, sender, args):
        """Atalhos de teclado: Ctrl+A (Tudo), Ctrl+Shift+A (Inverter)"""
        try:
            if args.Key == Input.Key.A and (Input.Keyboard.Modifiers & Input.ModifierKeys.Control) == Input.ModifierKeys.Control:
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
                    
                    self.lst_Sheets.Items.Refresh()
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
            
            self._hard_refresh()
        except:
            pass

    # --- LOGICA DE PERFIS (PRESETS) ---
    def load_profiles(self):
        """Preenche o ComboBox de perfis e seleciona o ultimo usado."""
        try:
            if not hasattr(self, 'cb_Profiles'): return
            self.cb_Profiles.Items.Clear()
            
            profiles = self.config.get("profiles", {})
            for p_name in sorted(profiles.keys()):
                self.cb_Profiles.Items.Add(p_name)
                
            last_profile = self.config.get("last_profile", "")
            if last_profile in profiles:
                self.cb_Profiles.SelectedItem = last_profile
            elif self.cb_Profiles.Items.Count > 0:
                self.cb_Profiles.SelectedIndex = 0
            
            self.dbg.info("Perfis carregados: {}".format(len(profiles)))
        except Exception as ex:
            self.dbg.error("Erro ao carregar perfis: " + str(ex))

    def apply_profile(self, sender, args):
        """Aplica as conf do perfil selecionado re-preenchendo a interface."""
        if not hasattr(self, 'cb_Profiles'): return
        p_name = self.cb_Profiles.SelectedItem
        if not p_name: return
        
        # Salva ultimo selecionado
        self.config["last_profile"] = p_name
        save_config(self.config)
        
        profiles = self.config.get("profiles", {})
        data = profiles.get(p_name)
        if not data: return
        
        try:
            # Aplica Checkboxes Principais
            if "export_dwg" in data:
                self.chk_ExportDWG.IsChecked = data["export_dwg"]
            if "export_pdf" in data:
                self.chk_ExportPDF.IsChecked = data["export_pdf"]
                
            # DWG Setup
            if "dwg_setup" in data:
                dwg_setup = data["dwg_setup"]
                if self.cb_DWGSetups.Items.Contains(dwg_setup):
                    self.cb_DWGSetups.SelectedItem = dwg_setup

            # PDF
            if "pdf_hide_crop" in data:
                self.chk_HideCrop.IsChecked = data["pdf_hide_crop"]
            if "pdf_combine" in data:
                self.chk_CombinePDF.IsChecked = data["pdf_combine"]
            if "pdf_zoom_type" in data:
                z_type = data["pdf_zoom_type"]
                if z_type == "fit":
                    self.rb_ZoomFit.IsChecked = True
                elif z_type == "zoom":
                    self.rb_Zoom100.IsChecked = True

            # Nome e Output
            if "name_pattern" in data:
                self.txt_NamePattern.Text = data["name_pattern"]
            if "output_folder" in data:
                self.txt_OutputFolder.Text = data["output_folder"]
            if "open_folder" in data:
                self.chk_OpenFolderAfter.IsChecked = data["open_folder"]
            if "debug_mode" in data:
                self.chk_DebugMode.IsChecked = data["debug_mode"]

            # Folhas selecionadas
            selected_numbers = data.get("selected_sheets", None)
            if selected_numbers is not None:
                self.lst_Sheets.UnselectAll()
                sel_set = set(selected_numbers)
                for item in self.sheet_items:
                    if item.Number in sel_set:
                        try:
                            self.lst_Sheets.SelectedItems.Add(item)
                        except:
                            pass
                self.force_refresh_sheets()
                
        except Exception as ex:
            print("Erro ao aplicar perfil: " + str(ex))
            
    def _collect_profile_data(self):
        """Coleta o estado atual da UI em um dicionario."""
        selected_sheet_numbers = [item.Number for item in self.sheet_items if item.IsSelected]
        return {
            "export_dwg": bool(self.chk_ExportDWG.IsChecked),
            "export_pdf": bool(self.chk_ExportPDF.IsChecked),
            "debug_mode": bool(self.chk_DebugMode.IsChecked),
            "dwg_setup": self.cb_DWGSetups.Text,
            "pdf_hide_crop": bool(self.chk_HideCrop.IsChecked),
            "pdf_combine": bool(self.chk_CombinePDF.IsChecked),
            "pdf_zoom_type": "fit" if self.rb_ZoomFit.IsChecked else "zoom",
            "name_pattern": self.txt_NamePattern.Text,
            "output_folder": self.txt_OutputFolder.Text,
            "open_folder": bool(self.chk_OpenFolderAfter.IsChecked),
            "selected_sheets": selected_sheet_numbers
        }

    def add_profile(self, sender, args):
        """Cria um perfil NOVO pedindo o nome ao usuario."""
        if not self.cb_Profiles: return
        
        p_name = forms.ask_for_string(prompt="Nome do novo perfil:", title="Criar Perfil")
        if not p_name: return
        
        profiles = self.config.get("profiles", {})
        
        if p_name in profiles:
            if not forms.alert("Já existe um perfil chamado '{}'. Deseja sobrescrever?".format(p_name), yes=True, no=True):
                return
        
        profiles[p_name] = self._collect_profile_data()
        self.config["profiles"] = profiles
        self.config["last_profile"] = p_name
        save_config(self.config)
        self.load_profiles()
        self.cb_Profiles.SelectedItem = p_name
        forms.alert("Perfil '{}' criado com sucesso!".format(p_name))

    def save_profile(self, sender, args):
        """Salva o estado atual NO PERFIL SELECIONADO (atualiza)."""
        if not self.cb_Profiles: return
        
        p_name = self.cb_Profiles.SelectedItem
        if not p_name:
            forms.alert("Nenhum perfil selecionado. Use o botão + para criar um novo.")
            return
        
        profiles = self.config.get("profiles", {})
        profiles[p_name] = self._collect_profile_data()
        self.config["profiles"] = profiles
        save_config(self.config)
        forms.alert("Perfil '{}' atualizado!".format(p_name))

    def delete_profile(self, sender, args):
        """Exclui perfil selecionado"""
        if not self.cb_Profiles: return
        p_name = self.cb_Profiles.SelectedItem
        if not p_name: return
        
        if not forms.alert("Excluir definitivamente o perfil '{}'?".format(p_name), yes=True, no=True):
            return
            
        profiles = self.config.get("profiles", {})
        if p_name in profiles:
            del profiles[p_name]
            self.config["profiles"] = profiles
            self.config["last_profile"] = ""
            save_config(self.config)
            
            self.load_profiles()
            forms.alert("Perfil deletado.", title="Excluído")

    # --- LOGICA DE SELECAO VISUAL ---
    def _hard_refresh(self):
        """Força a UI a descartar o cache redesenhando a DataGrid inteira via novo CollectionView."""
        try:
            # Desacopla da interface completamente
            self.lst_Sheets.ItemsSource = None
            
            # Recria o controlador de visualizacao
            self.view_source = CollectionViewSource.GetDefaultView(self.sheet_items)
            self.view_source.Filter = self.filter_sheets
            
            # Devolve pra DataGrid e forca atualizacao visual na marra
            self.lst_Sheets.ItemsSource = self.view_source
            self.lst_Sheets.Items.Refresh()
        except:
            pass

    def toggle_all_sheets(self, sender, args):
        """Seleciona ou desseleciona todos os itens visiveis na tabela."""
        try:
            new_state = bool(self.chk_SelectAll.IsChecked)
            
            if new_state:
                # Seleciona todos os itens que passam pelo filtro
                for item in self.sheet_items:
                    if self.filter_sheets(item):
                        if not self.lst_Sheets.SelectedItems.Contains(item):
                            self.lst_Sheets.SelectedItems.Add(item)
            else:
                # Limpa selecao (ou remove apenas os visiveis se preferir)
                self.lst_Sheets.UnselectAll()
                
            self.force_refresh_sheets()
        except Exception as ex:
            print("Erro ao selecionar tudo: " + str(ex))


    def on_list_selection_changed(self, sender, args):
        """Sincroniza a selecao do ListBox com a propriedade IsSelected do objeto."""
        try:
            # Para cada item que foi SELECIONADO agora:
            for item in args.AddedItems:
                item.IsSelected = True
            
            # Para cada item que foi DESSELECIONADO agora:
            for item in args.RemovedItems:
                item.IsSelected = False
        except:
            pass

    def on_row_clicked(self, sender, args):
        """Mantido apenas para compatibilidade, o ListBox ja lida com a selecao."""
        pass

    # --- FUNCOES DE UI E PROGRESSO ---

    def log_message(self, message):
        """Escreve no log detalhado"""
        try:
            timestamp = time.strftime("%H:%M:%S")
            if "===" in message:
                full_msg = "{}\n".format(message)
            else:
                full_msg = "[{}] {}\n".format(timestamp, message)
            
            # Envia também para o debug console se ativo
            self.dbg.info(message)
            
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
            self.lbl_TimeElapsed.Text = "--"
            self.lbl_TimeLabel.Text = "Restante"
            self.progressBar.Width = 0
            self.pnl_ExportItems.Children.Clear()
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
            
            # Atualiza largura da barra
            parent_width = 0
            if self.progressBar.Parent:
                parent_width = self.progressBar.Parent.ActualWidth
            if parent_width < 50:
                parent_width = 500  # fallback se o tab ainda nao renderizou
            self.progressBar.Width = (percent / 100.0) * parent_width
            
            WinFormsApp.DoEvents()
        except:
            pass
    
    def update_counters(self, success, errors, remaining_seconds=None):
        """Atualiza contadores de sucesso/erro/tempo restante"""
        try:
            self.lbl_SuccessCount.Text = str(success)
            self.lbl_ErrorCount.Text = str(errors)
            if remaining_seconds is not None:
                self.lbl_TimeElapsed.Text = "~" + self.format_time(remaining_seconds)
            WinFormsApp.DoEvents()
        except:
            pass
    
    def add_export_item(self, sheet_number, file_name, status, file_type=""):
        """Adiciona um item na lista de exportados com status visual"""
        try:
            # Container do item
            border = Controls.Border()
            border.CornerRadius = System.Windows.CornerRadius(4)
            border.Padding = System.Windows.Thickness(10, 8, 10, 8)
            border.Margin = System.Windows.Thickness(0, 0, 0, 4)
            
            if status == "success":
                border.Background = Media.SolidColorBrush(Media.Color.FromArgb(26, 78, 201, 155))  # Verde translucido
                icon = "✓"
                icon_color = Media.Color.FromRgb(78, 201, 155)
            elif status == "error":
                border.Background = Media.SolidColorBrush(Media.Color.FromArgb(26, 210, 85, 85))  # Vermelho translucido
                icon = "✕"
                icon_color = Media.Color.FromRgb(210, 85, 85)
            else:  # processing
                border.Background = Media.SolidColorBrush(Media.Color.FromArgb(26, 86, 156, 214))  # Azul translucido
                icon = "◐"
                icon_color = Media.Color.FromRgb(86, 156, 214)
            
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
            
            # Scroll para o final - percorrer arvore visual ate achar ScrollViewer
            parent = self.pnl_ExportItems.Parent
            while parent is not None:
                if hasattr(parent, 'ScrollToEnd'):
                    parent.ScrollToEnd()
                    break
                parent = getattr(parent, 'Parent', None)
            
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
        self.force_refresh_sheets()

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
            self.dbg.info("Folhas carregadas do projeto: {}".format(len(self.sheet_items)))
        except Exception as e:
            self.dbg.error("Erro ao carregar folhas: " + str(e))
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
        # Sincroniza o modo debug com a UI
        self.dbg.enabled = bool(self.chk_DebugMode.IsChecked)
        self.dbg.section("Iniciando Exportação de Folhas")
        self.dbg.info("Versão do Revit detectada: {}".format(REVIT_YEAR))
        
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
        self.IsEnabled = False  # Desativa a janela toda p/ evitar cliques
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
            self.dbg.sub("Processando Folha: {}".format(item.Number))
            self.dbg.debug("ElementId: {}".format(item.Element.Id))
            
            if self.is_cancelled: 
                self.dbg.warn("Exportação cancelada pelo usuário.")
                break
            
            current_num = idx + 1
            
            # Calculo de estimativa
            remaining_txt = ""
            est_remaining = None
            if len(last_times) > 0:
                avg_time = sum(last_times) / len(last_times)
                remaining_items = total_items - idx
                est_remaining = remaining_items * avg_time
                remaining_txt = " | Restante: ~{}".format(self.format_time(est_remaining))

            self.update_progress(current_num, total_items, "Exportando: {} - {}{}".format(item.Number, item.Name, remaining_txt))
            self.update_counters(success_count, error_count, est_remaining)
            
            # Pequeno respiro para manter a interface responsiva durante o loop
            if (idx + 1) % 5 == 0:
                WinFormsApp.DoEvents()
                time.sleep(0.01)

            try:
                self.log_message("PROCESSANDO: {} - {}".format(item.Number, item.Name))
                view_ids = List[DB.ElementId]()
                view_ids.Add(item.Element.Id)

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
                        success_count += 1
                        self.update_counters(success_count, error_count)
                    except Exception as e:
                        error_count += 1
                        self.add_export_item(item.Number, "Erro DWG", "error", "")
                        self.log_message("  > ERRO DWG: " + str(e))
                        self.update_counters(success_count, error_count)

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

                            expected_path = os.path.join(folder, "{}.pdf".format(item.PdfFileName))
                            try:
                                if os.path.exists(expected_path): os.remove(expected_path)
                            except: pass

                            # Revit ignora pdf_opts.FileName e gera com nome próprio,
                            # então detectamos o arquivo novo pela diferença de listagem.
                            before_files = set(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
                            doc.Export(folder, view_ids, pdf_opts)

                            final_path = None
                            for _ in range(30):  # até 10s
                                if self.is_cancelled: break
                                current_files = set(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
                                new_files = current_files - before_files
                                if new_files:
                                    generated = os.path.join(folder, next(iter(new_files)))
                                    if self.safe_rename_file(generated, expected_path):
                                        final_path = expected_path
                                    break
                                time.sleep(0.33)
                                WinFormsApp.DoEvents()

                            if final_path:
                                self.log_message("  [PDF] OK: {}".format(item.PdfFileName))
                                self.add_export_item(item.Number, item.FileName, "success", "PDF")
                                success_count += 1
                                self.update_counters(success_count, error_count)
                            else:
                                error_count += 1
                                self.log_message("  [PDF] ERRO: Nenhum arquivo novo detectado.")
                                self.add_export_item(item.Number, "PDF?", "error", "PDF")
                                self.update_counters(success_count, error_count)

                        except Exception as e:
                            error_count += 1
                            self.log_message("  > ERRO PDF: " + str(e))
                            self.add_export_item(item.Number, "Erro PDF", "error", "")
                            self.update_counters(success_count, error_count)
            
            except Exception as ex:
                error_count += 1
                self.add_export_item(item.Number, "Erro: {}".format(str(ex)[:30]), "error", "")
                self.log_message("  > ERRO GERAL LOOP: {}".format(str(ex)))
                self.update_counters(success_count, error_count)
            
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
        self.IsEnabled = True  # Reativa a janela
        self.btn_Start.IsEnabled = True
        self.btn_Cancel.Visibility = System.Windows.Visibility.Collapsed
        self.lbl_TimeLabel.Text = "Tempo Total"
        self.lbl_TimeElapsed.Text = self.format_time(total_time)
        self.update_counters(success_count, error_count)
        
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