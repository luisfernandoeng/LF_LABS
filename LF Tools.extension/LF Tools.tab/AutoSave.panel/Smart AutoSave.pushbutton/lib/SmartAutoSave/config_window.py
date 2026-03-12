# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationFramework')
from System.Windows import Window

from pyrevit import forms
from SmartAutoSave.config_manager import config
from SmartAutoSave.autosave_manager import AutoSaveManager

class ConfigWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        # Load values into UI
        try:
            self.EnableAutoSave.IsChecked = config.get("enabled", True)
            self.ShowToast.IsChecked = config.get("show_toast", True)
            # Radio Buttons for Interval
            interval = config.get("interval_minutes", 10)
            if interval == 1:
                self.Radio1min.IsChecked = True
            elif interval == 5:
                self.Radio5min.IsChecked = True
            elif interval == 15:
                self.Radio15min.IsChecked = True
            elif interval == 30:
                self.Radio30min.IsChecked = True
            else:
                self.Radio10min.IsChecked = True
                
            # Setup Pause Button state
            self.manager = AutoSaveManager(__revit__)
            self.update_pause_btn()
        except Exception as e:
            print("Erro ao carregar configurações: " + str(e))
            
    def update_pause_btn(self):
        if self.manager.is_paused:
            self.BtnPauseToggle.Content = "▶️ Retomar AutoSave"
        else:
            self.BtnPauseToggle.Content = "⏸️ Pausar Temporariamente"

    def on_pause_click(self, sender, e):
        enabled = bool(self.EnableAutoSave.IsChecked)
        if not enabled:
            forms.alert("O AutoSave está desligado na caixinha acima. Ligue-o para poder pausar/retomar o timer.")
            return

        if self.manager.is_paused:
            self.manager.resume()
        else:
            self.manager.pause()
            
        self.update_pause_btn()
        
    def on_save_config(self, sender, e):
        # Read from UI
        config.set("enabled", bool(self.EnableAutoSave.IsChecked))
        config.set("show_toast", bool(self.ShowToast.IsChecked))
        config.set("block_ui", False) # Forçado para false
        
        interval = 10
        if self.Radio1min.IsChecked: interval = 1
        elif self.Radio5min.IsChecked: interval = 5
        elif self.Radio15min.IsChecked: interval = 15
        elif self.Radio30min.IsChecked: interval = 30
        config.set("interval_minutes", interval)
        
        # Apply to Manager immediately
        try:
            manager = AutoSaveManager(__revit__)
            if config.get("enabled"):
                manager.start()
            else:
                manager.stop()
        except:
            pass
            
        self.Close()
        
def show_config():
    cur_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(cur_dir, 'config_window.xaml')
    w = ConfigWindow(xaml_path)
    w.ShowDialog()
