# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationFramework')
from System.Windows import Window
from System.Windows.Threading import DispatcherTimer
from System import TimeSpan

from pyrevit import forms

class ToastNotificationWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.Topmost = True
        self.position_bottom_right()
        
    def position_bottom_right(self):
        """Posiciona toast no canto inferior direito."""
        from System.Windows import SystemParameters
        
        screen_width = SystemParameters.WorkArea.Width
        screen_height = SystemParameters.WorkArea.Height
        
        # 320x90 is the approximate size of the Toast
        self.Left = screen_width - 320 - 20
        self.Top = screen_height - 90 - 20

    def show_toast(self, title, message, icon="💾", duration=3):
        """Mostra toast com auto-close."""
        self.ToastTitle.Text = title
        self.ToastMessage.Text = message
        self.ToastIcon.Text = icon
        
        self.Show()
        
        # Fade-in
        self.Opacity = 0
        self.fade_in()
        
        # Auto-close após duration segundos
        self.close_timer = DispatcherTimer()
        self.close_timer.Interval = TimeSpan.FromSeconds(duration)
        self.close_timer.Tick += self.trigger_fade_out
        self.close_timer.Start()
        
    def fade_in(self):
        """Animação fade-in suave."""
        self.fade_in_timer = DispatcherTimer()
        self.fade_in_timer.Interval = TimeSpan.FromMilliseconds(20)
        self.fade_in_timer.Tick += self._fade_in_tick
        self.fade_in_timer.Start()
        
    def _fade_in_tick(self, sender, e):
        self.Opacity += 0.1
        if self.Opacity >= 1.0:
            self.Opacity = 1.0
            self.fade_in_timer.Stop()
            
    def trigger_fade_out(self, sender, e):
        """Chama fade-out e para o timer principal."""
        self.close_timer.Stop()
        self.fade_out_timer = DispatcherTimer()
        self.fade_out_timer.Interval = TimeSpan.FromMilliseconds(20)
        self.fade_out_timer.Tick += self._fade_out_tick
        self.fade_out_timer.Start()
        
    def _fade_out_tick(self, sender, e):
        self.Opacity -= 0.1
        if self.Opacity <= 0:
            self.fade_out_timer.Stop()
            self.Close()

# Global instance reference to allow updating while saving
_current_toast = None

def show(title="Salvando projeto...", message="Aguarde...", icon="💾", duration=3):
    global _current_toast
    try:
        cur_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(cur_dir, 'toast_notification.xaml')
        _current_toast = ToastNotificationWindow(xaml_path)
        _current_toast.show_toast(title, message, icon, duration)
        return _current_toast
    except Exception as e:
        print("Toast Error: " + str(e))
        return None
