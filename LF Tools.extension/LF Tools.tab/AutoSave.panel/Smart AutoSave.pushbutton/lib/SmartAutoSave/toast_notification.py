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
        """Animação fade-in suave via WPF."""
        from System.Windows.Media.Animation import DoubleAnimation
        from System import TimeSpan
        anim = DoubleAnimation()
        anim.From = 0.0
        anim.To = 1.0
        anim.Duration = TimeSpan.FromMilliseconds(400)
        self.BeginAnimation(Window.OpacityProperty, anim)
            
    def trigger_fade_out(self, sender, e):
        """Chama fade-out e para o timer principal."""
        self.close_timer.Stop()
        from System.Windows.Media.Animation import DoubleAnimation
        from System import TimeSpan
        anim = DoubleAnimation()
        anim.From = self.Opacity
        anim.To = 0.0
        anim.Duration = TimeSpan.FromMilliseconds(400)
        anim.Completed += self._on_fade_out_completed
        self.BeginAnimation(Window.OpacityProperty, anim)
        
    def _on_fade_out_completed(self, sender, e):
        self.Close()

# Global instance reference to allow updating while saving
_current_toast = None

def show(title="Salvando projeto...", message="Aguarde...", icon="💾", duration=3):
    global _current_toast
    try:
        if _current_toast is not None:
            try:
                _current_toast.Close()
            except:
                pass
        cur_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(cur_dir, 'toast_notification.xaml')
        _current_toast = ToastNotificationWindow(xaml_path)
        _current_toast.show_toast(title, message, icon, duration)
        return _current_toast
    except Exception as e:
        print("Toast Error: " + str(e))
        return None
