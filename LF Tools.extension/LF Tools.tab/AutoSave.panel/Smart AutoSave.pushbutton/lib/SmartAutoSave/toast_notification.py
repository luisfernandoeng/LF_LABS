# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')
from System.Windows import Window
from System.Windows.Threading import DispatcherTimer
from System import TimeSpan
import System.Windows.Forms as WinForms

from pyrevit import forms


class ToastNotificationWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, hwnd=None):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.MouseLeftButtonDown += self.on_click_close
        self.Topmost = True
        self._hwnd = hwnd
        self.position_bottom_right()

    def on_click_close(self, sender, e):
        self.trigger_fade_out(None, None)

    def position_bottom_right(self):
        """Posiciona toast no canto inferior direito da tela principal (DPI-aware)."""
        try:
            from System.Windows import SystemParameters
            right  = SystemParameters.WorkArea.Right
            bottom = SystemParameters.WorkArea.Bottom
            self.Left = right  - 320 - 20
            self.Top  = bottom - 80 - 20
        except:
            pass

    def show_toast(self, title, message, icon=u"💾", duration=3):
        self.ToastTitle.Text   = title
        self.ToastMessage.Text = message
        self.ToastIcon.Text    = icon

        self.Show()

        self.Opacity = 0
        self.fade_in()

        self.close_timer = DispatcherTimer()
        self.close_timer.Interval = TimeSpan.FromSeconds(duration)
        self.close_timer.Tick += self.trigger_fade_out
        self.close_timer.Start()

    def restart_close_timer(self, seconds):
        """Reinicia o timer de auto-close com nova duração (usado após save bem-sucedido)."""
        try:
            self.close_timer.Stop()
        except:
            pass
        self.close_timer = DispatcherTimer()
        self.close_timer.Interval = TimeSpan.FromSeconds(seconds)
        self.close_timer.Tick += self.trigger_fade_out
        self.close_timer.Start()

    def fade_in(self):
        from System.Windows.Media.Animation import DoubleAnimation
        anim = DoubleAnimation()
        anim.From = 0.0
        anim.To   = 1.0
        anim.Duration = TimeSpan.FromMilliseconds(400)
        self.BeginAnimation(self.OpacityProperty, anim)

    def trigger_fade_out(self, sender, e):
        try: self.close_timer.Stop()
        except: pass
        from System.Windows.Media.Animation import DoubleAnimation
        anim = DoubleAnimation()
        anim.From = self.Opacity
        anim.To   = 0.0
        anim.Duration = TimeSpan.FromMilliseconds(400)
        anim.Completed += self._on_fade_out_completed
        self.BeginAnimation(self.OpacityProperty, anim)

    def _on_fade_out_completed(self, sender, e):
        self.Close()


_current_toast = None


def show(title=u"Salvando projeto...", message=u"Aguarde...", icon=u"💾", duration=3, hwnd=None):
    global _current_toast
    try:
        if _current_toast is not None:
            try:
                _current_toast.Close()
            except:
                pass
        cur_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(cur_dir, 'toast_notification.xaml')
        _current_toast = ToastNotificationWindow(xaml_path, hwnd=hwnd)
        _current_toast.show_toast(title, message, icon, duration)
        return _current_toast
    except Exception as e:
        print("Toast Error: " + str(e))
        return None
