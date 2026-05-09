# -*- coding: utf-8 -*-
import os
import clr
import ctypes
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
from System.Windows import Window
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Threading import DispatcherTimer
from System import TimeSpan
import System.Windows.Forms as WinForms

from pyrevit import forms


class ToastNotificationWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, hwnd=None):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.MouseLeftButtonDown += self.on_click_close
        self.Topmost = False
        self.ShowActivated = False
        self._hwnd = hwnd
        self.SourceInitialized += self._on_source_initialized
        self.position_bottom_right()

    def on_click_close(self, sender, e):
        self.trigger_fade_out(None, None)

    def _on_source_initialized(self, sender, e):
        try:
            hwnd = WindowInteropHelper(self).Handle.ToInt32()
            GWL_EXSTYLE = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style = style | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        except:
            pass

    def position_bottom_right(self):
        """Posiciona o toast no monitor onde a janela do Revit esta."""
        try:
            if self._hwnd:
                screen = WinForms.Screen.FromHandle(self._hwnd)
            else:
                screen = WinForms.Screen.PrimaryScreen

            area = screen.WorkingArea
            scale = self._get_dpi_scale()
            margin = 20.0

            self.Left = (float(area.Right) - (float(self.Width) * scale) - margin) / scale
            self.Top  = (float(area.Bottom) - (float(self.Height) * scale) - margin) / scale
        except:
            try:
                from System.Windows import SystemParameters
                self.Left = SystemParameters.WorkArea.Right  - self.Width  - 20
                self.Top  = SystemParameters.WorkArea.Bottom - self.Height - 20
            except:
                pass

    def _get_dpi_scale(self):
        try:
            hwnd = int(self._hwnd) if self._hwnd else 0
            if hwnd and hasattr(ctypes.windll.user32, "GetDpiForWindow"):
                dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
                if dpi:
                    return float(dpi) / 96.0
        except:
            pass
        return 1.0

    def show_toast(self, title, message, icon=u"💾", duration=3):
        self.ToastTitle.Text   = title
        self.ToastMessage.Text = message
        self.ToastIcon.Text    = icon

        self.Show()
        self._restore_revit_focus()

        self.Opacity = 0
        self.fade_in()

        self.close_timer = DispatcherTimer()
        self.close_timer.Interval = TimeSpan.FromSeconds(duration)
        self.close_timer.Tick += self.trigger_fade_out
        self.close_timer.Start()

    def _restore_revit_focus(self):
        try:
            if self._hwnd and not ctypes.windll.user32.IsIconic(int(self._hwnd)):
                ctypes.windll.user32.SetForegroundWindow(int(self._hwnd))
        except:
            pass

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
