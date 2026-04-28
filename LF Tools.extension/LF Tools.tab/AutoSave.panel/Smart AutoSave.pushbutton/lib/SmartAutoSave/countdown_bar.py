# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
import ctypes

from pyrevit import forms

_current_bar = None


class _RECT(ctypes.Structure):
    _fields_ = [
        ('left',   ctypes.c_long),
        ('top',    ctypes.c_long),
        ('right',  ctypes.c_long),
        ('bottom', ctypes.c_long),
    ]


def _get_window_rect(hwnd):
    rect = _RECT()
    ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(rect))
    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top


class CountdownBarWindow(forms.WPFWindow):
    def __init__(self, xaml_path, hwnd):
        forms.WPFWindow.__init__(self, xaml_path)
        self._hwnd = hwnd
        self._position()

    def _position(self):
        try:
            x, y, w, h = _get_window_rect(self._hwnd)
            self.Left  = float(x)
            self.Top   = float(y)
            self.Width = float(max(w, 100))
        except Exception as e:
            print("CountdownBar position error: " + str(e))

    def set_fraction(self, fraction):
        try:
            new_w = max(0.0, min(1.0, fraction)) * self.Width
            from System.Windows import FrameworkElement
            from System.Windows.Media.Animation import DoubleAnimation
            from System import TimeSpan
            anim = DoubleAnimation()
            anim.To = new_w
            anim.Duration = TimeSpan.FromMilliseconds(380)
            self.ProgressFill.BeginAnimation(FrameworkElement.WidthProperty, anim)
        except:
            try:
                self.ProgressFill.Width = max(0.0, min(1.0, fraction)) * self.Width
            except:
                pass


def show(hwnd, total_seconds):
    global _current_bar
    hide()
    try:
        cur_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(cur_dir, 'countdown_bar.xaml')
        _current_bar = CountdownBarWindow(xaml_path, hwnd)
        _current_bar.Show()
        _current_bar.set_fraction(1.0)
    except Exception as e:
        print("CountdownBar show error: " + str(e))


def update(remaining, total):
    global _current_bar
    if _current_bar is None:
        return
    try:
        frac = float(remaining) / float(total) if total > 0 else 0.0
        _current_bar.set_fraction(frac)
    except:
        pass


def hide():
    global _current_bar
    if _current_bar is not None:
        try:
            _current_bar.Close()
        except:
            pass
        _current_bar = None
