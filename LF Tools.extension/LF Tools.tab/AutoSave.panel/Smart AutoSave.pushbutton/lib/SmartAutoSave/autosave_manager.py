# -*- coding: utf-8 -*-
import os
import io
import time
import datetime
import clr
import ctypes
from System import AppDomain
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
from System.Windows.Threading import DispatcherTimer
from System import TimeSpan
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from SmartAutoSave.config_manager import config
import SmartAutoSave.toast_notification as toast
import SmartAutoSave.countdown_bar as countdown_bar

_APPDOMAIN_TIMER_KEY = "LFTools_AutoSaveTimer"
_RETRY_DELAY_SECONDS = 30
_USER_IDLE_SECONDS_BEFORE_SAVE = 8


class AutoSaveManager(object):
    _instance = None

    def __new__(cls, uiapp):
        if cls._instance is None:
            cls._instance = super(AutoSaveManager, cls).__new__(cls)
            cls._instance.initialized = False
        return cls._instance

    def __init__(self, uiapp):
        if self.initialized:
            return

        self.uiapp             = uiapp
        self.timer             = None
        self._retry_timer      = None
        self._countdown_timer     = None
        self._countdown_remaining = 0
        self._countdown_total     = 5
        self._countdown_toast     = None
        self.is_paused         = False

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.log_path = os.path.join(desktop, "pyRevit_AutoSave_Log.txt")

        self.save_handler = AutoSaveHandler(self)
        self.save_event   = ExternalEvent.Create(self.save_handler)

        self.initialized = True

        self._stop_appdomain_timer()

        if config.get("enabled"):
            self.start()

    # ------------------------------------------------------------------ #
    # Timer principal                                                      #
    # ------------------------------------------------------------------ #

    def start(self):
        self._stop_appdomain_timer()
        if self.timer:
            self.timer.Stop()

        interval = config.get("interval_minutes") or 10

        self.timer = DispatcherTimer()
        self.timer.Interval = TimeSpan.FromMinutes(interval)
        self.timer.Tick    += self.on_timer_tick
        self.timer.Start()

        AppDomain.CurrentDomain.SetData(_APPDOMAIN_TIMER_KEY, self.timer)

    def stop(self):
        self._stop_appdomain_timer()
        if self.timer:
            self.timer.Stop()
        self._cancel_retry()
        self._cancel_countdown()

    def pause(self):
        self.is_paused = True
        self._cancel_retry()
        self._cancel_countdown()

    def resume(self):
        self.is_paused = False
        if config.get("enabled"):
            self.start()

    def trigger_save_now(self):
        """Força salvamento imediato (sem countdown)."""
        if self.is_paused:
            return
        self._cancel_countdown()
        self.save_event.Raise()

    def on_timer_tick(self, sender, args):
        if self.is_paused or not config.get("enabled"):
            return
        if self._countdown_timer is not None:
            return  # já está em countdown
        if not self._is_revit_available_for_autosave():
            self.log(u"Revit minimizado ou fora de foco. AutoSave adiado.")
            return
        if self._is_user_active():
            self.log(u"Usuario ativo no Revit. AutoSave adiado.")
            return
        self._start_countdown()

    # ------------------------------------------------------------------ #
    # Countdown 5-4-3-2-1                                                  #
    # ------------------------------------------------------------------ #

    def _get_hwnd(self):
        try:
            hwnd = int(self.uiapp.MainWindowHandle)
            return hwnd if hwnd != 0 else None
        except:
            return None

    def _is_revit_available_for_autosave(self):
        hwnd = self._get_hwnd()
        if hwnd is None:
            return False
        try:
            if ctypes.windll.user32.IsIconic(int(hwnd)):
                return False
            foreground = ctypes.windll.user32.GetForegroundWindow()
            if int(foreground) != int(hwnd):
                return False
        except:
            pass
        return True

    def _get_idle_seconds(self):
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [
                ('cbSize', ctypes.c_uint),
                ('dwTime', ctypes.c_uint),
            ]

        try:
            info = LASTINPUTINFO()
            info.cbSize = ctypes.sizeof(LASTINPUTINFO)
            if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info)):
                tick = ctypes.windll.kernel32.GetTickCount()
                elapsed_ms = (int(tick) - int(info.dwTime)) & 0xFFFFFFFF
                return float(elapsed_ms) / 1000.0
        except:
            pass
        return None

    def _is_user_active(self):
        idle_seconds = self._get_idle_seconds()
        if idle_seconds is None:
            return False
        return idle_seconds < _USER_IDLE_SECONDS_BEFORE_SAVE

    def _start_countdown(self):
        if not self._is_revit_available_for_autosave() or self._is_user_active():
            return

        secs = config.get("countdown_seconds", 5)
        if secs <= 0:
            self.save_event.Raise()
            return

        self._countdown_total     = secs
        self._countdown_remaining = secs
        self._show_countdown_ui()

        self._countdown_timer = DispatcherTimer()
        self._countdown_timer.Interval = TimeSpan.FromSeconds(1)
        self._countdown_timer.Tick    += self._on_countdown_tick
        self._countdown_timer.Start()

    def _on_countdown_tick(self, sender, args):
        if not self._is_revit_available_for_autosave() or self._is_user_active():
            self._cancel_countdown()
            self.schedule_retry()
            self.log(u"AutoSave adiado durante countdown.")
            return

        self._countdown_remaining -= 1
        if self._countdown_remaining <= 0:
            sender.Stop()
            self._countdown_timer = None
            self._pre_save_ui()
            self.save_event.Raise()
        else:
            self._update_countdown_ui()

    def _cancel_countdown(self):
        if self._countdown_timer:
            try:
                self._countdown_timer.Stop()
            except:
                pass
            self._countdown_timer = None
        countdown_bar.hide()
        if self._countdown_toast:
            try:
                self._countdown_toast.Close()
            except:
                pass
            self._countdown_toast = None

    # ------------------------------------------------------------------ #
    # UI do countdown                                                      #
    # ------------------------------------------------------------------ #

    def _show_countdown_ui(self):
        hwnd  = self._get_hwnd()
        n     = self._countdown_remaining
        total = self._countdown_total

        if config.get("show_toast", True):
            self._countdown_toast = toast.show(
                u"⏱ Salvamento Automático",
                u"Salvando em {}s...".format(n),
                u"⏱",
                duration=total + 60,
                hwnd=hwnd,
            )

        if hwnd is not None:
            countdown_bar.show(hwnd, total)

    def _update_countdown_ui(self):
        n     = self._countdown_remaining
        total = self._countdown_total

        if self._countdown_toast:
            try:
                self._countdown_toast.ToastMessage.Text = u"Salvando em {}s...".format(n)
            except:
                pass

        countdown_bar.update(n, total)

    def _pre_save_ui(self):
        """Transição: countdown zerou, agora salva de verdade."""
        countdown_bar.hide()
        if self._countdown_toast:
            try:
                self._countdown_toast.Opacity = 1.0
                self._countdown_toast.ToastTitle.Text   = u"Salvando projeto..."
                self._countdown_toast.ToastMessage.Text = u"Aguarde..."
                self._countdown_toast.ToastIcon.Text    = u"💾"
                self._countdown_toast.fade_in()
            except:
                pass

    # ------------------------------------------------------------------ #
    # Retry quando Revit está ocupado                                      #
    # ------------------------------------------------------------------ #

    def schedule_retry(self):
        self._cancel_retry()
        self._retry_timer = DispatcherTimer()
        self._retry_timer.Interval = TimeSpan.FromSeconds(_RETRY_DELAY_SECONDS)
        self._retry_timer.Tick    += self._on_retry_tick
        self._retry_timer.Start()

    def _on_retry_tick(self, sender, args):
        sender.Stop()
        self._retry_timer = None
        if not self.is_paused and config.get("enabled"):
            if self._is_revit_available_for_autosave() and not self._is_user_active():
                self.save_event.Raise()
            else:
                self.schedule_retry()

    def _cancel_retry(self):
        if self._retry_timer:
            try:
                self._retry_timer.Stop()
            except:
                pass
            self._retry_timer = None

    # ------------------------------------------------------------------ #
    # Helpers                                                              #
    # ------------------------------------------------------------------ #

    def _stop_appdomain_timer(self):
        try:
            old = AppDomain.CurrentDomain.GetData(_APPDOMAIN_TIMER_KEY)
            if old is not None:
                old.Stop()
                AppDomain.CurrentDomain.SetData(_APPDOMAIN_TIMER_KEY, None)
        except:
            pass

    def log(self, message):
        if not config.get("log_history", False):
            return
        try:
            with io.open(self.log_path, "a", encoding="utf-8") as f:
                now = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
                f.write(u"{} {}\n".format(now, message))
        except:
            pass


class AutoSaveHandler(IExternalEventHandler):
    def __init__(self, manager):
        self.manager = manager

    def Execute(self, uiapp):
        start_time = time.time()
        show_toast = config.get("show_toast", True)
        toast_inst = None

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self.manager.log(u"⚠️ Nenhum documento ativo. Salvamento ignorado.")
            return
        doc = uidoc.Document
        if doc is None or doc.IsFamilyDocument:
            return

        try:
            if not doc.PathName:
                self.manager.log(u"⚠️ Projeto ainda não salvo em disco. Salvamento ignorado.")
                if self.manager._countdown_toast:
                    try: self.manager._countdown_toast.Close()
                    except: pass
                self.manager._countdown_toast = None
                return

            if not self._is_safe_to_save(doc):
                self.manager.log(u"⚠️ Revit ocupado — retry em {}s.".format(_RETRY_DELAY_SECONDS))
                self.manager.schedule_retry()
                return

            if not doc.IsModified:
                self.manager.log(u"— Sem alterações, salvamento ignorado.")
                if self.manager._countdown_toast:
                    try: self.manager._countdown_toast.Close()
                    except: pass
                self.manager._countdown_toast = None
                return

            is_cloud = self._is_cloud_model(doc)

            # Reaproveita o toast do countdown; senão cria um novo (trigger_save_now)
            toast_inst = self.manager._countdown_toast
            if show_toast and toast_inst is None:
                try:
                    hwnd = int(uiapp.MainWindowHandle)
                    hwnd = hwnd if hwnd != 0 else None
                except:
                    hwnd = None
                msg      = u"Salvando na nuvem (pode demorar)..." if is_cloud else u"Aguarde..."
                duration = 30 if is_cloud else 10
                toast_inst = toast.show(u"Salvando projeto...", msg, u"💾", duration, hwnd=hwnd)

            doc.Save()

            elapsed  = time.time() - start_time
            time_str = datetime.datetime.now().strftime("%H:%M")
            self.manager.log(u"✅ Salvo em {:.1f}s".format(elapsed))

            if show_toast and toast_inst:
                try:
                    toast_inst.Opacity = 1.0
                    toast_inst.ToastTitle.Text   = u"✅ Projeto salvo"
                    toast_inst.ToastMessage.Text = u"Salvo às " + time_str
                    toast_inst.ToastIcon.Text    = u"✅"
                    toast_inst.fade_in()
                    toast_inst.restart_close_timer(5)
                except:
                    pass

            self.manager._countdown_toast = None
            self.manager.start()

        except Exception as e:
            self.manager.log(u"❌ Erro ao salvar: " + str(e))
            effective_toast = toast_inst or self.manager._countdown_toast
            if show_toast and effective_toast:
                try:
                    effective_toast.Opacity = 1.0
                    effective_toast.ToastTitle.Text   = u"⚠️ Erro ao salvar"
                    effective_toast.ToastMessage.Text = u"Tente novamente mais tarde"
                    effective_toast.ToastIcon.Text    = u"⚠️"
                    effective_toast.fade_in()
                    effective_toast.restart_close_timer(5)
                except:
                    pass
            self.manager._countdown_toast = None

    def _is_safe_to_save(self, doc):
        if not config.get("wait_safe_state"):
            return True
        if not self.manager._is_revit_available_for_autosave():
            return False
        if self.manager._is_user_active():
            return False
        return not doc.IsModifiable

    def _is_cloud_model(self, doc):
        try:
            return doc.IsModelInCloud
        except:
            return False

    def GetName(self):
        return "AutoSaveHandler"
