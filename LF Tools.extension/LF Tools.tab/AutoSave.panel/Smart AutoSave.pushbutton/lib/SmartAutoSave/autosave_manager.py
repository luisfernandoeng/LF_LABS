# -*- coding: utf-8 -*-
import os
import io
import time
import datetime
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase') # Required for Dispatcher
clr.AddReference('System.Windows.Forms')
from System.Windows.Threading import DispatcherTimer, DispatcherPriority
from System import TimeSpan
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from Autodesk.Revit.DB import SaveAsOptions

# Import parts of the lib
from SmartAutoSave.config_manager import config
import SmartAutoSave.toast_notification as toast

# Import to load XAML
from pyrevit import forms

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
            
        self.uiapp = uiapp
        self.timer = None
        self.overlay_window = None
        self.is_paused = False
        
        # Desktop path for logs
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.log_path = os.path.join(desktop, "pyRevit_AutoSave_Log.txt")
        
        # Load logic
        self.save_handler = AutoSaveHandler(self)
        self.save_event = ExternalEvent.Create(self.save_handler)
        
        self.initialized = True
        
        if config.get("enabled"):
            self.start()

    def start(self):
        """Starts or restarts the timer based on config."""
        if self.timer:
            self.timer.Stop()
            
        interval = config.get("interval_minutes")
        if not interval or interval <= 0:
            interval = 10
            
        self.timer = DispatcherTimer()
        self.timer.Interval = TimeSpan.FromMinutes(interval)
        self.timer.Tick += self.on_timer_tick
        self.timer.Start()

    def stop(self):
        """Stops the autosave entirely."""
        if self.timer:
            self.timer.Stop()
            
    def pause(self):
        self.is_paused = True
        
    def resume(self):
        self.is_paused = False
        # Reset timer immediately
        if config.get("enabled"):
            self.start()

    def trigger_save_now(self):
        """Manually triggers the event."""
        if self.is_paused:
            return
        self.save_event.Raise()

    def on_timer_tick(self, sender, args):
        """Tick from the dispatcher."""
        if self.is_paused or not config.get("enabled"):
            return
        self.save_event.Raise()

    def log(self, message):
        """Logs history to Desktop if enabled."""
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
    
    def is_cloud_model(self, doc):
        """Verifica se documento está na nuvem (ACC)."""
        try:
            return doc.IsModelInCloud
        except:
            return False

    def Execute(self, uiapp):
        start_time = time.time()
        doc = uiapp.ActiveUIDocument.Document
        
        try:
            # Check if project is saved at least once
            if not doc.PathName:
                self.manager.log(u"⚠️ Salvamento cancelado: O projeto ainda não foi salvo em disco.")
                return

            if not self.is_safe_to_save(doc):
                self.manager.log(u"⚠️ Salvamento adiado (Revit ocupado).")
                return

            show_toast = config.get("show_toast", True)

            is_cloud = self.is_cloud_model(doc)
            toast_inst = None
            
            if show_toast:
                msg = "Salvando na nuvem (pode demorar)..." if is_cloud else "Aguarde..."
                duration = 30 if is_cloud else 10
                toast_inst = toast.show("Salvando projeto...", msg, "💾", duration)

            # Revit Native Save (handles backups natively)
            doc.Save()
            
            elapsed = time.time() - start_time
            time_str = datetime.datetime.now().strftime("%H:%M")
            msg_log = u"Salvamento concluído ({:.1f}s)".format(elapsed)
            
            self.manager.log(u"✅ " + msg_log)
            
            if show_toast and toast_inst:
                toast_inst.ToastTitle.Text = "✅ Projeto salvo"
                toast_inst.ToastMessage.Text = "Salvo às " + time_str
                toast_inst.ToastIcon.Text = "✅"
            
            # Restart timer
            self.manager.start()

        except Exception as e:
            self.manager.log(u"❌ Erro ao salvar: " + str(e))
            if 'toast_inst' in locals() and toast_inst:
                toast_inst.ToastTitle.Text = "⚠️ Erro ao salvar"
                toast_inst.ToastMessage.Text = "Tente novamente mais tarde"
                toast_inst.ToastIcon.Text = "⚠️"
    
    def is_safe_to_save(self, doc):
        if not config.get("wait_safe_state"):
            return True
        if doc.IsModifiable:
            return False
        return True
    
    def GetName(self):
        return "AutoSaveHandler"
