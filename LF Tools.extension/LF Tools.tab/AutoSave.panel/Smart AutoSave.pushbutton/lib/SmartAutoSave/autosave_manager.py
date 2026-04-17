# -*- coding: utf-8 -*-
import os
import io
import time
import datetime
import clr
from System import AppDomain
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
from System.Windows.Threading import DispatcherTimer, DispatcherPriority
from System import TimeSpan
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from Autodesk.Revit.DB import SaveAsOptions

from SmartAutoSave.config_manager import config
import SmartAutoSave.toast_notification as toast

from pyrevit import forms

_APPDOMAIN_TIMER_KEY = "LFTools_AutoSaveTimer"
_RETRY_DELAY_SECONDS  = 30   # tempo de retry quando Revit está ocupado


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

        self.uiapp        = uiapp
        self.timer        = None
        self._retry_timer = None
        self.is_paused    = False

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.log_path = os.path.join(desktop, "pyRevit_AutoSave_Log.txt")

        self.save_handler = AutoSaveHandler(self)
        self.save_event   = ExternalEvent.Create(self.save_handler)

        self.initialized = True

        # BUG 1 — Para o timer anterior que ficou no AppDomain de um
        # reload anterior ANTES de criar um novo, evitando timers duplicados.
        self._stop_appdomain_timer()

        if config.get("enabled"):
            self.start()

    # ------------------------------------------------------------------ #
    # Controle do timer principal                                          #
    # ------------------------------------------------------------------ #

    def start(self):
        """Inicia ou reinicia o timer com o intervalo configurado."""
        # Para qualquer timer anterior (local ou herdado do AppDomain)
        self._stop_appdomain_timer()
        if self.timer:
            self.timer.Stop()

        interval = config.get("interval_minutes") or 10

        self.timer = DispatcherTimer()
        self.timer.Interval = TimeSpan.FromMinutes(interval)
        self.timer.Tick    += self.on_timer_tick
        self.timer.Start()

        # Persiste referência no AppDomain para o próximo reload poder pará-lo
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_TIMER_KEY, self.timer)

    def stop(self):
        """Para o autosave completamente."""
        self._stop_appdomain_timer()
        if self.timer:
            self.timer.Stop()
        self._cancel_retry()

    def pause(self):
        self.is_paused = True
        self._cancel_retry()

    def resume(self):
        self.is_paused = False
        if config.get("enabled"):
            self.start()

    def trigger_save_now(self):
        """Força salvamento manual imediato."""
        if self.is_paused:
            return
        self.save_event.Raise()

    def on_timer_tick(self, sender, args):
        if self.is_paused or not config.get("enabled"):
            return
        self.save_event.Raise()

    # ------------------------------------------------------------------ #
    # BUG 3 — Retry quando Revit está ocupado                             #
    # ------------------------------------------------------------------ #

    def schedule_retry(self):
        """Agenda nova tentativa de save em _RETRY_DELAY_SECONDS segundos."""
        self._cancel_retry()
        self._retry_timer = DispatcherTimer()
        self._retry_timer.Interval = TimeSpan.FromSeconds(_RETRY_DELAY_SECONDS)
        self._retry_timer.Tick    += self._on_retry_tick
        self._retry_timer.Start()

    def _on_retry_tick(self, sender, args):
        sender.Stop()
        self._retry_timer = None
        if not self.is_paused and config.get("enabled"):
            self.save_event.Raise()

    def _cancel_retry(self):
        if self._retry_timer:
            try:
                self._retry_timer.Stop()
            except:
                pass
            self._retry_timer = None

    # ------------------------------------------------------------------ #
    # Helpers internos                                                     #
    # ------------------------------------------------------------------ #

    def _stop_appdomain_timer(self):
        """Para o timer guardado no AppDomain (sobrevive a reloads do engine)."""
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

        # BUG 2 — Guarda contra crash quando não há documento aberto
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
                return

            # BUG 3 — Se Revit está ocupado, agenda retry em vez de simplesmente desistir
            if not self._is_safe_to_save(doc):
                self.manager.log(u"⚠️ Revit ocupado — retry em {}s.".format(_RETRY_DELAY_SECONDS))
                self.manager.schedule_retry()
                return

            # Melhoria 4 — Pula salvamento se não houve alterações desde o último save
            if not doc.IsModified:
                self.manager.log(u"— Sem alterações, salvamento ignorado.")
                return

            show_toast = config.get("show_toast", True)
            is_cloud   = self._is_cloud_model(doc)
            toast_inst = None

            if show_toast:
                msg      = "Salvando na nuvem (pode demorar)..." if is_cloud else "Aguarde..."
                duration = 30 if is_cloud else 10
                toast_inst = toast.show("Salvando projeto...", msg, "💾", duration)

            doc.Save()

            elapsed  = time.time() - start_time
            time_str = datetime.datetime.now().strftime("%H:%M")
            self.manager.log(u"✅ Salvo em {:.1f}s".format(elapsed))

            if show_toast and toast_inst:
                toast_inst.ToastTitle.Text   = "✅ Projeto salvo"
                toast_inst.ToastMessage.Text = "Salvo às " + time_str
                toast_inst.ToastIcon.Text    = "✅"

            # Reinicia o timer a partir do momento do save bem-sucedido
            self.manager.start()

        except Exception as e:
            self.manager.log(u"❌ Erro ao salvar: " + str(e))
            if 'toast_inst' in locals() and toast_inst:
                toast_inst.ToastTitle.Text   = "⚠️ Erro ao salvar"
                toast_inst.ToastMessage.Text = "Tente novamente mais tarde"
                toast_inst.ToastIcon.Text    = "⚠️"

    def _is_safe_to_save(self, doc):
        """Seguro salvar = nenhuma transação aberta no documento."""
        if not config.get("wait_safe_state"):
            return True
        # IsModifiable == True → transação aberta → NÃO é seguro
        return not doc.IsModifiable

    def _is_cloud_model(self, doc):
        try:
            return doc.IsModelInCloud
        except:
            return False

    def GetName(self):
        return "AutoSaveHandler"
