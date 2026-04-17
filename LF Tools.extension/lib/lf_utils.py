#! python3
# -*- coding: utf-8 -*-
"""
lf_utils.py — Biblioteca compartilhada LF Tools
================================================
Utilitários comuns para todos os scripts pyRevit CPython da extensão.

Uso nos scripts:
    from lf_utils import DebugLogger, get_revit_context, safe_execution
    from lf_utils import make_warning_swallower, WPFWindowCPy

IMPORTANTE: Aplicar o monkeypatch de forms logo após o import:
    from lf_utils import patch_forms
    patch_forms(forms)          # forms já importado do pyrevit
"""

# ── Revit / pythonnet ─────────────────────────────────────────────────────────
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *                          # noqa: F401,F403
from Autodesk.Revit.DB.Electrical import *               # noqa: F401,F403
from Autodesk.Revit.UI import *                          # noqa: F401,F403
from Autodesk.Revit.UI.Selection import *                # noqa: F401,F403
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: F401

import System
import traceback
import time
import os
import re

# ── LAZY IMPORT: evita disparar events.py do pyRevit no momento do import ─────
# NÃO fazer: from pyrevit import script as _pyscript, EXEC_PARAMS  (causa bug)
# Os imports abaixo são feitos sob demanda nas funções que precisam deles.

# ── Windows.Forms (para dialogs CPython-safe) ─────────────────────────────────
clr.AddReference('System.Windows.Forms')
import System.Windows.Forms as _WF


# =============================================================================
#  LIMPEZA DE CACHE DO PYREVIT
# =============================================================================

def clear_pyrevit_cache(silent=False):
    """
    Remove o cache do pyRevit para o usuário atual.
    Equivalente a fechar o Revit e deletar a pasta de cache manualmente.

    Parâmetros:
        silent — se True, não exibe MessageBox ao final

    Retorna:
        True se limpou com sucesso, False se houve erro.

    Uso:
        from lf_utils import clear_pyrevit_cache
        clear_pyrevit_cache()

    ATENÇÃO: Recomenda-se reiniciar o Revit após executar esta função.
    """
    import shutil

    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        if not silent:
            _cpy_alert("Variável APPDATA não encontrada.", title="Limpar Cache")
        return False

    cache_paths = [
        os.path.join(appdata, "pyRevit", "cache"),
        os.path.join(appdata, "pyRevit-Master", "cache"),
    ]

    cleaned = []
    errors  = []

    for path in cache_paths:
        if os.path.isdir(path):
            try:
                shutil.rmtree(path)
                cleaned.append(path)
            except Exception as e:
                errors.append("{}: {}".format(path, e))

    if not silent:
        if cleaned:
            msg = "Cache removido com sucesso!\n\n" + "\n".join(cleaned)
            if errors:
                msg += "\n\nAtenção — falha ao remover:\n" + "\n".join(errors)
            msg += "\n\nReinicie o Revit para aplicar."
            _cpy_alert(msg, title="Limpar Cache pyRevit")
        elif errors:
            _cpy_alert(
                "Erro ao limpar cache:\n" + "\n".join(errors),
                title="Limpar Cache pyRevit"
            )
        else:
            _cpy_alert(
                "Nenhuma pasta de cache encontrada.\n"
                "(Pode já estar limpa ou o pyRevit está em local diferente.)",
                title="Limpar Cache pyRevit"
            )

    return len(errors) == 0


# =============================================================================
#  MONKEYPATCH DE FORMS
# =============================================================================
# Chamada OBRIGATÓRIA nos scripts que usam forms.alert / forms.CommandSwitchWindow.
# Exemplo:
#     from pyrevit import forms
#     from lf_utils import patch_forms
#     patch_forms(forms)

class _CPyCommandSwitchWindow:
    """Substituto CPython-safe para forms.CommandSwitchWindow."""
    @staticmethod
    def show(options, message="", title="Selecione"):
        form = _WF.Form()
        form.Text = title
        form.Width = 400
        form.Height = 100 + len(options) * 40
        form.StartPosition = _WF.FormStartPosition.CenterScreen
        form.FormBorderStyle = _WF.FormBorderStyle.FixedDialog

        lbl = _WF.Label()
        lbl.Text = message
        lbl.SetBounds(10, 10, 370, 30)
        form.Controls.Add(lbl)

        result = [None]
        y = 50

        def _make_handler(opt):
            def _handler(s, a):
                result[0] = opt
                form.Close()
            return _handler

        for opt in options:
            btn = _WF.Button()
            btn.Text = str(opt)
            btn.SetBounds(10, y, 370, 30)
            btn.Click += _make_handler(opt)
            form.Controls.Add(btn)
            y += 35

        form.ShowDialog()
        return result[0]


def _cpy_alert(msg, title="LF Tools"):
    _WF.MessageBox.Show(str(msg), str(title))


def _cpy_toast(msg, title="LF Tools"):
    _WF.MessageBox.Show(str(msg), str(title))


def _cpy_ask_for_string(prompt="", title="LF Tools", default=""):
    form = _WF.Form()
    form.Text = title
    form.Width = 400
    form.Height = 150
    form.StartPosition = _WF.FormStartPosition.CenterScreen

    lbl = _WF.Label()
    lbl.Text = prompt
    lbl.SetBounds(10, 10, 370, 40)

    txt = _WF.TextBox()
    txt.Text = default
    txt.SetBounds(10, 50, 370, 30)

    btn = _WF.Button()
    btn.Text = "OK"
    btn.SetBounds(150, 85, 80, 25)
    btn.DialogResult = _WF.DialogResult.OK

    form.Controls.Add(lbl)
    form.Controls.Add(txt)
    form.Controls.Add(btn)
    form.AcceptButton = btn

    return txt.Text if form.ShowDialog() == _WF.DialogResult.OK else None


def patch_forms(forms_module):
    """
    Aplica substituições CPython-safe nos helpers do módulo forms do pyRevit.
    Chamar logo após `from pyrevit import forms`.
    """
    forms_module.alert             = _cpy_alert
    forms_module.toast             = _cpy_toast
    forms_module.ask_for_string    = _cpy_ask_for_string
    forms_module.CommandSwitchWindow = _CPyCommandSwitchWindow


# =============================================================================
#  DEBUG LOGGER
# =============================================================================

class DebugLogger:
    """
    Logger de debug para scripts pyRevit CPython.
    Controlado pela flag DEBUG_MODE no topo de cada script.

    Métodos:
      dbg.section(titulo)    — separador visual de seção
      dbg.sub(titulo)        — sub-separador
      dbg.info(msg)          — informação geral
      dbg.debug(msg)         — detalhe técnico
      dbg.warn(msg)          — aviso (sempre impresso)
      dbg.error(msg)         — erro (sempre impresso)
      dbg.dump(label, obj)   — despeja repr de um objeto
      dbg.xyz(label, pt)     — imprime coordenadas XYZ formatadas
      dbg.timer_start(label) — inicia cronômetro nomeado
      dbg.timer_end(label)   — encerra e imprime tempo decorrido
      dbg.result(ok, msg)    — OK/FAIL visual
    """

    _SECTION = "=" * 60
    _SUB     = "-" * 40

    def __init__(self, enabled):
        self.enabled = enabled
        self._timers = {}

    def _print(self, prefix, msg):
        import sys
        line = "{} {}".format(prefix, msg)
        # Tenta stdout do pyRevit; se ScriptIO falhar, cai no __stdout__ bruto.
        try:
            print(line)
        except Exception:
            try:
                sys.__stdout__.write(line + "\n")
                sys.__stdout__.flush()
            except Exception:
                pass  # sem saída disponível — silencioso

    def section(self, titulo):
        if not self.enabled:
            return
        self._print("", "")
        self._print("", self._SECTION)
        self._print("", "  {}".format(titulo.upper()))
        self._print("", self._SECTION)

    def sub(self, titulo):
        if not self.enabled:
            return
        self._print("", "  {} {}".format(self._SUB, titulo))

    def info(self, msg):
        if self.enabled:
            self._print("[INFO ]", msg)

    def debug(self, msg):
        if self.enabled:
            self._print("[DEBUG]", msg)

    def warn(self, msg):
        self._print("[WARN ]", msg)

    def error(self, msg):
        self._print("[ERROR]", msg)

    def dump(self, label, obj):
        if self.enabled:
            self._print("[DUMP ]", "{} = {!r}".format(label, obj))

    def xyz(self, label, pt):
        if self.enabled:
            if pt is None:
                self._print("[XYZ  ]", "{} = None".format(label))
            else:
                self._print("[XYZ  ]", "{} = ({:.4f}, {:.4f}, {:.4f}) ft".format(
                    label, pt.X, pt.Y, pt.Z))

    def timer_start(self, label):
        self._timers[label] = time.time()
        if self.enabled:
            self._print("[TIMER]", "START  '{}'".format(label))

    def timer_end(self, label):
        if label in self._timers:
            elapsed = time.time() - self._timers.pop(label)
            if self.enabled:
                self._print("[TIMER]", "END    '{}' — {:.3f}s".format(label, elapsed))
        elif self.enabled:
            self._print("[TIMER]", "END    '{}' (sem start registrado)".format(label))

    def result(self, ok, msg):
        if self.enabled:
            tag = "[  OK ]" if ok else "[FAIL ]"
            self._print(tag, msg)


# =============================================================================
#  CONFIG JSON (substitui script.get_config() — sem dependência de pyrevit.script)
# =============================================================================

def get_script_config(script_path, defaults=None):
    """
    Lê configurações de um arquivo JSON ao lado do script.
    Substitui script.get_config() sem importar pyrevit.script.

    Uso:
        settings = get_script_config(__commandpath__, defaults={
            'chave': 'valor_padrão',
        })
        valor = settings.get('chave')
    """
    import json
    config_path = os.path.join(os.path.dirname(str(script_path)), 'config.json')
    data = dict(defaults or {})
    if os.path.isfile(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            data.update(saved)
        except Exception:
            pass
    return data


def save_script_config(script_path, settings):
    """
    Salva configurações em JSON ao lado do script.
    Substitui script.save_config() sem importar pyrevit.script.

    Uso:
        save_script_config(__commandpath__, settings)
    """
    import json
    config_path = os.path.join(os.path.dirname(str(script_path)), 'config.json')
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =============================================================================
#  CONTEXTO REVIT (lazy init — seguro para CPython/pythonnet)
# =============================================================================

def get_revit_context():
    """
    Retorna (uidoc, doc) inicializados de forma lazy.
    Lançar RuntimeError se o contexto não estiver disponível.

    Uso:
        uidoc, doc = get_revit_context()
    """
    try:
        _uidoc = __revit__.ActiveUIDocument   # __revit__ injetado pelo pyRevit
    except AttributeError:
        raise RuntimeError(
            "__revit__ não disponível. Execute o script pela Ribbon do Revit."
        )
    if _uidoc is None:
        raise RuntimeError("Nenhum documento aberto no Revit.")
    return _uidoc, _uidoc.Document


# =============================================================================
#  WARNING SWALLOWER (factory — evita Duplicate type name entre execuções)
# =============================================================================

def make_warning_swallower():
    """
    Cria e retorna uma instância de IFailuresPreprocessor que suprime warnings.
    Usar factory em vez de classe direta evita o erro
    'Duplicate type name within an assembly' na segunda execução.

    Uso:
        t.SetFailureHandlingOptions(
            t.GetFailureHandlingOptions()
             .SetClearAfterRollback(True)
             .SetForcedModalHandling(False)
             .SetFailuresPreprocessor(make_warning_swallower())
        )

    NOTA: EXEC_PARAMS é importado de forma lazy aqui para evitar o bug
    'interface takes exactly one argument' no events.py do pyRevit.
    """
    from Autodesk.Revit.DB import FailureSeverity, FailureProcessingResult
    # Import lazy — não fazer no topo do módulo!
    from pyrevit import EXEC_PARAMS

    class _WarningSwallower(IFailuresPreprocessor):
        __namespace__ = EXEC_PARAMS.exec_id   # OBRIGATÓRIO: evita Duplicate type name

        def PreprocessFailures(self, failuresAccessor):
            for f in failuresAccessor.GetFailureMessages():
                if f.GetSeverity() == FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(f)
            return FailureProcessingResult.Continue

    return _WarningSwallower()


# =============================================================================
#  SAFE EXECUTION WRAPPER
# =============================================================================

def safe_execution(fn, title="LF Tools", dbg=None):
    """
    Executa `fn()` com tratamento de erros padronizado.
    - OperationCanceledException: silencioso (usuário cancelou)
    - Qualquer outra exceção: exibe alert com traceback

    Parâmetros:
        fn    — callable sem argumentos (ex: execute_connection)
        title — título das janelas de erro
        dbg   — DebugLogger opcional; se fornecido, loga o crash
    """
    try:
        fn()
    except OperationCanceledException:
        pass
    except Exception as e:
        err_tb = traceback.format_exc()
        if dbg is not None:
            dbg.error("CRASH FATAL:\n{}".format(err_tb))
        _cpy_alert(
            "Erro: {}\n\n{}".format(e, err_tb) if (dbg and dbg.enabled) else str(e),
            title=title
        )


# =============================================================================
#  WPF WINDOW (CPython-safe — remove eventos do XAML antes de carregar)
# =============================================================================

class WPFWindowCPy:
    """
    Carregador de janelas WPF compatível com CPython/pythonnet.

    O XamlReader do pythonnet não resolve atributos de evento XAML
    (Click=, TextChanged=, etc.) — causam crash ao carregar.
    Esta classe remove esses atributos via regex ANTES de carregar,
    e expõe os elementos por nome via __getattr__.

    Uso:
        class MinhaJanela(WPFWindowCPy):
            def __init__(self, uiapp=None):
                super().__init__(
                    os.path.join(os.path.dirname(__commandpath__), 'ui.xaml'),
                    uiapp
                )
                # Reconectar eventos aqui (NÃO no XAML):
                self.btn_Ok.Click += self._on_ok

            def _on_ok(self, sender, args):
                self.Close()

    IMPORTANTE: Todos os eventos devem ser reconectados via += no __init__.
    Qualquer Click=, TextChanged= no XAML será removido e nunca chamado.
    """

    _XAML_EVENTS = re.compile(
        r'\s+(?:x:Class|Click|DoubleClick|'
        r'Mouse(?:Down|Up|Move|Enter|Leave|Wheel)|'
        r'Preview(?:Mouse(?:Down|Up|Move|LeftButtonDown|LeftButtonUp)|Key(?:Down|Up)|TextInput)|'
        r'Key(?:Down|Up)|TextInput|TextChanged|SelectionChanged|'
        r'Checked|Unchecked|Loaded|Unloaded|Clos(?:ing|ed)|Activated|Deactivated)'
        r'\s*=\s*(?:"[^"]*"|\'[^\']*\')'
    )

    def __init__(self, xaml_path, uiapp=None):
        from System.IO import StringReader
        from System.Windows.Markup import XamlReader
        import System.Xml

        with open(str(xaml_path), 'r', encoding='utf-8') as f:
            raw = f.read()
        clean = self._XAML_EVENTS.sub('', raw)
        rdr = System.Xml.XmlReader.Create(StringReader(clean))
        self._window = XamlReader.Load(rdr)

        if uiapp is not None:
            try:
                from System.Windows.Interop import WindowInteropHelper
                WindowInteropHelper(self._window).Owner = uiapp.MainWindowHandle
            except Exception:
                pass

    def __getattr__(self, name):
        if name.startswith('_'):
            raise AttributeError(name)
        el = object.__getattribute__(self, '_window').FindName(name)
        if el is not None:
            return el
        return getattr(object.__getattribute__(self, '_window'), name)

    def ShowDialog(self):
        return self._window.ShowDialog()

    def Show(self):
        self._window.Show()

    def Close(self):
        self._window.Close()
