#! python3
# -*- coding: utf-8 -*-
"""Transferir Circuitos - Desconecta circuitos do quadro de origem e reconecta no destino especificado."""
__title__ = "Transferir\nCircuitos"
__author__ = "Luís Fernando"

import clr
import os
import re
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')

import System
import System.Windows.Forms as _WF
from System.IO import StringReader
from System.Windows.Markup import XamlReader
import System.Xml
from System.Windows import Visibility, Thickness
from System.Windows.Controls import CheckBox
from System.Windows.Media import SolidColorBrush
from System.Windows.Media import Color as WpfColor
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction, TransactionStatus
)

_BUNDLE_DIR = os.path.dirname(__file__)

# ==================== CPython Compat ====================

def _alert(msg, title="LF Tools", yes=False, no=False, exitscript=False, **kw):
    try:
        if yes and no:
            r = _WF.MessageBox.Show(str(msg), str(title), _WF.MessageBoxButtons.YesNo)
            ans = (r == _WF.DialogResult.Yes)
            if exitscript and not ans:
                import sys; sys.exit(0)
            return ans
        _WF.MessageBox.Show(str(msg), str(title))
    except Exception:
        print("{}: {}".format(title, msg))
    if exitscript:
        import sys; sys.exit(0)

# (no monkey-patch needed — _alert used directly)


class _WPFWindowCPy:
    """CPython drop-in para pyrevit.forms.WPFWindow."""

    _XAML_EVENTS = re.compile(
        r'\s+(?:x:Class|'
        r'Click|DoubleClick|'
        r'Mouse(?:Down|Up|Move|Enter|Leave|Wheel)|'
        r'Preview(?:Mouse(?:Down|Up|Move|LeftButtonDown|LeftButtonUp)|'
        r'Key(?:Down|Up)|TextInput)|'
        r'Key(?:Down|Up)|TextInput|TextChanged|SelectionChanged|'
        r'SelectedItemChanged|ValueChanged|ScrollChanged|'
        r'Got(?:Focus|KeyboardFocus)|Lost(?:Focus|KeyboardFocus)|'
        r'Checked|Unchecked|Indeterminate|'
        r'Loaded|Unloaded|Initialized|'
        r'Clos(?:ing|ed)|Activated|Deactivated|'
        r'SizeChanged|LayoutUpdated|ContentRendered|'
        r'Drag(?:Enter|Leave|Over)|Drop|'
        r'ContextMenu(?:Opening|Closing)|'
        r'ToolTip(?:Opening|Closing)|'
        r'DataContextChanged|IsVisibleChanged|IsEnabledChanged|'
        r'RequestBringIntoView|SourceUpdated|TargetUpdated)'
        r'\s*=\s*(?:"[^"]*"|\'[^\']*\')'
    )

    def __init__(self, xaml_source, literal_string=None):
        stripped = str(xaml_source).strip()
        is_inline = (literal_string is True or
                     (literal_string is None and stripped.startswith('<')))
        if not is_inline:
            with open(str(xaml_source), 'r', encoding='utf-8') as _f:
                stripped = _f.read().strip()
        xaml_clean = self._XAML_EVENTS.sub('', stripped)
        rdr = System.Xml.XmlReader.Create(StringReader(xaml_clean))
        self._window = XamlReader.Load(rdr)

    def __getattr__(self, name):
        if name.startswith('_'):
            raise AttributeError(name)
        win = object.__getattribute__(self, '_window')
        el = win.FindName(name)
        if el is not None:
            return el
        return getattr(win, name)

    def ShowDialog(self):
        return self._window.ShowDialog()

    def Show(self):
        return self._window.Show()

    def Close(self):
        self._window.Close()


# ==================== Fim CPython Compat ====================

# ==================== Init Doc ====================

uidoc = __revit__.ActiveUIDocument  # noqa: F821
doc   = uidoc.Document if uidoc else None

if not doc:
    _alert("Nenhum projeto aberto.", exitscript=True)

app = doc.Application

# ==================== Fim Init Doc ====================

# ==================== Funções de Apoio ====================

def get_electrical_panels():
    """Retorna lista de painéis (ElectricalEquipment) ordenados pelo nome."""
    collector = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                 .WhereElementIsNotElementType())
    panels = []
    for p in collector:
        name = p.Name if p.Name else "Sem Nome"
        panels.append({
            "element": p,
            "id": p.Id,
            "name": name,
            "display": "{} (ID: {})".format(name, p.Id)
        })
    return sorted(panels, key=lambda x: x["name"])


def get_circuits_from_panel(panel_element):
    """Pega os circuitos onde o panel_element atua como painel."""
    circuits = []
    try:
        mep = panel_element.MEPModel
        if mep is None:
            return circuits
        systems = mep.GetElectricalSystems()
        if systems:
            for sys in systems:
                base_eq = sys.BaseEquipment
                if base_eq and base_eq.Id == panel_element.Id:
                    circuits.append(sys)
    except Exception:
        pass
    return circuits


# ==================== Classe UI ====================

class TransferCircuitsWindow(_WPFWindowCPy):
    def __init__(self, xaml_file):
        _WPFWindowCPy.__init__(self, xaml_file)
        self.panels = get_electrical_panels()
        self.dest_panels = []
        self.circuit_checkboxes = []

        self.init_ui()
        self.bind_events()

    def init_ui(self):
        if not self.panels:
            self.lbl_Info.Text = "Nenhum quadro elétrico encontrado no projeto."
            self.btn_Transfer.IsEnabled = False
            return

        displays = [p["display"] for p in self.panels]
        
        # Converte para lista .NET de forma segura para o CPython 3
        net_displays = List[System.Object]()
        for d in displays:
            net_displays.Add(d)
            
        self.cb_SourcePanel.ItemsSource = net_displays
        # O destino comeca vazio ou filtrado
        self._update_dest_list(-1)

        self.lbl_Info.Text = "Selecione o quadro de origem."

        if displays:
            self.cb_SourcePanel.SelectedIndex = 0

    def bind_events(self):
        self.btn_Cancel.Click += lambda s, a: self.Close()
        self.btn_Transfer.Click += self.on_transfer
        self.btn_SelectAll.Click += self.select_all
        self.btn_SelectNone.Click += self.select_none
        self.cb_SourcePanel.SelectionChanged += self.on_source_changed

    def on_source_changed(self, sender, args):
        self.sp_Circuits.Children.Clear()
        self.circuit_checkboxes = []

        idx = self.cb_SourcePanel.SelectedIndex
        self._update_dest_list(idx)
        
        if idx < 0:
            return

        panel = self.panels[idx]["element"]
        circuits = get_circuits_from_panel(panel)

        if not circuits:
            self.lbl_CircuitsCount.Text = "Nenhum circuito neste quadro."
            return

        self.lbl_CircuitsCount.Text = "{} circuito(s) disponívei(s)".format(len(circuits))

        def get_sort_key(c):
            circ_str = c.CircuitNumber
            nums = re.findall(r'\d+', circ_str)
            if nums:
                return [int(n) for n in nums]
            return [circ_str]

        circuits = sorted(circuits, key=get_sort_key)

        for circ in circuits:
            cb = CheckBox()
            c_name = circ.LoadName if hasattr(circ, 'LoadName') and circ.LoadName else "Sem Nome"
            c_num = circ.CircuitNumber

            poles = circ.Poles if hasattr(circ, 'Poles') else ""

            try:
                load_va = circ.ApparentLoad
                load_str = "{:.0f}W".format(load_va) if load_va < 1000 else "{:.1f}kW".format(load_va / 1000)
            except Exception:
                load_str = ""

            try:
                voltage = circ.Voltage
                voltage_str = "{:.0f}V".format(voltage)
            except Exception:
                voltage_str = ""

            lbl = "C{} — {}".format(c_num, c_name)
            extras = []
            if load_str:
                extras.append(load_str)
            if voltage_str:
                extras.append(voltage_str)
            if poles:
                extras.append("{}P".format(poles))
            if extras:
                lbl += "  [{}]".format(" | ".join(extras))

            cb.Content = lbl
            cb.IsChecked = True
            cb.Margin = Thickness(0, 3, 0, 3)
            try:
                cb.Foreground = SolidColorBrush(WpfColor.FromArgb(255, 241, 241, 241))
            except Exception:
                pass

            self.sp_Circuits.Children.Add(cb)
            self.circuit_checkboxes.append((cb, circ))

        self.lbl_Info.Text = "{} circuito(s) — selecione e escolha o destino.".format(len(circuits))

    def _update_dest_list(self, source_idx):
        """Atualiza a lista do destino removendo o quadro selecionado na origem."""
        dest_display_list = []
        self.dest_panels = []
        for i, p in enumerate(self.panels):
            if i != source_idx:
                dest_display_list.append(p["display"])
                self.dest_panels.append(p)
        
        net_dest = List[System.Object]()
        for d in dest_display_list:
            net_dest.Add(d)
        
        # Guardar seleçao atual se possivel
        prev_sel = self.cb_DestPanel.SelectedItem
        
        self.cb_DestPanel.ItemsSource = net_dest
        
        # Tentar restaurar seleçao
        if prev_sel and prev_sel in dest_display_list:
            self.cb_DestPanel.SelectedItem = prev_sel
        else:
            self.cb_DestPanel.SelectedIndex = -1

    def select_all(self, sender, args):
        for cb, _ in self.circuit_checkboxes:
            cb.IsChecked = True
        self._update_selection_count()

    def select_none(self, sender, args):
        for cb, _ in self.circuit_checkboxes:
            cb.IsChecked = False
        self._update_selection_count()

    def _update_selection_count(self):
        count = sum(1 for cb, _ in self.circuit_checkboxes if cb.IsChecked)
        total = len(self.circuit_checkboxes)
        self.lbl_Info.Text = "{} de {} circuito(s) selecionado(s).".format(count, total)

    def on_transfer(self, sender, args):
        s_idx = self.cb_SourcePanel.SelectedIndex
        d_idx = self.cb_DestPanel.SelectedIndex

        if s_idx < 0:
            _alert("Escolha o quadro de origem.")
            return
        if d_idx < 0:
            _alert("Escolha o quadro de destino.")
            return

        source_panel = self.panels[s_idx]["element"]
        dest_panel = self.dest_panels[d_idx]["element"]

        selected_circs = [circ for cb, circ in self.circuit_checkboxes if cb.IsChecked]
        if not selected_circs:
            _alert("Selecione pelo menos um circuito.")
            return

        confirma = _alert(
            "Transferir {} circuito(s) para o quadro '{}'?".format(
                len(selected_circs), dest_panel.Name),
            yes=True, no=True
        )
        if not confirma:
            return

        self.Close()

        sucessos = 0
        erros = []

        t = Transaction(doc, "Transferir Circuitos Inteligente")
        t.Start()
        try:
            for circ in selected_circs:
                circ_num_original = circ.CircuitNumber
                try:
                    circ.SelectPanel(dest_panel)
                    sucessos += 1
                except Exception as e:
                    erros.append((circ_num_original, str(e)))
            t.Commit()
        except Exception as e:
            try:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
            except:
                pass
            pass

        msg = "Transferência Concluída!\n\n"
        msg += "Circuitos movidos com sucesso: {}\n".format(sucessos)

        if erros:
            msg += "\nFalhas: {}\n".format(len(erros))
            for num, err in erros:
                err_lower = err.lower()
                if any(kw in err_lower for kw in ["voltage", "poles", "phase", "wire", "distribution"]):
                    msg += " - Circuito {}: Incompatibilidade de tensão/fases/fios com o quadro de destino.\n".format(num)
                elif "space" in err_lower or "slot" in err_lower or "full" in err_lower:
                    msg += " - Circuito {}: Quadro de destino sem espaço disponível.\n".format(num)
                elif "selectpanel" in err_lower or "member" in err_lower or "attribute" in err_lower:
                    msg += " - Circuito {}: Método SelectPanel não suportado — tente versão mais recente do Revit.\n".format(num)
                else:
                    msg += " - Circuito {}: {}.\n".format(num, err.split("\n")[0][:120])

        _alert(msg, title="Resumo da Transferência")


ui_obj = TransferCircuitsWindow(os.path.join(_BUNDLE_DIR, 'ui.xaml'))
ui_obj.ShowDialog()
