# -*- coding: utf-8 -*-
"""Transferir Circuitos — Desconecta circuitos do quadro de origem
e reconecta no destino com o número de polos escolhido.

Fluxo:
  1. Tenta SelectPanel (rápido, quando tensão/fases são compatíveis).
  2. Se falhar, recria o circuito no destino preservando propriedades.
"""
__title__ = "Transferir\nCircuitos"
__author__ = "Luís Fernando"

# ╔══════════════════════════════════════════════════════════╗
# ║  DEBUG_MODE                                              ║
# ║  True  = imprime detalhes no console pyRevit             ║
# ║  False = silencioso                                      ║
# ╚══════════════════════════════════════════════════════════╝
DEBUG_MODE = False

import os
import re
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Collections')

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, TransactionStatus, ElementSet, ElementId, StorageType
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType

from pyrevit import forms, script
from lf_utils import DebugLogger, make_warning_swallower

# ══════════════════════════════════════════════════════════════
#  INIT
# ══════════════════════════════════════════════════════════════

dbg   = DebugLogger(DEBUG_MODE)
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

_BUNDLE_DIR = os.path.dirname(__file__)


# ══════════════════════════════════════════════════════════════
#  FUNÇÕES DE APOIO
# ══════════════════════════════════════════════════════════════

def _safe_name(el):
    """Lê .Name de forma segura (IronPython e pythonnet)."""
    try:
        return el.Name
    except Exception:
        pass
    try:
        p = el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.HasValue:
            return p.AsString()
    except Exception:
        pass
    return "ID " + str(el.Id)


def get_electrical_panels():
    """Retorna lista de painéis (ElectricalEquipment) ordenados pelo nome."""
    collector = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                 .WhereElementIsNotElementType())
    panels = []
    for p in collector:
        name = _safe_name(p)
        panels.append({
            "element": p,
            "id": p.Id,
            "name": name,
            "display": "{} (ID: {})".format(name, p.Id)
        })
    return sorted(panels, key=lambda x: x["name"])


def get_circuits_from_panel(panel_element):
    """Pega os circuitos onde panel_element atua como painel (BaseEquipment)."""
    circuits = []
    try:
        mep = panel_element.MEPModel
        if mep is None:
            return circuits
        systems = mep.GetElectricalSystems()
        if systems:
            for sys in systems:
                try:
                    base_eq = sys.BaseEquipment
                    if base_eq and base_eq.Id == panel_element.Id:
                        circuits.append(sys)
                except Exception:
                    pass
    except Exception as e:
        dbg.error("get_circuits_from_panel: {}".format(e))
    return circuits


def _get_circuit_number(circ):
    """Retorna o numero do circuito como texto, tolerante a falhas."""
    try:
        num = circ.CircuitNumber
        return str(num or "")
    except Exception:
        return ""


def _get_circuit_start_slot(circ):
    """Retorna o slot inicial do circuito no quadro, quando disponivel."""
    try:
        slot = int(circ.StartSlot)
        if slot > 0:
            return slot
    except Exception:
        pass
    return None


def _circuit_sort_key(circ):
    """Ordenacao pela posicao real no quadro de origem.

    StartSlot e a ordem que aparece no painel. CircuitNumber entra apenas
    como fallback para casos em que a API nao exponha o slot.
    """
    slot = _get_circuit_start_slot(circ)
    try:
        eid = circ.Id.IntegerValue
    except Exception:
        eid = 0

    if slot is not None:
        return (0, slot, eid)

    cnum = _get_circuit_number(circ).strip()
    parts = re.split(r'(\d+)', cnum)
    natural = []
    for part in parts:
        if not part:
            continue
        if part.isdigit():
            natural.append((0, int(part)))
        else:
            natural.append((1, part.lower()))

    return (1, natural, eid)


# ══════════════════════════════════════════════════════════════
#  SNAPSHOT / RESTORE DE PROPRIEDADES
# ══════════════════════════════════════════════════════════════

# (label, BuiltInParameter ou None, [nomes lookup fallback])
_CIRCUIT_PROPS = [
    ("LoadName",        BuiltInParameter.RBS_ELEC_CIRCUIT_NAME, []),
    ("Rating",          BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM, []),
    ("Comments",        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, []),
    ("LoadClassification", None, ["Tipo de Carga"]),
    ("Description",     None, ["Descrição", "Description"]),
    ("WireSize",        None, ["Seção do Condutor Adotado (mm²)", "Condutor Adotado"]),
    ("L Considerado",   None, ["L Considerado", "L Considerado (m)"]),
    ("FCA",             None, ["FCA"]),
    ("FCT",             None, ["FCT"]),
]


def _read_param(elem, bip, names):
    """Lê o valor de um parâmetro (BuiltInParameter → LookupParameter)."""
    if bip is not None:
        try:
            p = elem.get_Parameter(bip)
            if p and p.HasValue:
                st = p.StorageType
                if st == StorageType.String:
                    return ("str", p.AsString())
                elif st == StorageType.Double:
                    return ("dbl", p.AsDouble())
                elif st == StorageType.Integer:
                    return ("int", p.AsInteger())
                elif st == StorageType.ElementId:
                    return ("eid", p.AsElementId())
        except Exception:
            pass
    for n in names:
        try:
            p = elem.LookupParameter(n)
            if p and p.HasValue:
                st = p.StorageType
                if st == StorageType.String:
                    return ("str", p.AsString())
                elif st == StorageType.Double:
                    return ("dbl", p.AsDouble())
                elif st == StorageType.Integer:
                    return ("int", p.AsInteger())
                elif st == StorageType.ElementId:
                    return ("eid", p.AsElementId())
        except Exception:
            pass
    return None


def _write_param(elem, bip, names, typed_val):
    """Escreve (tipo, valor) num parâmetro. Retorna True se conseguiu."""
    if typed_val is None:
        return False
    kind, val = typed_val

    def _do_set(p):
        if p is None or p.IsReadOnly:
            return False
        try:
            if kind == "str" and p.StorageType == StorageType.String:
                p.Set(str(val) if val else "")
                return True
            elif kind == "dbl" and p.StorageType == StorageType.Double:
                p.Set(float(val))
                return True
            elif kind == "int" and p.StorageType == StorageType.Integer:
                p.Set(int(val))
                return True
            elif kind == "eid" and p.StorageType == StorageType.ElementId:
                p.Set(val)
                return True
        except Exception:
            pass
        return False

    if bip is not None:
        try:
            if _do_set(elem.get_Parameter(bip)):
                return True
        except Exception:
            pass
    for n in names:
        try:
            if _do_set(elem.LookupParameter(n)):
                return True
        except Exception:
            pass
    return False


def snapshot_circuit(circ):
    """Captura todas as propriedades de um circuito elétrico."""
    snap = {}
    for label, bip, names in _CIRCUIT_PROPS:
        snap[label] = _read_param(circ, bip, names)
    try:
        snap["_CircuitNumber"] = circ.CircuitNumber
    except Exception:
        snap["_CircuitNumber"] = ""
    try:
        snap["_LoadName"] = circ.LoadName
    except Exception:
        snap["_LoadName"] = ""
    try:
        snap["_Poles"] = circ.PolesNumber
    except Exception:
        snap["_Poles"] = 1
    return snap


def restore_circuit(new_circ, snap):
    """Aplica as propriedades capturadas no circuito novo."""
    for label, bip, names in _CIRCUIT_PROPS:
        val = snap.get(label)
        if val is not None:
            _write_param(new_circ, bip, names, val)


# ══════════════════════════════════════════════════════════════
#  TRANSFERÊNCIA INTELIGENTE
# ══════════════════════════════════════════════════════════════

def transfer_one_circuit(circ, dest_panel, target_poles):
    """Transfere um circuito para dest_panel.

    Estratégia:
      1. SelectPanel (rápido, compatível).
      2. Se falhar, recria: snapshot → delete → create → restore.

    Retorna (sucesso, mensagem).
    """
    circ_num = ""
    try:
        circ_num = circ.CircuitNumber
    except Exception:
        pass

    # ── 1. SelectPanel direto ──
    try:
        circ.SelectPanel(dest_panel)
        # Ajustar polos se necessário
        try:
            if circ.PolesNumber != target_poles:
                p = circ.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
                if p and not p.IsReadOnly:
                    p.Set(target_poles)
        except Exception:
            pass
        dbg.info("C{}: SelectPanel OK".format(circ_num))
        return True, "SelectPanel OK"
    except Exception as e:
        dbg.debug("C{}: SelectPanel falhou ({}), tentando recriar".format(circ_num, e))

    # ── 2. Recriar circuito ──
    snap = snapshot_circuit(circ)

    # Coletar membros
    members = []
    try:
        if circ.Elements:
            for el in circ.Elements:
                members.append(el)
    except Exception:
        pass

    if not members:
        return False, "Circuito sem membros"

    # Deletar circuito antigo
    try:
        doc.Delete(circ.Id)
    except Exception as e:
        return False, "Erro ao deletar: {}".format(e)

    doc.Regenerate()

    # Criar novo circuito a partir do primeiro membro
    new_circ = None
    try:
        first_ids = List[ElementId]()
        first_ids.Add(members[0].Id)
        new_circ = ElectricalSystem.Create(doc, first_ids, ElectricalSystemType.PowerCircuit)
    except Exception as e:
        return False, "Erro ao criar: {}".format(e)

    if new_circ is None:
        return False, "ElectricalSystem.Create retornou None"

    # Adicionar membros restantes
    for m in members[1:]:
        try:
            add_ids = List[ElementId]()
            add_ids.Add(m.Id)
            new_circ.AddToCircuit(add_ids)
        except Exception as e:
            dbg.warn("C{}: membro {} não adicionado: {}".format(circ_num, m.Id, e))

    # Conectar ao painel destino
    try:
        new_circ.SelectPanel(dest_panel)
    except Exception as e:
        return False, "Recriado mas falhou ao conectar: {}".format(e)

    # Ajustar número de polos
    try:
        p = new_circ.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
        if p and not p.IsReadOnly:
            p.Set(target_poles)
    except Exception:
        pass

    # Restaurar propriedades
    restore_circuit(new_circ, snap)

    dbg.info("C{}: Recriado e transferido".format(circ_num))
    return True, "Recriado"


# ══════════════════════════════════════════════════════════════
#  INTERFACE (WPF / pyrevit.forms)
# ══════════════════════════════════════════════════════════════

class TransferCircuitsWindow(forms.WPFWindow):

    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.panels = get_electrical_panels()
        self.dest_panels = []
        # Cada item: (checkbox, combobox_polos, circuito)
        self.circuit_rows = []

        self._init_ui()
        self._bind_events()

    # ── Inicialização ──

    def _init_ui(self):
        if not self.panels:
            self.lbl_Info.Text = "Nenhum quadro elétrico encontrado no projeto."
            self.btn_Transfer.IsEnabled = False
            return

        for p in self.panels:
            self.cb_SourcePanel.Items.Add(p["display"])

        self._update_dest_list(-1)
        self.lbl_Info.Text = "Selecione o quadro de origem."

        if self.panels:
            self.cb_SourcePanel.SelectedIndex = 0

    def _bind_events(self):
        self.btn_Cancel.Click    += self._on_cancel
        self.btn_Transfer.Click  += self._on_transfer
        self.btn_SelectAll.Click += self._on_select_all
        self.btn_SelectNone.Click += self._on_select_none
        self.cb_SourcePanel.SelectionChanged += self._on_source_changed

    # ── Eventos ──

    def _on_cancel(self, sender, args):
        self.Close()

    def _on_source_changed(self, sender, args):
        from System.Windows.Controls import CheckBox, ComboBox as WpfComboBox
        from System.Windows import Thickness

        self.sp_Circuits.Children.Clear()
        self.circuit_rows = []

        idx = self.cb_SourcePanel.SelectedIndex
        self._update_dest_list(idx)

        if idx < 0:
            return

        panel = self.panels[idx]["element"]
        circuits = get_circuits_from_panel(panel)

        if not circuits:
            self.lbl_CircuitsCount.Text = "Nenhum circuito."
            return

        self.lbl_CircuitsCount.Text = "{} circuito(s)".format(len(circuits))

        circuits = sorted(circuits, key=_circuit_sort_key)

        for circ in circuits:
            self._add_circuit_row(circ)

        self.lbl_Info.Text = "{} circuito(s) — selecione e escolha o destino.".format(len(circuits))

    def _add_circuit_row(self, circ):
        """Cria uma linha: [CheckBox (info)] + [ComboBox (polos)]."""
        from System.Windows.Controls import (
            CheckBox, Grid as WpfGrid, ColumnDefinition,
            ComboBox as WpfComboBox
        )
        from System.Windows import Thickness, GridLength, GridUnitType

        # ── Dados ──
        c_name = ""
        try:
            c_name = circ.LoadName or ""
        except Exception:
            pass
        if not c_name:
            c_name = "Sem Nome"

        c_num = ""
        try:
            c_num = circ.CircuitNumber
        except Exception:
            pass

        curr_poles = 1
        try:
            curr_poles = circ.PolesNumber
        except Exception:
            pass

        load_str = ""
        try:
            load_va = circ.ApparentLoad
            load_str = "{:.0f}VA".format(load_va) if load_va < 1000 else "{:.1f}kVA".format(load_va / 1000)
        except Exception:
            pass

        voltage_str = ""
        try:
            voltage_str = "{:.0f}V".format(circ.Voltage)
        except Exception:
            pass

        lbl = "C{} — {}".format(c_num, c_name)
        extras = []
        if load_str:
            extras.append(load_str)
        if voltage_str:
            extras.append(voltage_str)
        if curr_poles:
            extras.append("{}P".format(curr_poles))
        if extras:
            lbl += "  [{}]".format(" | ".join(extras))

        # ── Visual: Grid com CheckBox + ComboBox ──
        row = WpfGrid()
        col0 = ColumnDefinition()
        col0.Width = GridLength(1.0, GridUnitType.Star)
        col1 = ColumnDefinition()
        col1.Width = GridLength(90.0, GridUnitType.Pixel)
        row.ColumnDefinitions.Add(col0)
        row.ColumnDefinitions.Add(col1)
        row.Margin = Thickness(0, 2, 0, 2)

        cb = CheckBox()
        cb.Content = lbl
        cb.IsChecked = True
        cb.FontSize = 13
        WpfGrid.SetColumn(cb, 0)
        row.Children.Add(cb)

        poles_cb = WpfComboBox()
        poles_cb.Items.Add("1 Polo")
        poles_cb.Items.Add("2 Polos")
        poles_cb.Items.Add("3 Polos")
        if curr_poles == 3:
            poles_cb.SelectedIndex = 2
        elif curr_poles == 2:
            poles_cb.SelectedIndex = 1
        else:
            poles_cb.SelectedIndex = 0
        poles_cb.Width = 80
        poles_cb.Height = 22
        poles_cb.FontSize = 11
        WpfGrid.SetColumn(poles_cb, 1)
        row.Children.Add(poles_cb)

        self.sp_Circuits.Children.Add(row)
        self.circuit_rows.append((cb, poles_cb, circ))

    def _update_dest_list(self, source_idx):
        """Atualiza a ComboBox destino excluindo o quadro de origem."""
        prev_sel = self.cb_DestPanel.SelectedItem

        self.cb_DestPanel.Items.Clear()
        self.dest_panels = []

        for i, p in enumerate(self.panels):
            if i != source_idx:
                self.cb_DestPanel.Items.Add(p["display"])
                self.dest_panels.append(p)

        if prev_sel and prev_sel in [p["display"] for p in self.dest_panels]:
            self.cb_DestPanel.SelectedItem = prev_sel
        else:
            self.cb_DestPanel.SelectedIndex = -1

    def _on_select_all(self, sender, args):
        for cb, _, _ in self.circuit_rows:
            cb.IsChecked = True
        self._update_count()

    def _on_select_none(self, sender, args):
        for cb, _, _ in self.circuit_rows:
            cb.IsChecked = False
        self._update_count()

    def _update_count(self):
        count = sum(1 for cb, _, _ in self.circuit_rows if cb.IsChecked)
        total = len(self.circuit_rows)
        self.lbl_Info.Text = "{} de {} selecionado(s).".format(count, total)

    # ── Transferir ──

    def _on_transfer(self, sender, args):
        s_idx = self.cb_SourcePanel.SelectedIndex
        d_idx = self.cb_DestPanel.SelectedIndex

        if s_idx < 0:
            forms.alert("Escolha o quadro de origem.", title="Transferir Circuitos")
            return
        if d_idx < 0:
            forms.alert("Escolha o quadro de destino.", title="Transferir Circuitos")
            return

        dest_panel = self.dest_panels[d_idx]["element"]

        # Coletar selecionados + polos desejados
        selected = []
        for cb, poles_cb, circ in self.circuit_rows:
            if cb.IsChecked:
                pi = poles_cb.SelectedIndex
                target_poles = [1, 2, 3][pi] if 0 <= pi <= 2 else 1
                selected.append((circ, target_poles))

        selected = sorted(selected, key=lambda item: _circuit_sort_key(item[0]))

        if not selected:
            forms.alert("Selecione pelo menos um circuito.", title="Transferir Circuitos")
            return

        dest_name = _safe_name(dest_panel)

        confirma = forms.alert(
            "Transferir {} circuito(s) para '{}'?\n\n"
            "Os circuitos serão desconectados do quadro atual\n"
            "e reconectados no destino com os polos escolhidos.".format(
                len(selected), dest_name),
            title="Transferir Circuitos",
            yes=True, no=True
        )
        if not confirma:
            return

        self.Close()

        # ── Executar transferência ──
        dbg.section("Transferir Circuitos")
        dbg.info("Destino: {} ({} circuitos)".format(dest_name, len(selected)))

        sucessos = 0
        erros = []

        t = Transaction(doc, "Transferir Circuitos")
        preprocessor = make_warning_swallower()
        if preprocessor:
            opts = t.GetFailureHandlingOptions()
            opts.SetFailuresPreprocessor(preprocessor)
            t.SetFailureHandlingOptions(opts)
        t.Start()

        try:
            for circ, target_poles in selected:
                circ_num = ""
                try:
                    circ_num = circ.CircuitNumber
                except Exception:
                    pass

                ok, msg = transfer_one_circuit(circ, dest_panel, target_poles)
                if ok:
                    sucessos += 1
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                else:
                    erros.append((circ_num, msg))
                    dbg.error("C{}: {}".format(circ_num, msg))

            t.Commit()
        except Exception as e:
            try:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
            except Exception:
                pass
            erros.append(("GERAL", str(e)))
            dbg.error("Exceção geral: {}".format(e))

        # ── Resumo ──
        dbg.section("Resultado")
        dbg.info("Sucesso: {}  |  Falhas: {}".format(sucessos, len(erros)))

        msg = "Transferência Concluída!\n\n"
        msg += "Circuitos movidos com sucesso: {}\n".format(sucessos)

        if erros:
            msg += "\nFalhas: {}\n".format(len(erros))
            for num, err in erros:
                msg += " • C{}: {}\n".format(num, err[:120])

        forms.alert(msg, title="Resumo da Transferência")


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    win = TransferCircuitsWindow(os.path.join(_BUNDLE_DIR, 'ui.xaml'))
    win.ShowDialog()
