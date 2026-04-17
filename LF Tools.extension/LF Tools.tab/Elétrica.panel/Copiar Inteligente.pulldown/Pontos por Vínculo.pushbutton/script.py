#! python3
# -*- coding: utf-8 -*-
"""Pontos por Vínculo v6"""
__title__ = "Pontos por\nVínculo"
__author__ = "Luís Fernando"

# ╔══════════════════════════════════════════════════════════════╗
# ║                    MODO DEBUG                                ║
# ║  True  = imprime detalhes no console pyRevit                 ║
# ║  False = silencioso                                          ║
# ╚══════════════════════════════════════════════════════════════╝
DEBUG_MODE = False
import os
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Thickness, GridLength, GridUnitType,
    VerticalAlignment, HorizontalAlignment, TextWrapping, Visibility, CornerRadius
)
from System.Windows import MessageBox
from System.Windows.Controls import (
    CheckBox, Grid, ColumnDefinition, RowDefinition, TextBlock, ComboBox, TextBox, Border, ScrollViewer, StackPanel
)
from System.Windows.Media import SolidColorBrush, Color
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    RevitLinkInstance, Level, Transaction, TransactionStatus,
    BuiltInParameter,
)
from Autodesk.Revit.DB.Structure import StructuralType

from lf_utils import DebugLogger, WPFWindowCPy, get_revit_context

dbg = DebugLogger(DEBUG_MODE)

# ==================== Helpers ====================

try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

SCAN_CATS = [
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_SpecialityEquipment,
]

ELEC_CATS = [
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_ElectricalFixtures,
]

DATA_CATS = [
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_ElectricalFixtures,
]

# Palavras-chave para filtrar famílias NÃO-ELÉTRICAS
NON_ELECTRIC_KEYWORDS = [
    "bancada", "mesa", "torneira", "pia", "carrinho", "armario", "armário",
    "cadeira", "prateleira", "balcao", "balcão", "gabinete",
    "mobiliario", "mobiliário", "estante", "banheira", "vaso", "sanitario",
    "sanitário", "lavatório", "lavatorio", "cuba", "tanque", "chuveiro",
    "porta", "janela", "escada", "rampa",
    "divisoria", "divisória", "guarda-corpo", "corrimao", "corrimão",
    "maca", "leito", "cama", "sofa", "sofá", "poltrona", "cilindro", "cofre"
]

_ACCENT_MAP = {'ã':'a','â':'a','á':'a','à':'a','ç':'c','é':'e','ê':'e',
               'í':'i','ó':'o','ô':'o','õ':'o','ú':'u','ü':'u'}

def _normalize(s):
    s = s.lower()
    for acc, base in _ACCENT_MAP.items():
        s = s.replace(acc, base)
    return s

def _is_non_electric(display_name):
    name = _normalize(display_name)
    return any(_normalize(kw) in name for kw in NON_ELECTRIC_KEYWORDS)


def get_link_instances():
    result = []
    for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        try:
            ldoc = link.GetLinkDocument()
            if ldoc:
                result.append({"display": ldoc.Title or link.Name, "link_inst": link, "link_doc": ldoc})
        except:
            pass
    return sorted(result, key=lambda x: x["display"])


# ── Nomes conhecidos do parâmetro de potência nas famílias do projeto ──
_POWER_PARAM_NAMES = [
    "Potência Aparente (VA)",
    "Potência Aparente",
    "Apparent Load",
    "Potência",
    "Power",
    "Wattage",
    "Potencia Aparente",
    "Potencia",
    "Carga Aparente",
]

# GUID real do shared parameter "Potência Aparente (VA)" das famílias LF
_POWER_GUID_STR = "44b19786-a579-4490-bcfe-f9ee378c8811"


def _try_set_value(p, val_float, val_str):
    """Tenta setar um parâmetro numérico. Tenta SetValueString e Set(double)."""
    if not p or p.IsReadOnly:
        return False
    from Autodesk.Revit.DB import StorageType
    if p.StorageType == StorageType.Double:
        # Tenta SetValueString (Revit converte unidades)
        try:
            p.SetValueString(str(val_str))
            return True
        except:
            pass
        # Fallback: valor bruto (unidades internas = pés, mas VA não tem conversão)
        try:
            p.Set(val_float)
            return True
        except:
            pass
    elif p.StorageType == StorageType.Integer:
        try:
            p.Set(int(val_float))
            return True
        except:
            pass
    return False


def _set_power(element, pow_str):
    """Seta a potência aparente no elemento usando múltiplas estratégias:
    1. LookupParameter por nomes conhecidos (instância)
    2. GUID do shared parameter
    3. BuiltInParameter RBS_ELEC_APPARENT_LOAD
    4. LookupParameter por nomes conhecidos (tipo)
    5. Conector elétrico MEPModel
    """
    import re
    from System import Guid

    pow_str = str(pow_str).strip().replace(',', '.')
    if not pow_str:
        return
    match = re.search(r"[-+]?[0-9]*\.?[0-9]+", pow_str)
    if not match:
        return
    val = float(match.group())

    # ── 1. LookupParameter por nome (instância) ──
    for pname in _POWER_PARAM_NAMES:
        try:
            p = element.LookupParameter(pname)
            if p and p.HasValue is not None:
                if _try_set_value(p, val, pow_str):
                    return
        except:
            pass

    # ── 2. GUID do shared parameter ──
    try:
        guid = Guid(_POWER_GUID_STR)
        p = element.get_Parameter(guid)
        if p:
            if _try_set_value(p, val, pow_str):
                return
    except:
        pass

    # ── 3. BuiltInParameter RBS_ELEC_APPARENT_LOAD ──
    try:
        p = element.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD)
        if _try_set_value(p, val, pow_str):
            return
    except:
        pass

    # ── 4. LookupParameter por nome (tipo da família) ──
    try:
        elem_type = element.Document.GetElement(element.GetTypeId())
        if elem_type:
            for pname in _POWER_PARAM_NAMES:
                try:
                    p = elem_type.LookupParameter(pname)
                    if p:
                        if _try_set_value(p, val, pow_str):
                            return
                except:
                    continue
    except:
        pass

    # ── 5. Conector elétrico (MEPModel) — última tentativa ──
    try:
        from Autodesk.Revit.DB.Electrical import ElectricalSystemType
        if hasattr(element, 'MEPModel') and element.MEPModel:
            cm = element.MEPModel.ConnectorManager
            if cm:
                from Autodesk.Revit.DB import Domain
                for c in cm.Connectors:
                    if c.Domain == Domain.DomainElectrical:
                        # Tenta setar a carga via propriedade do conector
                        try:
                            c.ElectricalApparentLoad = val
                            return
                        except:
                            pass
    except:
        pass


# ==================== UI ====================

class PontosVinculoWindow(WPFWindowCPy):
    def __init__(self, xaml_file):
        WPFWindowCPy.__init__(self, xaml_file)
        # doc fresco a cada abertura — o módulo é cacheado pelo pyRevit,
        # então o doc module-level pode ficar stale após múltiplas execuções.
        _, self._doc = get_revit_context()
        self._links = get_link_instances()
        self._host_levels = sorted(
            FilteredElementCollector(self._doc).OfClass(Level),
            key=lambda lv: lv.Elevation
        )
        self._elec_symbols = self._collect_symbols(ELEC_CATS)
        self._data_symbols = self._collect_symbols(DATA_CATS)
        self._family_rows = []
        self._filter_non_electric = True
        self._init_ui()

    def _collect_symbols(self, categories):
        result = []
        seen = set()
        for bic in categories:
            try:
                for s in FilteredElementCollector(self._doc).OfCategory(bic).WhereElementIsElementType():
                    try:
                        full_name = "{} : {}".format(s.Family.Name, s.Name)
                        if full_name not in seen:
                            seen.add(full_name)
                            result.append({"full_name": full_name, "display": s.Name, "symbol": s})
                    except:
                        pass
            except:
                pass
        result.sort(key=lambda x: x["display"])
        return result

    def _find_match_idx(self, display_name, symbol_list):
        """Índice 1-based (0 = não colocar). Tenta exato, depois por família via full_name."""
        if not symbol_list:
            return 0
        name_lower = display_name.lower()
        for i, sym in enumerate(symbol_list):
            if sym.get("full_name", sym["display"]).lower() == name_lower:
                return i + 1
        fam_part = display_name.split(" : ")[0].strip().lower()
        for i, sym in enumerate(symbol_list):
            if sym.get("full_name", sym["display"]).split(" : ")[0].strip().lower() == fam_part:
                return i + 1
        return 0

    def _init_ui(self):
        if not self._links:
            self.lbl_Status.Text = "Nenhum vínculo carregado no projeto."
            MessageBox.Show("Nenhum vínculo RVT carregado no projeto.", "Pontos por Vínculo")
            return

        net_links = List[System.Object]()
        for lnk in self._links:
            net_links.Add(lnk["display"])
        self.cb_Link.ItemsSource = net_links
        self.cb_Link.SelectedIndex = 0

        net_levels = List[System.Object]()
        for lv in self._host_levels:
            net_levels.Add(lv.Name)
        self.cb_DefaultLevel.ItemsSource = net_levels
        if self._host_levels:
            self.cb_DefaultLevel.SelectedIndex = 0

        self.chk_FilterNonElec.IsChecked = True
        self.chk_FilterNonElec.Checked   += self._on_filter_toggle
        self.chk_FilterNonElec.Unchecked += self._on_filter_toggle

        self.btn_Scan.Click       += self._on_scan
        self.btn_Place.Click      += self._on_place
        self.btn_Cancel.Click     += lambda s, a: self.Close()
        self.btn_SelectAll.Click  += lambda s, a: self._set_all_checked(True)
        self.btn_SelectNone.Click += lambda s, a: self._set_all_checked(False)

    def _on_filter_toggle(self, sender, args):
        self._filter_non_electric = (self.chk_FilterNonElec.IsChecked == True)
        self._apply_visibility_filter()

    def _apply_visibility_filter(self):
        hidden = 0
        for row in self._family_rows:
            if self._filter_non_electric and row["is_non_electric"]:
                row["grid"].Visibility = Visibility.Collapsed
                hidden += 1
            else:
                row["grid"].Visibility = Visibility.Visible
        visible = len(self._family_rows) - hidden
        suffix = " ({} oculto(s) por filtro)".format(hidden) if hidden else ""
        self.lbl_Status.Text = "{} tipo(s) visível(is){}.".format(visible, suffix)

    def _set_all_checked(self, state):
        for row in self._family_rows:
            if row["grid"].Visibility == Visibility.Visible:
                row["cb"].IsChecked = state

    def _on_scan(self, sender, args):
        idx = self.cb_Link.SelectedIndex
        if idx < 0:
            return
        link_doc = self._links[idx]["link_doc"]
        dbg.section("Pontos por Vínculo — Scan")
        dbg.info("Vínculo: {}".format(self._links[idx]["display"]))

        # Guarda o transform do vínculo para corrigir coordenadas ao colocar pontos
        try:
            self._link_transform = self._links[idx]["link_inst"].GetTotalTransform()
            dbg.debug("Transform obtido com sucesso.")
        except:
            self._link_transform = None
            dbg.warn("Transform não disponível — usando coordenadas locais.")
        self.sp_Families.Children.Clear()
        self._family_rows = []

        mapa = {}
        for bic in SCAN_CATS:
            try:
                count_cat = 0
                for el in FilteredElementCollector(link_doc).OfCategory(bic).WhereElementIsNotElementType():
                    fam, type_n = "", ""
                    try:
                        _sym = el.Symbol
                        fam, type_n = _sym.Family.Name, _sym.Name
                    except:
                        fam = el.Name
                    display = "{} : {}".format(fam, type_n) if type_n else fam
                    if display not in mapa:
                        mapa[display] = {"display": display, "instances": []}
                    mapa[display]["instances"].append(el)
                    count_cat += 1
                dbg.debug("Categoria {}: {} instâncias".format(str(bic).split('.')[-1], count_cat))
            except:
                pass

        for key in sorted(mapa.keys()):
            row = self._build_row(mapa[key])
            self.sp_Families.Children.Add(row["grid"])
            self._family_rows.append(row)

        dbg.info("Famílias encontradas: {}".format(len(self._family_rows)))
        self._apply_visibility_filter()

    # ------------------------------------------------------------------
    # ComboBox editável com filtragem em tempo real
    # ------------------------------------------------------------------

    def _make_editable_combo(self, symbol_list, default_idx):
        """ComboBox com estilo e Autocomplete Nativo WPF (To Excel style)."""
        cb = ComboBox()
        cb.IsEditable = True  # Ativa campo digitável
        # Em WPF, deixar IsTextSearchEnabled omitido/True + IsEditable=True 
        # aciona o AutoComplete natural de Windows (começa pela letra etc.)
        cb.Height = 28
        cb.Margin = Thickness(2, 2, 2, 2)
        cb.Padding = Thickness(6, 0, 0, 0)
        cb.VerticalContentAlignment = VerticalAlignment.Center

        # ── Visual: cores dark do tema ──────────────────────────────────
        INPUT_BG   = SolidColorBrush(Color.FromRgb(0x13, 0x16, 0x1C))
        BORDER_CLR = SolidColorBrush(Color.FromRgb(0x31, 0x38, 0x44))
        FG_CLR     = SolidColorBrush(Color.FromRgb(0xF2, 0xF4, 0xF8))
        cb.Background   = INPUT_BG
        cb.Foreground   = FG_CLR
        cb.BorderBrush  = BORDER_CLR
        cb.BorderThickness = Thickness(1)

        # Lista completa — índice 0 = "não colocar"
        all_items = ["— Não colocar —"] + [s["display"] for s in symbol_list]

        net = List[System.Object]()
        for it in all_items:
            net.Add(it)
        cb.ItemsSource = net
        cb.SelectedIndex = default_idx

        return cb

    def _get_symbol_from_combo(self, cb, symbol_list):
        """Retorna FamilySymbol inteligente com fallback e insensitive."""
        if cb.SelectedIndex > 0:
            idx = cb.SelectedIndex - 1
            if idx < len(symbol_list):
                return symbol_list[idx]["symbol"]

        text = str(cb.Text).strip() if cb.Text else ""
        if not text or text == "— Não colocar —":
            return None
            
        t_low = text.lower()
        for sym in symbol_list:
            if sym["display"].lower() == t_low or sym.get("full_name", "").lower() == t_low:
                return sym["symbol"]
                
        # Fallback: primeira ocorrência que comece ou contenha
        for sym in symbol_list:
            if t_low in sym["display"].lower() or t_low in sym.get("full_name", "").lower():
                return sym["symbol"]
                
        return None

    # ------------------------------------------------------------------

    def _build_row(self, fam):
        # Container com borda suave entre linhas
        grid = Grid()
        grid.Margin = Thickness(0, 1, 0, 1)
        grid.MinHeight = 34

        widths = [(28, 0), (130, 0), (36, 0), (0, 2), (0, 2), (72, 0)]
        for w, star in widths:
            cd = ColumnDefinition()
            cd.Width = GridLength(star, GridUnitType.Star) if star > 0 else GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        # CheckBox de seleção
        cb = CheckBox()
        cb.IsChecked = True
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin = Thickness(4, 0, 0, 0)
        cb.Foreground = SolidColorBrush(Color.FromRgb(0xF2, 0xF4, 0xF8))
        Grid.SetColumn(cb, 0)

        # Nome da família
        tb_name = TextBlock()
        tb_name.Text = fam["display"]
        tb_name.Margin = Thickness(5, 0, 5, 0)
        tb_name.VerticalAlignment = VerticalAlignment.Center
        tb_name.TextWrapping = TextWrapping.NoWrap
        from System.Windows import TextTrimming as _TT
        tb_name.TextTrimming = _TT.CharacterEllipsis
        tb_name.ToolTip = fam["display"]
        tb_name.Foreground = SolidColorBrush(Color.FromRgb(0xF2, 0xF4, 0xF8))
        Grid.SetColumn(tb_name, 1)

        # Quantidade
        tb_qty = TextBlock()
        tb_qty.Text = str(len(fam["instances"]))
        tb_qty.VerticalAlignment = VerticalAlignment.Center
        tb_qty.HorizontalAlignment = HorizontalAlignment.Center
        tb_qty.Foreground = SolidColorBrush(Color.FromRgb(0xC4, 0xCA, 0xD4))
        Grid.SetColumn(tb_qty, 2)

        # ComboBox — Ponto Elétrico
        elec_default = self._find_match_idx(fam["display"], self._elec_symbols)
        c_elec = self._make_editable_combo(self._elec_symbols, elec_default)
        Grid.SetColumn(c_elec, 3)

        # ComboBox — Ponto de Dados
        c_dados = self._make_editable_combo(self._data_symbols, 0)
        Grid.SetColumn(c_dados, 4)

        # TextBox — Potência (VA)
        INPUT_BG   = SolidColorBrush(Color.FromRgb(0x13, 0x16, 0x1C))
        BORDER_CLR = SolidColorBrush(Color.FromRgb(0x31, 0x38, 0x44))
        FG_CLR     = SolidColorBrush(Color.FromRgb(0xF2, 0xF4, 0xF8))
        t_pow = TextBox()
        t_pow.Height = 26
        t_pow.Margin = Thickness(2)
        t_pow.VerticalContentAlignment = VerticalAlignment.Center
        t_pow.Background   = INPUT_BG
        t_pow.Foreground   = FG_CLR
        t_pow.BorderBrush  = BORDER_CLR
        t_pow.BorderThickness = Thickness(1)
        t_pow.Padding = Thickness(4, 0, 4, 0)
        Grid.SetColumn(t_pow, 5)

        for child in [cb, tb_name, tb_qty, c_elec, c_dados, t_pow]:
            grid.Children.Add(child)

        return {
            "grid": grid,
            "cb": cb,
            "c_elec": c_elec,
            "c_dados": c_dados,
            "t_pow": t_pow,
            "instances": fam["instances"],
            "is_non_electric": _is_non_electric(fam["display"]),
        }

    def _on_place(self, sender, args):
        to_place = []
        for r in self._family_rows:
            if not r["cb"].IsChecked:
                continue
            elec_sym  = self._get_symbol_from_combo(r["c_elec"],  self._elec_symbols)
            dados_sym = self._get_symbol_from_combo(r["c_dados"], self._data_symbols)
            if not elec_sym and not dados_sym:
                continue
            to_place.append({
                "instances": r["instances"],
                "elec_sym":  elec_sym,
                "dados_sym": dados_sym,
                "pow":       r["t_pow"].Text.strip() or None,
            })

        if not to_place:
            MessageBox.Show("Nenhuma linha com ponto elétrico ou de dados selecionada.", "Pontos por Vínculo")
            return

        dbg.section("Pontos por Vínculo — Colocação")
        dbg.info("Famílias a processar: {}".format(len(to_place)))

        self.Close()

        _doc = self._doc
        t = Transaction(_doc, "Pontos por Vínculo")
        try:
            status = t.Start()
            if status != TransactionStatus.Started:
                MessageBox.Show("Não foi possível iniciar a transação (status: {}).".format(status), "Pontos por Vínculo")
                return

            levels = list(FilteredElementCollector(_doc).OfClass(Level))

            # Ativar todos os símbolos necessários e regenerar UMA VEZ antes de criar instâncias.
            # Revit exige Regenerate() após Activate() — sem isso NewFamilyInstance falha silenciosamente.
            any_activated = False
            for item in to_place:
                for sym in [item["elec_sym"], item["dados_sym"]]:
                    if sym and not sym.IsActive:
                        sym.Activate()
                        any_activated = True
            if any_activated:
                _doc.Regenerate()
                dbg.debug("Regenerate() após Activate().")

            transform = getattr(self, '_link_transform', None)

            # Offset do ponto de dados em relação ao elétrico:
            #   -0.10 m em Z (10 cm abaixo)  →  em pés: -0.10 / 0.3048
            #   +0.15 m em X (15 cm lateral) →  em pés: +0.15 / 0.3048
            DADOS_OFFSET_Z = -0.10 / 0.3048
            DADOS_OFFSET_X = +0.15 / 0.3048

            from Autodesk.Revit.DB import XYZ

            placed_count = 0
            power_ok = 0
            power_fail = 0
            skip_count = 0
            power_elements = []  # acumular (elemento, valor) para setar potência em lote

            for item in to_place:
                elec_name  = item["elec_sym"].Name  if item["elec_sym"]  else "—"
                dados_name = item["dados_sym"].Name if item["dados_sym"] else "—"
                dbg.sub("{} ({} inst.) → elec:{} dados:{} pot:{}".format(
                    item["instances"][0].Symbol.Family.Name if item["instances"] else "?",
                    len(item["instances"]), elec_name, dados_name, item["pow"] or "—"
                ))
                for inst in item["instances"]:
                    # ── Coordenadas ──
                    try:
                        pt_local = inst.Location.Point
                        pt = transform.OfPoint(pt_local) if transform else pt_local
                        lvl = min(levels, key=lambda l: abs(l.Elevation - pt.Z))
                    except:
                        skip_count += 1
                        dbg.warn("  inst.Id={} — sem localização, pulado.".format(inst.Id))
                        continue

                    # ── Ponto Elétrico ──
                    new_elec = None
                    if item["elec_sym"]:
                        try:
                            new_elec = _doc.Create.NewFamilyInstance(
                                pt, item["elec_sym"], lvl, StructuralType.NonStructural
                            )
                            placed_count += 1
                            if item["pow"]:
                                power_elements.append((new_elec, item["pow"]))
                        except Exception as ex:
                            dbg.warn("  elétrico falhou inst.Id={}: {}".format(inst.Id, ex))

                    # ── Ponto de Dados (independente do elétrico) ──
                    if item["dados_sym"]:
                        try:
                            pt_dados = XYZ(
                                pt.X + DADOS_OFFSET_X,
                                pt.Y,
                                pt.Z + DADOS_OFFSET_Z
                            )
                            _doc.Create.NewFamilyInstance(
                                pt_dados, item["dados_sym"], lvl, StructuralType.NonStructural
                            )
                            placed_count += 1
                        except Exception as ex:
                            dbg.warn("  dados falhou inst.Id={}: {}".format(inst.Id, ex))

            # ── Potência em lote: um único Regenerate() para todos ──
            if power_elements:
                _doc.Regenerate()
                for new_elec, pow_str in power_elements:
                    try:
                        _set_power(new_elec, pow_str)
                        power_ok += 1
                    except Exception as ex:
                        power_fail += 1
                        dbg.warn("  potência falhou Id={}: {}".format(new_elec.Id, ex))

            t.Commit()
            dbg.section("Resultado")
            dbg.info("Pontos criados:  {}".format(placed_count))
            dbg.info("Potência OK:     {}".format(power_ok))
            if power_fail:
                dbg.warn("Potência falhou: {}".format(power_fail))
            if skip_count:
                dbg.warn("Instâncias puladas (sem localização): {}".format(skip_count))
            MessageBox.Show("Pontos colocados com sucesso!", "Pontos por Vínculo")

        except Exception as e:
            try:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
            except:
                pass
            dbg.error("Exceção em _on_place: {}".format(e))
            MessageBox.Show("Erro ao colocar pontos:\n" + str(e), "Pontos por Vínculo")


if __name__ == "__main__":
    win = PontosVinculoWindow(os.path.join(os.path.dirname(__file__), 'ui.xaml'))
    win.ShowDialog()