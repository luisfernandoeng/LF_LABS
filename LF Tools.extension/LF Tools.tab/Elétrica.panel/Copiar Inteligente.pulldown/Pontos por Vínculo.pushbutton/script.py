# -*- coding: utf-8 -*-
"""Pontos por Vínculo v6"""
__title__ = "Pontos por\nVínculo"
__author__ = "Luís Fernando"

# ╔══════════════════════════════════════════════════════════════╗
# ║                    MODO DEBUG                                ║
# ║  True  = imprime detalhes no console pyRevit                 ║
# ║  False = silencioso                                          ║
# ╚══════════════════════════════════════════════════════════════╝
DEBUG_MODE = True
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
from System.Windows.Data import CollectionViewSource
from System.Windows.Controls import (
    CheckBox, Grid, ColumnDefinition, RowDefinition, TextBlock, ComboBox, TextBox, Border, ScrollViewer, StackPanel
)
from System.Windows.Media import SolidColorBrush, Color
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    RevitLinkInstance, Level, Transaction, TransactionStatus,
    BuiltInParameter, Phase
)
from Autodesk.Revit.DB.Structure import StructuralType

from pyrevit import forms, script
from lf_utils import DebugLogger

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
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
]

ELEC_CATS = [
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_ConduitFitting,
]

DATA_CATS = [
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_ConduitFitting,
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


def _is_non_electric_smart(display_name, std_cat=""):
    """Usa STD_CATEGORIA se disponível, senão cai para palavras-chave."""
    if std_cat:
        norm = _normalize(std_cat)
        if norm in ("eletrica", "elétrica"):
            return False
        return True
    return _is_non_electric(display_name)


def _get_safe_name(el):
    if not el: return "—"
    try:
        return el.Name
    except:
        pass
    try:
        p = el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.HasValue: return p.AsString()
    except:
        pass
    try:
        return "ID " + str(el.Id)
    except:
        return "Unknown"


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


def _collect_load_classifications(doc, extra_docs=None):
    """Retorna lista ordenada de {name, id, is_native} das Classificações de Carga.
    Tenta LoadClassification nativa; se vazia, coleta valores distintos do
    parâmetro de texto 'Tipo de Carga' no projeto e nos vínculos."""
    result = []
    lc_class = None
    try:
        from Autodesk.Revit.DB.Electrical import LoadClassification
        lc_class = LoadClassification
    except:
        pass
    if lc_class is None:
        try:
            from Autodesk.Revit.DB import LoadClassification
            lc_class = LoadClassification
        except:
            pass
    if lc_class is not None:
        try:
            for lc in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalLoadClassifications).ToElements():
                try:
                    lc_name = ""
                    try:
                        lc_name = lc.Name
                    except:
                        pass
                    if not lc_name:
                        try:
                            p = lc.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                            if p and p.HasValue: lc_name = p.AsString()
                        except: pass
                    if not lc_name:
                        lc_name = "ID " + str(lc.Id)

                    result.append({"name": lc_name, "id": lc.Id, "is_native": True})
                except Exception as e:
                    dbg.error("Erro lendo LoadClassification: {}".format(e))
            result.sort(key=lambda x: x["name"])
        except Exception as e:
            dbg.error("Erro iterando LoadClassifications nativas: {}".format(e))
        if result:
            return result

    # Fallback: coletar valores distintos do parâmetro texto "Tipo de Carga"
    TEXT_SCAN_CATS = [
        BuiltInCategory.OST_ConduitFitting,
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_SpecialityEquipment,
    ]
    seen = set()
    all_docs = [doc] + (list(extra_docs) if extra_docs else [])
    for d in all_docs:
        if not d:
            continue
        for bic in TEXT_SCAN_CATS:
            try:
                for el in FilteredElementCollector(d).OfCategory(bic).WhereElementIsNotElementType():
                    try:
                        p = el.LookupParameter("Tipo de Carga")
                        if p and p.HasValue:
                            val = (p.AsString() or "").strip()
                            if val and val not in seen:
                                seen.add(val)
                                result.append({"name": val, "id": None, "is_native": False})
                    except:
                        pass
            except:
                pass
    result.sort(key=lambda x: x["name"])
    return result


def _set_load_classification(element, lc_info):
    """Atribui Classificação de Carga ao elemento.
    lc_info: {"name": str, "id": ElementId|None, "is_native": bool}"""
    if not lc_info:
        return False
    from Autodesk.Revit.DB import StorageType
    # Abordagem nativa: via ElementId no conector ou BuiltInParameter
    if lc_info.get("is_native") and lc_info.get("id"):
        lc_id = lc_info["id"]
        try:
            from Autodesk.Revit.DB import Domain
            if hasattr(element, 'MEPModel') and element.MEPModel:
                cm = element.MEPModel.ConnectorManager
                if cm:
                    for conn in cm.Connectors:
                        if conn.Domain == Domain.DomainElectrical:
                            try:
                                conn.LoadClassificationId = lc_id
                                return True
                            except:
                                pass
        except:
            pass
        try:
            p = element.get_Parameter(BuiltInParameter.RBS_ELEC_LOAD_CLASSIFICATION)
            if p and not p.IsReadOnly:
                p.Set(lc_id)
                return True
        except:
            pass
            
        try:
            p = element.LookupParameter("Tipo de Carga")
            if p and not p.IsReadOnly:
                if p.StorageType == StorageType.ElementId:
                    p.Set(lc_id)
                    return True
        except:
            pass

    # Abordagem texto: setar parâmetro "Tipo de Carga" por nome
    text_val = lc_info.get("name", "")
    if text_val:
        try:
            p = element.LookupParameter("Tipo de Carga")
            if p and not p.IsReadOnly and p.StorageType == StorageType.String:
                p.Set(text_val)
                return True
        except:
            pass
    return False


# ── Nomes conhecidos do parâmetro de potência nas famílias do projeto ──
_POWER_PARAM_NAMES = [
    "Potência Ativa (W)",
    "Potência Ativa",
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

class PontosVinculoWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self._doc = __revit__.ActiveUIDocument.Document
        self._links = get_link_instances()
        self._host_levels = sorted(
            FilteredElementCollector(self._doc).OfClass(Level),
            key=lambda lv: lv.Elevation
        )
        self._elec_symbols = self._collect_symbols(ELEC_CATS)
        self._data_symbols = self._collect_symbols(DATA_CATS)
        link_docs = [l["link_doc"] for l in self._links if l.get("link_doc")]
        self._load_classifications = _collect_load_classifications(self._doc, link_docs)
        self._family_rows = []
        self._filter_non_electric = True
        self._filter_only_new = True
        self._init_ui()

    def _collect_symbols(self, categories):
        result = []
        seen = set()
        for bic in categories:
            try:
                for s in FilteredElementCollector(self._doc).OfCategory(bic).WhereElementIsElementType().ToElements():
                    try:
                        fam_name = ""
                        try:
                            fam_name = s.FamilyName
                        except:
                            try:
                                fam_name = s.Family.Name
                            except:
                                fam_name = "Unknown"
                        
                        sym_name = ""
                        try:
                            sym_name = s.Name
                        except:
                            pass
                            
                        if not sym_name:
                            try:
                                p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                                if p and p.HasValue: sym_name = p.AsString()
                            except:
                                pass
                                
                        if not sym_name:
                            sym_name = "ID " + str(s.Id)
                            
                        full_name = "{} : {}".format(fam_name, sym_name)
                        if full_name not in seen:
                            seen.add(full_name)
                            result.append({"full_name": full_name, "display": sym_name, "symbol": s})
                    except Exception as e:
                        dbg.error("Erro no simbolo {}: {}".format(getattr(s, 'Id', '?'), e))
            except Exception as e:
                dbg.error("Erro iterando categoria {}: {}".format(bic, e))
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
            forms.alert("Nenhum vínculo RVT carregado no projeto.", title="Pontos por Vínculo")
            return

        net_links = List[System.Object]()
        for lnk in self._links:
            net_links.Add(lnk["display"])
        self.cb_Link.ItemsSource = net_links
        self.cb_Link.SelectedIndex = 0
        self.cb_Link.SelectionChanged += self._on_link_changed
        self.cb_Link.KeyUp += self._on_cb_keyup

        net_levels = List[System.Object]()
        for lv in self._host_levels:
            net_levels.Add(lv.Name)
        self.cb_DefaultLevel.ItemsSource = net_levels
        if self._host_levels:
            self.cb_DefaultLevel.SelectedIndex = 0

        self.chk_FilterNonElec.IsChecked = True
        self.chk_FilterNonElec.Checked   += self._on_filter_toggle
        self.chk_FilterNonElec.Unchecked += self._on_filter_toggle

        self.cb_Phase.KeyUp += self._on_cb_keyup
        self._on_link_changed(None, None)

        self.btn_Scan.Click       += self._on_scan
        self.btn_Place.Click      += self._on_place
        self.btn_Cancel.Click     += lambda s, a: self.Close()
        self.btn_SelectAll.Click  += lambda s, a: self._set_all_checked(True)
        self.btn_SelectNone.Click += lambda s, a: self._set_all_checked(False)

    def _on_link_changed(self, sender, args):
        idx = self.cb_Link.SelectedIndex
        if idx < 0:
            self.cb_Phase.ItemsSource = None
            return
        link_doc = self._links[idx]["link_doc"]
        if link_doc:
            try:
                phases = list(FilteredElementCollector(link_doc).OfClass(Phase))
                net_phases = List[System.Object]()
                net_phases.Add("Todas as Fases")
                for p in phases:
                    net_phases.Add(p.Name)
                self.cb_Phase.ItemsSource = net_phases
                if phases:
                    self.cb_Phase.SelectedIndex = len(phases)
                else:
                    self.cb_Phase.SelectedIndex = 0
            except:
                self.cb_Phase.ItemsSource = None

    def _on_cb_keyup(self, sender, args):
        try:
            # Ignorar teclas de navegação para permitir o uso normal do combobox
            if str(args.Key) in ("Up", "Down", "Left", "Right", "Enter", "Tab", "Escape"):
                return
                
            tb = sender.Template.FindName("PART_EditableTextBox", sender)
            if not tb: return
            
            txt = tb.Text
            caret = tb.CaretIndex
            txt_lower = txt.lower()
            
            view = CollectionViewSource.GetDefaultView(sender.ItemsSource)
            if not view: return
            
            def filter_func(item):
                if not txt_lower: return True
                return txt_lower in str(item).lower()
                
            view.Filter = System.Predicate[System.Object](filter_func)
            view.Refresh()
            sender.IsDropDownOpen = True
            
            # Restaurar texto e posição do cursor que se perdem no Refresh do WPF
            tb.Text = txt
            tb.CaretIndex = caret
        except:
            pass

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

        selected_phase = self.cb_Phase.Text.strip() if self.cb_Phase.Text else ""
        filter_phase = bool(selected_phase) and selected_phase != "Todas as Fases"

        mapa = {}
        for bic in SCAN_CATS:
            try:
                count_cat = 0
                for el in FilteredElementCollector(link_doc).OfCategory(bic).WhereElementIsNotElementType():
                    # Filtro de fase: ignorar elementos se marcado
                    if filter_phase:
                        try:
                            phase_p = el.get_Parameter(BuiltInParameter.PHASE_CREATED)
                            if phase_p and phase_p.HasValue:
                                phase_name = phase_p.AsValueString() or ""
                                if selected_phase.lower() not in phase_name.lower():
                                    continue
                        except:
                            pass

                    fam, type_n = "", ""
                    std_cat = ""
                    load_type_str = ""
                    try:
                        _sym = el.Symbol
                        fam, type_n = _sym.Family.Name, _sym.Name
                        try:
                            p_cat = _sym.LookupParameter("STD_CATEGORIA")
                            if p_cat and p_cat.HasValue:
                                std_cat = (p_cat.AsString() or "").strip()
                        except:
                            pass
                    except:
                        fam = el.Name
                    try:
                        p_lt = el.LookupParameter("Tipo de Carga")
                        if p_lt and p_lt.HasValue:
                            load_type_str = (p_lt.AsString() or "").strip()
                    except:
                        pass
                    display = "{} : {}".format(fam, type_n) if type_n else fam
                    if display not in mapa:
                        mapa[display] = {"display": display, "instances": [], "std_cat": std_cat, "load_type": load_type_str}
                    elif not mapa[display].get("load_type") and load_type_str:
                        mapa[display]["load_type"] = load_type_str
                    mapa[display]["instances"].append(el)
                    count_cat += 1
                dbg.debug("Categoria {}: {} instâncias".format(str(bic).split('.')[-1], count_cat))
            except:
                pass

        # Atualizar classificações de carga com valores encontrados no vínculo
        lc_from_scan = {}
        for key, fam_data in mapa.items():
            lt = fam_data.get("load_type", "")
            if lt and lt not in lc_from_scan:
                lc_from_scan[lt] = {"name": lt, "id": None, "is_native": False}
        if lc_from_scan:
            # Mescla com classificações nativas (se houver); scan tem prioridade para novos valores
            existing_names = set(lc["name"] for lc in self._load_classifications)
            for name, lc_obj in lc_from_scan.items():
                if name not in existing_names:
                    self._load_classifications.append(lc_obj)
            self._load_classifications.sort(key=lambda x: x["name"])

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
        """ComboBox com estilo e Autocomplete com filtro."""
        cb = ComboBox()
        cb.IsEditable = True  # Ativa campo digitável
        cb.StaysOpenOnEdit = True # Mantém aberto ao digitar
        
        cb.Height = 28
        cb.Margin = Thickness(2, 2, 2, 2)
        cb.Padding = Thickness(6, 0, 0, 0)
        cb.VerticalContentAlignment = VerticalAlignment.Center

        # ── Visual: light theme ─────────────────────────────────────────
        INPUT_BG   = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        BORDER_CLR = SolidColorBrush(Color.FromRgb(0xC8, 0xC8, 0xC8))
        FG_CLR     = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
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
        
        # Adiciona o evento de pesquisa
        cb.KeyUp += self._on_cb_keyup

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

        widths = [(28, 0), (160, 0), (36, 0), (90, 0), (0, 2), (0, 2), (72, 0)]
        for w, star in widths:
            cd = ColumnDefinition()
            cd.Width = GridLength(star, GridUnitType.Star) if star > 0 else GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        # CheckBox de seleção
        cb = CheckBox()
        cb.IsChecked = True
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin = Thickness(4, 0, 0, 0)
        cb.Foreground = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
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
        tb_name.Foreground = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
        Grid.SetColumn(tb_name, 1)

        # Quantidade
        tb_qty = TextBlock()
        tb_qty.Text = str(len(fam["instances"]))
        tb_qty.VerticalAlignment = VerticalAlignment.Center
        tb_qty.HorizontalAlignment = HorizontalAlignment.Center
        tb_qty.Foreground = SolidColorBrush(Color.FromRgb(0x77, 0x77, 0x77))
        Grid.SetColumn(tb_qty, 2)

        # ComboBox — Tipo de Carga
        INPUT_BG   = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        BORDER_CLR = SolidColorBrush(Color.FromRgb(0xC8, 0xC8, 0xC8))
        FG_CLR     = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
        c_load = ComboBox()
        c_load.IsEditable = True
        c_load.StaysOpenOnEdit = True
        c_load.KeyUp += self._on_cb_keyup
        c_load.Height = 28
        c_load.Margin = Thickness(2, 2, 2, 2)
        c_load.Padding = Thickness(6, 0, 0, 0)
        c_load.VerticalContentAlignment = VerticalAlignment.Center
        c_load.Background  = INPUT_BG
        c_load.Foreground  = FG_CLR
        c_load.BorderBrush = BORDER_CLR
        c_load.BorderThickness = Thickness(1)
        load_items = List[System.Object]()
        load_items.Add("")
        for lc in self._load_classifications:
            load_items.Add(lc["name"])
        c_load.ItemsSource = load_items
        # Pré-seleciona o valor lido da instância no vínculo
        fam_load = fam.get("load_type", "")
        load_default = 0
        if fam_load:
            for i, lc in enumerate(self._load_classifications):
                if lc["name"] == fam_load:
                    load_default = i + 1
                    break
            if load_default == 0:
                c_load.Text = fam_load
        c_load.SelectedIndex = load_default
        Grid.SetColumn(c_load, 3)

        # ComboBox — Ponto Elétrico
        elec_default = self._find_match_idx(fam["display"], self._elec_symbols)
        c_elec = self._make_editable_combo(self._elec_symbols, elec_default)
        Grid.SetColumn(c_elec, 4)

        # ComboBox — Ponto de Dados
        c_dados = self._make_editable_combo(self._data_symbols, 0)
        Grid.SetColumn(c_dados, 5)

        # TextBox — Potência (VA)
        t_pow = TextBox()
        t_pow.Height = 26
        t_pow.Margin = Thickness(2)
        t_pow.VerticalContentAlignment = VerticalAlignment.Center
        t_pow.Background   = INPUT_BG
        t_pow.Foreground   = FG_CLR
        t_pow.BorderBrush  = BORDER_CLR
        t_pow.BorderThickness = Thickness(1)
        t_pow.Padding = Thickness(4, 0, 4, 0)
        Grid.SetColumn(t_pow, 6)

        for child in [cb, tb_name, tb_qty, c_load, c_elec, c_dados, t_pow]:
            grid.Children.Add(child)

        return {
            "grid": grid,
            "cb": cb,
            "c_load": c_load,
            "c_elec": c_elec,
            "c_dados": c_dados,
            "t_pow": t_pow,
            "instances": fam["instances"],
            "is_non_electric": _is_non_electric_smart(fam["display"], fam.get("std_cat", "")),
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
            lc_idx = r["c_load"].SelectedIndex
            lc_obj = None
            if lc_idx > 0 and lc_idx - 1 < len(self._load_classifications):
                lc_obj = self._load_classifications[lc_idx - 1]
            to_place.append({
                "instances": r["instances"],
                "elec_sym":  elec_sym,
                "dados_sym": dados_sym,
                "pow":       r["t_pow"].Text.strip() or None,
                "load_classification": lc_obj,
            })

        if not to_place:
            forms.alert("Nenhuma linha com ponto elétrico ou de dados selecionada.", title="Pontos por Vínculo")
            return

        dbg.section("Pontos por Vínculo — Colocação")
        dbg.info("Famílias a processar: {}".format(len(to_place)))

        self.Close()

        _doc = self._doc
        t = Transaction(_doc, "Pontos por Vínculo")
        try:
            status = t.Start()
            if status != TransactionStatus.Started:
                forms.alert("Não foi possível iniciar a transação (status: {}).".format(status), title="Pontos por Vínculo")
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
                elec_name  = _get_safe_name(item["elec_sym"]) if item["elec_sym"]  else "—"
                dados_name = _get_safe_name(item["dados_sym"]) if item["dados_sym"] else "—"
                
                fam_name = "?"
                if item["instances"]:
                    try:
                        fam_name = item["instances"][0].Symbol.FamilyName
                    except:
                        try:
                            fam_name = item["instances"][0].Symbol.Family.Name
                        except:
                            pass

                dbg.sub("{} ({} inst.) → elec:{} dados:{} pot:{}".format(
                    fam_name,
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
                            if item.get("load_classification"):
                                _set_load_classification(new_elec, item["load_classification"])
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
            forms.alert("Pontos colocados com sucesso!", title="Pontos por Vínculo")

        except Exception as e:
            try:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
            except:
                pass
            dbg.error("Exceção em _on_place: {}".format(e))
            forms.alert("Erro ao colocar pontos:\n" + str(e), title="Pontos por Vínculo")


if __name__ == "__main__":
    win = PontosVinculoWindow(script.get_bundle_file('ui.xaml'))
    win.ShowDialog()