# -*- coding: utf-8 -*-
"""Pontos por Vínculo v6"""
__title__ = "Pontos por\nVínculo"
__author__ = "Luís Fernando"

# ╔══════════════════════════════════════════════════════════════╗
# ║                    MODO DEBUG                                ║
# ║  True  = imprime detalhes no console pyRevit                 ║
# ║  False = silencioso                                          ║
# ╚══════════════════════════════════════════════════════════════╝
DEBUG_MODE = False  # padrão; o checkbox na UI sobrescreve e persiste
import os
import io
import json
import re
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
    BuiltInParameter, Phase,
    View3D, ReferenceIntersector, FindReferenceTarget,
    ElementCategoryFilter, XYZ, Line, ElementTransformUtils,
    FamilyPlacementType,
)
from Autodesk.Revit.DB.Structure import StructuralType

from pyrevit import forms, script
from lf_utils import DebugLogger, get_script_config, save_script_config
import auto_eletrica
import profile_manager

_debug_cfg = get_script_config(__file__, {'debug': DEBUG_MODE})
dbg = DebugLogger(_debug_cfg.get('debug', DEBUG_MODE))

# ==================== Perfis JSON ====================

PROFILES_DIR = os.path.join(os.path.dirname(__file__), 'profiles')


def _ensure_profiles_dir():
    if not os.path.isdir(PROFILES_DIR):
        os.makedirs(PROFILES_DIR)


def _list_profiles():
    try:
        _ensure_profiles_dir()
        return sorted(f[:-5] for f in os.listdir(PROFILES_DIR) if f.endswith('.json'))
    except:
        return []


def _profile_path(name):
    return os.path.join(PROFILES_DIR, name + '.json')


def _load_profile(name):
    """Carrega perfil do JSON. Retorna dict {display_name: {campos}} ou {}."""
    try:
        path = _profile_path(name)
        with io.open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}


def _save_profile(name, data):
    """Salva dict de perfil em JSON. Retorna True se OK."""
    try:
        _ensure_profiles_dir()
        path = _profile_path(name)
        with io.open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2,
                      sort_keys=True)
        return True
    except:
        return False


def _safe_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()


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

# ==================== Face Placement Helpers ====================

def _get_3d_view(doc):
    """Retorna a primeira View3D não-template disponível para ReferenceIntersector."""
    try:
        for v in FilteredElementCollector(doc).OfClass(View3D):
            if not v.IsTemplate:
                return v
    except:
        pass
    return None


def _is_luminaire_sym(symbol):
    """Retorna True se o símbolo for da categoria Luminárias (OST_LightingFixtures)."""
    try:
        return symbol.Category.Id.IntegerValue == int(BuiltInCategory.OST_LightingFixtures)
    except:
        return False


# ==================== Circuit helper ====================

def _read_circuit_info(el, _doc, circuit_cache):
    """
    Lê potência individual, descrição e número do circuito de um elemento.
    Usa circuit_cache {circuit_id → info} para evitar reprocessar o mesmo circuito.
    Retorna dict ou None se o elemento não tiver circuito.
    """
    try:
        systems = list(el.MEPModel.ElectricalSystems)
    except Exception:
        return None
    if not systems:
        return None

    for sys in systems:
        try:
            sid = sys.Id
            if sid in circuit_cache:
                return circuit_cache[sid]

            info = {"potencia_individual": 0.0, "descricao": u"", "nome": u""}

            try:
                info["nome"] = sys.CircuitNumber or u""
            except Exception:
                pass
            for dname in [u"Descrição", u"Descricao", u"Description"]:
                try:
                    p = sys.LookupParameter(dname)
                    if p and p.HasValue:
                        info["descricao"] = (p.AsString() or u"").strip()
                        if info["descricao"]:
                            break
                except Exception:
                    pass

            # Potência individual: carga aparente total / nº de elementos no circuito
            try:
                elems = sys.Elements
                count = elems.Size if elems else 0
                if count > 0:
                    try:
                        total_va = sys.ApparentLoad
                        info["potencia_individual"] = total_va / count
                    except Exception:
                        pass
            except Exception:
                pass

            circuit_cache[sid] = info
            return info
        except Exception:
            pass
    return None


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
        self._elec_face_symbols = [s for s in self._elec_symbols if s.get('is_face')]
        self._data_face_symbols = [s for s in self._data_symbols if s.get('is_face')]
        self._family_rows = []
        self._filter_non_electric = True
        self._filter_only_new = True
        self._profile = {}
        self._syncing_profile = False
        self._init_ui()
        self._ae = auto_eletrica.AutoEletricaController(
            win=self,
            doc=self._doc,
            uidoc=__revit__.ActiveUIDocument,
            dbg=dbg,
            profiles_dir=PROFILES_DIR,
        )
        load_types = []
        try:
            load_types = auto_eletrica._get_load_types(self._doc)
        except Exception:
            pass
        self._gp = profile_manager.ProfileManagerController(
            win=self,
            dbg=dbg,
            profiles_dir=PROFILES_DIR,
            load_types=load_types,
        )

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
                            is_face = False
                            try:
                                fpt = s.Family.FamilyPlacementType
                                is_face = fpt in (FamilyPlacementType.WorkPlaneBased,
                                                  FamilyPlacementType.OneLevelBasedHosted)
                            except:
                                pass
                            result.append({"full_name": full_name, "display": sym_name,
                                           "symbol": s, "is_face": is_face})
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

        self._refresh_profiles()
        self.btn_LoadProfile.Click      += self._on_load_profile
        self.btn_SaveProfile.Click      += self._on_save_profile
        self.btn_ReadProject.Click      += self._on_read_project
        self.cb_Profile.SelectionChanged += self._on_ppv_profile_combo_changed

        self.chk_Debug.IsChecked = dbg.enabled
        self.chk_Debug.Checked   += self._on_debug_toggle
        self.chk_Debug.Unchecked += self._on_debug_toggle

        face_cfg = _debug_cfg.get('face_placement', False)
        self.chk_FacePlacement.IsChecked = face_cfg
        self.chk_FacePlacement.Checked   += self._on_face_toggle
        self.chk_FacePlacement.Unchecked += self._on_face_toggle

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

    # ── Debug ────────────────────────────────────────────────────────────

    def _on_debug_toggle(self, _sender, _args):
        dbg.enabled = bool(self.chk_Debug.IsChecked)
        save_script_config(__file__, {
            'debug':         dbg.enabled,
            'face_placement': bool(self.chk_FacePlacement.IsChecked),
        })
        if dbg.enabled:
            dbg.section(u"Debug ativado")

    def _update_row_face_combos(self, row, use_face):
        """Troca ItemsSource dos combos elétrico/dados para a lista face ou completa."""
        pairs = [
            ("c_elec",  self._elec_face_symbols  if use_face else self._elec_symbols),
            ("c_dados", self._data_face_symbols   if use_face else self._data_symbols),
            ]
        sym_keys = ["elec_sym_list", "dados_sym_list"]
        for (combo_key, new_list), sym_key in zip(pairs, sym_keys):
            cb = row[combo_key]
            old_list = row[sym_key]
            # Guarda texto selecionado antes de trocar
            cur_text = u""
            if cb.SelectedIndex > 0:
                idx = cb.SelectedIndex - 1
                if idx < len(old_list):
                    cur_text = old_list[idx]["display"]
            elif cb.Text and cb.Text != u"— Não colocar —":
                cur_text = cb.Text

            net = List[System.Object]()
            net.Add(u"— Não colocar —")
            for s in new_list:
                net.Add(s["display"])
            cb.ItemsSource = net

            # Tenta restaurar seleção
            cb.SelectedIndex = 0
            if cur_text:
                for i, s in enumerate(new_list):
                    if s["display"].lower() == cur_text.lower():
                        cb.SelectedIndex = i + 1
                        break

            row[sym_key] = new_list

    def _on_face_toggle(self, _sender, _args):
        use_face = bool(self.chk_FacePlacement.IsChecked)
        save_script_config(__file__, {
            'debug':          dbg.enabled,
            'face_placement': use_face,
        })
        for row in self._family_rows:
            self._update_row_face_combos(row, use_face)

    def _find_nearest_wall_face(self, ri, pt, max_dist_ft=6.56):
        """
        Lança raios horizontais (4 direções cardeais) a partir de pt.
        Retorna (Reference, XYZ ponto_na_face, XYZ direção_do_raio) da parede mais próxima,
        ou (None, None, None) se nenhuma parede encontrada dentro de max_dist_ft.
        max_dist_ft padrão ≈ 2 m.
        """
        dirs = [XYZ(1, 0, 0), XYZ(-1, 0, 0), XYZ(0, 1, 0), XYZ(0, -1, 0)]
        best_ref  = None
        best_pt   = None
        best_dir  = None
        best_dist = max_dist_ft

        for d in dirs:
            try:
                results = ri.Find(pt, d)
                if not results:
                    continue
                for r in results:
                    if r.Proximity < best_dist:
                        best_dist = r.Proximity
                        best_ref  = r.GetReference()
                        best_dir  = d
                        # Projeta o ponto original na face (mantém Z original)
                        best_pt   = XYZ(
                            pt.X + d.X * r.Proximity,
                            pt.Y + d.Y * r.Proximity,
                            pt.Z,
                        )
            except:
                pass

        return best_ref, best_pt, best_dir

    # ── Perfis ──────────────────────────────────────────────────────────

    def _on_ppv_profile_combo_changed(self, sender, args):
        if self._syncing_profile:
            return
        self._syncing_profile = True
        try:
            idx = self.cb_Profile.SelectedIndex
            for combo in [self.ae_CmbProfile, self.gp_CmbProfile]:
                try:
                    if combo.SelectedIndex != idx:
                        combo.SelectedIndex = idx
                except Exception:
                    pass
        finally:
            self._syncing_profile = False

    def _refresh_profiles(self, select_name=None):
        names = _list_profiles()
        net = List[System.Object]()
        net.Add(u"— Selecionar perfil —")
        for n in names:
            net.Add(n)
        self._syncing_profile = True
        try:
            self.cb_Profile.ItemsSource = net
            idx = 0
            if select_name:
                try:
                    idx = names.index(select_name) + 1
                except ValueError:
                    idx = 0
            self.cb_Profile.SelectedIndex = idx
        finally:
            self._syncing_profile = False

    def _refresh_all_profiles(self, select_name=None):
        """Re-popula os 3 combos de perfil e seleciona select_name."""
        self._refresh_profiles(select_name)
        try:
            self._ae._refresh_ae_profiles(select_name)
        except Exception:
            pass
        try:
            self._gp._refresh_combo()
            if select_name:
                names = _list_profiles()
                try:
                    idx = names.index(select_name) + 1
                except ValueError:
                    idx = 0
                self._syncing_profile = True
                try:
                    self.gp_CmbProfile.SelectedIndex = idx
                finally:
                    self._syncing_profile = False
        except Exception:
            pass

    def _on_load_profile(self, sender, args):
        idx = self.cb_Profile.SelectedIndex
        if idx <= 0:
            forms.alert(u"Selecione um perfil antes de carregar.", title=u"Perfil")
            return
        names = _list_profiles()
        if idx - 1 >= len(names):
            return
        name = names[idx - 1]
        self._profile = _load_profile(name)
        if not self._profile:
            forms.alert(u"Perfil '{}' está vazio ou corrompido.".format(name), title=u"Perfil")
            return
        self._apply_profile_to_rows()
        # Aplica também na aba AE
        try:
            if hasattr(self, '_ae') and self._ae:
                for w in self._ae._widgets:
                    entry = self._find_profile_entry(w.family_name)
                    if entry:
                        w.apply_profile(entry)
                for i in range(self.ae_CmbProfile.Items.Count):
                    if str(self.ae_CmbProfile.Items[i]) == name:
                        self.ae_CmbProfile.SelectedIndex = i
                        break
        except Exception:
            pass
        self.lbl_Status.Text = u"Perfil '{}' carregado ({} entrada(s)).".format(
            name, len(self._profile))

    def _on_save_profile(self, sender, args):
        if not self._family_rows:
            forms.alert(u"Escaneie o vínculo antes de salvar o perfil.", title=u"Perfil")
            return
        name = forms.ask_for_string(
            prompt=u"Nome do perfil:",
            title=u"Salvar Perfil",
            default=u"Meu Perfil"
        )
        if not name:
            return
        name = _safe_filename(name)

        data = {}
        for row in self._family_rows:
            display = row.get("display", u"")
            if not display:
                continue

            # ponto_eletrico — prefere display name do símbolo selecionado
            pe = u""
            try:
                elec_list = row.get("elec_sym_list", self._elec_symbols)
                pe_idx = row["c_elec"].SelectedIndex
                if pe_idx > 0 and pe_idx - 1 < len(elec_list):
                    pe = elec_list[pe_idx - 1]["display"]
                else:
                    pe = (row["c_elec"].Text or u"").strip()
                    if pe == u"— Não colocar —":
                        pe = u""
            except:
                pass

            # ponto_dados
            pd_val = u""
            try:
                data_list = row.get("dados_sym_list", self._data_symbols)
                pd_idx = row["c_dados"].SelectedIndex
                if pd_idx > 0 and pd_idx - 1 < len(data_list):
                    pd_val = data_list[pd_idx - 1]["display"]
                else:
                    pd_val = (row["c_dados"].Text or u"").strip()
                    if pd_val == u"— Não colocar —":
                        pd_val = u""
            except:
                pass

            qty_pts = 1
            try:
                qty_pts = max(1, int(str(row["tb_pts"].Text).strip() or u"1"))
            except Exception:
                qty_pts = 1
            qty_dados_pts = 1
            try:
                qty_dados_pts = max(1, int(str(row["tb_dados_pts"].Text).strip() or u"1"))
            except Exception:
                qty_dados_pts = 1
            data[display] = {
                "ponto_eletrico":   pe,
                "ponto_dados":      pd_val,
                "checked":          bool(row.get("cb") and row["cb"].IsChecked),
                "qty_per_inst":     qty_pts,
                "qty_dados_per_inst": qty_dados_pts,
            }

        # Mescla dados AE (altura, carga_va, tensao, prefixo, tipo_carga)
        try:
            if hasattr(self, '_ae') and self._ae and self._ae._widgets:
                for w in self._ae._widgets:
                    ae_vals = w.get_values()
                    fam_lower = w.family_name.lower()
                    matched_key = None
                    for k in data:
                        if k.split(u' : ')[0].strip().lower() == fam_lower:
                            matched_key = k
                            break
                    if matched_key is None:
                        matched_key = w.family_name
                        data[matched_key] = {}
                    data[matched_key].update(ae_vals)
        except Exception:
            pass

        if _save_profile(name, data):
            self._refresh_all_profiles(name)
            forms.toast(u"Perfil '{}' salvo com {} elemento(s)!".format(name, len(data)))
        else:
            forms.alert(u"Erro ao salvar o perfil.", title=u"Perfil")

    def _on_read_project(self, _sender, _args):
        """
        Lê elementos elétricos já colocados no projeto.

        Se um vínculo estiver selecionado em cb_Link, usa-o como fonte de arquitetura:
        para cada família do vínculo, procura elementos de projeto próximos (≤ 50 cm)
        e lê seus parâmetros elétricos (tipo_carga, potência, tensão, altura).
        Caso contrário, varre diretamente os elementos do projeto.
        """
        READ_CATS = [
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_SpecialityEquipment,
        ]

        levels = list(FilteredElementCollector(self._doc).OfClass(Level))

        def _read_elec_params(el):
            """Lê tipo_carga, tensão, altura (m) e display do símbolo de um elemento."""
            result = {"tipo_carga": u"", "tensao": u"", "altura_m": u"", "sym_display": u""}
            try:
                p = el.LookupParameter(u"Tipo de Carga")
                if p and p.HasValue:
                    result["tipo_carga"] = (p.AsString() or u"").strip()
            except: pass
            for vname in [u"Tensão", u"Tensao", u"Voltage", u"Voltagem"]:
                try:
                    p = el.LookupParameter(vname)
                    if p and p.HasValue:
                        v = p.AsValueString() or p.AsString() or u""
                        if v:
                            result["tensao"] = v.strip()
                            break
                except: pass
            try:
                pt = el.Location.Point
                if levels:
                    lvl = min(levels, key=lambda l: abs(l.Elevation - pt.Z))
                    offset_m = (pt.Z - lvl.Elevation) * 0.3048
                    result["altura_m"] = u"{:.2f}".format(offset_m)
            except: pass
            try:
                sym = el.Symbol
                fam = sym.Family.Name
                typ = sym.Name
                result["sym_display"] = u"{} : {}".format(fam, typ) if typ else fam
            except: pass
            return result

        # ── Modo 1: vínculo de arquitetura selecionado ───────────────────
        link_idx = self.cb_Link.SelectedIndex
        use_arch = (link_idx >= 0 and len(self._links) > 0)

        if use_arch:
            arch      = self._links[link_idx]
            link_doc  = arch["link_doc"]
            transform = arch["link_inst"].GetTotalTransform()
            TOLE_FT   = 0.50 / 0.3048  # 50 cm

            # Coleta espacial de todos os elementos elétricos do projeto
            proj_elec = []
            for bic in READ_CATS:
                try:
                    for el in FilteredElementCollector(self._doc).OfCategory(bic).WhereElementIsNotElementType():
                        try:
                            proj_elec.append({"el": el, "pt": el.Location.Point})
                        except: pass
                except: pass

            mapa = {}  # display_name_arch → acumulador
            for bic in SCAN_CATS:
                try:
                    for el in FilteredElementCollector(link_doc).OfCategory(bic).WhereElementIsNotElementType():
                        try:
                            sym = el.Symbol
                            fam, typ = sym.Family.Name, sym.Name
                            display = u"{} : {}".format(fam, typ) if typ else fam

                            pt_local = el.Location.Point
                            pt_world = transform.OfPoint(pt_local)

                            # Elemento de projeto mais próximo dentro da tolerância
                            best_el   = None
                            best_dist = TOLE_FT
                            for pe in proj_elec:
                                d = pt_world.DistanceTo(pe["pt"])
                                if d < best_dist:
                                    best_dist = d
                                    best_el   = pe["el"]

                            if display not in mapa:
                                mapa[display] = {
                                    "arch_instances":  [],
                                    "proj_instances":  [],
                                    "tipo_carga": u"", "tensao": u"",
                                    "altura_m": u"", "sym_display": u"",
                                    "potencia_media": 0.0,
                                    "circuito_descricao": u"", "circuito_nome": u"",
                                }
                            mapa[display]["arch_instances"].append(el)
                            if best_el:
                                mapa[display]["proj_instances"].append(best_el)
                                ep = _read_elec_params(best_el)
                                for k in ["tipo_carga", "tensao", "altura_m", "sym_display"]:
                                    if not mapa[display][k] and ep[k]:
                                        mapa[display][k] = ep[k]
                        except: pass
                except: pass

            # Enriquecer com dados de circuito
            circuit_cache = {}
            from collections import Counter as _Counter
            for display, entry in mapa.items():
                infos = []
                for el in entry["proj_instances"]:
                    ci = _read_circuit_info(el, self._doc, circuit_cache)
                    if ci: infos.append(ci)
                if infos:
                    pots = [c["potencia_individual"] for c in infos if c["potencia_individual"] > 0]
                    entry["potencia_media"] = sum(pots) / len(pots) if pots else 0.0
                    descs = [c["descricao"] for c in infos if c["descricao"]]
                    nomes = [c["nome"]      for c in infos if c["nome"]]
                    entry["circuito_descricao"] = _Counter(descs).most_common(1)[0][0] if descs else u""
                    entry["circuito_nome"]      = _Counter(nomes).most_common(1)[0][0] if nomes else u""

            com_matched = sum(1 for v in mapa.values() if v["proj_instances"])
            msg = (
                u"Leitura concluída! (modo vínculo de arquitetura)\n\n"
                u"  Tipos no vínculo:            {}\n"
                u"  Com ponto já aplicado:       {}\n\n"
                u"Deseja salvar como perfil?"
            ).format(len(mapa), com_matched)

        else:
            # ── Modo 2: varrer projeto diretamente ───────────────────────
            mapa = {}
            for bic in READ_CATS:
                try:
                    for el in FilteredElementCollector(self._doc).OfCategory(bic).WhereElementIsNotElementType():
                        try:
                            sym = el.Symbol
                            fam, typ = u"", u""
                            try: fam = sym.Family.Name
                            except: pass
                            try: typ = sym.Name
                            except: pass
                            if not fam:
                                try: fam = el.Name
                                except: fam = u"ID_{}".format(el.Id.IntegerValue)
                            display = u"{} : {}".format(fam, typ) if typ else fam

                            if display not in mapa:
                                mapa[display] = {
                                    "proj_instances": [],
                                    "arch_instances": [],
                                    "tipo_carga": u"", "tensao": u"",
                                    "altura_m": u"", "sym_display": display,
                                    "potencia_media": 0.0,
                                    "circuito_descricao": u"", "circuito_nome": u"",
                                }
                            mapa[display]["proj_instances"].append(el)
                            ep = _read_elec_params(el)
                            for k in ["tipo_carga", "tensao", "altura_m"]:
                                if not mapa[display][k] and ep[k]:
                                    mapa[display][k] = ep[k]
                        except: pass
                except: pass

            if not mapa:
                forms.alert(u"Nenhum elemento elétrico encontrado no projeto.", title=u"Ler Projeto")
                return

            circuit_cache = {}
            from collections import Counter as _Counter
            for display, entry in mapa.items():
                infos = []
                for el in entry["proj_instances"]:
                    ci = _read_circuit_info(el, self._doc, circuit_cache)
                    if ci: infos.append(ci)
                if infos:
                    pots = [c["potencia_individual"] for c in infos if c["potencia_individual"] > 0]
                    entry["potencia_media"] = sum(pots) / len(pots) if pots else 0.0
                    descs = [c["descricao"] for c in infos if c["descricao"]]
                    nomes = [c["nome"]      for c in infos if c["nome"]]
                    entry["circuito_descricao"] = _Counter(descs).most_common(1)[0][0] if descs else u""
                    entry["circuito_nome"]      = _Counter(nomes).most_common(1)[0][0] if nomes else u""

            com_pot  = sum(1 for v in mapa.values() if v["potencia_media"] > 0)
            msg = (
                u"Leitura concluída!\n\n"
                u"  Tipos encontrados:        {}\n"
                u"  Com potência calculada:   {}\n\n"
                u"Deseja salvar como perfil?"
            ).format(len(mapa), com_pot)

        if not mapa:
            forms.alert(u"Nenhum elemento encontrado.", title=u"Ler Projeto")
            return

        if not forms.alert(msg, title=u"Ler Projeto", yes=True, no=True):
            return

        # ── Montar profile dict ──────────────────────────────────────────
        profile_data = {}
        for display, entry in sorted(mapa.items()):
            pot = entry.get("potencia_media", 0.0)
            sym_disp = entry.get("sym_display") or display
            profile_data[display] = {
                "tipo_carga":         entry.get("tipo_carga", u""),
                "tensao":             entry.get("tensao", u""),
                "altura":             entry.get("altura_m", u""),
                "ponto_eletrico":     sym_disp,
                "ponto_dados":        u"",
                "potencia":           u"{:.0f}".format(pot) if pot > 0 else u"",
                "circuito_descricao": entry.get("circuito_descricao", u""),
                "circuito_nome":      entry.get("circuito_nome", u""),
                "qtd_encontrados":    len(entry["proj_instances"]),
            }

        name = forms.ask_for_string(
            prompt=u"Nome do perfil:",
            title=u"Salvar Perfil Lido",
            default=u"Lido do Projeto"
        )
        if not name:
            return
        name = _safe_filename(name)

        if _save_profile(name, profile_data):
            self._refresh_all_profiles(name)
            forms.toast(u"Perfil '{}' gerado com {} tipo(s)!".format(name, len(profile_data)))
        else:
            forms.alert(u"Erro ao salvar o perfil.", title=u"Ler Projeto")

    def _find_profile_entry(self, display):
        """Busca entrada no perfil: exato → por família (antes do ' : ')."""
        disp_low = display.lower()
        for k, v in self._profile.items():
            if k.lower() == disp_low:
                return v
        fam_part = display.split(u" : ")[0].strip().lower()
        for k, v in self._profile.items():
            if k.split(u" : ")[0].strip().lower() == fam_part:
                return v
        return None

    def _apply_profile_to_rows(self):
        for row in self._family_rows:
            entry = self._find_profile_entry(row.get("display", u""))
            if entry:
                self._apply_entry_to_row(row, entry)

    def _apply_entry_to_row(self, row, entry):
        # ponto_eletrico
        pe = entry.get("ponto_eletrico", u"")
        if pe:
            elec_list = row.get("elec_sym_list", self._elec_symbols)
            matched = False
            for i, sym in enumerate(elec_list):
                if (sym["display"].lower() == pe.lower() or
                        sym.get("full_name", "").lower() == pe.lower()):
                    row["c_elec"].SelectedIndex = i + 1
                    matched = True
                    break
            if not matched:
                row["c_elec"].Text = pe

        # ponto_dados
        pd_val = entry.get("ponto_dados", u"")
        if pd_val:
            data_list = row.get("dados_sym_list", self._data_symbols)
            matched = False
            for i, sym in enumerate(data_list):
                if (sym["display"].lower() == pd_val.lower() or
                        sym.get("full_name", "").lower() == pd_val.lower()):
                    row["c_dados"].SelectedIndex = i + 1
                    matched = True
                    break
            if not matched:
                row["c_dados"].Text = pd_val

        if u"checked" in entry:
            row["cb"].IsChecked = bool(entry[u"checked"])

        qty = entry.get(u"qty_per_inst", 1)
        try:
            row["tb_pts"].Text = str(max(1, int(qty)))
        except Exception:
            pass

        qty_dados = entry.get(u"qty_dados_per_inst", 1)
        try:
            row["tb_dados_pts"].Text = str(max(1, int(qty_dados)))
        except Exception:
            pass

    # ── Scan ────────────────────────────────────────────────────────────

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

        for key in sorted(mapa.keys()):
            row = self._build_row(mapa[key])
            # Aplica perfil carregado (se houver) antes de exibir a linha
            if self._profile:
                entry = self._find_profile_entry(mapa[key]["display"])
                if entry:
                    self._apply_entry_to_row(row, entry)
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
        grid = Grid()
        grid.Margin = Thickness(0, 1, 0, 1)
        grid.MinHeight = 34

        # cols: [check:28] [name:220] [qty-link:40] [qty-el:46] [elec:2*] [qty-dados:46] [dados:2*]
        widths = [(28, 0), (220, 0), (40, 0), (46, 0), (0, 2), (46, 0), (0, 2)]
        for w, star in widths:
            cd = ColumnDefinition()
            cd.Width = GridLength(star, GridUnitType.Star) if star > 0 else GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        cb = CheckBox()
        cb.IsChecked = True
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin = Thickness(4, 0, 0, 0)
        cb.Foreground = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
        Grid.SetColumn(cb, 0)

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

        tb_qty_link = TextBlock()
        tb_qty_link.Text = str(len(fam["instances"]))
        tb_qty_link.VerticalAlignment = VerticalAlignment.Center
        tb_qty_link.HorizontalAlignment = HorizontalAlignment.Center
        tb_qty_link.Foreground = SolidColorBrush(Color.FromRgb(0x77, 0x77, 0x77))
        Grid.SetColumn(tb_qty_link, 2)

        # TextBox para quantidade de pontos por instância
        from System.Windows.Controls import TextBox as _TB
        tb_pts = _TB()
        tb_pts.Text = u"1"
        tb_pts.Width = 36
        tb_pts.Height = 24
        tb_pts.TextAlignment = System.Windows.TextAlignment.Center
        tb_pts.VerticalAlignment = VerticalAlignment.Center
        tb_pts.HorizontalAlignment = HorizontalAlignment.Center
        tb_pts.Margin = Thickness(2, 0, 2, 0)
        tb_pts.ToolTip = u"Pontos elétricos por instância (deslocamento 5 cm)"
        tb_pts.Background = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        tb_pts.Foreground = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
        tb_pts.BorderBrush = SolidColorBrush(Color.FromRgb(0xC8, 0xC8, 0xC8))
        tb_pts.BorderThickness = Thickness(1)
        Grid.SetColumn(tb_pts, 3)

        use_face_now = (self.chk_FacePlacement.IsChecked == True)
        active_elec  = self._elec_face_symbols if use_face_now else self._elec_symbols
        active_data  = self._data_face_symbols if use_face_now else self._data_symbols
        elec_default = self._find_match_idx(fam["display"], active_elec)
        c_elec = self._make_editable_combo(active_elec, elec_default)
        Grid.SetColumn(c_elec, 4)

        tb_dados_pts = _TB()
        tb_dados_pts.Text = u"1"
        tb_dados_pts.Width = 36
        tb_dados_pts.Height = 24
        tb_dados_pts.TextAlignment = System.Windows.TextAlignment.Center
        tb_dados_pts.VerticalAlignment = VerticalAlignment.Center
        tb_dados_pts.HorizontalAlignment = HorizontalAlignment.Center
        tb_dados_pts.Margin = Thickness(2, 0, 2, 0)
        tb_dados_pts.ToolTip = u"Pontos de dados por instância (deslocamento 5 cm)"
        tb_dados_pts.Background = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        tb_dados_pts.Foreground = SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
        tb_dados_pts.BorderBrush = SolidColorBrush(Color.FromRgb(0xC8, 0xC8, 0xC8))
        tb_dados_pts.BorderThickness = Thickness(1)
        Grid.SetColumn(tb_dados_pts, 5)

        c_dados = self._make_editable_combo(active_data, 0)
        Grid.SetColumn(c_dados, 6)

        for child in [cb, tb_name, tb_qty_link, tb_pts, c_elec, tb_dados_pts, c_dados]:
            grid.Children.Add(child)

        return {
            "grid":          grid,
            "cb":            cb,
            "tb_pts":        tb_pts,
            "tb_dados_pts":  tb_dados_pts,
            "c_elec":        c_elec,
            "c_dados":       c_dados,
            "instances":        fam["instances"],
            "display":          fam["display"],
            "is_non_electric":  _is_non_electric_smart(fam["display"], fam.get("std_cat", "")),
            "elec_sym_list":    active_elec,
            "dados_sym_list":   active_data,
        }

    def _on_place(self, sender, args):
        to_place = []
        for r in self._family_rows:
            if not r["cb"].IsChecked:
                continue
            elec_sym  = self._get_symbol_from_combo(r["c_elec"],  r.get("elec_sym_list",  self._elec_symbols))
            dados_sym = self._get_symbol_from_combo(r["c_dados"], r.get("dados_sym_list", self._data_symbols))
            if not elec_sym and not dados_sym:
                continue
            qty = 1
            try:
                qty = max(1, int(str(r["tb_pts"].Text).strip() or u"1"))
            except Exception:
                qty = 1
            qty_dados = 1
            try:
                qty_dados = max(1, int(str(r["tb_dados_pts"].Text).strip() or u"1"))
            except Exception:
                qty_dados = 1
            to_place.append({
                "instances": r["instances"],
                "elec_sym":  elec_sym,
                "dados_sym": dados_sym,
                "qty":       qty,
                "qty_dados": qty_dados,
            })

        if not to_place:
            forms.alert("Nenhuma linha com ponto elétrico ou de dados selecionada.", title="Pontos por Vínculo")
            return

        dbg.section("Pontos por Vínculo — Colocação")
        dbg.info("Famílias a processar: {}".format(len(to_place)))

        _doc = self._doc
        t = Transaction(_doc, "Pontos por Vínculo")
        try:
            status = t.Start()
            if status != TransactionStatus.Started:
                forms.alert("Não foi possível iniciar a transação (status: {}).".format(status), title="Pontos por Vínculo")
                return

            levels = list(FilteredElementCollector(_doc).OfClass(Level))

            # ── Face-based placement setup ──
            use_face = (self.chk_FacePlacement.IsChecked == True)
            ri = None
            if use_face:
                view3d = _get_3d_view(_doc)
                if view3d:
                    wall_filter = ElementCategoryFilter(BuiltInCategory.OST_Walls)
                    ri = ReferenceIntersector(wall_filter, FindReferenceTarget.Face, view3d)
                    ri.FindReferencesInRevitLinks = True
                    dbg.debug("ReferenceIntersector criado (inclui vínculos).")
                else:
                    dbg.warn("Posicionar na face: nenhuma vista 3D disponível — usando placement normal.")
                    use_face = False

            # Ativar todos os símbolos necessários e regenerar UMA VEZ antes de criar instâncias.
            # Revit exige Regenerate() após Activate() — sem isso NewFamilyInstance falha silenciosamente.
            any_activated = False
            total_inst = sum(len(x["instances"]) for x in to_place)
            processed_inst = 0
            with forms.ProgressBar(title="Colocando Pontos...", cancellable=True) as pb:
                for item in to_place:
                    if pb.cancelled: break
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

                placed_count = 0
                skip_count = 0
                created_element_ids = []

                for item in to_place:
                    if pb.cancelled: break
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

                    dbg.sub("{} ({} inst.) → elec:{} dados:{}".format(
                        fam_name,
                        len(item["instances"]), elec_name, dados_name
                    ))
                    LATERAL_FT = 0.05 / 0.3048  # 5 cm em pés
                    qty_item   = item.get("qty", 1)

                    for inst in item["instances"]:
                        if pb.cancelled: break
                        processed_inst += 1
                        pb.update_progress(processed_inst, total_inst)
                        # ── Coordenadas base ──
                        try:
                            pt_local = inst.Location.Point
                            pt_base  = transform.OfPoint(pt_local) if transform else pt_local
                            lvl      = min(levels, key=lambda l: abs(l.Elevation - pt_base.Z))
                        except:
                            skip_count += 1
                            dbg.warn("  inst.Id={} — sem localização, pulado.".format(inst.Id))
                            continue

                        # ── Pontos Elétricos ──
                        for q in range(qty_item):
                            if pb.cancelled: break
                            # Desloca lateralmente em X: 0 cm no primeiro, +5 cm por cópia adicional
                            pt = XYZ(pt_base.X + q * LATERAL_FT, pt_base.Y, pt_base.Z)

                            new_elec = None
                            if item["elec_sym"]:
                                placed   = False
                                face_ref = None
                                face_pt  = None
                                face_dir = None

                                if use_face and ri and not _is_luminaire_sym(item["elec_sym"]):
                                    try:
                                        face_ref, face_pt, face_dir = self._find_nearest_wall_face(ri, pt)
                                    except:
                                        pass

                                    if face_ref and face_pt:
                                        try:
                                            new_elec = _doc.Create.NewFamilyInstance(
                                                face_ref, face_pt, XYZ(0, 0, 1), item["elec_sym"]
                                            )
                                            placed = True
                                            dbg.debug("  [L1-face] inst.Id={} q={}".format(inst.Id, q))
                                        except Exception as ex:
                                            dbg.debug("  [L1] falhou: {} — tentando L2".format(ex))

                                        if not placed:
                                            try:
                                                new_elec = _doc.Create.NewFamilyInstance(
                                                    face_pt, item["elec_sym"], lvl, StructuralType.NonStructural
                                                )
                                                placed = True
                                                dbg.debug("  [L2-pos] inst.Id={} q={}".format(inst.Id, q))
                                                if face_dir:
                                                    try:
                                                        _doc.Regenerate()
                                                        outward = XYZ(-face_dir.X, -face_dir.Y, 0)
                                                        facing  = new_elec.FacingOrientation
                                                        if facing and facing.GetLength() > 1e-6:
                                                            import math as _math
                                                            angle = facing.AngleTo(outward)
                                                            cross = facing.CrossProduct(outward)
                                                            if cross.Z < 0:
                                                                angle = -angle
                                                            if abs(angle) > 1e-6:
                                                                axis = Line.CreateBound(
                                                                    face_pt,
                                                                    XYZ(face_pt.X, face_pt.Y, face_pt.Z + 1)
                                                                )
                                                                ElementTransformUtils.RotateElement(
                                                                    _doc, new_elec.Id, axis, angle
                                                                )
                                                    except Exception as rot_ex:
                                                        dbg.debug("  rotação falhou: {}".format(rot_ex))
                                            except Exception as ex:
                                                dbg.warn("  [L2] falhou inst.Id={}: {}".format(inst.Id, ex))
                                    else:
                                        dbg.debug("  sem parede próxima inst.Id={}".format(inst.Id))

                                if not placed:
                                    try:
                                        new_elec = _doc.Create.NewFamilyInstance(
                                            pt, item["elec_sym"], lvl, StructuralType.NonStructural
                                        )
                                        placed = True
                                    except Exception as ex:
                                        dbg.warn("  [L3] falhou inst.Id={}: {}".format(inst.Id, ex))

                                if placed:
                                    placed_count += 1
                                if new_elec:
                                    created_element_ids.append(new_elec.Id)
                                    try:
                                        _p = new_elec.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                                        if _p and not _p.IsReadOnly:
                                            _p.Set(u'PpV: {}'.format(fam_name))
                                    except: pass

                        # ── Pontos de Dados (loop independente) ──
                        qty_dados = item.get("qty_dados", 1)
                        for q in range(qty_dados):
                            if pb.cancelled: break
                            if item["dados_sym"]:
                                try:
                                    pt_dados = XYZ(
                                        pt_base.X + q * LATERAL_FT + DADOS_OFFSET_X,
                                        pt_base.Y,
                                        pt_base.Z + DADOS_OFFSET_Z
                                    )
                                    new_dados = _doc.Create.NewFamilyInstance(
                                        pt_dados, item["dados_sym"], lvl, StructuralType.NonStructural
                                    )
                                    placed_count += 1
                                    if new_dados:
                                        created_element_ids.append(new_dados.Id)
                                        try:
                                            _p = new_dados.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                                            if _p and not _p.IsReadOnly:
                                                _p.Set(u'PpV: {}'.format(fam_name))
                                        except: pass
                                except Exception as ex:
                                    dbg.warn("  dados falhou inst.Id={} q={}: {}".format(inst.Id, q, ex))

            # Carimbar para handshake com Auto-Elétrica (silencioso se parâmetro não existir)
            for eid in created_element_ids:
                try:
                    el = _doc.GetElement(eid)
                    if el:
                        p = el.LookupParameter('LF_StatusIntegracao')
                        if p and not p.IsReadOnly:
                            p.Set('Aguardando_Eletrica')
                except Exception:
                    pass

            t.Commit()

            # Seleciona os elementos criados
            if created_element_ids:
                from System.Collections.Generic import List
                from Autodesk.Revit.DB import ElementId
                id_list = List[ElementId](created_element_ids)
                try:
                    __revit__.ActiveUIDocument.Selection.SetElementIds(id_list)
                except:
                    pass

            # Passa para a aba Auto-Elétrica com os elementos já selecionados
            try:
                self.MainTabs.SelectedIndex = 1
                self._ae.refresh_from_selection()
            except Exception:
                pass

            dbg.section("Resultado")
            dbg.info("Pontos criados:  {}".format(placed_count))
            if skip_count:
                dbg.warn("Instâncias puladas (sem localização): {}".format(skip_count))

            # Feedback sem fechar a janela
            _msg = u"{} ponto(s) colocado(s)!".format(placed_count)
            if skip_count:
                _msg += u"  ({} pulado(s))".format(skip_count)
            self.lbl_Status.Text = _msg
            self.lbl_Status.Foreground = SolidColorBrush(Color.FromRgb(0x10, 0x7C, 0x10))
            try:
                forms.toast(_msg, title=u"Pontos por Vínculo")
            except Exception:
                pass

            # Troca para aba Auto-Elétrica com os elementos já carregados
            try:
                self.MainTabs.SelectedIndex = 1
                self._ae.refresh_from_selection(placed_ids=created_element_ids)
            except Exception:
                pass

        except Exception as e:
            try:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
            except:
                pass
            dbg.error("Exceção em _on_place: {}".format(e))
            self.lbl_Status.Text = u"Erro: {}".format(str(e)[:140])
            self.lbl_Status.Foreground = SolidColorBrush(Color.FromRgb(0xC4, 0x2B, 0x1A))


if __name__ == "__main__":
    win = PontosVinculoWindow(script.get_bundle_file('ui.xaml'))
    win.ShowDialog()