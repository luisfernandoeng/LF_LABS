# -*- coding: utf-8 -*-
"""auto_eletrica.py — módulo de parametrização elétrica em lote.

Usado como aba "Auto-Elétrica" dentro da janela do Pontos por Vínculo.
Recebe referência à janela principal, ao doc, ao uidoc e ao dbg.
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

import io
import json
import os
import re

import System
from System.Collections.Generic import List
from System.Windows import (
    CornerRadius, GridLength, GridUnitType,
    HorizontalAlignment, Thickness, VerticalAlignment, Visibility
)
from System.Windows.Controls import (
    Border, Button, ColumnDefinition, ComboBox,
    Grid, StackPanel, TextBlock, TextBox
)
from System.Windows.Media import Color, SolidColorBrush

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, Domain,
    ElementId, FilteredElementCollector, StorageType,
    Transaction
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType

# ── Paleta (espelha os valores do ui.xaml) ────────────────────────────────────

def _rgb(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))

C_WHITE   = _rgb(0xFF, 0xFF, 0xFF)
C_BG      = _rgb(0xF8, 0xF8, 0xF8)
C_BORDER  = _rgb(0xC8, 0xC8, 0xC8)
C_DIVIDER = _rgb(0xE0, 0xE0, 0xE0)
C_TEXT1   = _rgb(0x33, 0x33, 0x33)
C_TEXT2   = _rgb(0x77, 0x77, 0x77)
C_HINT    = _rgb(0x99, 0x99, 0x99)
C_BLUE    = _rgb(0x00, 0x78, 0xD7)
C_SELBG   = _rgb(0xE6, 0xF2, 0xFA)
C_GREEN   = _rgb(0x10, 0x7C, 0x10)

FEET_PER_METER = 1.0 / 0.3048
TENSOES = [u'127V', u'220V', u'380V']

_ELEC_CATEGORIES = [
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_SecurityDevices,
]

# ── Profile store (ae_data.json na pasta profiles/) ───────────────────────────

class AEProfileStore(object):

    def __init__(self, profiles_dir):
        self._path = os.path.join(profiles_dir, u'ae_data.json')
        self._data = self._load()

    def _load(self):
        try:
            with io.open(self._path, u'r', encoding=u'utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def load(self, family_name):
        return dict(self._data.get(family_name, {}))

    def save(self, family_name, ae_cfg):
        self._data[family_name] = ae_cfg
        try:
            if not os.path.isdir(os.path.dirname(self._path)):
                os.makedirs(os.path.dirname(self._path))
            with io.open(self._path, u'w', encoding=u'utf-8') as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2, sort_keys=True)
        except Exception:
            pass


# ── Revit helpers ─────────────────────────────────────────────────────────────

def _meters_to_feet(m):
    return float(m) * FEET_PER_METER


def _set_param(elem, names, value):
    for n in names:
        p = elem.LookupParameter(n)
        if p and not p.IsReadOnly:
            try:
                st = p.StorageType
                if st == StorageType.Double:
                    p.Set(float(value))
                elif st == StorageType.String:
                    p.Set(unicode(value) if isinstance(value, str) else value)
                elif st == StorageType.Integer:
                    p.Set(int(value))
                return True
            except Exception:
                pass
    return False


def _electrical_connectors(elem):
    try:
        cm = elem.MEPModel.ConnectorManager
        if cm:
            return [c for c in cm.Connectors if c.Domain == Domain.DomainElectrical]
    except Exception:
        pass
    return []


def _has_electrical_connector(elem):
    return bool(_electrical_connectors(elem))


def _disconnect_from_circuits(elem, dbg):
    try:
        systems = list(elem.MEPModel.ElectricalSystems)
        for s in systems:
            try:
                ids = List[ElementId]()
                ids.Add(elem.Id)
                s.RemoveFromCircuit(ids)
            except Exception as ex:
                dbg.warn(u'  RemoveFromCircuit falhou: {}'.format(ex))
    except Exception:
        pass


def _get_family_name(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p and p.HasValue:
            val = p.AsString()
            if val and val.startswith(u'PpV: '):
                return val.replace(u'PpV: ', u'')
    except Exception:
        pass
    try:
        return elem.Symbol.FamilyName
    except Exception:
        pass
    try:
        return elem.Symbol.Family.Name
    except Exception:
        pass
    return u'(sem familia)'
def _group_by_family(elements):
    groups = {}
    order  = []
    for elem in elements:
        fname = _get_family_name(elem)
        if fname not in groups:
            groups[fname] = []
            order.append(fname)
        groups[fname].append(elem)
    return order, groups


def _get_panels(doc):
    panels = []
    col = (FilteredElementCollector(doc)
           .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
           .WhereElementIsNotElementType())
    for p in col:
        mark = u''
        mp = p.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if mp:
            mark = mp.AsString() or u''
        label = u'{} — {}'.format(mark, p.Name) if mark else p.Name
        panels.append((label.strip(u' —'), p))
    return sorted(panels, key=lambda x: x[0])


def _get_load_types(doc):
    names = set()
    try:
        col = (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_ElectricalLoadClassifications)
               .ToElements())
        for lc in col:
            lc_name = u''
            try:
                lc_name = lc.Name
            except Exception:
                pass
            if not lc_name:
                try:
                    p = lc.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p and p.HasValue:
                        lc_name = p.AsString()
                except Exception:
                    pass
            if lc_name:
                names.add(lc_name)
    except Exception:
        pass

    # Fallback: coletar valores distintos do parâmetro de texto "Tipo de Carga"
    cats = [
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_SpecialityEquipment,
    ]
    for cat in cats:
        try:
            col = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
            for el in col:
                p = el.LookupParameter(u'Tipo de Carga')
                if p and p.HasValue:
                    val = (p.AsString() or u'').strip()
                    if val:
                        names.add(val)
        except Exception:
            pass

    return sorted(list(names))


def _get_electrical_elements_from_ids(ids, doc):
    """Retorna apenas elementos com conector elétrico (para seleção ativa)."""
    result = []
    for eid in ids:
        elem = doc.GetElement(eid)
        if elem and _has_electrical_connector(elem):
            result.append(elem)
    return result


def _get_any_instances(ids, doc):
    """Retorna qualquer instância válida (para pickup de placement — sem filtro de conector)."""
    result = []
    for eid in ids:
        try:
            elem = doc.GetElement(eid)
            if elem is None:
                continue
            # Só instâncias (FamilyInstance, não tipos ou elementos de documento)
            if elem.Category is not None and hasattr(elem, 'Location'):
                result.append(elem)
        except Exception:
            pass
    return result


def _find_pending_elements(doc):
    result = []
    for cat in _ELEC_CATEGORIES:
        try:
            col = (FilteredElementCollector(doc)
                   .OfCategory(cat)
                   .WhereElementIsNotElementType())
            for elem in col:
                p = elem.LookupParameter(u'LF_StatusIntegracao')
                if p and p.AsString() == u'Aguardando_Eletrica':
                    result.append(elem)
        except Exception:
            pass
    return result


def _get_next_circuit_number(doc, panel, prefix):
    if not prefix:
        return 1
    used = set()
    try:
        for s in FilteredElementCollector(doc).OfClass(ElectricalSystem):
            try:
                base = s.BaseEquipment
                if base is None or base.Id != panel.Id:
                    continue
                for pname in (u'Nome da carga', u'Load Name'):
                    p = s.LookupParameter(pname)
                    if p:
                        name = p.AsString() or u''
                        if name.upper().startswith(prefix.upper()):
                            suffix = name[len(prefix):].lstrip(u'-_ ')
                            m = re.match(r'^(\d+)$', suffix)
                            if m:
                                used.add(int(m.group(1)))
                        break
            except Exception:
                pass
    except Exception:
        pass
    n = 1
    while n in used:
        n += 1
    return n


# ── WPF builder helpers ───────────────────────────────────────────────────────

def _tb(text, size=12, color=None, bold=False, margin=None):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = size
    tb.Foreground = color or C_TEXT2
    if bold:
        from System.Windows import FontWeights
        tb.FontWeight = FontWeights.SemiBold
    if margin:
        tb.Margin = margin
    return tb


def _make_textbox(text=u''):
    box = TextBox()
    box.Text = text
    box.Height = 30
    box.Padding = Thickness(8, 0, 8, 0)
    box.VerticalContentAlignment = VerticalAlignment.Center
    box.FontSize = 12
    box.Background = C_WHITE
    box.Foreground = C_TEXT1
    box.BorderBrush = C_BORDER
    box.BorderThickness = Thickness(1)
    return box


def _make_combobox(items, selected_idx=0):
    cb = ComboBox()
    cb.Height = 30
    cb.FontSize = 12
    cb.Foreground = C_TEXT1
    cb.Background = C_WHITE
    cb.BorderBrush = C_BORDER
    cb.IsEditable = True
    for item in items:
        cb.Items.Add(item)
    if items and 0 <= selected_idx < len(items):
        cb.SelectedIndex = selected_idx
    return cb


def _field_stack(label_text, widget):
    sp = StackPanel()
    sp.Children.Add(_tb(label_text, 11, C_TEXT2, margin=Thickness(0, 0, 0, 3)))
    sp.Children.Add(widget)
    return sp


def _divider():
    d = Border()
    d.Height = 1
    d.Background = C_DIVIDER
    d.Margin = Thickness(0, 8, 0, 10)
    return d


# ── FamilyGroupWidget ─────────────────────────────────────────────────────────

class FamilyGroupWidget(object):
    """Card WPF para um tipo de família — construído programaticamente."""

    def __init__(self, family_name, elements, profile, load_types, store, dbg):
        self.family_name = family_name
        self.elements    = elements
        self._store      = store
        self._load_types = load_types
        self._dbg        = dbg
        self.root        = self._build(profile)

    def _build(self, profile):
        card = Border()
        card.Background      = C_WHITE
        card.BorderBrush     = C_BORDER
        card.BorderThickness = Thickness(1)
        card.CornerRadius    = CornerRadius(6)
        card.Padding         = Thickness(14, 12, 14, 12)
        card.Margin          = Thickness(0, 0, 0, 10)

        body = StackPanel()

        # Título + badge
        title_grid = Grid()
        title_grid.Margin = Thickness(0, 0, 0, 2)
        tc0 = ColumnDefinition(); tc0.Width = GridLength(1, GridUnitType.Star)
        tc1 = ColumnDefinition(); tc1.Width = GridLength.Auto
        title_grid.ColumnDefinitions.Add(tc0)
        title_grid.ColumnDefinitions.Add(tc1)

        title_tb = _tb(self.family_name, 13, C_TEXT1, bold=True)
        Grid.SetColumn(title_tb, 0)

        badge = Border()
        badge.Background        = C_BLUE
        badge.CornerRadius      = CornerRadius(10)
        badge.Padding           = Thickness(8, 2, 8, 2)
        badge.VerticalAlignment = VerticalAlignment.Center
        badge_tb = _tb(u'{} pt'.format(len(self.elements)), 10, C_WHITE, bold=True)
        badge.Child = badge_tb
        Grid.SetColumn(badge, 1)

        title_grid.Children.Add(title_tb)
        title_grid.Children.Add(badge)
        body.Children.Add(title_grid)
        body.Children.Add(_divider())

        # Linha 1: Altura | Carga | Tensão | Prefixo
        row1 = Grid()
        row1.Margin = Thickness(0, 0, 0, 10)
        for i in range(7):
            cd = ColumnDefinition()
            cd.Width = GridLength(10) if i % 2 == 1 else GridLength(1, GridUnitType.Star)
            row1.ColumnDefinitions.Add(cd)

        saved_tensao = profile.get(u'tensao', u'220V')
        tensao_idx   = TENSOES.index(saved_tensao) if saved_tensao in TENSOES else 1

        self._txt_altura  = _make_textbox(str(profile.get(u'altura',   u'1.20')))
        self._txt_carga   = _make_textbox(str(profile.get(u'carga_va', u'600')))
        self._cmb_tensao  = _make_combobox(TENSOES, tensao_idx)
        self._txt_prefixo = _make_textbox(profile.get(u'prefixo', u''))

        fields = [
            (u'Altura (m)',  self._txt_altura),
            (u'Carga (VA)',  self._txt_carga),
            (u'Tensao',      self._cmb_tensao),
            (u'Prefixo',     self._txt_prefixo),
        ]
        for i, (lbl, widget) in enumerate(fields):
            sp = _field_stack(lbl, widget)
            Grid.SetColumn(sp, i * 2)
            row1.Children.Add(sp)

        body.Children.Add(row1)

        # Linha 2: Tipo de Carga
        tipo_items = [u'(não definido)'] + self._load_types
        saved_tipo = profile.get(u'tipo_carga', u'')
        tipo_idx   = tipo_items.index(saved_tipo) if saved_tipo in tipo_items else 0
        self._cmb_tipo = _make_combobox(tipo_items, tipo_idx)
        body.Children.Add(_field_stack(u'Tipo de Carga', self._cmb_tipo))

        card.Child = body
        return card

    def apply_profile(self, profile):
        if u'altura' in profile:
            self._txt_altura.Text = str(profile[u'altura'])
        if u'carga_va' in profile:
            self._txt_carga.Text = str(profile[u'carga_va'])
        if u'tensao' in profile:
            tensao = profile[u'tensao']
            if tensao in TENSOES:
                self._cmb_tensao.SelectedIndex = TENSOES.index(tensao)
        if u'prefixo' in profile:
            self._txt_prefixo.Text = profile.get(u'prefixo', u'')
        if u'tipo_carga' in profile:
            tipo = profile[u'tipo_carga']
            tipo_items = [u'(não definido)'] + self._load_types
            if tipo in tipo_items:
                self._cmb_tipo.SelectedIndex = tipo_items.index(tipo)
            elif tipo:
                self._cmb_tipo.Text = tipo

    def get_values(self):
        tensao = str(self._cmb_tensao.SelectedItem) if self._cmb_tensao.SelectedItem else u'220V'
        tipo   = str(self._cmb_tipo.SelectedItem)   if self._cmb_tipo.SelectedItem   else u''
        if tipo == u'(não definido)':
            tipo = u''
        return {
            u'altura':    self._txt_altura.Text.replace(u',', u'.'),
            u'carga_va':  self._txt_carga.Text.replace(u',', u'.'),
            u'tensao':    tensao,
            u'prefixo':   (self._txt_prefixo.Text or u'').strip(),
            u'tipo_carga': tipo,
        }


# ── AutoEletricaController ────────────────────────────────────────────────────

class AutoEletricaController(object):
    """Controla a aba Auto-Elétrica dentro da janela PontosVinculoWindow."""

    def __init__(self, win, doc, uidoc, dbg, profiles_dir):
        self._win          = win
        self._doc          = doc
        self._uidoc        = uidoc
        self._dbg          = dbg
        self._store        = AEProfileStore(profiles_dir)
        self._profiles_dir = profiles_dir
        self._panels       = []
        self._load_types   = []
        self._widgets      = []

        self._load_types = _get_load_types(doc)
        self._panels     = _get_panels(doc)
        self._populate_panels()
        self._refresh_ae_profiles()

        # Wire events
        win.ae_BtnBuscarPendentes.Click += self._on_buscar_pendentes
        win.ae_BtnExecutar.Click        += self._on_executar
        win.ae_BtnLoadProfile.Click          += self._on_ae_load_profile
        win.ae_BtnSaveProfile.Click          += self._on_ae_save_profile
        win.ae_CmbProfile.SelectionChanged   += self._on_ae_profile_combo_changed

        # Lê seleção inicial; fallback para último placement salvo
        self.refresh_from_selection()

    # ── Perfis AE ─────────────────────────────────────────────────────────────

    def _list_ae_profiles(self):
        try:
            return sorted(
                f[:-5] for f in os.listdir(self._profiles_dir)
                if f.endswith(u'.json') and f != u'ae_data.json'
            )
        except Exception:
            return []

    def _ae_profile_path(self, name):
        return os.path.join(self._profiles_dir, name + u'.json')

    def _load_ae_profile_data(self, name):
        try:
            with io.open(self._ae_profile_path(name), u'r', encoding=u'utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_ae_profile_data(self, name, data):
        try:
            if not os.path.isdir(self._profiles_dir):
                os.makedirs(self._profiles_dir)
            with io.open(self._ae_profile_path(name), u'w', encoding=u'utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
            return True
        except Exception:
            return False

    def _refresh_ae_profiles(self, select_name=None):
        names = self._list_ae_profiles()
        net = List[System.Object]()
        net.Add(u'— Selecionar perfil —')
        for n in names:
            net.Add(n)
        self._win._syncing_profile = True
        try:
            self._win.ae_CmbProfile.ItemsSource = net
            idx = 0
            if select_name:
                try:
                    idx = names.index(select_name) + 1
                except ValueError:
                    idx = 0
            self._win.ae_CmbProfile.SelectedIndex = idx
        finally:
            self._win._syncing_profile = False

    def _on_ae_profile_combo_changed(self, sender, e):
        if getattr(self._win, '_syncing_profile', False):
            return
        self._win._syncing_profile = True
        try:
            idx = self._win.ae_CmbProfile.SelectedIndex
            for combo in [self._win.cb_Profile, self._win.gp_CmbProfile]:
                try:
                    if combo.SelectedIndex != idx:
                        combo.SelectedIndex = idx
                except Exception:
                    pass
        finally:
            self._win._syncing_profile = False

    def _on_ae_load_profile(self, sender, e):
        from pyrevit import forms
        idx = self._win.ae_CmbProfile.SelectedIndex
        if idx <= 0:
            forms.alert(u'Selecione um perfil antes de carregar.', exitscript=False)
            return
        names = self._list_ae_profiles()
        if idx - 1 >= len(names):
            return
        name = names[idx - 1]
        data = self._load_ae_profile_data(name)
        if not data:
            forms.alert(u"Perfil '{}' está vazio ou corrompido.".format(name), exitscript=False)
            return
        # Aplica a ambas as abas usando _find_profile_entry para match parcial
        self._win._profile = data
        for w in self._widgets:
            entry = self._win._find_profile_entry(w.family_name)
            if entry:
                w.apply_profile(entry)
        self._win._apply_profile_to_rows()
        try:
            self._win._refresh_all_profiles(name)
        except Exception:
            pass
        self._win.ae_TxtStatus.Text = u"Perfil '{}' carregado.".format(name)
        self._dbg.info(u'[AE] Perfil carregado: {}'.format(name))

    def _on_ae_save_profile(self, sender, e):
        from pyrevit import forms
        if not self._widgets:
            forms.alert(u'Nenhuma família para salvar. Selecione elementos primeiro.', exitscript=False)
            return
        name = forms.ask_for_string(
            prompt=u'Nome do perfil:',
            title=u'Salvar Perfil',
            default=u'Meu Perfil'
        )
        if not name:
            return
        name = re.sub(r'[<>:"/\\|?*]', u'_', name).strip()

        # Começa com dados das linhas PpV (ponto_eletrico, ponto_dados)
        data = {}
        try:
            for row in self._win._family_rows:
                disp = row.get(u'display', u'')
                if not disp:
                    continue
                pe = u''
                try:
                    elec_list = row.get(u'elec_sym_list', [])
                    pe_idx = row[u'c_elec'].SelectedIndex
                    if pe_idx > 0 and pe_idx - 1 < len(elec_list):
                        pe = elec_list[pe_idx - 1][u'display']
                    else:
                        pe = (row[u'c_elec'].Text or u'').strip()
                        if pe == u'— Não colocar —':
                            pe = u''
                except Exception:
                    pass
                pd_val = u''
                try:
                    dados_list = row.get(u'dados_sym_list', [])
                    pd_idx = row[u'c_dados'].SelectedIndex
                    if pd_idx > 0 and pd_idx - 1 < len(dados_list):
                        pd_val = dados_list[pd_idx - 1][u'display']
                    else:
                        pd_val = (row[u'c_dados'].Text or u'').strip()
                        if pd_val == u'— Não colocar —':
                            pd_val = u''
                except Exception:
                    pass
                data[disp] = {
                    u'ponto_eletrico': pe,
                    u'ponto_dados':    pd_val,
                    u'checked':        bool(row.get(u'cb') and row[u'cb'].IsChecked),
                }
        except Exception:
            pass

        # Mescla dados AE (altura, carga_va, tensao, prefixo, tipo_carga)
        for w in self._widgets:
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

        if self._save_ae_profile_data(name, data):
            try:
                self._win._refresh_all_profiles(name)
            except Exception:
                self._refresh_ae_profiles(name)
            self._win.ae_TxtStatus.Text = u"Perfil '{}' salvo ({} família(s)).".format(name, len(data))
            self._dbg.info(u'[AE] Perfil salvo: {}'.format(name))
        else:
            forms.alert(u'Erro ao salvar o perfil.', exitscript=False)

    # ── Setup UI ──────────────────────────────────────────────────────────────

    def _populate_panels(self):
        self._win.ae_CmbQuadro.Items.Clear()
        for label, _ in self._panels:
            self._win.ae_CmbQuadro.Items.Add(label)
        if self._panels:
            self._win.ae_CmbQuadro.SelectedIndex = 0

    def _update_status_display(self, elements):
        n = len(elements)
        if n == 0:
            self._win.ae_TxtSelecao.Text = (
                u'Nenhum elemento com conector elétrico na seleção.'
            )
            self._win.ae_BtnBuscarPendentes.Visibility = Visibility.Visible
        else:
            self._win.ae_TxtSelecao.Text = (
                u'{} elemento(s) com conector elétrico — {} tipo(s) de família.'.format(
                    n, len(set(_get_family_name(e) for e in elements))
                )
            )
            self._win.ae_BtnBuscarPendentes.Visibility = Visibility.Collapsed

    def _build_cards(self, elements):
        self._win.ae_PanelGroups.Children.Clear()
        self._widgets = []
        if not elements:
            return
        order, groups = _group_by_family(elements)
        self._dbg.section(u'[AE] Montando cards: {} familia(s)'.format(len(order)))
        for fname in order:
            # Carrega perfil AE, com fallback para dados do perfil PpV
            profile = self._store.load(fname)
            if not profile:
                profile = self._fallback_from_ppv_profile(fname)
            self._dbg.debug(u'  {} → perfil: {}'.format(fname, profile))
            w = FamilyGroupWidget(
                fname, groups[fname], profile,
                self._load_types, self._store, self._dbg
            )
            self._widgets.append(w)
            self._win.ae_PanelGroups.Children.Add(w.root)

    def _fallback_from_ppv_profile(self, family_name):
        """Usa potência e tipo_carga do perfil PpV ativo como valores iniciais."""
        try:
            entry = self._win._find_profile_entry(family_name)
            if entry:
                return {
                    u'tipo_carga': entry.get(u'tipo_carga', u''),
                    u'carga_va':   entry.get(u'potencia', u'600'),
                    u'prefixo':    entry.get(u'circuito_nome', u''),
                }
        except Exception:
            pass
        return {}

    # ── Eventos ───────────────────────────────────────────────────────────────

    def refresh_from_selection(self, placed_ids=None):
        """Atualiza a aba com elementos.

        Prioridade:
          1. placed_ids — passados diretamente após placement (ElementId list)
          2. Seleção ativa no Revit (filtra por conector elétrico)
        """
        if placed_ids:
            elements = _get_any_instances(placed_ids, self._doc)
            self._dbg.info(u'[AE] refresh: {} elem do placement atual.'.format(len(elements)))
        else:
            sel_ids  = list(self._uidoc.Selection.GetElementIds())
            elements = _get_electrical_elements_from_ids(sel_ids, self._doc)
            if elements:
                self._dbg.info(u'[AE] refresh: {} elem da seleção ativa.'.format(len(elements)))
            else:
                self._dbg.info(u'[AE] refresh: sem elementos.')
        self._build_cards(elements)
        self._update_status_display(elements)
        self._win.lbl_Status.Text = (
            u'Auto-Elétrica: {} ponto(s) prontos para parametrizar.'.format(len(elements))
            if elements else u'Auto-Elétrica: selecione elementos elétricos.'
        )

    def _on_buscar_pendentes(self, sender, e):
        self._dbg.section(u'[AE] Buscando pendentes...')
        found = _find_pending_elements(self._doc)
        self._dbg.info(u'  {} pendente(s) encontrado(s).'.format(len(found)))
        self._build_cards(found)
        self._update_status_display(found)
        self._win.ae_TxtStatus.Text = (
            u'{} pendente(s) encontrado(s).'.format(len(found))
            if found else u'Nenhum pendente encontrado.'
        )

    def _on_executar(self, sender, e):
        if not self._widgets:
            from pyrevit import forms
            forms.alert(u'Nenhum elemento para processar.', exitscript=False)
            return
        if self._win.ae_CmbQuadro.SelectedIndex < 0:
            from pyrevit import forms
            forms.alert(u'Selecione um quadro de distribuição.', exitscript=False)
            return

        _, panel_elem = self._panels[self._win.ae_CmbQuadro.SelectedIndex]

        # Validar campos
        groups = []
        for w in self._widgets:
            vals = w.get_values()
            try:
                altura = float(vals[u'altura'])
            except ValueError:
                from pyrevit import forms
                forms.alert(
                    u'Altura inválida em "{}": {}'.format(w.family_name, vals[u'altura']),
                    exitscript=False
                )
                return
            try:
                carga = float(vals[u'carga_va'])
            except ValueError:
                from pyrevit import forms
                forms.alert(
                    u'Carga inválida em "{}": {}'.format(w.family_name, vals[u'carga_va']),
                    exitscript=False
                )
                return
            groups.append({
                u'elements':   w.elements,
                u'family':     w.family_name,
                u'altura':     altura,
                u'carga_va':   carga,
                u'tensao':     vals[u'tensao'],
                u'prefixo':    vals[u'prefixo'],
                u'tipo_carga': vals[u'tipo_carga'],
                u'panel':      panel_elem,
            })

        self._dbg.section(u'[AE] Executando: {} grupo(s)'.format(len(groups)))
        created = []
        for group in groups:
            name = self._apply_group(group)
            if name:
                created.append(u'{} ({} pts)'.format(name, len(group[u'elements'])))

        if created:
            self._win.ae_TxtStatus.Text = u'OK: ' + u', '.join(created)
            self._dbg.info(u'[AE] Circuitos criados: {}'.format(u', '.join(created)))
            try:
                from pyrevit import forms
                forms.toast(u'Circuitos criados: ' + u', '.join(created), title=u'Auto-Elétrica')
            except Exception:
                pass
        else:
            self._win.ae_TxtStatus.Text = u'Nenhum circuito criado.'

    # ── Core: criar circuito por grupo ────────────────────────────────────────

    def _apply_group(self, group):
        elements  = group[u'elements']
        altura_ft = _meters_to_feet(group[u'altura'])
        potencia  = group[u'carga_va']
        prefix    = group[u'prefixo']
        panel     = group[u'panel']

        tensao_n = 0.0
        try:
            tensao_n = float(group[u'tensao'].replace(u'V', u''))
        except ValueError:
            pass

        next_num     = _get_next_circuit_number(self._doc, panel, prefix)
        circuit_name = u'{}-{:02d}'.format(prefix, next_num) if prefix else None

        self._dbg.enter(u'[AE] Grupo "{}": {} elem → circuito {}'.format(
            group[u'family'], len(elements), circuit_name
        ))

        with Transaction(self._doc, u'Auto-Elétrica: {}'.format(group[u'family'])) as t:
            t.Start()
            try:
                for elem in elements:
                    # 1. Elevação
                    ep = elem.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                    if ep and not ep.IsReadOnly:
                        ep.Set(altura_ft)
                        self._dbg.step(u'  elevação setada: {:.2f} m → {:.3f} ft'.format(
                            group[u'altura'], altura_ft
                        ))

                    # 2. Potência por conector
                    for c in _electrical_connectors(elem):
                        try:
                            c.ElectricalApparentLoad = potencia
                        except Exception as ex:
                            self._dbg.warn(u'  ElectricalApparentLoad falhou: {}'.format(ex))

                    # 3. Tensão
                    if tensao_n:
                        _set_param(elem,
                                   [u'Tensão (V)', u'Tensão', u'Voltage', u'Voltagem', u'Volts'],
                                   tensao_n)

                    # 4. Tipo de carga
                    if group[u'tipo_carga']:
                        _set_param(elem,
                                   [u'Tipo de carga', u'Load Classification', u'Load Type'],
                                   group[u'tipo_carga'])

                    # 5. Desconectar circuitos existentes
                    _disconnect_from_circuits(elem, self._dbg)

                self._doc.Regenerate()

                # 6. Criar ElectricalSystem
                id_list = List[ElementId]()
                for elem in elements:
                    id_list.Add(elem.Id)

                circuit = ElectricalSystem.Create(
                    self._doc, id_list, ElectricalSystemType.PowerCircuit
                )
                self._dbg.info(u'  ElectricalSystem criado: Id={}'.format(circuit.Id))

                # 7. Conectar ao quadro
                if panel:
                    circuit.SelectPanel(panel)
                    self._dbg.info(u'  Conectado ao painel: {}'.format(panel.Name))

                # 8. Nomear
                if circuit_name:
                    _set_param(circuit, [u'Nome da carga', u'Load Name'], circuit_name)

                # 9. Limpar carimbo de integração
                for elem in elements:
                    p = elem.LookupParameter(u'LF_StatusIntegracao')
                    if p and not p.IsReadOnly:
                        p.Set(u'')

                t.Commit()
                self._dbg.exit(u'[AE] Grupo "{}" → OK'.format(group[u'family']))
                return circuit_name

            except Exception as ex:
                t.RollBack()
                self._dbg.fail(u'[AE] Grupo "{}" → ERRO: {}'.format(group[u'family'], ex))
                try:
                    from pyrevit import forms
                    forms.alert(
                        u'Erro ao criar circuito para "{}":\n{}'.format(
                            group[u'family'], str(ex)
                        ),
                        exitscript=False
                    )
                except Exception:
                    pass
                return None
