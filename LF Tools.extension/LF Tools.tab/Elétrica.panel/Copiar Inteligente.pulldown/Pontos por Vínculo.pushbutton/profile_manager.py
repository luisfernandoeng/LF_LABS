# -*- coding: utf-8 -*-
"""profile_manager.py — Aba de gerenciamento de perfis."""

import clr
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
    HorizontalAlignment, Thickness, VerticalAlignment
)
from System.Windows.Controls import (
    Border, Button, CheckBox, ColumnDefinition,
    ComboBox, Grid, StackPanel, TextBlock, TextBox
)
from System.Windows.Media import Color, SolidColorBrush

from pyrevit import forms

TENSOES = [u'127V', u'220V', u'380V']


def _rgb(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))

C_WHITE   = _rgb(0xFF, 0xFF, 0xFF)
C_BG      = _rgb(0xF8, 0xF8, 0xF8)
C_BORDER  = _rgb(0xC8, 0xC8, 0xC8)
C_DIVIDER = _rgb(0xE0, 0xE0, 0xE0)
C_TEXT1   = _rgb(0x33, 0x33, 0x33)
C_TEXT2   = _rgb(0x77, 0x77, 0x77)
C_RED     = _rgb(0xC4, 0x2B, 0x1A)


def _lbl(text, size=10):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = size
    tb.Foreground = C_TEXT2
    tb.VerticalAlignment = VerticalAlignment.Center
    tb.Margin = Thickness(0, 0, 0, 2)
    return tb


def _input(text=u''):
    b = TextBox()
    b.Text = text
    b.Height = 28
    b.Padding = Thickness(6, 0, 6, 0)
    b.VerticalContentAlignment = VerticalAlignment.Center
    b.FontSize = 11
    b.Background = C_WHITE
    b.Foreground = C_TEXT1
    b.BorderBrush = C_BORDER
    b.BorderThickness = Thickness(1)
    return b


def _combo(items, selected=None):
    cb = ComboBox()
    cb.Height = 28
    cb.FontSize = 11
    cb.Foreground = C_TEXT1
    cb.Background = C_WHITE
    cb.BorderBrush = C_BORDER
    cb.IsEditable = True
    for item in items:
        cb.Items.Add(item)
    if selected in items:
        cb.SelectedIndex = items.index(selected)
    elif items:
        cb.SelectedIndex = 0
    return cb


def _field(label, widget):
    sp = StackPanel()
    sp.Children.Add(_lbl(label))
    sp.Children.Add(widget)
    return sp


# ── ProfileEntryRow ───────────────────────────────────────────────────────────

class ProfileEntryRow(object):

    def __init__(self, family_name, entry, load_types, on_delete):
        self._load_types = load_types
        self._on_delete  = on_delete
        self.root        = self._build(family_name, entry)

    def _build(self, family_name, entry):
        card = Border()
        card.Background      = C_WHITE
        card.BorderBrush     = C_BORDER
        card.BorderThickness = Thickness(1)
        card.CornerRadius    = CornerRadius(4)
        card.Padding         = Thickness(12, 10, 12, 10)
        card.Margin          = Thickness(0, 0, 0, 6)

        outer = Grid()
        for w in [GridLength.Auto, GridLength(1, GridUnitType.Star), GridLength.Auto]:
            cd = ColumnDefinition(); cd.Width = w
            outer.ColumnDefinitions.Add(cd)

        # Checkbox
        self._chk = CheckBox()
        self._chk.IsChecked         = entry.get(u'checked', True)
        self._chk.VerticalAlignment = VerticalAlignment.Top
        self._chk.Margin            = Thickness(0, 8, 10, 0)
        self._chk.ToolTip           = u'Marcar/desmarcar esta família'
        Grid.SetColumn(self._chk, 0)

        # Botão excluir
        btn = Button()
        btn.Content          = u'✕'
        btn.Width            = 26
        btn.Height           = 26
        btn.FontSize         = 11
        btn.BorderThickness  = Thickness(1)
        btn.BorderBrush      = C_BORDER
        btn.Background       = C_BG
        btn.Foreground       = C_RED
        btn.VerticalAlignment = VerticalAlignment.Top
        btn.Margin           = Thickness(8, 4, 0, 0)
        btn.Cursor           = System.Windows.Input.Cursors.Hand
        btn.ToolTip          = u'Remover esta entrada'
        btn.Click           += lambda s, e: self._on_delete(self)
        Grid.SetColumn(btn, 2)

        # Conteúdo
        body = StackPanel()
        Grid.SetColumn(body, 1)

        # Nome da família (chave)
        self._txt_key = _input(family_name)
        self._txt_key.FontSize   = 12
        self._txt_key.FontWeight = System.Windows.FontWeights.SemiBold
        self._txt_key.Margin     = Thickness(0, 0, 0, 8)
        self._txt_key.ToolTip    = u'Nome da família — chave do perfil'
        body.Children.Add(self._txt_key)

        # Linha PpV: Ponto Elétrico | Ponto de Dados
        row_ppv = Grid()
        for w in [GridLength(1, GridUnitType.Star), GridLength(10), GridLength(1, GridUnitType.Star)]:
            cd = ColumnDefinition(); cd.Width = w
            row_ppv.ColumnDefinitions.Add(cd)

        self._txt_pe = _input(entry.get(u'ponto_eletrico', u''))
        self._txt_pd = _input(entry.get(u'ponto_dados',    u''))
        f_pe = _field(u'Ponto Elétrico', self._txt_pe)
        f_pd = _field(u'Ponto de Dados', self._txt_pd)
        Grid.SetColumn(f_pe, 0)
        Grid.SetColumn(f_pd, 2)
        row_ppv.Children.Add(f_pe)
        row_ppv.Children.Add(f_pd)
        body.Children.Add(row_ppv)

        div = Border()
        div.Height = 1; div.Background = C_DIVIDER
        div.Margin = Thickness(0, 8, 0, 8)
        body.Children.Add(div)

        # Linha AE: Altura | Carga | Tensão | Prefixo | Tipo de Carga
        row_ae = Grid()
        for i in range(9):
            cd = ColumnDefinition()
            cd.Width = GridLength(8) if i % 2 == 1 else GridLength(1, GridUnitType.Star)
            row_ae.ColumnDefinitions.Add(cd)

        self._txt_altura  = _input(str(entry.get(u'altura',   u'1.20')))
        self._txt_carga   = _input(str(entry.get(u'carga_va', u'600')))
        self._cmb_tensao  = _combo(TENSOES, entry.get(u'tensao', u'220V'))
        self._txt_prefixo = _input(entry.get(u'prefixo', u''))
        tipo_items = [u''] + self._load_types
        tipo_val   = entry.get(u'tipo_carga', u'')
        self._cmb_tipo = _combo(tipo_items, tipo_val if tipo_val in tipo_items else None)
        if tipo_val and tipo_val not in tipo_items:
            self._cmb_tipo.Text = tipo_val

        ae_defs = [
            (u'Altura (m)',    self._txt_altura),
            (u'Carga (VA)',    self._txt_carga),
            (u'Tensão',        self._cmb_tensao),
            (u'Prefixo',       self._txt_prefixo),
            (u'Tipo de Carga', self._cmb_tipo),
        ]
        for i, (lbl, widget) in enumerate(ae_defs):
            col = _field(lbl, widget)
            Grid.SetColumn(col, i * 2)
            row_ae.Children.Add(col)
        body.Children.Add(row_ae)

        outer.Children.Add(self._chk)
        outer.Children.Add(body)
        outer.Children.Add(btn)
        card.Child = outer
        return card

    def get_key(self):
        return (self._txt_key.Text or u'').strip()

    def get_entry(self):
        tensao = str(self._cmb_tensao.SelectedItem) if self._cmb_tensao.SelectedItem else u'220V'
        tipo   = str(self._cmb_tipo.SelectedItem)   if self._cmb_tipo.SelectedItem   else u''
        return {
            u'ponto_eletrico': (self._txt_pe.Text    or u'').strip(),
            u'ponto_dados':    (self._txt_pd.Text    or u'').strip(),
            u'checked':        bool(self._chk.IsChecked),
            u'altura':         (self._txt_altura.Text.replace(u',', u'.') or u'1.20'),
            u'carga_va':       (self._txt_carga.Text.replace(u',', u'.') or u'600'),
            u'tensao':         tensao,
            u'prefixo':        (self._txt_prefixo.Text or u'').strip(),
            u'tipo_carga':     tipo,
        }


# ── ProfileManagerController ──────────────────────────────────────────────────

class ProfileManagerController(object):

    def __init__(self, win, dbg, profiles_dir, load_types):
        self._win          = win
        self._dbg          = dbg
        self._profiles_dir = profiles_dir
        self._load_types   = load_types
        self._rows         = []
        self._current_name = None

        self._refresh_combo()

        win.gp_BtnNovo.Click               += self._on_novo
        win.gp_BtnRenomear.Click           += self._on_renomear
        win.gp_BtnExcluir.Click            += self._on_excluir
        win.gp_BtnCarregar.Click           += self._on_carregar
        win.gp_BtnAddEntry.Click           += self._on_add_entry
        win.gp_BtnSalvar.Click             += self._on_salvar
        win.gp_CmbProfile.SelectionChanged += self._on_gp_combo_changed

    # ── Filesystem ────────────────────────────────────────────────────────

    def _list(self):
        try:
            return sorted(
                f[:-5] for f in os.listdir(self._profiles_dir)
                if f.endswith(u'.json') and f != u'ae_data.json'
            )
        except Exception:
            return []

    def _path(self, name):
        return os.path.join(self._profiles_dir, name + u'.json')

    def _load(self, name):
        try:
            with io.open(self._path(name), u'r', encoding=u'utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_file(self, name, data):
        try:
            if not os.path.isdir(self._profiles_dir):
                os.makedirs(self._profiles_dir)
            with io.open(self._path(name), u'w', encoding=u'utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
            return True
        except Exception:
            return False

    def _delete_file(self, name):
        try:
            p = self._path(name)
            if os.path.isfile(p):
                os.remove(p)
        except Exception:
            pass

    # ── Combo ─────────────────────────────────────────────────────────────

    def _refresh_combo(self):
        names = self._list()
        net = List[System.Object]()
        net.Add(u'— Selecionar perfil —')
        for n in names:
            net.Add(n)
        self._win._syncing_profile = True
        try:
            self._win.gp_CmbProfile.ItemsSource = net
            self._win.gp_CmbProfile.SelectedIndex = 0
        finally:
            self._win._syncing_profile = False

    def sync_index(self, idx):
        self._win._syncing_profile = True
        try:
            if self._win.gp_CmbProfile.SelectedIndex != idx:
                self._win.gp_CmbProfile.SelectedIndex = idx
        finally:
            self._win._syncing_profile = False

    # ── Rows ──────────────────────────────────────────────────────────────

    def _build_entries(self, data):
        self._win.gp_PanelEntries.Children.Clear()
        self._rows = []
        for key in sorted(data.keys()):
            self._add_row(key, data[key])

    def _add_row(self, family_name=u'', entry=None):
        row = ProfileEntryRow(
            family_name, entry or {}, self._load_types, self._on_delete_row
        )
        self._rows.append(row)
        self._win.gp_PanelEntries.Children.Add(row.root)
        return row

    def _on_delete_row(self, row):
        try:
            self._win.gp_PanelEntries.Children.Remove(row.root)
            self._rows.remove(row)
            self._status(u'{} entrada(s) — não esqueça de salvar'.format(len(self._rows)))
        except Exception:
            pass

    # ── Eventos ───────────────────────────────────────────────────────────

    def _on_gp_combo_changed(self, sender, e):
        if getattr(self._win, '_syncing_profile', False):
            return
        idx = self._win.gp_CmbProfile.SelectedIndex
        # Sincroniza os outros dois combos
        self._win._syncing_profile = True
        try:
            for combo in [self._win.cb_Profile, self._win.ae_CmbProfile]:
                try:
                    if combo.SelectedIndex != idx:
                        combo.SelectedIndex = idx
                except Exception:
                    pass
        finally:
            self._win._syncing_profile = False
        # Carrega entradas para edição
        if idx <= 0:
            self._win.gp_PanelEntries.Children.Clear()
            self._rows = []
            self._current_name = None
            self._status(u'')
            return
        names = self._list()
        if idx - 1 >= len(names):
            return
        name = names[idx - 1]
        self._current_name = name
        data = self._load(name)
        self._build_entries(data)
        self._status(u"'{}' — {} entrada(s). Edite e clique em Salvar Alterações.".format(name, len(data)))

    def _on_novo(self, sender, e):
        name = forms.ask_for_string(
            prompt=u'Nome do novo perfil:', title=u'Novo Perfil', default=u'Novo Perfil'
        )
        if not name:
            return
        name = re.sub(r'[<>:"/\\|?*]', u'_', name).strip()
        if self._save_file(name, {}):
            self._current_name = name
            self._win.gp_PanelEntries.Children.Clear()
            self._rows = []
            self._win._refresh_all_profiles(name)
            self._status(u"Perfil '{}' criado — adicione entradas e salve.".format(name))

    def _on_renomear(self, sender, e):
        if not self._current_name:
            forms.alert(u'Selecione um perfil para renomear.', exitscript=False)
            return
        new_name = forms.ask_for_string(
            prompt=u'Novo nome:', title=u'Renomear', default=self._current_name
        )
        if not new_name or new_name == self._current_name:
            return
        new_name = re.sub(r'[<>:"/\\|?*]', u'_', new_name).strip()
        data = self._load(self._current_name)
        if self._save_file(new_name, data):
            self._delete_file(self._current_name)
            self._current_name = new_name
            self._win._refresh_all_profiles(new_name)
            self._status(u"Renomeado para '{}'.".format(new_name))

    def _on_excluir(self, sender, e):
        if not self._current_name:
            forms.alert(u'Selecione um perfil para excluir.', exitscript=False)
            return
        if not forms.alert(
            u"Excluir o perfil '{}'?".format(self._current_name),
            title=u'Confirmar', yes=True, no=True
        ):
            return
        name = self._current_name
        self._delete_file(name)
        self._current_name = None
        self._win.gp_PanelEntries.Children.Clear()
        self._rows = []
        self._win._refresh_all_profiles(u'')
        self._status(u"Perfil '{}' excluído.".format(name))

    def _on_add_entry(self, sender, e):
        self._add_row(u'Nova Família', {})
        self._status(u'{} entrada(s) — não esqueça de salvar'.format(len(self._rows)))

    def _on_salvar(self, sender, e):
        if not self._current_name:
            name = forms.ask_for_string(
                prompt=u'Nome do perfil:', title=u'Salvar', default=u'Meu Perfil'
            )
            if not name:
                return
            self._current_name = re.sub(r'[<>:"/\\|?*]', u'_', name).strip()
        data = {row.get_key(): row.get_entry() for row in self._rows if row.get_key()}
        if self._save_file(self._current_name, data):
            self._win._refresh_all_profiles(self._current_name)
            self._status(u"'{}' salvo com {} entrada(s).".format(self._current_name, len(data)))
        else:
            forms.alert(u'Erro ao salvar.', exitscript=False)

    def _on_carregar(self, sender, e):
        if not self._current_name:
            forms.alert(u'Selecione um perfil para carregar.', exitscript=False)
            return
        data = self._load(self._current_name)
        if not data:
            forms.alert(u'Perfil vazio. Adicione entradas e salve primeiro.', exitscript=False)
            return
        try:
            self._win._profile = data
            self._win._apply_profile_to_rows()
        except Exception:
            pass
        try:
            for w in self._win._ae._widgets:
                entry = self._win._find_profile_entry(w.family_name)
                if entry:
                    w.apply_profile(entry)
        except Exception:
            pass
        self._win._refresh_all_profiles(self._current_name)
        self._status(u"'{}' carregado nas abas Pontos por Vínculo e Auto-Elétrica.".format(
            self._current_name))

    def _status(self, msg):
        try:
            self._win.gp_TxtStatus.Text = msg
        except Exception:
            pass
