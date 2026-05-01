# -*- coding: utf-8 -*-
"""Gerenciar Folhas — LF Tools
Gerencia componentes de folhas em lote: parâmetros do carimbo, vistas e revisões.
"""
__title__ = "Gerenciar\nFolhas"
__author__ = "Luís Fernando"

import clr, os, io, json, re

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, FamilyInstance, View, Viewport,
    Transaction, BuiltInCategory, BuiltInParameter, ElementId, StorageType
)
from System.Data import DataTable, DataRowState, DataRowVersion
from System.Windows.Controls import (
    DataGridTextColumn, DataGridTemplateColumn, DataGridLength,
    DataGridLengthUnitType, DataGridRow, CheckBox
)
from System.Windows import (
    DataTemplate, FrameworkElementFactory, HorizontalAlignment, VerticalAlignment
)
from System.Windows.Data import Binding, BindingMode, RelativeSource, RelativeSourceMode
from pyrevit import forms

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ── Perfis ────────────────────────────────────────────────────────────────────

_PROFILES_DIR = os.path.join(os.getenv('APPDATA'), 'pyRevit', 'LFTools', 'sheet_profiles')


def _safe_col(name):
    return re.sub(r'[^\w]', u'_', name)


def _ensure_profiles_dir():
    if not os.path.exists(_PROFILES_DIR):
        try:
            os.makedirs(_PROFILES_DIR)
        except Exception:
            pass


def _load_profiles():
    _ensure_profiles_dir()
    profiles = []
    try:
        for f in sorted(os.listdir(_PROFILES_DIR)):
            if not f.endswith('.json'):
                continue
            try:
                with io.open(os.path.join(_PROFILES_DIR, f), 'r', encoding='utf-8') as fp:
                    profiles.append(json.load(fp))
            except Exception:
                pass
    except Exception:
        pass
    return profiles


def _save_profile(profile):
    _ensure_profiles_dir()
    nome = profile.get('nome', 'sem_nome')
    fname = re.sub(r'[^\w\-]', '_', nome) + '.json'
    with io.open(os.path.join(_PROFILES_DIR, fname), 'w', encoding='utf-8') as fp:
        json.dump(profile, fp, ensure_ascii=False, indent=2)


def _delete_profile(nome):
    fname = re.sub(r'[^\w\-]', '_', nome) + '.json'
    path = os.path.join(_PROFILES_DIR, fname)
    if os.path.exists(path):
        os.remove(path)


# ── Revit helpers ─────────────────────────────────────────────────────────────

def _get_all_sheets():
    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    return sorted(list(sheets), key=lambda s: s.SheetNumber)


def _get_title_block(sheet):
    col = (FilteredElementCollector(doc, sheet.Id)
           .OfCategory(BuiltInCategory.OST_TitleBlocks)
           .OfClass(FamilyInstance))
    return col.FirstElement()


def _get_tb_editable_params(tb):
    """[(display_name, safe_col, value), ...]"""
    if not tb:
        return []
    result = []
    for p in tb.Parameters:
        if p.IsReadOnly:
            continue
        if p.StorageType != StorageType.String:
            continue
        name = p.Definition.Name
        result.append((name, _safe_col(name), p.AsString() or u""))
    return result


def _get_issue_date(sheet):
    try:
        p = sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE)
        return p.AsString() or u"" if p else u""
    except Exception:
        return u""


def _get_all_views():
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    result = []
    for v in views:
        try:
            if v.IsTemplate:
                continue
            if isinstance(v, ViewSheet):
                continue
            result.append(v)
        except Exception:
            pass
    return sorted(result, key=lambda v: v.Name)


def _get_viewports(sheet):
    result = []
    try:
        for vid in sheet.GetAllViewports():
            vp = doc.GetElement(vid)
            if not vp:
                continue
            view = doc.GetElement(vp.ViewId)
            c = vp.GetBoxCenter()
            result.append({
                'vp_id':      str(vid.IntegerValue),
                'view_id':    str(vp.ViewId.IntegerValue),
                'view_name':  view.Name if view else u"",
                'view_type':  view.ViewType.ToString() if view else u"",
                'cx': str(c.X), 'cy': str(c.Y), 'cz': str(c.Z),
            })
    except Exception:
        pass
    return result


def _get_revisions_str(sheet):
    try:
        ids = sheet.GetAllRevisionIds()
        parts = []
        for rid in ids:
            rev = doc.GetElement(rid)
            if rev:
                try:
                    parts.append(u"Rev {} — {}".format(rev.SequenceNumber, rev.Description or u""))
                except Exception:
                    parts.append(u"Rev")
        return u"; ".join(parts) if parts else u"—"
    except Exception:
        return u"—"


def _set_str_param(elem, name, value):
    if not elem:
        return
    try:
        p = elem.LookupParameter(name)
        if p and not p.IsReadOnly and p.StorageType == StorageType.String:
            p.Set(value)
    except Exception:
        pass


def _apply_carimbo_changes(changes, builtin_safe):
    """
    changes = [{sheet_id, tb_id, params:{safe:(display,val)}}]
    builtin_safe = set of safe col names that are BuiltIn ViewSheet params
    """
    with Transaction(doc, u"Gerenciar Folhas — Carimbo") as t:
        t.Start()
        for ch in changes:
            sheet = doc.GetElement(ch['sheet_id'])
            tb    = doc.GetElement(ch['tb_id']) if ch.get('tb_id') else None
            for safe, (display, val) in ch['params'].items():
                if safe in builtin_safe:
                    _set_str_param(sheet, display, val)
                else:
                    _set_str_param(tb, display, val)
        t.Commit()


def _swap_viewport(sheet, old_vp_id, new_view_id):
    with Transaction(doc, u"Gerenciar Folhas — Trocar Vista") as t:
        t.Start()
        try:
            old_vp = doc.GetElement(old_vp_id)
            center = old_vp.GetBoxCenter()
            doc.Delete(old_vp_id)
            doc.Regenerate()
            if Viewport.CanAddViewToSheet(doc, sheet.Id, new_view_id):
                Viewport.Create(doc, sheet.Id, new_view_id, center)
                t.Commit()
                return True
            t.RollBack()
            return False
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass
            return False


# ── Window ────────────────────────────────────────────────────────────────────

# Built-in ViewSheet params: safe_col → (display_name, BuiltInParameter)
_BUILTIN = {
    u"Num":  (u"Número",         BuiltInParameter.SHEET_NUMBER),
    u"Nome": (u"Nome da Folha",  BuiltInParameter.SHEET_NAME),
    u"Data": (u"Data de Emissão", BuiltInParameter.SHEET_ISSUE_DATE),
}


class SheetManagerWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        self._sheets      = []
        self._all_views   = []
        self._profiles    = []
        self._param_map   = {}   # safe → display_name
        self._tb_cols     = []   # [(display, safe)] from title block
        self._carimbo_dt  = None
        self._conteudo_dt = None
        self._revisoes_dt = None

        self._wire_events()
        self._load_data()

    # ── Event wiring ──────────────────────────────────────────────────────────

    def _wire_events(self):
        self.SelectAllBtn.Click          += self._on_select_all
        self.SelectNoneBtn.Click         += self._on_select_none
        self.SetFieldBtn.Click           += self._on_set_field
        self.ApplyProfileBtn.Click       += self._on_apply_profile
        self.ExportBtn.Click             += self._on_export_excel
        self.ImportBtn.Click             += self._on_import_excel
        self.RefreshBtn.Click            += self._on_refresh
        self.ManageProfilesBtn.Click     += self._on_manage_profiles
        self.SwapViewBtn.Click           += self._on_swap_view
        self.ApplyBtn.Click              += self._on_apply

    # ── Data loading ──────────────────────────────────────────────────────────

    def _load_data(self):
        try:
            self._sheets    = _get_all_sheets()
            self._all_views = _get_all_views()
            self._profiles  = _load_profiles()
        except Exception as ex:
            forms.alert(u"Erro ao carregar dados: " + str(ex))
            return

        self._populate_profile_combo()
        self._setup_carimbo_grid()
        self._setup_conteudo_grid()
        self._setup_revisoes_grid()
        self._update_status()

    # ── Profile combo ─────────────────────────────────────────────────────────

    def _populate_profile_combo(self):
        self.ProfileCombo.Items.Clear()
        self.ProfileCombo.Items.Add(u"— Todos —")
        for p in self._profiles:
            self.ProfileCombo.Items.Add(p.get('nome', u'?'))
        self.ProfileCombo.SelectedIndex = 0

    # ── DataGrid column builders ──────────────────────────────────────────────

    def _add_check_col(self, grid):
        """Checkbox that reflects row IsSelected state (read-only visual)."""
        factory = FrameworkElementFactory(CheckBox)
        b = Binding(u"IsSelected")
        b.RelativeSource = RelativeSource(RelativeSourceMode.FindAncestor)
        b.RelativeSource.AncestorType = DataGridRow
        b.Mode = BindingMode.OneWay
        factory.SetBinding(CheckBox.IsCheckedProperty, b)
        factory.SetValue(CheckBox.IsHitTestVisibleProperty, False)
        factory.SetValue(CheckBox.HorizontalAlignmentProperty, HorizontalAlignment.Center)
        factory.SetValue(CheckBox.VerticalAlignmentProperty, VerticalAlignment.Center)

        tmpl = DataTemplate()
        tmpl.VisualTree = factory

        col = DataGridTemplateColumn()
        col.Header = u""
        col.CellTemplate = tmpl
        col.Width = DataGridLength(32.0)
        col.CanUserSort = False
        col.CanUserResize = False
        grid.Columns.Add(col)

    def _add_text_col(self, grid, safe, display, width=120, editable=True, star=False):
        col = DataGridTextColumn()
        col.Header = display
        col.Binding = Binding(safe)
        col.Width = (DataGridLength(1.0, DataGridLengthUnitType.Star)
                     if star else DataGridLength(float(width)))
        col.IsReadOnly = not editable
        grid.Columns.Add(col)

    # ── Carimbo grid ──────────────────────────────────────────────────────────

    def _setup_carimbo_grid(self):
        dt = DataTable()
        dt.Columns.Add(u"_SheetId", str)
        dt.Columns.Add(u"_TbId",    str)

        # Built-in columns always present
        for safe, (display, _) in _BUILTIN.items():
            dt.Columns.Add(safe, str)
            self._param_map[safe] = display

        # TB-specific columns from first sheet that has a title block
        self._tb_cols = []
        for sheet in self._sheets:
            tb = _get_title_block(sheet)
            if tb:
                for disp, safe, _ in _get_tb_editable_params(tb):
                    if safe not in dt.Columns:
                        dt.Columns.Add(safe, str)
                        self._tb_cols.append((disp, safe))
                        self._param_map[safe] = disp
                break

        for sheet in self._sheets:
            tb  = _get_title_block(sheet)
            row = dt.NewRow()
            row[u"_SheetId"] = str(sheet.Id.IntegerValue)
            row[u"_TbId"]    = str(tb.Id.IntegerValue) if tb else u""
            row[u"Num"]      = sheet.SheetNumber or u""
            row[u"Nome"]     = sheet.Name        or u""
            row[u"Data"]     = _get_issue_date(sheet)
            if tb:
                for disp, safe, val in _get_tb_editable_params(tb):
                    if safe in dt.Columns:
                        row[safe] = val
            dt.Rows.Add(row)

        dt.AcceptChanges()
        self._carimbo_dt = dt

        self.CarimboGrid.Columns.Clear()
        self._add_check_col(self.CarimboGrid)
        self._add_text_col(self.CarimboGrid, u"Num",  u"Número",        80,  True)
        self._add_text_col(self.CarimboGrid, u"Nome", u"Nome da Folha", 220, True)
        self._add_text_col(self.CarimboGrid, u"Data", u"Data Emissão",  110, True)
        for disp, safe in self._tb_cols:
            self._add_text_col(self.CarimboGrid, safe, disp, 150, True)

        self.CarimboGrid.ItemsSource = dt.DefaultView

    # ── Conteúdo grid ─────────────────────────────────────────────────────────

    def _setup_conteudo_grid(self):
        dt = DataTable()
        for col in (u"_SheetId", u"_VpId", u"_ViewId",
                    u"_CX", u"_CY", u"_CZ",
                    u"SheetNum", u"SheetName", u"ViewName", u"ViewType"):
            dt.Columns.Add(col, str)

        for sheet in self._sheets:
            for vp in _get_viewports(sheet):
                row = dt.NewRow()
                row[u"_SheetId"]   = str(sheet.Id.IntegerValue)
                row[u"_VpId"]      = vp['vp_id']
                row[u"_ViewId"]    = vp['view_id']
                row[u"_CX"]        = vp['cx']
                row[u"_CY"]        = vp['cy']
                row[u"_CZ"]        = vp['cz']
                row[u"SheetNum"]   = sheet.SheetNumber or u""
                row[u"SheetName"]  = sheet.Name        or u""
                row[u"ViewName"]   = vp['view_name']
                row[u"ViewType"]   = vp['view_type']
                dt.Rows.Add(row)

        dt.AcceptChanges()
        self._conteudo_dt = dt

        self.ConteudoGrid.Columns.Clear()
        self._add_check_col(self.ConteudoGrid)
        self._add_text_col(self.ConteudoGrid, u"SheetNum",  u"Folha",        80,  False)
        self._add_text_col(self.ConteudoGrid, u"SheetName", u"Nome",        200,  False)
        self._add_text_col(self.ConteudoGrid, u"ViewName",  u"Vista Atual", 260,  False)
        self._add_text_col(self.ConteudoGrid, u"ViewType",  u"Tipo",        130,  False)

        self.ConteudoGrid.ItemsSource = dt.DefaultView

    # ── Revisões grid ─────────────────────────────────────────────────────────

    def _setup_revisoes_grid(self):
        dt = DataTable()
        for col in (u"_SheetId", u"SheetNum", u"SheetName", u"Revisoes"):
            dt.Columns.Add(col, str)

        for sheet in self._sheets:
            row = dt.NewRow()
            row[u"_SheetId"]  = str(sheet.Id.IntegerValue)
            row[u"SheetNum"]  = sheet.SheetNumber or u""
            row[u"SheetName"] = sheet.Name        or u""
            row[u"Revisoes"]  = _get_revisions_str(sheet)
            dt.Rows.Add(row)

        dt.AcceptChanges()
        self._revisoes_dt = dt

        self.RevisoesGrid.Columns.Clear()
        self._add_check_col(self.RevisoesGrid)
        self._add_text_col(self.RevisoesGrid, u"SheetNum",  u"Folha",   80,  False)
        self._add_text_col(self.RevisoesGrid, u"SheetName", u"Nome",   220,  False)
        self._add_text_col(self.RevisoesGrid, u"Revisoes",  u"Revisões atribuídas", 0, False, star=True)

        self.RevisoesGrid.ItemsSource = dt.DefaultView

    # ── Status ────────────────────────────────────────────────────────────────

    def _update_status(self):
        total = len(self._sheets)
        tab   = self.MainTabs.SelectedIndex
        grids = [self.CarimboGrid, self.ConteudoGrid, self.RevisoesGrid]
        sel   = 0
        if 0 <= tab < len(grids):
            sel = grids[tab].SelectedItems.Count
        self.StatusLabel.Text = u"{} folha(s) no projeto  |  {} linha(s) selecionada(s)".format(total, sel)

    def _active_grid(self):
        idx = self.MainTabs.SelectedIndex
        return [self.CarimboGrid, self.ConteudoGrid, self.RevisoesGrid][max(0, min(idx, 2))]

    # ── Toolbar events ────────────────────────────────────────────────────────

    def _on_select_all(self, sender, args):
        self._active_grid().SelectAll()
        self._update_status()

    def _on_select_none(self, sender, args):
        self._active_grid().UnselectAll()
        self._update_status()

    def _on_refresh(self, sender, args):
        self._load_data()

    def _on_set_field(self, sender, args):
        if self.MainTabs.SelectedIndex != 0:
            forms.alert(u"'Setar Campo' está disponível na aba Carimbo.")
            return

        selected = list(self.CarimboGrid.SelectedItems)
        if not selected:
            forms.alert(u"Selecione ao menos uma linha na tabela.")
            return

        field_labels = [u"{} ({})".format(v, k)
                        for k, v in sorted(self._param_map.items())]
        chosen = forms.SelectFromList.show(
            field_labels,
            title=u"Setar Campo em Lote",
            button_name=u"Selecionar",
            multiselect=False,
        )
        if not chosen:
            return

        # Parse safe name from "(safe)" suffix
        chosen    = str(chosen)
        safe_name = chosen.split(u"(")[-1].rstrip(u")")
        display   = self._param_map.get(safe_name, safe_name)

        new_val = forms.ask_for_string(
            prompt=u"Novo valor para '{}'  ({} folhas):".format(display, len(selected)),
            title=u"Setar Campo em Lote",
        )
        if new_val is None:
            return

        for item in selected:
            try:
                item[safe_name] = new_val or u""
            except Exception:
                pass

        self._update_status()
        forms.toast(u"{} linha(s) atualizadas (staged — clique Aplicar).".format(len(selected)))

    def _on_apply_profile(self, sender, args):
        idx = self.ProfileCombo.SelectedIndex
        if idx <= 0:
            forms.toast(u"Selecione um perfil no dropdown.")
            return

        profile  = self._profiles[idx - 1]
        defaults = profile.get(u'parametros_padrao', {})
        if not defaults:
            forms.toast(u"Perfil não possui valores padrão definidos.")
            return

        selected = list(self.CarimboGrid.SelectedItems)
        if not selected:
            forms.alert(u"Selecione ao menos uma linha na aba Carimbo.")
            return

        for item in selected:
            for safe, val in defaults.items():
                try:
                    if safe in self._carimbo_dt.Columns:
                        item[safe] = val or u""
                except Exception:
                    pass

        forms.toast(u"Perfil aplicado a {} linha(s) (staged — clique Aplicar).".format(len(selected)))

    def _on_swap_view(self, sender, args):
        selected = list(self.ConteudoGrid.SelectedItems)
        if not selected:
            forms.alert(u"Selecione ao menos uma linha na aba Conteúdo.")
            return

        view_names = [v.Name for v in self._all_views]
        if not view_names:
            forms.alert(u"Nenhuma vista disponível no projeto.")
            return

        chosen = forms.SelectFromList.show(
            view_names,
            title=u"Selecionar Nova Vista",
            button_name=u"Confirmar",
            multiselect=False,
        )
        if not chosen:
            return

        new_view = next((v for v in self._all_views if v.Name == str(chosen)), None)
        if not new_view:
            return

        ok = 0
        fail = 0
        for item in selected:
            try:
                sheet_id = ElementId(int(item[u"_SheetId"]))
                vp_id    = ElementId(int(item[u"_VpId"]))
                sheet    = doc.GetElement(sheet_id)
                if _swap_viewport(sheet, vp_id, new_view.Id):
                    item[u"ViewName"] = new_view.Name
                    item[u"ViewType"] = new_view.ViewType.ToString()
                    item[u"_ViewId"]  = str(new_view.Id.IntegerValue)
                    ok += 1
                else:
                    fail += 1
            except Exception:
                fail += 1

        self._conteudo_dt.AcceptChanges()

        msg = u"✅ {} vista(s) trocadas".format(ok)
        if fail:
            msg += u" | ⚠️ {} falhou".format(fail)
        (forms.toast if ok else forms.alert)(msg)

    def _on_export_excel(self, sender, args):
        forms.toast(u"Export Excel — em desenvolvimento.")

    def _on_import_excel(self, sender, args):
        forms.toast(u"Import Excel — em desenvolvimento.")

    def _on_manage_profiles(self, sender, args):
        opts  = [u"⊕  Criar novo perfil"]
        opts += [u"🗑  Deletar: " + p.get('nome', '?') for p in self._profiles]

        chosen = forms.SelectFromList.show(
            opts, title=u"Gerenciar Perfis",
            button_name=u"Executar",
            multiselect=False,
        )
        if not chosen:
            return
        chosen = str(chosen)

        if u"Criar" in chosen:
            nome = forms.ask_for_string(
                prompt=u"Nome do perfil:", title=u"Novo Perfil")
            if not nome:
                return
            _save_profile({u'nome': nome, u'parametros_padrao': {}})
            self._profiles = _load_profiles()
            self._populate_profile_combo()
            forms.toast(u"Perfil '{}' criado.".format(nome))

        elif u"Deletar" in chosen:
            nome = chosen.split(u"Deletar: ", 1)[-1]
            if forms.alert(u"Deletar perfil '{}'?".format(nome), yes=True, no=True):
                _delete_profile(nome)
                self._profiles = _load_profiles()
                self._populate_profile_combo()
                forms.toast(u"Perfil deletado.")

    # ── Apply ─────────────────────────────────────────────────────────────────

    def _on_apply(self, sender, args):
        tab = self.MainTabs.SelectedIndex
        if tab == 0:
            self._apply_carimbo()
        elif tab == 1:
            forms.toast(u"Use 'Trocar Vista...' para alterar vistas.")
        else:
            forms.toast(u"Gerenciamento de revisões em desenvolvimento.")

    def _apply_carimbo(self):
        try:
            self.CarimboGrid.CommitEdit()
        except Exception:
            pass

        builtin_safe = set(_BUILTIN.keys())
        changes = []

        for row in self._carimbo_dt.Rows:
            if row.RowState == DataRowState.Unchanged:
                continue

            sheet_id_str = str(row[u"_SheetId"])
            tb_id_str    = str(row[u"_TbId"])
            changed      = {}

            for col in self._carimbo_dt.Columns:
                cname = col.ColumnName
                if cname.startswith(u"_"):
                    continue
                try:
                    curr = str(row[cname, DataRowVersion.Current]  or u"")
                    orig = str(row[cname, DataRowVersion.Original] or u"")
                    if curr != orig:
                        display = self._param_map.get(cname, cname)
                        changed[cname] = (display, curr)
                except Exception:
                    pass

            if changed:
                changes.append({
                    u'sheet_id': ElementId(int(sheet_id_str)),
                    u'tb_id':    ElementId(int(tb_id_str)) if tb_id_str else None,
                    u'params':   changed,
                })

        if not changes:
            forms.toast(u"Nenhuma alteração detectada.")
            return

        try:
            _apply_carimbo_changes(changes, builtin_safe)
            self._carimbo_dt.AcceptChanges()
            forms.toast(u"✅ {} folha(s) atualizadas.".format(len(changes)))
        except Exception as ex:
            forms.alert(u"Erro ao aplicar alterações:\n" + str(ex))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not doc or doc.IsFamilyDocument:
        forms.alert(u"Abra um projeto Revit para usar este comando.")
    else:
        cur_dir   = os.path.dirname(__file__)
        xaml_path = os.path.join(cur_dir, 'ui.xaml')
        win = SheetManagerWindow(xaml_path)
        win.show_dialog()
