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
    Transaction, BuiltInCategory, BuiltInParameter, ElementId, StorageType,
    Revision, ViewSchedule, ScheduleSheetInstance,
)
from System.Collections.Generic import List as ClrList
from System.Data import DataTable, DataRowState, DataRowVersion
from System.Windows.Controls import (
    DataGridTextColumn, DataGridTemplateColumn, DataGridCheckBoxColumn,
    DataGridComboBoxColumn, DataGridLength, DataGridLengthUnitType,
    DataGridRow, CheckBox, DataGridEditingUnit
)
from System.Windows import (
    DataTemplate, FrameworkElementFactory, HorizontalAlignment, VerticalAlignment,
    Style, Setter, Thickness
)
from System.Windows.Data import Binding, BindingMode, RelativeSource, RelativeSourceMode, UpdateSourceTrigger
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
            detail_param = getattr(BuiltInParameter, 'VIEWPORT_DETAIL_NUMBER', None)
            detail = _param_as_string(vp, detail_param) if detail_param else u""
            title = _lookup_param_as_string(vp, u"Title on Sheet")
            if not title and view:
                title = _lookup_param_as_string(view, u"Title on Sheet")
            result.append({
                'vp_id':      str(vid.IntegerValue),
                'view_id':    str(vp.ViewId.IntegerValue),
                'view_name':  view.Name if view else u"",
                'view_type':  view.ViewType.ToString() if view else u"",
                'detail_number': detail,
                'title_on_sheet': title,
                'scale':      str(view.Scale) if view and hasattr(view, 'Scale') else u"",
                'cx': str(c.X), 'cy': str(c.Y), 'cz': str(c.Z),
            })
    except Exception:
        pass
    return result


def _param_as_string(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if not p:
            return u""
        return p.AsString() or p.AsValueString() or u""
    except Exception:
        return u""


def _lookup_param_as_string(elem, name):
    try:
        p = elem.LookupParameter(name)
        if not p:
            return u""
        return p.AsString() or p.AsValueString() or u""
    except Exception:
        return u""


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
            new_view = doc.GetElement(new_view_id)
            if isinstance(new_view, ViewSchedule):
                ScheduleSheetInstance.Create(doc, sheet.Id, new_view_id, center)
                t.Commit()
                return True
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


def _diagnose_view_add(sheet, new_view):
    if not sheet or not new_view:
        return u"folha ou vista inválida."
    try:
        if isinstance(new_view, ViewSchedule):
            return u"vista é uma tabela/quadro; usando ScheduleSheetInstance."
    except Exception:
        pass
    try:
        if Viewport.CanAddViewToSheet(doc, sheet.Id, new_view.Id):
            return u"Viewport.CanAddViewToSheet retornou True antes da troca."
    except Exception as ex:
        return u"erro ao consultar CanAddViewToSheet: {}".format(ex)

    try:
        placed = []
        for sh in _get_all_sheets():
            for vp_id in sh.GetAllViewports():
                vp = doc.GetElement(vp_id)
                if vp and vp.ViewId.IntegerValue == new_view.Id.IntegerValue:
                    placed.append(u"{} - {}".format(sh.SheetNumber, sh.Name))
        if placed:
            return u"vista já está em folha: " + u"; ".join(placed[:5])
    except Exception:
        pass

    try:
        if new_view.IsTemplate:
            return u"vista é template."
    except Exception:
        pass

    try:
        return u"Revit recusou a vista para esta folha. Tipo: {}".format(new_view.ViewType)
    except Exception:
        return u"Revit recusou a vista para esta folha."


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

        self._sheets        = []
        self._all_views     = []
        self._profiles      = []
        self._param_map     = {}   # safe → display_name
        self._tb_cols       = []   # [(display, safe)] from title block
        self._carimbo_dt    = None
        self._conteudo_dt   = None
        self._revisoes_dt   = None
        self._doc_revisions = []   # [(label, ElementId)]
        self._debug_enabled = False
        self._debug_lines   = []

        self._wire_events()
        self._load_data()

    # ── Event wiring ──────────────────────────────────────────────────────────

    def _wire_events(self):
        self._bind_click('SelectAllBtn', self._on_select_all)
        self._bind_click('SelectNoneBtn', self._on_select_none)
        self._bind_click('SetFieldBtn', self._on_set_field)
        self._bind_click('ReplicateStampBtn', self._on_replicate_stamp)
        self._bind_click('ApplyProfileBtn', self._on_apply_profile)
        self._bind_click('ExportBtn', self._on_export_excel)
        self._bind_click('ImportBtn', self._on_import_excel)
        self._bind_click('RefreshBtn', self._on_refresh)
        self._bind_click('ManageProfilesBtn', self._on_manage_profiles)
        self._bind_click('SwapViewBtn', self._on_swap_view)
        self._bind_click('SaveViewProfileBtn', self._on_save_view_profile)
        self._bind_click('ApplyBtn', self._on_apply)
        self._bind_click('AddRevisionBtn', self._on_add_revision)
        self._bind_click('RemoveRevisionBtn', self._on_remove_revision)
        self._bind_click('NewRevisionBtn', self._on_new_revision)
        self._bind_click('DebugBtn', self._on_toggle_debug)
        try:
            self.MainTabs.SelectionChanged += self._on_tab_changed
        except Exception:
            pass
        try:
            self.SearchBox.TextChanged += self._on_search_changed
        except Exception:
            pass
        for grid_name in ('CarimboGrid', 'ConteudoGrid', 'RevisoesGrid'):
            try:
                getattr(self, grid_name).SelectionChanged += self._on_grid_selection_changed
            except Exception:
                pass
            try:
                getattr(self, grid_name).CurrentCellChanged += self._on_grid_selection_changed
            except Exception:
                pass

    def _bind_click(self, name, handler):
        try:
            getattr(self, name).Click += handler
        except Exception:
            pass

    def _commit_grid(self, grid):
        try:
            grid.CommitEdit(DataGridEditingUnit.Cell, True)
        except Exception:
            pass
        try:
            grid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass
        try:
            grid.CommitEdit()
        except Exception:
            pass

    def _debug(self, message):
        if not self._debug_enabled:
            return
        try:
            self._debug_lines.append(unicode(message))
        except Exception:
            self._debug_lines.append(str(message))
        try:
            print(u"[Gerenciar Folhas] " + unicode(message))
        except Exception:
            try:
                print("[Gerenciar Folhas] " + str(message))
            except Exception:
                pass

    def _debug_begin(self, title):
        self._debug_lines = []
        self._debug(title)

    def _debug_end(self, title=u"Debug"):
        if not self._debug_enabled or not self._debug_lines:
            return
        forms.alert(u"\n".join(self._debug_lines), title=title)

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
        self._populate_revision_combo()
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

    # ── Revision combo ────────────────────────────────────────────────────────

    def _populate_revision_combo(self):
        try:
            self.RevisionCombo.Items.Clear()
            self._doc_revisions = []
            for rid in Revision.GetAllRevisionIds(doc):
                rev = doc.GetElement(rid)
                if not rev:
                    continue
                try:
                    seq  = rev.SequenceNumber
                    date = rev.RevisionDate or u""
                    desc = rev.Description  or u"(sem descrição)"
                    label = u"Rev {} — {} {}".format(seq, date, desc).strip(" —")
                except Exception:
                    label = u"Revisão {}".format(rid.IntegerValue)
                self._doc_revisions.append((label, rid))
                self.RevisionCombo.Items.Add(label)
            if self.RevisionCombo.Items.Count > 0:
                self.RevisionCombo.SelectedIndex = 0
        except Exception:
            pass

    def _get_selected_revision_id(self):
        if not self._doc_revisions:
            return None
        idx = self.RevisionCombo.SelectedIndex
        if 0 <= idx < len(self._doc_revisions):
            return self._doc_revisions[idx][1]
        try:
            text = str(self.RevisionCombo.Text or u"").strip()
            for label, rid in self._doc_revisions:
                if label == text:
                    return rid
        except Exception:
            pass
        return None

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

    def _add_select_col(self, grid):
        col = DataGridCheckBoxColumn()
        col.Header = u"Sel."
        binding = Binding(u"Selected")
        binding.Mode = BindingMode.TwoWay
        binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        col.Binding = binding
        col.Width = DataGridLength(48.0)
        col.CanUserSort = False
        col.CanUserResize = False
        style = Style(CheckBox)
        style.Setters.Add(Setter(CheckBox.HorizontalAlignmentProperty, HorizontalAlignment.Center))
        style.Setters.Add(Setter(CheckBox.VerticalAlignmentProperty, VerticalAlignment.Center))
        style.Setters.Add(Setter(CheckBox.MarginProperty, Thickness(0)))
        col.ElementStyle = style
        col.EditingElementStyle = style
        grid.Columns.Add(col)

    def _add_text_col(self, grid, safe, display, width=120, editable=True, star=False):
        col = DataGridTextColumn()
        col.Header = (u"✎ " + display) if editable else display
        col.Binding = Binding(safe)
        col.Width = (DataGridLength(1.0, DataGridLengthUnitType.Star)
                     if star else DataGridLength(float(width)))
        col.IsReadOnly = not editable
        grid.Columns.Add(col)

    def _add_combo_col(self, grid, safe, display, items, width=180):
        col = DataGridComboBoxColumn()
        col.Header = display
        binding = Binding(safe)
        binding.Mode = BindingMode.TwoWay
        binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        col.SelectedItemBinding = binding
        col.ItemsSource = items
        col.Width = DataGridLength(float(width))
        try:
            style = self.FindResource("ComboBoxStyle")
            col.ElementStyle = style
            col.EditingElementStyle = style
        except Exception:
            pass
        grid.Columns.Add(col)

    # ── Carimbo grid ──────────────────────────────────────────────────────────

    def _setup_carimbo_grid(self):
        dt = DataTable()
        dt.Columns.Add(u"Selected", bool)
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
            row[u"Selected"] = False
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
        self._add_select_col(self.CarimboGrid)
        self._add_text_col(self.CarimboGrid, u"Num",  u"Número",        80,  True)
        self._add_text_col(self.CarimboGrid, u"Nome", u"Nome da Folha", 220, True)
        self._add_text_col(self.CarimboGrid, u"Data", u"Data Emissão",  110, True)
        for disp, safe in self._tb_cols:
            self._add_text_col(self.CarimboGrid, safe, disp, 150, True)

        self.CarimboGrid.ItemsSource = dt.DefaultView

    # ── Conteúdo grid ─────────────────────────────────────────────────────────

    def _setup_conteudo_grid(self):
        dt = DataTable()
        dt.Columns.Add(u"Selected", bool)
        for col in (u"_SheetId", u"SheetNum", u"SheetName"):
            dt.Columns.Add(col, str)

        sheet_viewports = []
        max_viewports = 0
        for sheet in self._sheets:
            vps = _get_viewports(sheet)
            vps = sorted(vps, key=lambda v: (v.get('detail_number') or u"", v.get('view_name') or u""))
            sheet_viewports.append((sheet, vps))
            max_viewports = max(max_viewports, len(vps))

        for idx in range(max_viewports):
            n = idx + 1
            for col in (
                u"_VpId_{0}".format(n), u"_ViewId_{0}".format(n),
                u"_CX_{0}".format(n), u"_CY_{0}".format(n), u"_CZ_{0}".format(n),
                u"Detail_{0}".format(n), u"View_{0}".format(n),
                u"Title_{0}".format(n), u"Type_{0}".format(n), u"Scale_{0}".format(n),
            ):
                dt.Columns.Add(col, str)

        for sheet, vps in sheet_viewports:
            row = dt.NewRow()
            row[u"Selected"] = False
            row[u"_SheetId"] = str(sheet.Id.IntegerValue)
            row[u"SheetNum"] = sheet.SheetNumber or u""
            row[u"SheetName"] = sheet.Name or u""
            for idx, vp in enumerate(vps):
                n = idx + 1
                row[u"_VpId_{0}".format(n)] = vp['vp_id']
                row[u"_ViewId_{0}".format(n)] = vp['view_id']
                row[u"_CX_{0}".format(n)] = vp['cx']
                row[u"_CY_{0}".format(n)] = vp['cy']
                row[u"_CZ_{0}".format(n)] = vp['cz']
                row[u"Detail_{0}".format(n)] = vp['detail_number']
                row[u"View_{0}".format(n)] = vp['view_name']
                row[u"Title_{0}".format(n)] = vp['title_on_sheet']
                row[u"Type_{0}".format(n)] = vp['view_type']
                row[u"Scale_{0}".format(n)] = vp['scale']
            dt.Rows.Add(row)

        dt.AcceptChanges()
        self._conteudo_dt = dt

        self.ConteudoGrid.Columns.Clear()
        self._add_select_col(self.ConteudoGrid)
        self._add_text_col(self.ConteudoGrid, u"SheetNum",  u"Sheet Number", 90,  False)
        self._add_text_col(self.ConteudoGrid, u"SheetName", u"Sheet Name",   220, False)
        view_names = [u""] + [v.Name for v in self._all_views]
        for idx in range(max_viewports):
            n = idx + 1
            self._add_combo_col(
                self.ConteudoGrid,
                u"View_{0}".format(n),
                u"View {0:02d}".format(n),
                view_names,
                220,
            )

        self.ConteudoGrid.ItemsSource = dt.DefaultView

    # ── Revisões grid ─────────────────────────────────────────────────────────

    def _setup_revisoes_grid(self):
        dt = DataTable()
        dt.Columns.Add(u"Selected", bool)
        for col in (u"_SheetId", u"SheetNum", u"SheetName", u"Revisoes"):
            dt.Columns.Add(col, str)

        for sheet in self._sheets:
            row = dt.NewRow()
            row[u"Selected"]  = False
            row[u"_SheetId"]  = str(sheet.Id.IntegerValue)
            row[u"SheetNum"]  = sheet.SheetNumber or u""
            row[u"SheetName"] = sheet.Name        or u""
            row[u"Revisoes"]  = _get_revisions_str(sheet)
            dt.Rows.Add(row)

        dt.AcceptChanges()
        self._revisoes_dt = dt

        self.RevisoesGrid.Columns.Clear()
        self._add_select_col(self.RevisoesGrid)
        self._add_text_col(self.RevisoesGrid, u"SheetNum",  u"Folha",   80,  False)
        self._add_text_col(self.RevisoesGrid, u"SheetName", u"Nome",   220,  False)
        self._add_text_col(self.RevisoesGrid, u"Revisoes",  u"Revisões atribuídas", 0, False, star=True)

        self.RevisoesGrid.ItemsSource = dt.DefaultView

    # ── Status ────────────────────────────────────────────────────────────────

    def _update_status(self):
        total = len(self._sheets)
        tab   = self.MainTabs.SelectedIndex
        dt = self._active_dt()
        grid = self._active_grid()
        sel = 0
        rows = 0
        if dt:
            try:
                rows = dt.DefaultView.Count
                for rv in dt.DefaultView:
                    try:
                        if bool(rv[u"Selected"]):
                            sel += 1
                    except Exception:
                        pass
            except Exception:
                rows = 0
        self.StatusLabel.Text = u"{} folha(s) | {} linha(s) visíveis | {} linha(s) selecionada(s)".format(total, rows, sel)

    def _active_grid(self):
        idx = self.MainTabs.SelectedIndex
        return [self.CarimboGrid, self.ConteudoGrid, self.RevisoesGrid][max(0, min(idx, 2))]

    def _active_dt(self):
        idx = self.MainTabs.SelectedIndex
        return [self._carimbo_dt, self._conteudo_dt, self._revisoes_dt][max(0, min(idx, 2))]

    def _selected_rows(self, grid, dt=None):
        self._commit_grid(grid)
        selected = []
        if dt:
            for rv in dt.DefaultView:
                try:
                    if bool(rv[u"Selected"]):
                        selected.append(rv)
                except Exception:
                    pass
        if selected:
            return selected
        return list(grid.SelectedItems)

    def _checked_rows(self, grid, dt=None):
        self._commit_grid(grid)
        selected = []
        if dt:
            for rv in dt.DefaultView:
                try:
                    if bool(rv[u"Selected"]):
                        selected.append(rv)
                except Exception:
                    pass
        return selected

    def _current_row(self, grid):
        self._commit_grid(grid)
        try:
            item = grid.CurrentItem
            if item:
                return item
        except Exception:
            pass
        try:
            if grid.SelectedItems.Count:
                return grid.SelectedItems[0]
        except Exception:
            pass
        return None

    def _set_visible_selected(self, value):
        dt = self._active_dt()
        if not dt:
            return
        for rv in dt.DefaultView:
            try:
                rv[u"Selected"] = value
            except Exception:
                pass

    def _escape_filter_text(self, text):
        return (text or u"").replace("'", "''").replace("[", "[[]").replace("%", "[%]").replace("*", "[*]")

    def _apply_search_filter(self):
        dt = self._active_dt()
        if not dt:
            return
        q = u""
        try:
            q = unicode(self.SearchBox.Text or u"").strip()
        except Exception:
            try:
                q = str(self.SearchBox.Text or u"").strip()
            except Exception:
                q = u""
        if not q:
            dt.DefaultView.RowFilter = u""
            self._update_status()
            return
        q = self._escape_filter_text(q)
        idx = self.MainTabs.SelectedIndex
        if idx == 1 and dt:
            cols = [c.ColumnName for c in dt.Columns
                    if c.ColumnName in (u"SheetNum", u"SheetName") or c.ColumnName.startswith(u"View_")]
        else:
            cols = [
                [u"Num", u"Nome", u"Data"],
                [u"SheetNum", u"SheetName"],
                [u"SheetNum", u"SheetName", u"Revisoes"],
            ][max(0, min(idx, 2))]
        clauses = [u"Convert({}, 'System.String') LIKE '%{}%'".format(c, q)
                   for c in cols if c in dt.Columns]
        dt.DefaultView.RowFilter = u" OR ".join(clauses)
        self._update_status()

    # ── Toolbar events ────────────────────────────────────────────────────────

    def _on_select_all(self, sender, args):
        self._set_visible_selected(True)
        self._active_grid().SelectAll()
        self._update_status()

    def _on_select_none(self, sender, args):
        self._set_visible_selected(False)
        self._active_grid().UnselectAll()
        self._update_status()

    def _on_tab_changed(self, sender, args):
        self._apply_search_filter()

    def _on_search_changed(self, sender, args):
        self._apply_search_filter()

    def _on_grid_selection_changed(self, sender, args):
        self._update_status()

    def _on_toggle_debug(self, sender, args):
        self._debug_enabled = not self._debug_enabled
        try:
            self.DebugBtn.Content = u"Debug ON" if self._debug_enabled else u"Debug"
        except Exception:
            pass
        forms.alert(
            u"Debug ligado. Ao clicar Aplicar, o plugin vai mostrar um relatório passo a passo."
            if self._debug_enabled else u"Debug desligado.",
            title=u"Debug",
        )
        self._debug(u"Debug habilitado.")

    def _on_refresh(self, sender, args):
        self._load_data()

    def _on_set_field(self, sender, args):
        if self.MainTabs.SelectedIndex != 0:
            forms.alert(u"'Setar Campo' está disponível na aba Carimbo.")
            return

        selected = self._selected_rows(self.CarimboGrid, self._carimbo_dt)
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

    def _on_replicate_stamp(self, sender, args):
        if self.MainTabs.SelectedIndex != 0:
            forms.alert(u"'Replicar Carimbo' está disponível na aba Carimbo.")
            return

        source = self._current_row(self.CarimboGrid)
        targets = self._checked_rows(self.CarimboGrid, self._carimbo_dt)
        if not source:
            forms.alert(u"Clique na linha que será usada como modelo.")
            return
        if not targets:
            forms.alert(u"Marque as folhas que receberão os dados do carimbo.")
            return

        source_id = self._row_value(source, u"_SheetId")
        fields = []
        for col in self._carimbo_dt.Columns:
            cname = col.ColumnName
            if cname.startswith(u"_") or cname == u"Selected":
                continue
            fields.append(cname)

        copied = 0
        for row in targets:
            if self._row_value(row, u"_SheetId") == source_id:
                continue
            for cname in fields:
                try:
                    row[cname] = source[cname] or u""
                except Exception:
                    pass
            copied += 1

        if copied:
            forms.toast(u"Carimbo replicado para {} folha(s) (clique Aplicar).".format(copied))
        else:
            forms.toast(u"Nenhuma folha de destino diferente da linha modelo.")

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

        selected = self._selected_rows(self.CarimboGrid, self._carimbo_dt)
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
        forms.toast(u"Use os dropdowns da aba Vistas e clique Aplicar para trocar as vistas.")

    def _on_add_revision(self, sender, args):
        rev_id = self._get_selected_revision_id()
        if not rev_id:
            forms.toast(u"Selecione uma revisão na lista acima.")
            return
        selected = self._selected_rows(self.RevisoesGrid, self._revisoes_dt)
        if not selected:
            forms.alert(u"Selecione ao menos uma folha na tabela.")
            return
        ok = 0
        with Transaction(doc, u"Gerenciar Folhas — Atribuir Revisão") as t:
            t.Start()
            try:
                for row in selected:
                    try:
                        sheet = doc.GetElement(ElementId(int(str(row[u"_SheetId"]))))
                        if not sheet:
                            continue
                        current_ids = list(sheet.GetAdditionalRevisionIds())
                        current_int = [e.IntegerValue for e in current_ids]
                        if rev_id.IntegerValue not in current_int:
                            current_ids.append(rev_id)
                            ids_list = ClrList[ElementId](current_ids)
                            sheet.SetAdditionalRevisionIds(ids_list)
                            ok += 1
                    except Exception:
                        pass
                t.Commit()
            except Exception:
                try:
                    t.RollBack()
                except Exception:
                    pass
        if ok > 0:
            self._setup_revisoes_grid()
            forms.toast(u"Revisão atribuída a {} folha(s).".format(ok))
        else:
            forms.toast(u"Nenhuma alteração realizada (revisão já atribuída ou sem folhas válidas).")

    def _on_new_revision(self, sender, args):
        selected = self._selected_rows(self.RevisoesGrid, self._revisoes_dt)
        if not selected:
            forms.alert(u"Selecione as folhas que receberão a nova revisão.")
            return

        desc = forms.ask_for_string(
            prompt=u"Descrição da nova revisão:",
            title=u"Nova Revisão",
        )
        if desc is None:
            return
        date = forms.ask_for_string(
            prompt=u"Data da revisão:",
            title=u"Nova Revisão",
        )
        if date is None:
            return

        ok = 0
        new_rev = None
        with Transaction(doc, u"Gerenciar Folhas — Nova Revisão") as t:
            t.Start()
            try:
                new_rev = Revision.Create(doc)
                try:
                    new_rev.Description = desc or u""
                except Exception:
                    pass
                try:
                    new_rev.RevisionDate = date or u""
                except Exception:
                    pass

                for row in selected:
                    try:
                        sheet = doc.GetElement(ElementId(int(str(row[u"_SheetId"]))))
                        if not sheet:
                            continue
                        current_ids = list(sheet.GetAdditionalRevisionIds())
                        current_ids.append(new_rev.Id)
                        sheet.SetAdditionalRevisionIds(ClrList[ElementId](current_ids))
                        ok += 1
                    except Exception as ex:
                        self._debug(u"Falha ao atribuir nova revisão: " + str(ex))
                t.Commit()
            except Exception as ex:
                self._debug(u"Falha ao criar nova revisão: " + str(ex))
                try:
                    t.RollBack()
                except Exception:
                    pass
                forms.alert(u"Não foi possível criar a nova revisão:\n" + str(ex))
                return

        self._populate_revision_combo()
        self._setup_revisoes_grid()
        forms.toast(u"Nova revisão criada e atribuída a {} folha(s).".format(ok))

    def _on_remove_revision(self, sender, args):
        rev_id = self._get_selected_revision_id()
        if not rev_id:
            forms.toast(u"Selecione uma revisão na lista acima.")
            return
        selected = self._selected_rows(self.RevisoesGrid, self._revisoes_dt)
        if not selected:
            forms.alert(u"Selecione ao menos uma folha na tabela.")
            return
        ok = 0
        with Transaction(doc, u"Gerenciar Folhas — Remover Revisão") as t:
            t.Start()
            try:
                for row in selected:
                    try:
                        sheet = doc.GetElement(ElementId(int(str(row[u"_SheetId"]))))
                        if not sheet:
                            continue
                        current_ids = list(sheet.GetAdditionalRevisionIds())
                        new_ids = [e for e in current_ids
                                   if e.IntegerValue != rev_id.IntegerValue]
                        if len(new_ids) < len(current_ids):
                            ids_list = ClrList[ElementId](new_ids)
                            sheet.SetAdditionalRevisionIds(ids_list)
                            ok += 1
                    except Exception:
                        pass
                t.Commit()
            except Exception:
                try:
                    t.RollBack()
                except Exception:
                    pass
        if ok > 0:
            self._setup_revisoes_grid()
            forms.toast(u"Revisão removida de {} folha(s).".format(ok))
        else:
            forms.toast(u"Nenhuma alteração realizada (revisão não encontrada nas folhas selecionadas).")

    def _row_value(self, row, key):
        try:
            return unicode(row[key] or u"")
        except Exception:
            try:
                return str(row[key] or u"")
            except Exception:
                return u""

    def _on_save_view_profile(self, sender, args):
        selected = self._selected_rows(self.ConteudoGrid, self._conteudo_dt)
        if not selected:
            forms.alert(u"Selecione as vistas/detalhes na View List antes de salvar o perfil.")
            return

        first_sheet = self._row_value(selected[0], u"SheetNum")
        default_name = u"Perfil_" + (first_sheet or u"folha")
        nome = forms.ask_for_string(
            default=default_name,
            prompt=u"Nome do perfil JSON desta folha/conjunto:",
            title=u"Salvar Perfil de View List",
        )
        if not nome:
            return

        sheets = {}
        views = []
        for row in selected:
            sheet_id = self._row_value(row, u"_SheetId")
            sheet_num = self._row_value(row, u"SheetNum")
            sheet_name = self._row_value(row, u"SheetName")
            if sheet_id not in sheets:
                sheets[sheet_id] = {
                    u"id": sheet_id,
                    u"number": sheet_num,
                    u"name": sheet_name,
                }
            for col in self._conteudo_dt.Columns:
                cname = col.ColumnName
                if not cname.startswith(u"View_"):
                    continue
                n = cname.split(u"_", 1)[-1]
                view_name = self._row_value(row, cname)
                vp_id = self._row_value(row, u"_VpId_" + n)
                if not view_name and not vp_id:
                    continue
                views.append({
                    u"sheet_id": sheet_id,
                    u"sheet_number": sheet_num,
                    u"sheet_name": sheet_name,
                    u"slot": n,
                    u"viewport_id": vp_id,
                    u"view_id": self._row_value(row, u"_ViewId_" + n),
                    u"detail_number": self._row_value(row, u"Detail_" + n),
                    u"view_name": view_name,
                    u"title_on_sheet": self._row_value(row, u"Title_" + n),
                    u"view_type": self._row_value(row, u"Type_" + n),
                    u"scale": self._row_value(row, u"Scale_" + n),
                    u"position": {
                        u"x": self._row_value(row, u"_CX_" + n),
                        u"y": self._row_value(row, u"_CY_" + n),
                        u"z": self._row_value(row, u"_CZ_" + n),
                    },
                })

        profile = {
            u"nome": nome,
            u"tipo": u"view_list_profile",
            u"sheet_profiles": list(sheets.values()),
            u"views": views,
            u"parametros_padrao": {},
        }
        try:
            _save_profile(profile)
            self._profiles = _load_profiles()
            self._populate_profile_combo()
            forms.toast(u"Perfil '{}' salvo com {} vista(s).".format(nome, len(views)))
        except Exception as ex:
            forms.alert(u"Erro ao salvar perfil JSON:\n" + str(ex))

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
            self._apply_view_list()
        else:
            forms.toast(u"Use '＋ Atribuir' ou '− Remover' para gerenciar revisões nas folhas selecionadas.")

    def _apply_view_list(self):
        self._debug_begin(u"Aplicar Vistas")
        self._commit_grid(self.ConteudoGrid)

        if not self._conteudo_dt:
            self._debug(u"DataTable de vistas não está carregado.")
            self._debug_end(u"Debug — Aplicar Vistas")
            return

        view_by_name = {}
        for view in self._all_views:
            if view.Name not in view_by_name:
                view_by_name[view.Name] = view
        self._debug(u"Vistas disponíveis no projeto: {}".format(len(view_by_name)))

        changes = []
        for row in self._conteudo_dt.Rows:
            if row.RowState == DataRowState.Unchanged:
                continue
            sheet_id_str = str(row[u"_SheetId"])
            sheet_label = u"{} - {}".format(row[u"SheetNum"], row[u"SheetName"])
            self._debug(u"Linha alterada: " + sheet_label)
            for col in self._conteudo_dt.Columns:
                cname = col.ColumnName
                if not cname.startswith(u"View_"):
                    continue
                try:
                    curr = str(row[cname, DataRowVersion.Current] or u"")
                    orig = str(row[cname, DataRowVersion.Original] or u"")
                except Exception:
                    continue
                if curr == orig or not curr:
                    continue
                n = cname.split(u"_", 1)[-1]
                vp_id_str = str(row[u"_VpId_" + n] or u"")
                if not vp_id_str:
                    self._debug(u"  Slot {} ignorado: sem viewport existente.".format(n))
                    continue
                new_view = view_by_name.get(curr)
                if not new_view:
                    self._debug(u"  Slot {} ignorado: vista '{}' não encontrada.".format(n, curr))
                    continue
                self._debug(u"  Slot {}: '{}' -> '{}'".format(n, orig, curr))
                changes.append((sheet_id_str, vp_id_str, new_view))

        if not changes:
            self._debug(u"Apply Vistas: nenhuma alteração detectada no DataTable.")
            self._debug_end(u"Debug — Aplicar Vistas")
            forms.toast(u"Nenhuma troca de vista detectada.")
            return

        ok = 0
        fail = 0
        for sheet_id_str, vp_id_str, new_view in changes:
            try:
                sheet = doc.GetElement(ElementId(int(sheet_id_str)))
                vp_id = ElementId(int(vp_id_str))
                self._debug(u"Executando: folha {} viewport {} nova vista '{}'".format(
                    sheet.SheetNumber if sheet else sheet_id_str,
                    vp_id_str,
                    new_view.Name,
                ))
                self._debug(u"  Diagnóstico prévio: " + _diagnose_view_add(sheet, new_view))
                if _swap_viewport(sheet, vp_id, new_view.Id):
                    self._debug(u"  OK")
                    ok += 1
                else:
                    self._debug(u"  Falhou: " + _diagnose_view_add(sheet, new_view))
                    fail += 1
            except Exception as ex:
                self._debug(u"  Exceção em sheet {} viewport {}: {}".format(sheet_id_str, vp_id_str, ex))
                fail += 1

        self._load_data()
        msg = u"{} vista(s) trocada(s).".format(ok)
        if fail:
            msg += u" {} falhou/falharam.".format(fail)
        self._debug(u"Resumo: {} OK | {} falha(s)".format(ok, fail))
        self._debug_end(u"Debug — Aplicar Vistas")
        (forms.toast if ok else forms.alert)(msg)

    def _apply_carimbo(self):
        self._debug_begin(u"Aplicar Carimbo")
        self._commit_grid(self.CarimboGrid)

        builtin_safe = set(_BUILTIN.keys())
        changes = []

        for row in self._carimbo_dt.Rows:
            if row.RowState == DataRowState.Unchanged:
                continue

            sheet_id_str = str(row[u"_SheetId"])
            tb_id_str    = str(row[u"_TbId"])
            changed      = {}
            self._debug(u"Linha alterada: {} - {}".format(row[u"Num"], row[u"Nome"]))

            for col in self._carimbo_dt.Columns:
                cname = col.ColumnName
                if cname.startswith(u"_") or cname == u"Selected":
                    continue
                try:
                    curr = str(row[cname, DataRowVersion.Current]  or u"")
                    orig = str(row[cname, DataRowVersion.Original] or u"")
                    if curr != orig:
                        display = self._param_map.get(cname, cname)
                        changed[cname] = (display, curr)
                        self._debug(u"  {}: '{}' -> '{}'".format(display, orig, curr))
                except Exception:
                    self._debug(u"  Falha ao ler coluna {}".format(cname))

            if changed:
                changes.append({
                    u'sheet_id': ElementId(int(sheet_id_str)),
                    u'tb_id':    ElementId(int(tb_id_str)) if tb_id_str else None,
                    u'params':   changed,
                })

        if not changes:
            self._debug(u"Nenhuma alteração real encontrada no carimbo.")
            self._debug_end(u"Debug — Aplicar Carimbo")
            forms.toast(u"Nenhuma alteração detectada.")
            return

        try:
            self._debug(u"Aplicando alterações em {} folha(s).".format(len(changes)))
            _apply_carimbo_changes(changes, builtin_safe)
            self._carimbo_dt.AcceptChanges()
            self._debug(u"Transação concluída.")
            self._debug_end(u"Debug — Aplicar Carimbo")
            forms.toast(u"✅ {} folha(s) atualizadas.".format(len(changes)))
        except Exception as ex:
            self._debug(u"Erro na transação: " + str(ex))
            self._debug_end(u"Debug — Aplicar Carimbo")
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
