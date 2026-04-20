# -*- coding: utf-8 -*-
"""
Auto-Cotas - Dimensionamento Automático por Eixo Central
=========================================================
LF Tools - pyRevit Extension

Dois modos:
  - Cotar elementos selecionados (misto)
  - Cotar por categorias: Referência (base) + Alvos (o que cotar)
    Ex: Paredes (ref) → Conduítes, Dispositivos (alvos)
    Cada categoria-alvo gera sua própria cadeia de cotas.
"""
import clr
import math
import System
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ElementId,
    Transaction, Line, XYZ,
    FamilyInstance, Wall,
    Reference, ReferenceArray, Options,
    Dimension, DimensionType, ViewPlan, ViewSection,
    LocationCurve, LocationPoint,
    Grid as RevitGrid
)
import Autodesk.Revit.DB as DB

from pyrevit import revit, forms, script

doc = revit.doc
uidoc = revit.uidoc
app = revit.HOST_APP.app
active_view = doc.ActiveView

MM_PER_FT = 304.8

# ====== Categorias ======
class CategoryOption(object):
    def __init__(self, name, bic, checked=False):
        self.name = name
        self.bic = bic
        self.is_checked = checked

# Referências (elementos de base)
REF_CATEGORIES = [
    ("Paredes",                  BuiltInCategory.OST_Walls),
    ("Pisos",                    BuiltInCategory.OST_Floors),
    ("Eixos (Grids)",            BuiltInCategory.OST_Grids),
    ("Pilares Estruturais",      BuiltInCategory.OST_StructuralColumns),
    ("Colunas Arquitetônicas",   BuiltInCategory.OST_Columns),
]

# Alvos (o que cotar em relação à referência)
TARGET_CATEGORIES = [
    ("Conduítes",                   BuiltInCategory.OST_Conduit),
    ("Eletrocalhas",                BuiltInCategory.OST_CableTray),
    ("Tubulações",                  BuiltInCategory.OST_PipeCurves),
    ("Dutos (HVAC)",                BuiltInCategory.OST_DuctCurves),
    ("Dispositivos Elétricos",      BuiltInCategory.OST_ElectricalFixtures),
    ("Dispositivos de Iluminação",  BuiltInCategory.OST_LightingDevices),
    ("Luminárias",                  BuiltInCategory.OST_LightingFixtures),
    ("Dispositivos de Comunicação", BuiltInCategory.OST_CommunicationDevices),
    ("Equipamentos Mecânicos",      BuiltInCategory.OST_MechanicalEquipment),
    ("Conexões de Conduíte",        BuiltInCategory.OST_ConduitFitting),
    ("Conexões de Tubulação",       BuiltInCategory.OST_PipeFitting),
]

# ====== Funções Geométricas ======

def get_element_curve(el):
    """Extrai a curva central (eixo) de um elemento."""
    if isinstance(el, RevitGrid):
        return el.Curve
    loc = el.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve
    return None


def get_element_reference(el):
    """Obtém uma Reference geométrica válida para o eixo central do elemento."""
    # 1. Grids
    if isinstance(el, RevitGrid):
        try: return Reference(el)
        except: pass

    # 2. Famílias (Luminárias, Dispositivos, etc.) - Tenta planos centrais primeiro
    if isinstance(el, FamilyInstance):
        for ref_type in [FamilyInstanceReferenceType.CenterLeftRight, 
                        FamilyInstanceReferenceType.CenterFrontBack,
                        FamilyInstanceReferenceType.Left,
                        FamilyInstanceReferenceType.Right]:
            try:
                refs = el.GetReferences(ref_type)
                if refs: return refs[0]
            except: pass

    # 3. Curvas (Conduítes, Tubos, Dutos)
    loc = el.Location
    if isinstance(loc, LocationCurve):
        try:
            # Em Conduítes/Tubos, o Reference do elemento costuma ser o eixo
            return Reference(el)
        except: pass

    # 4. Fallback: Geometria explícita
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    opt.View = active_view

    geom = el.get_Geometry(opt)
    if geom:
        for obj in geom:
            # Tenta pegar referências de instâncias
            if hasattr(obj, 'GetInstanceGeometry'):
                for sub in obj.GetInstanceGeometry():
                    if hasattr(sub, 'Reference') and sub.Reference: return sub.Reference
            # Tenta pegar referências diretas (linhas/faces)
            if hasattr(obj, 'Reference') and obj.Reference: return obj.Reference
            
    # 5. Último recurso: Referência direta do elemento
    try: return Reference(el)
    except: pass

    return None


def get_element_position_on_axis(el, axis):
    """Obtém posição do ponto médio do elemento projetada no eixo."""
    curve = get_element_curve(el)
    if curve:
        mid = curve.Evaluate(0.5, True)
        return mid.DotProduct(axis)

    loc = el.Location
    if isinstance(loc, LocationPoint):
        return loc.Point.DotProduct(axis)

    try:
        bb = el.get_BoundingBox(active_view)
        if bb:
            center = (bb.Min + bb.Max) * 0.5
            return center.DotProduct(axis)
    except:
        pass

    return None


def get_dim_types():
    """Coleta DimensionTypes de estilo Alinhado (Aligned) ou Linear."""
    col = FilteredElementCollector(doc).OfClass(DimensionType)
    types = []
    for dt in col:
        try:
            style = dt.StyleType.ToString()
            if style in ["Aligned", "Linear"]:
                types.append(dt)
        except:
            pass

    def sort_key(x):
        try:
            s = x.StyleType.ToString()
            if s == "Aligned": return 0
            if s == "Linear":  return 1
            return 2
        except: return 3

    types.sort(key=sort_key)
    return types


def create_individual_dimensions(elements, direction, offset_mm, dim_type=None, base_offset_ft=0.0):
    """
    Cria uma cadeia de cotas (multi-segmento) para os elementos dados.

    Retorna (n_elementos_cotados, final_offset_ft_usado).
    Retorna (0, base_offset_ft) se não foi possível criar a cota.

    base_offset_ft: posição mínima (em pés) para a linha de cota — usado para
                    empilhar cadeias de diferentes categorias sem sobreposição.
    """
    view = active_view
    offset_ft = offset_mm / MM_PER_FT

    # 1. Direção de ordenação
    master_dir = None
    for el in elements:
        c = get_element_curve(el)
        if c and isinstance(c, Line):
            master_dir = (c.GetEndPoint(1) - c.GetEndPoint(0)).Normalize()
            break

    view_right  = view.RightDirection
    view_up     = view.UpDirection
    view_normal = view.ViewDirection
    view_plane_depth = view.Origin.DotProduct(view_normal)

    # Melhoria de Robustez: Força eixos puros se master_dir estiver quase alinhado
    def _clean_dir(d):
        if abs(d.DotProduct(view_right)) > 0.99: return view_right
        if abs(d.DotProduct(view_up)) > 0.99:    return view_up
        return d

    if direction == "horizontal":
        sort_axis = _clean_dir(master_dir) if master_dir else view_right
        if abs(sort_axis.DotProduct(view_up)) > 0.5: # Inversão detectada
             sort_axis = view_right
             
        if sort_axis.DotProduct(view_right) < 0:
            sort_axis = sort_axis.Negate()

        # Eixo perpendicular para o offset da cota
        offset_axis = view_normal.CrossProduct(sort_axis).Normalize()
        if offset_axis.DotProduct(view_up) < 0:
            offset_axis = offset_axis.Negate()
    else:
        sort_axis = _clean_dir(master_dir) if master_dir else view_up
        if abs(sort_axis.DotProduct(view_right)) > 0.5: # Inversão detectada
             sort_axis = view_up

        if sort_axis.DotProduct(view_up) < 0:
            sort_axis = sort_axis.Negate()

        offset_axis = view_normal.CrossProduct(sort_axis).Normalize()
        if offset_axis.DotProduct(view_right) < 0:
            offset_axis = offset_axis.Negate()

    # 2. Coleta posições e referências
    valid_data = []
    for el in elements:
        pos = get_element_position_on_axis(el, sort_axis)
        ref = get_element_reference(el)
        if pos is not None and ref is not None:
            valid_data.append((pos, el, ref))

    if len(valid_data) < 2:
        return 0, base_offset_ft

    valid_data.sort(key=lambda x: x[0])

    # BUG 3 FIX: limiar de 10 mm em vez de ShortCurveTolerance (~0,008 mm)
    min_dist = 10.0 / MM_PER_FT

    filtered = []
    seen_refs = set()
    last_pos = None

    for pos, el, ref in valid_data:
        try:
            ref_stable = ref.ConvertToStableRepresentation(doc)
            if ref_stable in seen_refs: continue
            if last_pos is not None and abs(pos - last_pos) < min_dist: continue

            filtered.append((pos, el, ref))
            seen_refs.add(ref_stable)
            last_pos = pos
        except: pass

    if len(filtered) < 2:
        return 0, base_offset_ft

    # 3. Posição da linha de cota
    max_offset = None
    for _, el, _ in filtered:
        bb = el.get_BoundingBox(view)
        if bb:
            for corner in [bb.Min, bb.Max]:
                candidate = corner.DotProduct(offset_axis)
                if max_offset is None or candidate > max_offset:
                    max_offset = candidate

    if max_offset is None:
        max_offset = 0.0

    # Respeita o base_offset (para empilhar cadeias de categorias diferentes)
    final_offset_val = max(max_offset, base_offset_ft) + offset_ft

    # 4. Cria a cota
    ref_array = ReferenceArray()
    for _, _, ref in filtered:
        ref_array.Append(ref)

    def get_pt(s_val):
        return (sort_axis * s_val) + (offset_axis * final_offset_val) + (view_normal * view_plane_depth)

    p1 = get_pt(filtered[0][0])
    p2 = get_pt(filtered[-1][0])

    try:
        dim_line = Line.CreateBound(p1, p2)
        dim = None

        try:
            if dim_type:
                dim = DB.Dimension.Create(doc, view.Id, dim_line, ref_array, dim_type.Id)
            else:
                dim = DB.Dimension.Create(doc, view.Id, dim_line, ref_array)
        except:
            try:
                if dim_type:
                    dim = doc.Create.NewAlignedDimension(view, dim_line, ref_array, dim_type)
                else:
                    dim = doc.Create.NewAlignedDimension(view, dim_line, ref_array)
            except:
                try:
                    if dim_type:
                        dim = doc.Create.NewDimension(view, dim_line, ref_array, dim_type)
                    else:
                        dim = doc.Create.NewDimension(view, dim_line, ref_array)
                except Exception as e_final:
                    import traceback
                    print("Falha total na criação de cota: {}".format(traceback.format_exc()))
                    return 0, base_offset_ft

        if dim:
            return len(filtered), final_offset_val
        return 0, base_offset_ft

    except Exception as ex:
        print("Erro crítico ao criar cadeia: {}".format(ex))
        return 0, base_offset_ft


# ====== Janela WPF ======
class AutoCotasWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.ref_categories    = []
        self.target_categories = []
        self.dim_types         = []
        self._init_categories()
        self._init_dim_types()
        self._bind_events()
        self._update_info()

    def _init_categories(self):
        self.ref_categories    = [CategoryOption(n, b) for n, b in REF_CATEGORIES]
        self.target_categories = [CategoryOption(n, b) for n, b in TARGET_CATEGORIES]
        self.lb_RefCategories.ItemsSource    = self.ref_categories
        self.lb_TargetCategories.ItemsSource = self.target_categories

    def _init_dim_types(self):
        self.dim_types = get_dim_types()
        type_names = []
        for dt in self.dim_types:
            try:
                type_names.append(dt.Name if hasattr(dt, 'Name') else "Tipo {}".format(dt.Id.IntegerValue))
            except:
                type_names.append("Tipo {}".format(dt.Id.IntegerValue))

        if type_names:
            self.cb_DimType.ItemsSource  = type_names
            self.cb_DimType.SelectedIndex = 0

    def _bind_events(self):
        self.btn_Generate.Click       += self.generate_dimensions
        self.btn_Cancel.Click         += lambda s, a: self.Close()
        self.rb_ModeSelected.Checked  += self._update_info
        self.rb_ModeCategory.Checked  += self._update_info

    def _update_info(self, sender=None, args=None):
        if self.rb_ModeSelected.IsChecked:
            self.panel_Categories.Visibility = System.Windows.Visibility.Collapsed
            sel_count = len(uidoc.Selection.GetElementIds())
            if sel_count > 0:
                self.lbl_Info.Text = "{} elementos selecionados. Serão cotados todos juntos na direção escolhida.".format(sel_count)
            else:
                self.lbl_Info.Text = "Nenhum elemento selecionado. Feche a janela e selecione primeiro."
        else:
            self.panel_Categories.Visibility = System.Windows.Visibility.Visible
            self.lbl_Info.Text = (
                "Referências = base da cota. Alvos = o que será cotado. "
                "Cada categoria-alvo gera uma cadeia separada."
            )

    def generate_dimensions(self, sender, args):
        use_selected = self.rb_ModeSelected.IsChecked

        try:
            offset_mm = float(self.txt_Offset.Text)
        except:
            forms.alert("Offset inválido. Use um número (ex: 500).")
            return

        selected_dim_type = None
        if 0 <= self.cb_DimType.SelectedIndex < len(self.dim_types):
            selected_dim_type = self.dim_types[self.cb_DimType.SelectedIndex]

        do_horizontal = self.chk_DirHorizontal.IsChecked
        do_vertical   = self.chk_DirVertical.IsChecked

        if not do_horizontal and not do_vertical:
            forms.alert("Marque pelo menos uma direção.")
            return

        # --- Monta grupos de elementos ---
        # Cada grupo vira uma cadeia de cotas independente.
        element_groups = []

        if use_selected:
            sel_ids = uidoc.Selection.GetElementIds()
            if not sel_ids:
                forms.alert("Nenhum elemento selecionado.")
                return
            elements = [doc.GetElement(eid) for eid in sel_ids]
            elements = [e for e in elements if e is not None]
            if len(elements) < 2:
                forms.alert("São necessários pelo menos 2 elementos.\nSelecionados: {}".format(len(elements)))
                return
            element_groups.append(elements)

        else:
            ref_bics    = [c.bic for c in self.ref_categories    if c.is_checked]
            target_bics = [c.bic for c in self.target_categories if c.is_checked]

            if not ref_bics and not target_bics:
                forms.alert("Marque pelo menos uma categoria de referência ou alvo.")
                return

            # Coleta elementos de referência (âncoras comuns a todas as cadeias)
            ref_elements = []
            for bic in ref_bics:
                ref_elements.extend(list(
                    FilteredElementCollector(doc, active_view.Id)
                    .OfCategory(bic).WhereElementIsNotElementType()
                ))

            if target_bics:
                # MELHORIA 4: uma cadeia por categoria-alvo, cada uma ancorada nas referências
                for bic in target_bics:
                    target_els = list(
                        FilteredElementCollector(doc, active_view.Id)
                        .OfCategory(bic).WhereElementIsNotElementType()
                    )
                    group = ref_elements + target_els
                    if len(group) >= 2:
                        element_groups.append(group)
            else:
                if len(ref_elements) >= 2:
                    element_groups.append(ref_elements)

            if not element_groups:
                forms.alert(
                    "Nenhum grupo com elementos suficientes encontrado na vista.\n"
                    "Verifique se as categorias selecionadas têm elementos visíveis."
                )
                return

        self.Close()

        total_elements = 0
        total_chains   = 0

        with revit.Transaction("Auto-Cotas", doc):
            for direction in ([d for d in ["horizontal", "vertical"]
                               if (d == "horizontal" and do_horizontal) or
                                  (d == "vertical"   and do_vertical)]):
                # Empilha as cadeias: cada nova cadeia começa depois da anterior
                next_base_ft = 0.0
                for group in element_groups:
                    res, used_offset = create_individual_dimensions(
                        group, direction, offset_mm, selected_dim_type,
                        base_offset_ft=next_base_ft
                    )
                    if res > 0:
                        total_elements += res
                        total_chains   += 1
                        # Próxima cadeia começa depois desta + um gap extra
                        next_base_ft = used_offset + offset_mm / MM_PER_FT

        if total_chains > 0:
            forms.alert(
                "✅ {} cadeia(s) criada(s) abrangendo {} elemento(s)!".format(
                    total_chains, total_elements),
                title="Auto-Cotas"
            )
        else:
            forms.alert(
                "Nenhuma cota criada.\n\n"
                "Dicas:\n"
                "• Verifique se os elementos selecionados são paralelos entre si\n"
                "• Aumente o offset se a cota estiver oculta atrás do elemento\n"
                "• Garanta que as categorias possuem referências geométricas válidas",
                title="Auto-Cotas"
            )


# ====== Entry Point ======
if __name__ == "__main__":
    if doc:
        if not isinstance(active_view, (ViewPlan, ViewSection)):
            forms.alert(
                "Auto-Cotas funciona apenas em Plantas e Cortes.\n"
                "Abra uma Vista de Planta e execute novamente.",
                title="Auto-Cotas"
            )
        else:
            AutoCotasWindow("AutoCotasWindow.xaml").show(modal=True)
