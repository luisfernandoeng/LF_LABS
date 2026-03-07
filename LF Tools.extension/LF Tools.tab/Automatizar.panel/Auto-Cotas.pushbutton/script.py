# -*- coding: utf-8 -*-
"""
Auto-Cotas - Dimensionamento Automático por Eixo Central
=========================================================
LF Tools - pyRevit Extension

Dois modos:
  - Cotar elementos selecionados (misto)
  - Cotar por categorias: Referência (base) + Alvos (o que cotar)
    Ex: Paredes (ref) → Conduítes, Dispositivos (alvos)
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
    DimensionType, ViewPlan, ViewSection,
    LocationCurve, LocationPoint,
    Grid as RevitGrid
)

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
    ("Paredes", BuiltInCategory.OST_Walls),
    ("Pisos", BuiltInCategory.OST_Floors),
    ("Eixos (Grids)", BuiltInCategory.OST_Grids),
    ("Pilares Estruturais", BuiltInCategory.OST_StructuralColumns),
    ("Colunas Arquitetônicas", BuiltInCategory.OST_Columns),
]

# Alvos (o que cotar em relação à referência)
TARGET_CATEGORIES = [
    ("Conduítes", BuiltInCategory.OST_Conduit),
    ("Eletrocalhas", BuiltInCategory.OST_CableTray),
    ("Tubulações", BuiltInCategory.OST_PipeCurves),
    ("Dutos (HVAC)", BuiltInCategory.OST_DuctCurves),
    ("Dispositivos Elétricos", BuiltInCategory.OST_ElectricalFixtures),
    ("Dispositivos de Iluminação", BuiltInCategory.OST_LightingDevices),
    ("Luminárias", BuiltInCategory.OST_LightingFixtures),
    ("Dispositivos de Comunicação", BuiltInCategory.OST_CommunicationDevices),
    ("Equipamentos Mecânicos", BuiltInCategory.OST_MechanicalEquipment),
    ("Conexões de Conduíte", BuiltInCategory.OST_ConduitFitting),
    ("Conexões de Tubulação", BuiltInCategory.OST_PipeFitting),
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
    """Obtém uma Reference válida para criar dimensão."""
    if isinstance(el, RevitGrid):
        try:
            return Reference(el)
        except:
            pass
    
    try:
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.View = active_view
        
        geom = el.get_Geometry(opt)
        if geom:
            for geom_obj in geom:
                if hasattr(geom_obj, 'Reference') and geom_obj.Reference:
                    return geom_obj.Reference
                if hasattr(geom_obj, 'GetInstanceGeometry'):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        for sub in inst_geom:
                            if hasattr(sub, 'Reference') and sub.Reference:
                                return sub.Reference
    except:
        pass
    
    try:
        return Reference(el)
    except:
        pass
    
    return None


def project_point_on_axis(point, axis_direction):
    """Projeta um ponto no eixo dado."""
    return point.X * axis_direction.X + point.Y * axis_direction.Y + point.Z * axis_direction.Z


def get_element_position_on_axis(el, axis_direction):
    """Obtém posição do ponto médio do elemento no eixo."""
    curve = get_element_curve(el)
    if curve:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        mid = XYZ((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, (p0.Z + p1.Z) / 2.0)
        return project_point_on_axis(mid, axis_direction)
    
    loc = el.Location
    if isinstance(loc, LocationPoint):
        return project_point_on_axis(loc.Point, axis_direction)
    
    # Fallback: bounding box center
    try:
        bb = el.get_BoundingBox(active_view)
        if bb:
            center = XYZ((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0, (bb.Min.Z + bb.Max.Z) / 2.0)
            return project_point_on_axis(center, axis_direction)
    except:
        pass
    
    return None


def get_dim_types():
    """Coleta DimensionTypes lineares."""
    col = FilteredElementCollector(doc).OfClass(DimensionType)
    types = []
    for dt in col:
        try:
            if dt.StyleType.ToString() == "Linear":
                types.append(dt)
        except:
            types.append(dt)
    return types


def create_individual_dimensions(elements, direction, offset_mm, dim_type=None):
    """
    Cria cotas individuais entre pares adjacentes de elementos.
    Junta todos os elementos (referência + alvos), ordena e cota.
    """
    view = active_view
    offset_ft = offset_mm / MM_PER_FT
    
    if direction == "horizontal":
        sort_axis = XYZ(1, 0, 0)
        offset_dir = XYZ(0, 1, 0)
    else:
        sort_axis = XYZ(0, 1, 0)
        offset_dir = XYZ(1, 0, 0)
    
    # Filtra elementos válidos
    valid_elements = []
    for el in elements:
        pos = get_element_position_on_axis(el, sort_axis)
        ref = get_element_reference(el)
        if pos is not None and ref is not None:
            valid_elements.append((pos, el, ref))
    
    if len(valid_elements) < 2:
        return 0
    
    valid_elements.sort(key=lambda x: x[0])
    
    # Calcula offset da linha de cota
    max_offset_point = None
    for _, el, _ in valid_elements:
        bb = el.get_BoundingBox(view)
        if bb:
            if direction == "horizontal":
                candidate = bb.Max.Y
            else:
                candidate = bb.Max.X
            if max_offset_point is None or candidate > max_offset_point:
                max_offset_point = candidate
    
    if max_offset_point is None:
        max_offset_point = 0
    
    dim_offset = max_offset_point + offset_ft
    
    # Tolerância mínima do Revit
    try:
        min_length = app.ShortCurveTolerance * 2.0
    except:
        min_length = 0.01
    
    count = 0
    
    for i in range(len(valid_elements) - 1):
        pos_a, el_a, ref_a = valid_elements[i]
        pos_b, el_b, ref_b = valid_elements[i + 1]
        
        # Pula pares muito próximos
        dist_ft = abs(pos_b - pos_a)
        if dist_ft < min_length:
            continue
        
        ref_array = ReferenceArray()
        ref_array.Append(ref_a)
        ref_array.Append(ref_b)
        
        if direction == "horizontal":
            p1 = XYZ(pos_a, dim_offset, 0)
            p2 = XYZ(pos_b, dim_offset, 0)
        else:
            p1 = XYZ(dim_offset, pos_a, 0)
            p2 = XYZ(dim_offset, pos_b, 0)
        
        try:
            dim_line = Line.CreateBound(p1, p2)
            if dim_type:
                dim = doc.Create.NewDimension(view, dim_line, ref_array, dim_type)
            else:
                dim = doc.Create.NewDimension(view, dim_line, ref_array)
            if dim:
                count += 1
        except Exception as ex:
            pass  # Silenciosamente pula falhas individuais
    
    return count


# ====== Janela WPF ======
class AutoCotasWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.ref_categories = []
        self.target_categories = []
        self.dim_types = []
        self._init_categories()
        self._init_dim_types()
        self._bind_events()
        self._update_info()
    
    def _init_categories(self):
        """Popula as duas listas de categorias."""
        self.ref_categories = [CategoryOption(n, b, False) for n, b in REF_CATEGORIES]
        self.target_categories = [CategoryOption(n, b, False) for n, b in TARGET_CATEGORIES]
        self.lb_RefCategories.ItemsSource = self.ref_categories
        self.lb_TargetCategories.ItemsSource = self.target_categories
    
    def _init_dim_types(self):
        """Popula ComboBox de tipos de cota."""
        self.dim_types = get_dim_types()
        type_names = []
        for dt in self.dim_types:
            try:
                type_names.append(dt.Name if hasattr(dt, 'Name') else "Tipo {}".format(dt.Id.IntegerValue))
            except:
                type_names.append("Tipo {}".format(dt.Id.IntegerValue))
        
        if type_names:
            self.cb_DimType.ItemsSource = type_names
            self.cb_DimType.SelectedIndex = 0
    
    def _bind_events(self):
        self.btn_Generate.Click += self.generate_dimensions
        self.btn_Cancel.Click += lambda s, a: self.Close()
        self.rb_ModeSelected.Checked += self._update_info
        self.rb_ModeCategory.Checked += self._update_info
    
    def _update_info(self, sender=None, args=None):
        """Atualiza texto informativo e visibilidade do painel."""
        if self.rb_ModeSelected.IsChecked:
            self.panel_Categories.Visibility = System.Windows.Visibility.Collapsed
            sel_count = len(uidoc.Selection.GetElementIds())
            if sel_count > 0:
                self.lbl_Info.Text = "{} elementos selecionados. Serão cotados todos juntos na direção escolhida.".format(sel_count)
            else:
                self.lbl_Info.Text = "Nenhum elemento selecionado. Feche a janela e selecione primeiro."
        else:
            self.panel_Categories.Visibility = System.Windows.Visibility.Visible
            self.lbl_Info.Text = "Marque as categorias de REFERÊNCIA (base) e ALVOS (o que cotar). Todos serão cotados juntos na vista."
    
    def generate_dimensions(self, sender, args):
        """Ação principal."""
        use_selected = self.rb_ModeSelected.IsChecked
        
        try:
            offset_mm = float(self.txt_Offset.Text)
        except:
            forms.alert("Offset inválido. Use um número (ex: 500).")
            return
        
        # Tipo de cota
        selected_dim_type = None
        if self.cb_DimType.SelectedIndex >= 0 and self.cb_DimType.SelectedIndex < len(self.dim_types):
            selected_dim_type = self.dim_types[self.cb_DimType.SelectedIndex]
        
        # Direção
        do_horizontal = self.chk_DirHorizontal.IsChecked
        do_vertical = self.chk_DirVertical.IsChecked
        
        if not do_horizontal and not do_vertical:
            forms.alert("Marque pelo menos uma direção.")
            return
        
        # Coletar elementos
        elements = []
        
        if use_selected:
            sel_ids = uidoc.Selection.GetElementIds()
            if not sel_ids:
                forms.alert("Nenhum elemento selecionado.")
                return
            elements = [doc.GetElement(eid) for eid in sel_ids]
            elements = [e for e in elements if e is not None]
        else:
            # Por categorias: junta referências + alvos
            ref_bics = [c.bic for c in self.ref_categories if c.is_checked]
            target_bics = [c.bic for c in self.target_categories if c.is_checked]
            
            if not ref_bics and not target_bics:
                forms.alert("Marque pelo menos uma categoria de referência ou alvo.")
                return
            
            all_bics = ref_bics + target_bics
            for bic in all_bics:
                col = FilteredElementCollector(doc, active_view.Id)
                col.OfCategory(bic).WhereElementIsNotElementType()
                elements.extend(list(col))
        
        if len(elements) < 2:
            forms.alert("São necessários pelo menos 2 elementos.\nEncontrados: {}".format(len(elements)))
            return
        
        self.Close()
        
        total_created = 0
        
        with revit.Transaction("Auto-Cotas", doc):
            if do_horizontal:
                total_created += create_individual_dimensions(
                    elements, "horizontal", offset_mm, selected_dim_type
                )
            if do_vertical:
                total_created += create_individual_dimensions(
                    elements, "vertical", offset_mm, selected_dim_type
                )
        
        if total_created > 0:
            forms.alert("✅ {} cotas criadas!".format(total_created), title="Auto-Cotas")
        else:
            forms.alert(
                "Nenhuma cota criada.\n\n"
                "Dicas:\n"
                "• Eixos (Grids) têm referências mais simples\n"
                "• Verifique se a vista é uma planta\n"
                "• Aumente o offset se elementos estão sobrepostos",
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
