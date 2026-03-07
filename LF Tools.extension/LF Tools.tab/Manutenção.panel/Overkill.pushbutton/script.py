# -*- coding: utf-8 -*-
import clr
import math
import re 
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.DB import LocationCurve, LocationPoint, Line, Arc, FamilyInstance, TextNote
import System

# pyRevit
from pyrevit import revit, forms, script

# ====== Contexto ======
doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# ====== Classes de Apoio ======
class CategoryOption(object):
    def __init__(self, name, bic, checked=False):
        self.name = name
        self.bic = bic
        self.is_checked = checked

class DuplicateItem(object):
    """Representa UM elemento duplicado para o DataGrid."""
    def __init__(self, el_id, cat_name, description):
        self.ElementId = el_id
        self.ElementIdStr = str(el_id.IntegerValue)
        self.CategoryName = cat_name
        self.Description = description
        self._is_selected = True

    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value

# ====== Funções Geométricas ======
MM_PER_FT = 304.8

def quantize_feet_to_mm_int(val_ft, step_mm):
    return int(round((val_ft * MM_PER_FT) / float(step_mm)))

def qxyz(pt, step_mm):
    return (
        quantize_feet_to_mm_int(pt.X, step_mm),
        quantize_feet_to_mm_int(pt.Y, step_mm),
        quantize_feet_to_mm_int(pt.Z, step_mm)
    )

def safe_level_id(el):
    try:
        lid = getattr(el, 'LevelId', None)
        if lid and lid.IntegerValue > 0: return lid.IntegerValue
    except: pass
    try:
        p = el.LookupParameter("Level") or el.LookupParameter("Nível") or el.LookupParameter("Reference Level")
        if p and p.AsElementId().IntegerValue > 0: return p.AsElementId().IntegerValue
    except: pass
    return -1

def get_location_signature(el, step_mm_curve=1.0, step_mm_point=2.0):
    """Gera uma assinatura única baseada na geometria e tipo do elemento."""
    try:
        # 1. NOTAS DE TEXTO
        if isinstance(el, TextNote):
            pos = qxyz(el.Coord, step_mm_point)
            content = el.Text.strip()
            return ("TEXT", pos, content)

        loc = el.Location

        # 2. CURVAS (Linhas, Dutos, Tubos)
        if isinstance(loc, LocationCurve):
            c = loc.Curve
            if isinstance(c, Line):
                p1 = qxyz(c.GetEndPoint(0), step_mm_curve)
                p2 = qxyz(c.GetEndPoint(1), step_mm_curve)
                return ("LINE", tuple(sorted((p1, p2))))

            if isinstance(c, Arc):
                center = qxyz(c.Center, step_mm_curve)
                radius = quantize_feet_to_mm_int(c.Radius, step_mm_curve)
                return ("ARC", center, radius)

        # 3. PONTOS (Famílias pontuais e Anotações Genéricas)
        if isinstance(el, FamilyInstance) and isinstance(loc, LocationPoint):
            return ("POINT", qxyz(loc.Point, step_mm_point))

        # 4. Fallback (Bounding Box)
        bb = el.get_BoundingBox(None)
        if bb:
            center = (bb.Min + bb.Max) * 0.5
            return ("BBOX", qxyz(center, step_mm_point))

    except Exception:
        pass
    return None

def get_element_description(el):
    """Retorna uma descrição legível do elemento para exibição."""
    try:
        # TextNote → mostra o texto
        if isinstance(el, TextNote):
            text = el.Text.strip()
            if len(text) > 70:
                text = text[:67] + "..."
            return text

        # FamilyInstance → "Família: Tipo"
        if isinstance(el, FamilyInstance):
            sym = el.Symbol
            fam_name = sym.Family.Name if sym and sym.Family else "?"
            type_name = sym.get_Parameter(
                clr.GetClrType(type(el)).Assembly.GetType(
                    "Autodesk.Revit.DB.BuiltInParameter"
                ).GetField("SYMBOL_NAME_PARAM").GetValue(None)
            )
            # Simple fallback
            type_name_str = ""
            try:
                type_name_str = el.Name
            except:
                type_name_str = "?"
            return "{}: {}".format(fam_name, type_name_str)
        
        # Curvas → tipo + comprimento
        loc = el.Location
        if isinstance(loc, LocationCurve):
            length_mm = loc.Curve.Length * MM_PER_FT
            type_name = ""
            try:
                type_name = el.Name
            except:
                type_name = "Elemento"
            return "{} ({:.0f} mm)".format(type_name, length_mm)

        # Fallback
        try:
            return el.Name
        except:
            return "Elemento {}".format(el.Id.IntegerValue)
    except:
        return "Elemento {}".format(el.Id.IntegerValue)

# ====== Janela Principal ======
class OverkillWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.categories = []
        self.duplicate_items = []
        self._init_categories()
        self._bind_events()

    def _init_categories(self):
        cats = [
            ("Notas de Texto", BuiltInCategory.OST_TextNotes, False),
            ("Anotação Genérica", BuiltInCategory.OST_GenericAnnotation, False),
            ("Identificadores (Tags)", BuiltInCategory.OST_Tags, False),
            ("Conduítes", BuiltInCategory.OST_Conduit, False),
            ("Conexões de Conduíte", BuiltInCategory.OST_ConduitFitting, False),
            ("Eletrocalhas", BuiltInCategory.OST_CableTray, False),
            ("Conexões de Eletrocalha", BuiltInCategory.OST_CableTrayFitting, False),
            ("Dutos", BuiltInCategory.OST_DuctCurves, False),
            ("Conexões de Duto", BuiltInCategory.OST_DuctFitting, False),
            ("Tubulações", BuiltInCategory.OST_PipeCurves, False),
            ("Conexões de Tubulação", BuiltInCategory.OST_PipeFitting, False),
            ("Equipamentos Mecânicos", BuiltInCategory.OST_MechanicalEquipment, False),
            ("Luminárias", BuiltInCategory.OST_LightingFixtures, False),
            ("Dispositivos de Iluminação", BuiltInCategory.OST_LightingDevices, False),
            ("Dispositivos Elétricos", BuiltInCategory.OST_ElectricalFixtures, False),
            ("Mobiliário", BuiltInCategory.OST_Furniture, False),
            ("Linhas de Detalhe", BuiltInCategory.OST_Lines, False)
        ]
        
        self.categories = [CategoryOption(name, bic, chk) for name, bic, chk in cats]
        self.categories_lb.ItemsSource = self.categories

    def _bind_events(self):
        self.run_btn.Click += self.run_analysis
        self.cancel_btn.Click += lambda s, a: self.Close()
        self.cat_all_btn.Click += self.select_all_cats
        self.cat_none_btn.Click += self.select_none_cats
        self.confirm_btn.Click += self.execute_deletion
        self.res_all_btn.Click += self.select_all_results
        self.res_none_btn.Click += self.select_none_results
        self.back_btn.Click += self.go_back

    def select_all_cats(self, sender, args):
        for c in self.categories: c.is_checked = True
        self.categories_lb.Items.Refresh()

    def select_none_cats(self, sender, args):
        for c in self.categories: c.is_checked = False
        self.categories_lb.Items.Refresh()

    def select_all_results(self, sender, args):
        for item in self.duplicate_items: item._is_selected = True
        self.results_dg.Items.Refresh()
    
    def select_none_results(self, sender, args):
        for item in self.duplicate_items: item._is_selected = False
        self.results_dg.Items.Refresh()

    def go_back(self, sender, args):
        """Volta para a fase de configuração."""
        self.config_grid.Visibility = System.Windows.Visibility.Visible
        self.results_grid.Visibility = System.Windows.Visibility.Collapsed
        self.run_btn.Visibility = System.Windows.Visibility.Visible
        self.confirm_btn.Visibility = System.Windows.Visibility.Collapsed

    def run_analysis(self, sender, args):
        selected_bics = [c.bic for c in self.categories if c.is_checked]
        if not selected_bics:
            forms.alert("Selecione pelo menos uma categoria.")
            return

        is_active_view = self.scope_view_rb.IsChecked
        
        try:
            tol_curve = float(self.tol_curve_tb.Text)
            tol_point = float(self.tol_point_tb.Text)
        except:
            forms.alert("Tolerâncias inválidas. Use números (ex: 1.0).")
            return

        self.duplicate_items = []

        # Coleta de Elementos
        elements_to_check = []
        for bic in selected_bics:
            if is_active_view:
                col = FilteredElementCollector(doc, doc.ActiveView.Id)
            else:
                col = FilteredElementCollector(doc)
            
            col.OfCategory(bic).WhereElementIsNotElementType()
            elements_to_check.extend(list(col))

        if not elements_to_check:
            forms.alert("Nenhum elemento encontrado para análise.")
            return

        # Processamento
        signatures_seen = {}  # {signature: element_id}
        count_analyzed = 0

        for el in elements_to_check:
            try:
                cat_id = el.Category.Id.IntegerValue if el.Category else -1
                cat_name = el.Category.Name if el.Category else "Desconhecido"
                type_id = el.GetTypeId().IntegerValue if el.GetTypeId() else -1
                lvl_id = safe_level_id(el)
                
                geo_sig = get_location_signature(el, tol_curve, tol_point)
                
                if geo_sig:
                    full_sig = (cat_id, type_id, lvl_id, geo_sig)
                    
                    if full_sig in signatures_seen:
                        existing_id = signatures_seen[full_sig]
                        # Mantém o ID menor (mais antigo)
                        if el.Id.IntegerValue > existing_id.IntegerValue:
                            dup_id = el.Id
                        else:
                            dup_id = existing_id
                            signatures_seen[full_sig] = el.Id
                        
                        # Gera descrição do elemento duplicado
                        dup_el = doc.GetElement(dup_id)
                        desc = get_element_description(dup_el) if dup_el else "?"
                        self.duplicate_items.append(DuplicateItem(dup_id, cat_name, desc))
                    else:
                        signatures_seen[full_sig] = el.Id
                    
                count_analyzed += 1
            except:
                pass

        if not self.duplicate_items:
            forms.alert("Nenhum duplicado encontrado em {} elementos.".format(count_analyzed))
            return

        # Popula o DataGrid com itens individuais
        self.results_dg.ItemsSource = self.duplicate_items
        
        # Contabiliza por categoria para o resumo
        cat_counts = {}
        for item in self.duplicate_items:
            cat_counts[item.CategoryName] = cat_counts.get(item.CategoryName, 0) + 1
        
        summary_parts = ["{}: {}".format(cat, count) for cat, count in cat_counts.items()]
        summary_text = "Encontrados {} duplicados ({})".format(
            len(self.duplicate_items), " | ".join(summary_parts)
        )
        self.summary_tb.Text = summary_text

        # Troca para fase de resultados
        self.config_grid.Visibility = System.Windows.Visibility.Collapsed
        self.results_grid.Visibility = System.Windows.Visibility.Visible
        self.run_btn.Visibility = System.Windows.Visibility.Collapsed
        self.confirm_btn.Visibility = System.Windows.Visibility.Visible

    def execute_deletion(self, sender, args):
        """Executa exclusão dos itens marcados."""
        items_to_process = [item for item in self.duplicate_items if item.IsSelected]
                
        if not items_to_process:
            forms.alert("Nenhum elemento selecionado para exclusão.")
            return
            
        delete_mode = self.mode_delete_rb.IsChecked
        ids_collection = List[ElementId]([item.ElementId for item in items_to_process])
        
        self.Close()

        if delete_mode:
            with revit.Transaction("Overkill - Deletar", doc):
                for el_id in ids_collection:
                    try:
                        doc.Delete(el_id)
                    except:
                        pass
            forms.alert("Sucesso! {} elementos removidos.".format(len(items_to_process)))
        else:
            uidoc.Selection.SetElementIds(ids_collection)
            forms.alert("{} elementos duplicados foram SELECIONADOS.".format(len(items_to_process)))

if __name__ == "__main__":
    if doc:
        OverkillWindow("OverkillWindow.xaml").show(modal=True)