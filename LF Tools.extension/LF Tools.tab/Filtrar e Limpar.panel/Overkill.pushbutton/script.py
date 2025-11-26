# -*- coding: utf-8 -*-
import clr
import math
import re 
from System.Collections.Generic import List # Importação necessária para listas do Revit
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.DB import LocationCurve, LocationPoint, Line, Arc, FamilyInstance, TextNote

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
            # Assinatura: Posição (Coord) + Conteúdo do Texto
            pos = qxyz(el.Coord, step_mm_point)
            content = el.Text.strip() # Remove espaços extras
            return ("TEXT", pos, content)

        loc = el.Location

        # 2. CURVAS (Linhas, Dutos, Tubos)
        if isinstance(loc, LocationCurve):
            c = loc.Curve
            if isinstance(c, Line):
                p1 = qxyz(c.GetEndPoint(0), step_mm_curve)
                p2 = qxyz(c.GetEndPoint(1), step_mm_curve)
                return ("LINE", tuple(sorted((p1, p2)))) # Ordem não importa

            if isinstance(c, Arc):
                center = qxyz(c.Center, step_mm_curve)
                radius = quantize_feet_to_mm_int(c.Radius, step_mm_curve)
                return ("ARC", center, radius)

        # 3. PONTOS (Famílias pontuais)
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

# ====== Janela Principal ======
class OverkillWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.categories = []
        self._init_categories()
        self._bind_events()

    def _init_categories(self):
        # Lista de categorias suportadas
        cats = [
            ("Notas de Texto", BuiltInCategory.OST_TextNotes, True),
            ("Identificadores (Tags)", BuiltInCategory.OST_Tags, False),
            ("Conduítes", BuiltInCategory.OST_Conduit, True),
            ("Conexões de Conduíte", BuiltInCategory.OST_ConduitFitting, False),
            ("Eletrocalhas", BuiltInCategory.OST_CableTray, True),
            ("Conexões de Eletrocalha", BuiltInCategory.OST_CableTrayFitting, False),
            ("Dutos", BuiltInCategory.OST_DuctCurves, True),
            ("Conexões de Duto", BuiltInCategory.OST_DuctFitting, False),
            ("Tubulações", BuiltInCategory.OST_PipeCurves, True),
            ("Conexões de Tubulação", BuiltInCategory.OST_PipeFitting, False),
            ("Equipamentos Mecânicos", BuiltInCategory.OST_MechanicalEquipment, False),
            ("Dispositivos de Iluminação", BuiltInCategory.OST_LightingFixtures, False),
            ("Dispositivos Elétricos", BuiltInCategory.OST_ElectricalFixtures, False),
            ("Mobiliário", BuiltInCategory.OST_Furniture, False),
            ("Linhas de Detalhe", BuiltInCategory.OST_Lines, True)
        ]
        
        self.categories = [CategoryOption(name, bic, chk) for name, bic, chk in cats]
        self.categories_lb.ItemsSource = self.categories

    def _bind_events(self):
        self.run_btn.Click += self.run_analysis
        self.cancel_btn.Click += lambda s, a: self.Close()
        self.cat_all_btn.Click += self.select_all_cats
        self.cat_none_btn.Click += self.select_none_cats

    def select_all_cats(self, sender, args):
        for c in self.categories: c.is_checked = True
        self.categories_lb.Items.Refresh()

    def select_none_cats(self, sender, args):
        for c in self.categories: c.is_checked = False
        self.categories_lb.Items.Refresh()

    def run_analysis(self, sender, args):
        # 1. Configurações da UI
        selected_bics = [c.bic for c in self.categories if c.is_checked]
        if not selected_bics:
            forms.alert("Selecione pelo menos uma categoria.")
            return

        is_active_view = self.scope_view_rb.IsChecked
        delete_mode = self.mode_delete_rb.IsChecked
        
        try:
            tol_curve = float(self.tol_curve_tb.Text)
            tol_point = float(self.tol_point_tb.Text)
        except:
            forms.alert("Tolerâncias inválidas. Use números (ex: 1.0).")
            return

        self.Close()

        # 2. Coleta de Elementos
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

        # 3. Processamento (Core Logic)
        duplicates = []
        signatures_seen = {} # {signature: element_id}
        count_analyzed = 0

        for el in elements_to_check:
            try:
                cat_id = el.Category.Id.IntegerValue if el.Category else -1
                type_id = el.GetTypeId().IntegerValue if el.GetTypeId() else -1
                lvl_id = safe_level_id(el)
                
                geo_sig = get_location_signature(el, tol_curve, tol_point)
                
                if geo_sig:
                    full_sig = (cat_id, type_id, lvl_id, geo_sig)
                    
                    if full_sig in signatures_seen:
                        existing_id = signatures_seen[full_sig]
                        # Mantém o ID menor (mais antigo)
                        if el.Id.IntegerValue > existing_id.IntegerValue:
                            duplicates.append(el.Id)
                        else:
                            duplicates.append(existing_id)
                            signatures_seen[full_sig] = el.Id
                    else:
                        signatures_seen[full_sig] = el.Id
                    
                count_analyzed += 1
            except:
                pass

        # 4. Resultado
        if not duplicates:
            forms.alert("Limpeza concluída!\nNenhum duplicado encontrado em {} elementos.".format(count_analyzed))
            return

        # Converter lista Python para List[ElementId] do .NET
        ids_collection = List[ElementId](duplicates)

        if delete_mode:
            # MODO DELETAR
            msg = "Foram encontrados {} duplicados em {} analisados.\n\nDeseja deletá-los agora?".format(len(duplicates), count_analyzed)
            if forms.alert(msg, yes=True, no=True):
                with revit.Transaction("Overkill - Deletar", doc):
                    doc.Delete(ids_collection) # CORRIGIDO AQUI
                forms.alert("Sucesso! {} elementos removidos.".format(len(duplicates)))
        else:
            # MODO SELEÇÃO
            uidoc.Selection.SetElementIds(ids_collection) # CORRIGIDO AQUI
            forms.alert("Análise concluída!\n\n{} elementos duplicados foram SELECIONADOS na tela.\nVerifique as propriedades para decidir o que fazer.".format(len(duplicates)))

if __name__ == "__main__":
    if doc:
        OverkillWindow("OverkillWindow.xaml").show(modal=True)