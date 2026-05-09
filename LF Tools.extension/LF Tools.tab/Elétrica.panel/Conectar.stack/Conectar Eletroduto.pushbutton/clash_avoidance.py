# coding: utf-8
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

def get_obstacles_in_path(doc, pt_start, pt_end, margin=2.0):
    """
    Fase 1: Retorna elementos que estão na região da rota entre pt_start e pt_end.
    margin: Folga em pés (padrão ~60cm)
    """
    min_x = min(pt_start.X, pt_end.X) - margin
    min_y = min(pt_start.Y, pt_end.Y) - margin
    min_z = min(pt_start.Z, pt_end.Z) - margin
    
    max_x = max(pt_start.X, pt_end.X) + margin
    max_y = max(pt_start.Y, pt_end.Y) + margin
    max_z = max(pt_start.Z, pt_end.Z) + margin
    
    outline = Outline(XYZ(min_x, min_y, min_z), XYZ(max_x, max_y, max_z))
    bb_filter = BoundingBoxIntersectsFilter(outline)
    
    # Categorias de interesse para colisão
    cat_list = [
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_CableTray,
        BuiltInCategory.OST_Conduit
    ]
    
    cat_filter_list = [ElementCategoryFilter(c) for c in cat_list]
    or_cat_filter = LogicalOrFilter(cat_filter_list)
    
    collector = FilteredElementCollector(doc) \
        .WherePasses(or_cat_filter) \
        .WherePasses(bb_filter) \
        .WhereElementIsNotElementType()
        
    return list(collector)

def create_route_solid(pt_start, pt_end, diameter):
    """
    Cria um cilindro representando o eletroduto/tubo.
    """
    try:
        line = Line.CreateBound(pt_start, pt_end)
        # Cria um perfil circular
        # Isso exige transações ou rotinas complexas no Revit, 
        # para simulação rápida usamos SolidOptions ou extrusão.
        pass
    except:
        pass
    return None

def check_collision(doc, pt_start, pt_end, diameter=0.1):
    """
    Fase 2: Retorna o primeiro obstáculo que colide com a linha de A para B.
    """
    obstacles = get_obstacles_in_path(doc, pt_start, pt_end)
    if not obstacles:
        return None
        
    line = Line.CreateBound(pt_start, pt_end)
    curve_filter = ElementIntersectsCurveFilter(line)
    
    # Refina a busca
    colliding_elements = []
    for obs in obstacles:
        if curve_filter.PassesFilter(obs):
            colliding_elements.append(obs)
            
    return colliding_elements[0] if colliding_elements else None

def calculate_saddle_bypass(pt_start, pt_end, obstacle, clearance=0.5):
    """
    Fase 3: Calcula os pontos P1, P2, P3, P4 para desviar do obstáculo.
    """
    # Lógica paramétrica de desvio
    # Retorna a nova lista de pontos
    return [pt_start, pt_end]

