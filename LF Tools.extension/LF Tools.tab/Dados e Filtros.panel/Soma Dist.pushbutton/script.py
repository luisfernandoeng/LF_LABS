# -*- coding: utf-8 -*-

"""
Script para contar comprimento total de eletrodutos, dutos, tubulacoes
e conexoes no Revit (incluindo curvas/cotovelos).
"""

__title__ = 'Contar Comprimento'
__author__ = 'Luis Fernando'

# Importar bibliotecas do Revit
from Autodesk.Revit.DB import (
    UnitUtils,
    UnitTypeId,
    StorageType,
    BuiltInCategory,
    FilteredElementCollector,
    Curve,
    BuiltInParameter,
    LocationCurve,
    Options,
    GeometryInstance,
    ViewDetailLevel
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

# Importar bibliotecas do pyRevit
from pyrevit import revit, script

import math


# Funcao auxiliar para obter o comprimento
def get_length(element):
    """
    Obtém o comprimento do elemento em pés (unidades internas do Revit).
    Funciona para eletrodutos, tubos, dutos e suas conexões.
    """
    length = 0.0
    
    # PRIORIDADE 1: Tentar CURVE_ELEM_LENGTH (Built-in parameter padrão)
    try:
        length_param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if length_param and length_param.HasValue:
            val = length_param.AsDouble()
            if val > 0.0:
                return val
    except:
        pass
    
    # PRIORIDADE 2: Lista de nomes de parametros personalizados comuns
    # IMPORTANTE: Não incluir "Conduit Length" pois para fittings retorna valor incorreto
    param_names = [
        "Length",
        "Centerline Length",
        "Comprimento",
        "Comprimento do eletroduto",
        "Comprimento da linha de centro"
    ]
    
    for name in param_names:
        param = element.LookupParameter(name)
        if param and param.StorageType == StorageType.Double and param.HasValue:
            val = param.AsDouble()
            if val > 0.0:
                return val
    
    # PRIORIDADE 3: LocationCurve (para elementos lineares)
    try:
        location = element.Location
        if isinstance(location, LocationCurve):
            curve = location.Curve
            if curve:
                return curve.Length
    except:
        pass
    
    # PRIORIDADE 4: FITTINGS (Conexões) - tratamento especial
    try:
        cat = element.Category
        if cat:
            cat_id = cat.Id.IntegerValue
            
            is_fitting = cat_id in [
                int(BuiltInCategory.OST_PipeFitting),
                int(BuiltInCategory.OST_ConduitFitting),
                int(BuiltInCategory.OST_DuctFitting)
            ]
            
            if is_fitting:
                # A) COTOVELOS: Calcular via raio e ângulo
                radius = None
                radius_param_names = [
                    "Radius", 
                    "Raio", 
                    "Raio de curvatura", 
                    "Bend Radius",
                    "Bend Radius Label"
                ]
                
                for radius_name in radius_param_names:
                    radius_param = element.LookupParameter(radius_name)
                    if radius_param and radius_param.HasValue:
                        radius = radius_param.AsDouble()
                        if radius > 0:
                            break
                
                # Tentar vários nomes possíveis para ângulo
                angle_rad = None
                angle_param_names = ["Angle", "Ângulo", "Angulo"]
                
                for angle_name in angle_param_names:
                    angle_param = element.LookupParameter(angle_name)
                    if angle_param and angle_param.HasValue:
                        angle_rad = angle_param.AsDouble()
                        if angle_rad > 0:
                            break
                
                # Se encontrou raio e ângulo, calcular comprimento do arco
                if radius and angle_rad and radius > 0 and angle_rad > 0:
                    return radius * angle_rad
                
                # B) TEEs, CROSSES: Usar "Center to End"
                cte = None
                cte_param_names = ["Center to End", "Centro até Extremidade", "Centro para Extremidade"]
                
                for cte_name in cte_param_names:
                    cte_param = element.LookupParameter(cte_name)
                    if cte_param and cte_param.HasValue:
                        cte = cte_param.AsDouble()
                        if cte > 0:
                            break
                
                if cte and cte > 0:
                    # Tentar contar conectores para ajustar cálculo
                    try:
                        if hasattr(element, 'ConnectorManager'):
                            connectors = element.ConnectorManager.Connectors
                            num_connectors = sum(1 for _ in connectors)
                            
                            # TEE = 3 conectores, Cross = 4 conectores
                            if num_connectors == 3:
                                return cte * 2  # Aproximação para TEE
                            elif num_connectors == 4:
                                return cte * 3  # Aproximação para Cross
                    except:
                        pass
                    
                    return cte
                
                # C) GEOMETRIA: Buscar curvas na geometria do fitting
                try:
                    options = Options()
                    options.ComputeReferences = False
                    options.DetailLevel = ViewDetailLevel.Coarse
                    
                    geom_elem = element.get_Geometry(options)
                    if geom_elem:
                        total_curve_length = 0.0
                        
                        for geom_obj in geom_elem:
                            # Verificar GeometryInstance (comum em famílias)
                            if isinstance(geom_obj, GeometryInstance):
                                inst_geom = geom_obj.GetInstanceGeometry()
                                if inst_geom:
                                    for inst_obj in inst_geom:
                                        if isinstance(inst_obj, Curve):
                                            total_curve_length += inst_obj.Length
                            elif isinstance(geom_obj, Curve):
                                total_curve_length += geom_obj.Length
                        
                        if total_curve_length > 0:
                            return total_curve_length
                except:
                    pass
    except:
        pass
    
    return length


# Main script
if __name__ == '__main__':
    doc = revit.doc
    uidoc = revit.uidoc

    # Categorias permitidas
    allowed_categories = [
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_ConduitFitting,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_PipeFitting
    ]

    selected_elements = []

    try:
        # Selecao multipla
        selection = uidoc.Selection.PickObjects(
            ObjectType.Element,
            'Selecione os eletrodutos, tubos, dutos, curvas e conexoes.'
        )
        for element_id in selection:
            element = doc.GetElement(element_id)
            selected_elements.append(element)

    except Exception as e:
        TaskDialog.Show("Cancelado", "A selecao foi cancelada.")
        script.exit()

    if not selected_elements:
        TaskDialog.Show("Aviso", "Nenhum elemento foi selecionado.")
        script.exit()

    total_length_feet = 0.0
    valid_elements = 0
    elements_by_type = {}

    for element in selected_elements:
        if element.Category and element.Category.Id.IntegerValue in [int(cat) for cat in allowed_categories]:
            length = get_length(element)
            if length > 0.0:
                total_length_feet += length
                valid_elements += 1
                
                # Agrupar por tipo
                type_name = element.Category.Name
                if type_name not in elements_by_type:
                    elements_by_type[type_name] = {"count": 0, "length": 0.0}
                elements_by_type[type_name]["count"] += 1
                elements_by_type[type_name]["length"] += length

    # Converter pes -> metros
    total_length_meters = UnitUtils.ConvertFromInternalUnits(
        total_length_feet, UnitTypeId.Meters
    )

    # Montar mensagem limpa e profissional
    message = "RELATORIO DE COMPRIMENTO\n"
    message += "=" * 50 + "\n\n"
    message += "Total de elementos selecionados: {}\n".format(len(selected_elements))
    message += "Elementos validos: {}\n\n".format(valid_elements)
    
    if elements_by_type:
        message += "DETALHAMENTO POR CATEGORIA:\n"
        message += "-" * 50 + "\n"
        for type_name, data in sorted(elements_by_type.items()):
            length_m = UnitUtils.ConvertFromInternalUnits(data["length"], UnitTypeId.Meters)
            message += "{:<30} {:>3} un.  {:>8.2f} m\n".format(
                type_name + ":", 
                data["count"], 
                length_m
            )
        
        message += "-" * 50 + "\n"
        message += "{:<30} {:>3} un.  {:>8.2f} m\n".format(
            "TOTAL:",
            valid_elements,
            total_length_meters
        )
    else:
        message += "Nenhum elemento com comprimento foi encontrado."

    TaskDialog.Show("Resultado da Contagem", message)