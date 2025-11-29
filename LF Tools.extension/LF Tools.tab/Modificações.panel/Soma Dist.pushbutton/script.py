# -*- coding: utf-8 -*-

"""
Script para contar comprimento total de eletrodutos, dutos, tubulacoes
e conexoes no Revit (incluindo curvas/cotovelos).
"""

__title__ = 'Contar Comprimento'
__author__ = 'Luis - Engenheiro Eletricista'

# Importar bibliotecas do Revit
from Autodesk.Revit.DB import (
    UnitUtils,
    UnitTypeId,
    StorageType,
    BuiltInCategory,
    FilteredElementCollector,
    Curve
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

# Importar bibliotecas do pyRevit
from pyrevit import revit, script

# Funcao auxiliar para obter o comprimento
def get_length(element):
    """Tenta obter o comprimento do elemento em pes (unidades internas do Revit)."""
    length = 0.0
    
    # Lista de nomes de parametros comuns para comprimento
    param_names = [
        "Length",
        "Conduit Length", 
        "Centerline Length",
        "Comprimento",
        "Comprimento do eletroduto",
        "Center to End"
    ]
    
    # Tentar obter comprimento via parametros
    for name in param_names:
        param = element.LookupParameter(name)
        if param and param.StorageType == StorageType.Double:
            val = param.AsDouble()
            if val > 0.0:
                return val
    
    # Se nao encontrou via parametro, tentar pela geometria (Location Curve)
    try:
        location = element.Location
        if hasattr(location, 'Curve'):
            curve = location.Curve
            if curve:
                return curve.Length
    except:
        pass
    
    # Para conexoes (fittings), tentar obter dimensoes geometricas
    cat_id = element.Category.Id.IntegerValue
    if cat_id in [int(BuiltInCategory.OST_PipeFitting), 
                  int(BuiltInCategory.OST_ConduitFitting),
                  int(BuiltInCategory.OST_DuctFitting)]:
        
        # Tentar obter o comprimento desenvolvido de uma curva
        try:
            # Para cotovelos, tentar pegar o raio e calcular
            radius_param = element.LookupParameter("Radius")
            angle_param = element.LookupParameter("Angle")
            
            if radius_param and angle_param:
                radius = radius_param.AsDouble()
                angle = angle_param.AsDouble()
                if radius > 0 and angle > 0:
                    # Comprimento do arco = raio * angulo (em radianos)
                    import math
                    length = radius * angle
                    return length
        except:
            pass
        
        # Alternativamente, usar geometria do elemento
        try:
            options = Autodesk.Revit.DB.Options()
            geom_elem = element.get_Geometry(options)
            if geom_elem:
                for geom_obj in geom_elem:
                    if isinstance(geom_obj, Autodesk.Revit.DB.Solid):
                        # Para fittings, podemos estimar pelo volume ou edges
                        edges = geom_obj.Edges
                        for edge in edges:
                            curve = edge.AsCurve()
                            if curve:
                                length += curve.Length
                        if length > 0:
                            # Retornar o comprimento medio das arestas
                            return length / max(1, edges.Size) if edges.Size > 0 else 0
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
                
                # Agrupar por tipo para debug
                type_name = element.Category.Name
                if type_name not in elements_by_type:
                    elements_by_type[type_name] = {"count": 0, "length": 0.0}
                elements_by_type[type_name]["count"] += 1
                elements_by_type[type_name]["length"] += length

    # Converter pes -> metros
    total_length_meters = UnitUtils.ConvertFromInternalUnits(
        total_length_feet, UnitTypeId.Meters
    )
    total_length_meters_rounded = round(total_length_meters, 2)

    # Exibir resultado detalhado
    message = "Comprimento total dos elementos selecionados:\n\n"
    message += "TOTAL: {} m\n".format(total_length_meters_rounded)
    message += "Elementos validos: {}\n\n".format(valid_elements)
    
    message += "Detalhamento por tipo:\n"
    for type_name, data in elements_by_type.items():
        length_m = UnitUtils.ConvertFromInternalUnits(data["length"], UnitTypeId.Meters)
        message += "- {}: {} un. ({:.2f} m)\n".format(
            type_name, 
            data["count"], 
            round(length_m, 2)
        )

    TaskDialog.Show("Resultado da Contagem", message)
