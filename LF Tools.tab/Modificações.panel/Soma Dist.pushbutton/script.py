# -*- coding: utf-8 -*-

"""
Script para contar comprimento total de eletrodutos, dutos, tubulacoes
e conexoes no Revit.
"""

__title__ = 'Contar Comprimento'
__author__ = 'Luis - Engenheiro Eletricista'

# Importar bibliotecas do Revit
from Autodesk.Revit.DB import (
    UnitUtils,
    UnitTypeId,
    StorageType,
    BuiltInCategory,
    FilteredElementCollector
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

# Importar bibliotecas do pyRevit
from pyrevit import revit, script

# Funcao auxiliar para obter o comprimento
def get_length(element):
    """Tenta obter o comprimento do elemento em pes."""
    length = 0.0
    param_names = [
        "Length",
        "Conduit Length", 
        "Centerline Length",
        "Comprimento",
        "Comprimento do eletroduto"
    ]
    
    for name in param_names:
        param = element.LookupParameter(name)
        if param and param.StorageType == StorageType.Double:
            val = param.AsDouble()
            if val > 0.0:
                return val

    # Caso especial: conexoes de tubo
    if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
        for alt in ["Center to End", "Length", "Comprimento"]:
            param = element.LookupParameter(alt)
            if param and param.StorageType == StorageType.Double:
                val = param.AsDouble()
                if val > 0.0:
                    return val

    return length

# Main script
if __name__ == '__main__':
    doc = revit.doc
    uidoc = revit.uidoc

    # Categorias permitidas
    allowed_categories = [
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_DuctCurves, 
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_PipeFitting
    ]

    selected_elements = []

    try:
        # Selecao multipla
        selection = uidoc.Selection.PickObjects(
            ObjectType.Element,
            'Selecione os eletrodutos, tubos, dutos e conexoes.'
        )
        for element_id in selection:
            element = doc.GetElement(element_id)
            selected_elements.append(element)

    except Exception as e:
        TaskDialog.Show("Cancelado", "A selecao foi cancelada.")
        script.exit()

    total_length_feet = 0.0
    valid_elements = 0

    for element in selected_elements:
        if element.Category and element.Category.Id.IntegerValue in [int(cat) for cat in allowed_categories]:
            length = get_length(element)
            if length > 0.0:
                total_length_feet += length
                valid_elements += 1

    # Converter pes -> metros
    total_length_meters = UnitUtils.ConvertFromInternalUnits(
        total_length_feet, UnitTypeId.Meters
    )
    total_length_meters_rounded = round(total_length_meters, 2)

    # Exibir resultado
    message = "Comprimento total dos elementos selecionados:\n\n"
    message += " - Comprimento: {} m\n".format(total_length_meters_rounded)
    message += " - Quantidade de elementos: {}".format(valid_elements)

    TaskDialog.Show("Resultado da Contagem", message)