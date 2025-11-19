# -*- coding: utf-8 -*-
# pyRevit

"""
Script para contar o comprimento total de eletrodutos, dutos, tubulações
e também conexões (Pipe Fittings) no Revit.
"""

__title__ = 'Contar Comprimento'
__author__ = 'Luís - Engenheiro Eletricista'

# Importar bibliotecas do Revit
from Autodesk.Revit.DB import (
    UnitUtils,
    UnitTypeId,
    StorageType,
    BuiltInCategory
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

# Importar bibliotecas do pyRevit
from pyrevit import revit, script


# Função auxiliar para obter o comprimento
def get_length(element):
    """Tenta obter o comprimento do elemento em pés."""
    length = 0.0
    param_names = [
        "Length",                  # Inglês padrão
        "Conduit Length",          # Para conduítes
        "Centerline Length",       # Para dutos
        "Comprimento",             # PT-BR genérico
        "Comprimento do eletroduto"
    ]
    for name in param_names:
        param = element.LookupParameter(name)
        if param and param.StorageType == StorageType.Double:
            val = param.AsDouble()
            if val > 0.0:
                return val

    # Caso especial: conexões de tubo podem ter parâmetro "Center to End" ou "Length"
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
        BuiltInCategory.OST_Conduit,      # Eletrodutos
        BuiltInCategory.OST_DuctCurves,   # Dutos
        BuiltInCategory.OST_PipeCurves,   # Tubos
        BuiltInCategory.OST_PipeFitting   # Conexões de Tubo
    ]

    selected_elements = []

    try:
        # Seleção múltipla
        selection = uidoc.Selection.PickObjects(
            ObjectType.Element,
            'Selecione os eletrodutos, tubos, dutos e conexões.'
        )
        for element_id in selection:
            element = doc.GetElement(element_id)
            selected_elements.append(element)

    except Exception:
        TaskDialog.Show("Cancelado", "A seleção foi cancelada.")
        script.exit()

    total_length_feet = 0.0
    valid_elements = 0

    for element in selected_elements:
        if element.Category and element.Category.Id.IntegerValue in [int(cat) for cat in allowed_categories]:
            length = get_length(element)
            if length > 0.0:
                total_length_feet += length
                valid_elements += 1

    # Converter pés -> metros
    total_length_meters = UnitUtils.ConvertFromInternalUnits(
        total_length_feet, UnitTypeId.Meters
    )
    total_length_meters_rounded = round(total_length_meters, 2)

    # Exibir resultado
    message = (
        "Comprimento total dos elementos selecionados:\n\n"
        " - Comprimento: {} m\n"
        " - Quantidade de elementos: {}"
    ).format(total_length_meters_rounded, valid_elements)

    TaskDialog.Show("Resultado da Contagem", message)
