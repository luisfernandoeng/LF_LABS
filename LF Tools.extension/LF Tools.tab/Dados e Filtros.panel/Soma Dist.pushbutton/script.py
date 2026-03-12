# -*- coding: utf-8 -*-
"""
Contar Comprimento - Soma total de eletrodutos, dutos, tubulações,
conexões e bandejas de cabos. Agrupa por categoria e diâmetro/tamanho.
"""

__title__ = 'Contar\nComprimento'
__author__ = 'Luis Fernando'

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
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, script, forms

import math


# =====================================================================
#  CATEGORIAS SUPORTADAS
# =====================================================================
ALLOWED_CATEGORIES = [
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting,
]

ALLOWED_CAT_IDS = [int(cat) for cat in ALLOWED_CATEGORIES]

# Nomes amigáveis para categorias
CATEGORY_NAMES = {
    int(BuiltInCategory.OST_Conduit): "Eletroduto",
    int(BuiltInCategory.OST_ConduitFitting): "Conexão de Eletroduto",
    int(BuiltInCategory.OST_DuctCurves): "Duto",
    int(BuiltInCategory.OST_DuctFitting): "Conexão de Duto",
    int(BuiltInCategory.OST_PipeCurves): "Tubulação",
    int(BuiltInCategory.OST_PipeFitting): "Conexão de Tubulação",
    int(BuiltInCategory.OST_CableTray): "Bandeja de Cabos",
    int(BuiltInCategory.OST_CableTrayFitting): "Conexão de Bandeja",
}

# Agrupamento: fittings pertencem ao grupo da categoria principal
CATEGORY_GROUP = {
    int(BuiltInCategory.OST_Conduit): "Eletroduto",
    int(BuiltInCategory.OST_ConduitFitting): "Eletroduto",
    int(BuiltInCategory.OST_DuctCurves): "Duto",
    int(BuiltInCategory.OST_DuctFitting): "Duto",
    int(BuiltInCategory.OST_PipeCurves): "Tubulação",
    int(BuiltInCategory.OST_PipeFitting): "Tubulação",
    int(BuiltInCategory.OST_CableTray): "Bandeja de Cabos",
    int(BuiltInCategory.OST_CableTrayFitting): "Bandeja de Cabos",
}


# =====================================================================
#  FUNCOES DE COMPRIMENTO
# =====================================================================
def get_length(element):
    """Obtém o comprimento do elemento em pés (unidades internas)."""
    length = 0.0
    
    # PRIORIDADE 1: CURVE_ELEM_LENGTH (Built-in parameter padrão)
    try:
        length_param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if length_param and length_param.HasValue:
            val = length_param.AsDouble()
            if val > 0.0:
                return val
    except:
        pass
    
    # PRIORIDADE 2: Parametros personalizados comuns
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
    
    # PRIORIDADE 3: LocationCurve
    try:
        location = element.Location
        if isinstance(location, LocationCurve):
            curve = location.Curve
            if curve:
                return curve.Length
    except:
        pass
    
    # PRIORIDADE 4: FITTINGS - tratamento especial
    try:
        cat = element.Category
        if cat:
            cat_id = cat.Id.IntegerValue
            
            is_fitting = cat_id in [
                int(BuiltInCategory.OST_PipeFitting),
                int(BuiltInCategory.OST_ConduitFitting),
                int(BuiltInCategory.OST_DuctFitting),
                int(BuiltInCategory.OST_CableTrayFitting)
            ]
            
            if is_fitting:
                # A) COTOVELOS: raio × ângulo
                radius = None
                for radius_name in ["Radius", "Raio", "Raio de curvatura", "Bend Radius", "Bend Radius Label"]:
                    radius_param = element.LookupParameter(radius_name)
                    if radius_param and radius_param.HasValue:
                        radius = radius_param.AsDouble()
                        if radius > 0:
                            break
                
                angle_rad = None
                for angle_name in ["Angle", "Ângulo", "Angulo"]:
                    angle_param = element.LookupParameter(angle_name)
                    if angle_param and angle_param.HasValue:
                        angle_rad = angle_param.AsDouble()
                        if angle_rad > 0:
                            break
                
                if radius and angle_rad and radius > 0 and angle_rad > 0:
                    return radius * angle_rad
                
                # B) TEEs, CROSSES: Center to End
                cte = None
                for cte_name in ["Center to End", "Centro até Extremidade", "Centro para Extremidade"]:
                    cte_param = element.LookupParameter(cte_name)
                    if cte_param and cte_param.HasValue:
                        cte = cte_param.AsDouble()
                        if cte > 0:
                            break
                
                if cte and cte > 0:
                    try:
                        if hasattr(element, 'ConnectorManager'):
                            connectors = element.ConnectorManager.Connectors
                            num_connectors = sum(1 for _ in connectors)
                            if num_connectors == 3:
                                return cte * 2
                            elif num_connectors == 4:
                                return cte * 3
                    except:
                        pass
                    return cte
                
                # C) GEOMETRIA: Curvas na geometria
                try:
                    options = Options()
                    options.ComputeReferences = False
                    options.DetailLevel = ViewDetailLevel.Coarse
                    
                    geom_elem = element.get_Geometry(options)
                    if geom_elem:
                        total_curve_length = 0.0
                        for geom_obj in geom_elem:
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


def get_size_label(element):
    """Obtém o diâmetro/tamanho do elemento para agrupamento."""
    # Tentar diâmetro (eletrodutos, tubos)
    diameter_params = [
        "Diameter", "Diâmetro", "Diametro",
        "Outside Diameter", "Diâmetro Externo",
        "Nominal Diameter", "Diâmetro Nominal",
        "Size", "Tamanho"
    ]
    
    for name in diameter_params:
        param = element.LookupParameter(name)
        if param and param.HasValue:
            if param.StorageType == StorageType.Double:
                val_mm = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
                if val_mm > 0:
                    # Arredondar para inteiro pra limpar valores como 24.99999
                    val_mm_rounded = int(round(val_mm))
                    return "{}mm".format(val_mm_rounded)
            elif param.StorageType == StorageType.String:
                val = param.AsString()
                if val:
                    return val.strip()
    
    # Tentar tamanho via Type (para bandejas, dutos etc.)
    try:
        elem_type = element.Document.GetElement(element.GetTypeId())
        if elem_type:
            for name in diameter_params:
                param = elem_type.LookupParameter(name)
                if param and param.HasValue:
                    if param.StorageType == StorageType.Double:
                        val_mm = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
                        if val_mm > 0:
                            return "{}mm".format(int(round(val_mm)))
                    elif param.StorageType == StorageType.String:
                        val = param.AsString()
                        if val:
                            return val.strip()
    except:
        pass
    
    # Para dutos/bandejas: tentar largura x altura
    try:
        w_param = element.LookupParameter("Width") or element.LookupParameter("Largura")
        h_param = element.LookupParameter("Height") or element.LookupParameter("Altura")
        if w_param and h_param and w_param.HasValue and h_param.HasValue:
            w_mm = int(round(UnitUtils.ConvertFromInternalUnits(w_param.AsDouble(), UnitTypeId.Millimeters)))
            h_mm = int(round(UnitUtils.ConvertFromInternalUnits(h_param.AsDouble(), UnitTypeId.Millimeters)))
            if w_mm > 0 and h_mm > 0:
                return "{}x{}mm".format(w_mm, h_mm)
    except:
        pass
    
    return "N/D"


# =====================================================================
#  MAIN
# =====================================================================
if __name__ == '__main__':
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    output.close_others()

    selected_elements = []

    # 1. Verificar seleção prévia
    pre_selected = revit.get_selection()
    if pre_selected:
        for el in pre_selected:
            try:
                if el.Category and el.Category.Id.IntegerValue in ALLOWED_CAT_IDS:
                    selected_elements.append(el)
            except:
                pass
    
    # 2. Se não tem seleção válida, pedir ao usuário
    if not selected_elements:
        try:
            selection = uidoc.Selection.PickObjects(
                ObjectType.Element,
                'Selecione eletrodutos, tubos, dutos, bandejas e conexões.'
            )
            for ref in selection:
                element = doc.GetElement(ref)
                if element.Category and element.Category.Id.IntegerValue in ALLOWED_CAT_IDS:
                    selected_elements.append(element)
        except:
            script.exit()

    if not selected_elements:
        forms.alert("Nenhum elemento válido selecionado.")
        script.exit()

    # 3. Processar elementos
    # Estrutura: { grupo: { tamanho: { count, length, is_fitting_count } } }
    data = {}
    total_length_feet = 0.0
    total_valid = 0
    total_fittings = 0

    for element in selected_elements:
        cat_id = element.Category.Id.IntegerValue
        length = get_length(element)
        
        if length > 0.0:
            total_length_feet += length
            total_valid += 1
            
            group = CATEGORY_GROUP.get(cat_id, "Outros")
            size = get_size_label(element)
            
            is_fitting = cat_id in [
                int(BuiltInCategory.OST_ConduitFitting),
                int(BuiltInCategory.OST_DuctFitting),
                int(BuiltInCategory.OST_PipeFitting),
                int(BuiltInCategory.OST_CableTrayFitting)
            ]
            
            if is_fitting:
                total_fittings += 1
            
            if group not in data:
                data[group] = {}
            if size not in data[group]:
                data[group][size] = {"count": 0, "length": 0.0, "fittings": 0}
            
            data[group][size]["count"] += 1
            data[group][size]["length"] += length
            if is_fitting:
                data[group][size]["fittings"] += 1

    # 4. Exibir resultado com pyRevit output
    total_meters = UnitUtils.ConvertFromInternalUnits(total_length_feet, UnitTypeId.Meters)

    output.print_md("# 📏 Relatório de Comprimento")
    output.print_md("---")
    output.print_md("**Elementos selecionados:** {} &nbsp;|&nbsp; **Válidos:** {} &nbsp;|&nbsp; **Conexões:** {}".format(
        len(selected_elements), total_valid, total_fittings))
    output.print_md("")

    if not data:
        output.print_md("> Nenhum elemento com comprimento encontrado.")
        script.exit()

    # Tabela por grupo
    for group_name in sorted(data.keys()):
        sizes = data[group_name]
        
        group_total_feet = sum(s["length"] for s in sizes.values())
        group_total_m = UnitUtils.ConvertFromInternalUnits(group_total_feet, UnitTypeId.Meters)
        group_total_count = sum(s["count"] for s in sizes.values())
        
        output.print_md("## {} &nbsp; ({} un. — {:.2f} m)".format(
            group_name, group_total_count, group_total_m))
        
        # Montar tabela
        table_data = []
        for size_name in sorted(sizes.keys()):
            s = sizes[size_name]
            length_m = UnitUtils.ConvertFromInternalUnits(s["length"], UnitTypeId.Meters)
            fitting_info = ""
            if s["fittings"] > 0:
                straight = s["count"] - s["fittings"]
                if straight > 0:
                    fitting_info = "{} retos + {} conexões".format(straight, s["fittings"])
                else:
                    fitting_info = "{} conexões".format(s["fittings"])
            else:
                fitting_info = "{} retos".format(s["count"])
            
            table_data.append([size_name, s["count"], "{:.2f} m".format(length_m), fitting_info])
        
        output.print_table(
            table_data,
            columns=["Tamanho", "Qtd", "Comprimento", "Detalhe"],
            title=""
        )
        output.print_md("")

    # Resumo final
    output.print_md("---")
    output.print_md("## 📊 TOTAL GERAL: &nbsp; {} elementos &nbsp;→&nbsp; **{:.2f} m**".format(
        total_valid, total_meters))
    
    if total_meters >= 1000:
        output.print_md("&nbsp;&nbsp;&nbsp;&nbsp; *(= {:.3f} km)*".format(total_meters / 1000.0))