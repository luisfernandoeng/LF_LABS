# -*- coding: utf-8 -*-
"""Lê a posição (X,Y) dos elementos selecionados e gera um relatório para ajudar na programação."""

from pyrevit import revit, script, DB
import json

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

selection_ids = uidoc.Selection.GetElementIds()

if not selection_ids:
    script.exit("Por favor, selecione os elementos (linhas, textos, blocos) antes de rodar o script.")

data = []

# Converter Pés para Milímetros para ficar mais fácil de entender
def to_mm(feet):
    return round(feet * 304.8, 1)

for eid in selection_ids:
    elem = doc.GetElement(eid)
    
    info = {
        "Categoria": elem.Category.Name if elem.Category else "Desconhecida",
        "Id": str(elem.Id)
    }
    
    # Se for uma Linha de Detalhe (Detail Line / Detail Curve)
    if isinstance(elem, DB.DetailCurve) or isinstance(elem, DB.ModelCurve):
        curve = elem.GeometryCurve
        if isinstance(curve, DB.Line):
            info["Tipo"] = "Linha Reta"
            info["Inicio_XY_mm"] = [to_mm(curve.GetEndPoint(0).X), to_mm(curve.GetEndPoint(0).Y)]
            info["Fim_XY_mm"] = [to_mm(curve.GetEndPoint(1).X), to_mm(curve.GetEndPoint(1).Y)]
            info["Comprimento_mm"] = to_mm(curve.Length)
        else:
            info["Tipo"] = "Curva Complexa"
            
    # Se for Texto (Text Note)
    elif isinstance(elem, DB.TextNote):
        info["Tipo"] = "Texto"
        info["Texto"] = elem.Text
        pos = elem.Coord
        info["Posicao_XY_mm"] = [to_mm(pos.X), to_mm(pos.Y)]
        
    # Se for uma Família / Anotação
    elif isinstance(elem, DB.FamilyInstance):
        info["Tipo"] = "Família/Bloco"
        info["Nome"] = elem.Symbol.Family.Name
        if elem.Location and isinstance(elem.Location, DB.LocationPoint):
            pos = elem.Location.Point
            info["Posicao_XY_mm"] = [to_mm(pos.X), to_mm(pos.Y)]
            
    data.append(info)

# Exibir resultado
output.print_md("### 📊 Relatório de Geometria para o ChatGPT")
output.print_md("Copie o texto na caixa preta abaixo e envie no chat para a IA. Assim ela saberá as coordenadas exatas para programar o desenho!")
output.print_html("<br>")

# JSON formatado
json_str = json.dumps(data, indent=2, ensure_ascii=False)
output.print_code(json_str)
