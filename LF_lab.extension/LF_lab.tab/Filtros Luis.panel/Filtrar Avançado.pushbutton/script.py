# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ElementId, StorageType
)
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

try:
    categoria_opcoes = {
        "Eletrodutos (segmentos + curvas reais)": {
            "categorias": [
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            ],
            "classes": ["Conduit"],
            "requires_param": "RN_optional",  # só fittings precisam de RN
            "requires_param_absent": None
        },
        "Dutos": {
            "categorias": [BuiltInCategory.OST_DuctCurves],
            "classes": ["Duct"],
            "requires_param": None,
            "requires_param_absent": None
        },
        "Eletrocalhas": {
            "categorias": [BuiltInCategory.OST_CableTray],
            "classes": ["CableTray"],
            "requires_param": None,
            "requires_param_absent": None
        },
        "Tubulações": {
            "categorias": [BuiltInCategory.OST_PipeCurves],
            "classes": ["Pipe"],
            "requires_param": None,
            "requires_param_absent": None
        },
        "Conexões de Conduíte (sem RN)": {
            "categorias": [BuiltInCategory.OST_ConduitFitting],
            "classes": [],
            "requires_param": None,
            "requires_param_absent": "RN"
        }
    }

    # Etapa 1: Escolher categoria
    categoria_nome = forms.SelectFromList.show(
        sorted(categoria_opcoes.keys()),
        title="Escolha o que deseja filtrar:",
        multiselect=False
    )
    if not categoria_nome:
        script.exit()

    config = categoria_opcoes[categoria_nome]
    categorias = config["categorias"]
    classes = config["classes"]
    required_param = config.get("requires_param", None)
    required_param_absent = config.get("requires_param_absent", None)

    # Etapa 2: Escopo
    escopo = forms.SelectFromList.show(
        ["Somente na Vista Atual", "Projeto Inteiro"],
        title="Onde aplicar o filtro?",
        multiselect=False
    )
    if not escopo:
        script.exit()
    usar_vista_atual = escopo == "Somente na Vista Atual"

    # Etapa 3: Nome do parâmetro
    parametro = forms.ask_for_string(
        default="Comentários",
        prompt="Nome do parâmetro a filtrar:"
    )
    if parametro is None:
        script.exit()

    # Etapa 4: Valor buscado
    valor = forms.ask_for_string(
        default="",
        prompt="Valor a buscar (deixe vazio para valores em branco):"
    )
    if valor is None:
        script.exit()

    elementos_filtrados = []

    for cat in categorias:
        collector = FilteredElementCollector(doc, revit.active_view.Id) if usar_vista_atual else FilteredElementCollector(doc)
        collector = collector.OfCategory(cat).WhereElementIsNotElementType()

        for el in collector:
            tipo = el.GetType().Name

            # 🧠 Filtro por classe
            if classes and tipo not in classes:
                # Se exige RN opcional e não for classe esperada, verifica parâmetro
                if required_param == "RN_optional" and not el.LookupParameter("RN"):
                    continue

            # 🧠 Filtro por presença obrigatória de parâmetro
            if required_param and required_param != "RN_optional":
                if not el.LookupParameter(required_param):
                    continue

            # 🧠 Filtro por ausência de parâmetro
            if required_param_absent and el.LookupParameter(required_param_absent):
                continue

            # 🔍 Checar valor do parâmetro alvo
            p = el.LookupParameter(parametro)
            if not p:
                continue

            val = None
            if p.StorageType == StorageType.String:
                val = p.AsString()
            elif p.StorageType == StorageType.Integer:
                val = str(p.AsInteger())
            elif p.StorageType == StorageType.Double:
                val = str(p.AsDouble())

            if valor == "":
                if not val:
                    elementos_filtrados.append(el.Id)
            else:
                if val and val.strip().lower() == valor.strip().lower():
                    elementos_filtrados.append(el.Id)

    # Resultado
    if elementos_filtrados:
        uidoc.Selection.SetElementIds(List[ElementId](elementos_filtrados))
        forms.alert("{} elementos selecionados.".format(len(elementos_filtrados)))
    else:
        forms.alert("Nenhum elemento encontrado com esse critério.")

except Exception as e:
    forms.alert("Erro inesperado:\n{}".format(e))
    logger.error("Erro fatal: {}".format(e))

finally:
    logger.info("Script finalizado com segurança.")
