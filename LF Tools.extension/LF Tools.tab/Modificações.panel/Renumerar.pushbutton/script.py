# -*- coding: UTF-8 -*-
# pyRevit + Revit API
from pyrevit import revit, forms
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.UI.Selection import ObjectType
import re
from Autodesk.Revit.Exceptions import OperationCanceledException

__title__ = "Renumerar"
__author__ = "Luís Fernando - Eng. Eletricista"
__doc__ = "Seleciona elementos um por um (com append para preservar ordem), define prefixo, número inicial, incremento e dígitos, escolhe parâmetro comum e renumera (String e Integer)."

uidoc = revit.uidoc
doc = revit.doc

# --- Funções auxiliares ---
def limpar_texto(valor):
    """Remove caracteres inválidos para caminhos/nome de arquivo"""
    return re.sub(r'[<>:"/\\|?*]', '', valor)

def get_editable_params(elements):
    """Lista parâmetros de texto ou inteiro editáveis comuns a todos os elementos."""
    if not elements:
        return []
    
    # Obtém os parâmetros do primeiro elemento como base
    first_elem = elements[0]
    params_set = set()
    for p in first_elem.Parameters:
        if p.StorageType in (StorageType.String, StorageType.Integer) and not p.IsReadOnly:
            params_set.add(p.Definition.Name)
    
    # Interseção com os parâmetros dos demais elementos
    for elem in elements[1:]:
        elem_params = set()
        for p in elem.Parameters:
            if p.StorageType in (StorageType.String, StorageType.Integer) and not p.IsReadOnly:
                elem_params.add(p.Definition.Name)
        params_set = params_set.intersection(elem_params)
    
    return sorted(params_set)

def get_parameter(elem, param_name):
    """Busca o parâmetro no elemento ou no tipo"""
    param = elem.LookupParameter(param_name)
    if not param:
        type_id = elem.GetTypeId()
        if type_id and type_id != DB.ElementId.InvalidElementId:
            elem_type = doc.GetElement(type_id)
            if elem_type:
                param = elem_type.LookupParameter(param_name)
    return param

try:
    # Passo 1: Seleção dos elementos um por um para preservar ordem
    selected_refs = []
    forms.alert("Selecione os elementos na ordem desejada (pressione ESC para finalizar).")
    while True:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element)  # Seleciona um por um
            selected_refs.append(ref)  # Append na lista para manter a ordem exata
        except OperationCanceledException:
            break  # Sai do loop ao pressionar ESC
    
    if not selected_refs:
        forms.alert("Nenhum elemento selecionado.")
        raise Exception("Nenhum elemento selecionado.")

    selected_elements = [doc.GetElement(ref) for ref in selected_refs]

    # Passo 2: Prefixo
    prefixo = forms.ask_for_string(default="EL-", prompt="Digite o prefixo (apenas para parâmetros String):") or "EL-"
    prefixo = limpar_texto(prefixo)

    # Passo 3: Número inicial
    try:
        numero_inicial = int(forms.ask_for_string(default="1", prompt="Digite o número inicial:") or "1")
    except ValueError:
        forms.alert("Número inicial inválido. Usando 1.")
        numero_inicial = 1

    # Passo 4: Incremento
    try:
        incremento = int(forms.ask_for_string(default="1", prompt="Digite o incremento:") or "1")
    except ValueError:
        forms.alert("Incremento inválido. Usando 1.")
        incremento = 1

    # Passo 5: Número de dígitos
    try:
        digitos = int(forms.ask_for_string(default="2", prompt="Digite o número de dígitos (ex.: 2 para 01, 03):") or "2")
        if digitos < 1:
            forms.alert("Número de dígitos inválido. Usando 2.")
            digitos = 2
    except ValueError:
        forms.alert("Número de dígitos inválido. Usando 2.")
        digitos = 2

    # Passo 6: Escolha do parâmetro
    param_list = get_editable_params(selected_elements)
    if not param_list:
        forms.alert("Nenhum parâmetro editável comum encontrado (String ou Integer) entre os elementos.")
        raise Exception("Sem parâmetros comuns válidos.")

    selected_param_name = forms.SelectFromList.show(
        param_list,
        title="Selecione o parâmetro para renumeração",
        button_name="Confirmar"
    )
    if not selected_param_name:
        forms.alert("Nenhum parâmetro selecionado.")
        raise Exception("Nenhum parâmetro selecionado.")

    # Passo 7: Renumeração
    erros = []
    with revit.Transaction("Renumerar Elementos Selecionados"):
        contador = numero_inicial
        for elem in selected_elements:
            param = get_parameter(elem, selected_param_name)
            if param and not param.IsReadOnly:
                try:
                    numero_formatado = str(contador).zfill(digitos)  # Adiciona zeros à esquerda
                    valor_final = limpar_texto(prefixo + numero_formatado) if param.StorageType == StorageType.String else contador
                    if param.StorageType == StorageType.String:
                        param.Set(valor_final)
                    elif param.StorageType == StorageType.Integer:
                        param.Set(int(valor_final))  # Remove zeros para Integer
                    contador += incremento
                except Exception as e:
                    erros.append(elem.Id.IntegerValue)
                    forms.alert("Erro ao renumerar elemento ID {0}: {1}".format(elem.Id.IntegerValue, str(e)))
            else:
                erros.append(elem.Id.IntegerValue)

    # Resultado final
    if erros:
        forms.alert("Renumeração concluída, mas alguns elementos não puderam ser alterados.\nIDs com erro: " + ", ".join(map(str, erros)))
    else:
        forms.alert("Renumeração concluída com sucesso!")

except Exception as e:
    forms.alert("Operação cancelada ou erro ocorrido.\n" + str(e))