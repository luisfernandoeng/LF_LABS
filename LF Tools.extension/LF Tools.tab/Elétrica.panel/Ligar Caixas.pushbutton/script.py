# -*- coding: utf-8 -*-
import clr
import os
import datetime
import io

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException

from pyrevit import forms, revit

# Documento e UIDocumento ativos
doc = revit.doc
uidoc = revit.uidoc

# ===== Funções auxiliares =====

def log(msg):
    """Escreve logs no Desktop com codificação UTF-8."""
    log_path = os.path.expanduser("~/Desktop/ligar_caixas_log.txt")
    with io.open(log_path, "a", encoding="utf-8") as f:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(u"{} - {}\n".format(ts, msg))

def obter_nome_tipo(elemento):
    """Obtém o nome do tipo de um elemento com segurança."""
    try:
        tipo = doc.GetElement(elemento.GetTypeId())
        return tipo.Name if tipo and hasattr(tipo, "Name") else "Sem tipo"
    except:
        return "Erro ao obter tipo"

def obter_conectores_validos(elemento):
    """Retorna conectores válidos para conduítes."""
    conectores = []
    try:
        conectores = elemento.MEPModel.ConnectorManager.Connectors
    except AttributeError:
        try:
            conectores = elemento.ConnectorManager.Connectors
        except AttributeError:
            log("Elemento {} não possui conectores válidos.".format(elemento.Id))
            return []
    validos = [c for c in conectores if c.Domain == Domain.DomainCableTrayConduit and not c.IsConnected]
    log("Elemento {}: {} conectores válidos encontrados.".format(elemento.Id, len(validos)))
    return validos

def conector_mais_proximo(conectores, ponto_ref):
    """Retorna o conector mais próximo de um ponto de referência."""
    if not conectores:
        return None
    return min(conectores, key=lambda c: c.Origin.DistanceTo(ponto_ref))

def encontrar_tipo_conduite(nome_procurado="Eletroduto Flexivel Corrugado de PVC amarelo_Tigreflex"):
    """Encontra um tipo de conduíte pelo nome ou retorna o primeiro disponível."""
    log("Buscando tipo de conduíte: {}".format(nome_procurado))
    tipos = FilteredElementCollector(doc).OfClass(ConduitType).ToElements()
    if not tipos:
        log("Nenhum tipo de conduíte encontrado no projeto.")
        return None
    
    valid_types = []
    for t in tipos:
        try:
            if hasattr(t, "Name") and t.Name:
                valid_types.append(t)
                if nome_procurado.lower() in t.Name.lower():
                    log("Tipo encontrado: {} (ID: {})".format(t.Name, t.Id))
                    return t
            else:
                log("Elemento ID {} não possui propriedade Name válida.".format(t.Id))
        except Exception as e:
            log("Erro ao acessar tipo de conduíte ID {}: {}".format(t.Id, str(e)))
    
    if valid_types:
        log("Tipo '{}' não encontrado. Usando o primeiro disponível: {} (ID: {})".format(
            nome_procurado, valid_types[0].Name, valid_types[0].Id))
        return valid_types[0]
    
    log("Nenhum tipo de conduíte válido encontrado no projeto.")
    return None

def validar_caixa(elemento):
    """Verifica se o elemento está em uma categoria válida para caixas elétricas."""
    try:
        categoria = elemento.Category.Name if elemento.Category else "Sem categoria"
        log("Validando elemento {} | Categoria: {}".format(elemento.Id, categoria))
        return categoria in ["Electrical Equipment", "Electrical Fixtures", "Luminárias", "Dispositivos elétricos", "Dispositivos de telefonia"]
    except:
        log("Erro ao validar categoria do elemento {}".format(elemento.Id))
        return False

def obter_subcomponente_caixa(elemento):
    """Verifica se o elemento ou seus subcomponentes são caixas elétricas válidas."""
    if validar_caixa(elemento):
        log("Elemento {} é uma caixa elétrica válida.".format(elemento.Id))
        return elemento
    try:
        subcomponentes = elemento.GetSubComponentIds()
        log("Elemento {} tem {} subcomponentes.".format(elemento.Id, len(subcomponentes)))
        for sub_id in subcomponentes:
            sub_elem = doc.GetElement(sub_id)
            if validar_caixa(sub_elem):
                log("Subcomponente {} é uma caixa elétrica válida.".format(sub_elem.Id))
                return sub_elem
    except:
        log("Erro ao verificar subcomponentes do elemento {}".format(elemento.Id))
    return None

def obter_nivel_valido():
    """Obtém um nível válido para criar o conduíte."""
    try:
        nivel = doc.ActiveView.GenLevel
        if nivel and nivel.Id != ElementId.InvalidElementId:
            log("Nível encontrado: {}".format(nivel.Name))
            return nivel.Id
        else:
            niveis = FilteredElementCollector(doc).OfClass(Level).ToElements()
            if niveis:
                log("Nível padrão selecionado: {}".format(niveis[0].Name))
                return niveis[0].Id
            log("Nenhum nível encontrado no projeto.")
            return None
    except:
        log("Erro ao obter nível da vista ativa.")
        return None

def criar_conduite_entre_pontos(start_point, end_point, tipo_conduite, nivel_id, diametro=None):
    """Cria um conduíte entre dois pontos, semelhante ao Conduit.ByLine do Dynamo."""
    try:
        cond = Conduit.Create(doc, tipo_conduite.Id, start_point, end_point, nivel_id)
        if diametro:
            param = cond.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
            if param and diametro > 0:
                param.Set(diametro)
                log("Diâmetro do conduíte definido: {:.2f} mm".format(diametro * 304.8))
            else:
                log("Diâmetro inválido ou parâmetro não encontrado.")
        log("Conduíte criado com sucesso. ID: {}".format(cond.Id))
        return cond
    except Exception as e:
        log("Erro ao criar conduíte: {}".format(str(e)))
        raise

# ===== Execução principal =====

def ligar_caixas():
    log("\n=== Início da execução ===")

    # Seleção da primeira caixa
    with forms.WarningBar(title="Selecione a primeira caixa elétrica"):
        try:
            ref1 = uidoc.Selection.PickObject(ObjectType.Element, "Selecione a primeira caixa elétrica")
            if not ref1:
                log("Seleção da primeira caixa cancelada.")
                return
        except InvalidOperationException:
            log("Seleção da primeira caixa cancelada ou erro na seleção.")
            return
    caixa1 = doc.GetElement(ref1.ElementId)
    caixa1 = obter_subcomponente_caixa(caixa1)
    if not caixa1:
        log("Elemento {} não é uma caixa elétrica ou não contém subcomponentes válidos.".format(ref1.ElementId))
        forms.alert("O primeiro elemento selecionado não é uma caixa elétrica ou não contém uma caixa válida.", title="Erro")
        return

    # Seleção da segunda caixa
    with forms.WarningBar(title="Selecione a segunda caixa elétrica"):
        try:
            ref2 = uidoc.Selection.PickObject(ObjectType.Element, "Selecione a segunda caixa elétrica")
            if not ref2:
                log("Seleção da segunda caixa cancelada.")
                return
        except InvalidOperationException:
            log("Seleção da segunda caixa cancelada ou erro na seleção.")
            return
    caixa2 = doc.GetElement(ref2.ElementId)
    caixa2 = obter_subcomponente_caixa(caixa2)
    if not caixa2:
        log("Elemento {} não é uma caixa elétrica ou não contém subcomponentes válidos.".format(ref2.ElementId))
        forms.alert("O segundo elemento selecionado não é uma caixa elétrica ou não contém uma caixa válida.", title="Erro")
        return

    log("Caixa 1: {} | Tipo: {} | Categoria: {}".format(caixa1.Id, obter_nome_tipo(caixa1), caixa1.Category.Name))
    log("Caixa 2: {} | Tipo: {} | Categoria: {}".format(caixa2.Id, obter_nome_tipo(caixa2), caixa2.Category.Name))

    # Obtém conectores válidos
    conectores1 = obter_conectores_validos(caixa1)
    conectores2 = obter_conectores_validos(caixa2)
    if not conectores1 or not conectores2:
        log("Uma ou ambas as caixas não possuem conectores válidos.")
        forms.alert("Uma ou ambas as caixas não possuem conectores válidos. Verifique se as caixas têm conectores MEP configurados.", title="Erro")
        return

    # Seleciona os conectores mais próximos
    con1 = conector_mais_proximo(conectores1, caixa2.Location.Point)
    con2 = conector_mais_proximo(conectores2, caixa1.Location.Point)
    if not con1 or not con2:
        log("Não foi possível encontrar conectores válidos para conexão.")
        forms.alert("Não foi possível encontrar conectores válidos para conexão.", title="Erro")
        return

    # Encontra o tipo de conduíte
    tipo_conduite = encontrar_tipo_conduite()
    if not tipo_conduite:
        log("Nenhum tipo de conduíte válido disponível no projeto.")
        forms.alert("Nenhum tipo de conduíte encontrado. Carregue um tipo de conduíte no projeto.", title="Erro")
        return

    # Obtém um nível válido
    nivel_id = obter_nivel_valido()
    if not nivel_id:
        log("Nenhum nível válido encontrado no projeto.")
        forms.alert("Nenhum nível válido encontrado. Veja o log no Desktop.", title="Erro")
        return

    # Criação do conduíte
    with Transaction(doc, "Ligar Caixas com Conduíte") as t:
        try:
            t.Start()
            # Cria conduíte entre os pontos dos conectores
            diametro = con1.Radius * 2 if con1.Radius > 0 else None
            cond = criar_conduite_entre_pontos(con1.Origin, con2.Origin, tipo_conduite, nivel_id, diametro)

            # Conecta o conduíte às caixas
            cond_connectors = [c for c in cond.ConnectorManager.Connectors if not c.IsConnected]
            if len(cond_connectors) >= 2:
                con1.ConnectTo(cond_connectors[0])
                con2.ConnectTo(cond_connectors[1])
            else:
                log("Conduíte não possui conectores suficientes para conexão.")
                raise Exception("Conduíte não possui conectores suficientes.")

            t.Commit()
            log("Conduíte criado com sucesso. ID: {}".format(cond.Id))
            forms.alert("Conduíte criado com sucesso!", title="Sucesso")
        except Exception as e:
            t.RollBack()
            log("Erro ao criar conduíte: {}".format(str(e)))
            forms.alert("Erro ao criar conduíte: {}. Veja o log no Desktop.".format(str(e)), title="Erro")

    log("=== Fim da execução ===")

if __name__ == "__main__":
    ligar_caixas()