# -*- coding: utf-8 -*-
#pylint: disable=E0401,W0703,C0103,W0622
"""
LF Labs - Ligar Caixas (Conectar Eletrodutos)

Ferramenta para conectar eletrodutos entre dois elementos selecionados
(caixas, equipamentos, etc.) que possuam conectores MEP válidos.
Tenta criar um eletroduto e conectar os elementos de forma programática.

Autor: [Luís Fernando/LF Labs]
Versão: 1.1
Data: 2024-07-09
"""
from pyrevit import forms, script
from pyrevit.revit import doc, uidoc, active_view
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
import os
import datetime
import traceback

# Configuração do Logger para depuração
logger = script.get_logger()
logger.setLevel("DEBUG")

# Caminho do arquivo de log - ATENÇÃO: Mude este caminho para um local acessível!
# Ex: log_path = os.path.join(os.environ["TEMP"], "log_ligar_caixas.txt")
log_path = r"C:\Users\Ian\Desktop\log_ligar_caixas.txt"

def setup_log():
    """Configura o arquivo de log, limpando-o no início da execução."""
    try:
        with open(log_path, "w") as f:
            f.write("=== LOG LIGAR CAIXAS - {} ===\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    except IOError as e:
        forms.alert("Não foi possível criar/limpar o arquivo de log em: {}\nErro: {}".format(log_path, e), title="Erro de Log")
        # Se não conseguir escrever no log, pelo menos tentamos logar no console do PyRevit
        logger.error("Não foi possível criar/limpar o arquivo de log: %s", e)

def log_message(msg, level="INFO"):
    """
    Registra mensagens no arquivo de log e no console do PyRevit.
    Níveis: INFO, DEBUG, WARNING, ERROR.
    """
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = "{} - {} - {}".format(timestamp, level, msg)
    try:
        with open(log_path, "a") as f:
            f.write(log_line + "\n")
    except IOError as e:
        # Se não conseguir escrever no log, imprime no console do PyRevit
        logger.error("Erro ao escrever no arquivo de log: %s - %s", log_line, e)

    if level == "ERROR":
        logger.error(msg)
    elif level == "DEBUG":
        logger.debug(msg)
    elif level == "WARNING":
        logger.warning(msg)
    else:
        logger.info(msg)

def get_mep_connectors(element):
    """
    Obtém todos os conectores MEP de um dado elemento.
    Retorna uma lista de conectores.
    """
    connectors = []
    if hasattr(element, 'MEPModel') and element.MEPModel and element.MEPModel.ConnectorManager:
        for conn in element.MEPModel.ConnectorManager.Connectors:
            if conn:
                connectors.append(conn)
    elif hasattr(element, 'ConnectorManager') and element.ConnectorManager:
        # Alguns elementos podem ter ConnectorManager direto (e.g., Conduit)
        for conn in element.ConnectorManager.Connectors:
            if conn:
                connectors.append(conn)
    
    if not connectors:
        log_message("Elemento {} (ID: {}) não possui conectores MEP válidos.".format(element.Name, element.Id), "WARNING")
    return connectors

def get_closest_connector(element, target_point):
    """
    Encontra o conector mais próximo de um elemento em relação a um ponto alvo.
    """
    mep_connectors = get_mep_connectors(element)
    if not mep_connectors:
        return None

    closest = None
    min_dist = float("inf")
    for c in mep_connectors:
        # Verifica se o conector está ativo e tem uma origem válida
        if c.IsConnected and c.Origin: # Prioriza conectores livres
            # Considera apenas conectores MEP que podem aceitar novas conexões
            if c.Domain == Domain.DomainCableTrayConduit and c.ConnectorType == ConnectorType.End:
                dist = c.Origin.DistanceTo(target_point)
                if dist < min_dist:
                    min_dist = dist
                    closest = c
    
    # Fallback: Se nenhum conector livre for encontrado, tenta pegar o mais próximo mesmo que esteja conectado
    if not closest and mep_connectors:
        for c in mep_connectors:
            if c.Origin:
                dist = c.Origin.DistanceTo(target_point)
                if dist < min_dist:
                    min_dist = dist
                    closest = c
    
    if closest:
        log_message("Conector encontrado para {}: {}".format(element.Name, closest.Origin), "DEBUG")
    else:
        log_message("Nenhum conector adequado encontrado para {} perto de {}".format(element.Name, target_point), "WARNING")
    return closest

def conduit_exists_between(conn1, conn2):
    """
    Verifica se já existe um eletroduto conectando os dois conectores.
    Isso é uma verificação aproximada baseada nos pontos de origem dos conectores.
    """
    # A tolerância é importante para comparações de ponto flutuante no Revit API
    tolerance = 0.01 # Pés, equivalente a ~3mm
    
    # Coleta todos os eletrodutos
    all_conduits = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType()
    
    for conduit in all_conduits:
        # Eletrodutos devem ter LocationCurve
        if conduit.Location and isinstance(conduit.Location, LocationCurve):
            curve = conduit.Location.Curve
            p_start = curve.GetEndPoint(0)
            p_end = curve.GetEndPoint(1)

            # Verifica se os endpoints do eletroduto estão próximos dos origins dos conectores
            # Verifica ambas as direções (p1-p2 ou p2-p1)
            if (p_start.IsAlmostEqualTo(conn1.Origin, tolerance) and p_end.IsAlmostEqualTo(conn2.Origin, tolerance)) or \
               (p_start.IsAlmostEqualTo(conn2.Origin, tolerance) and p_end.IsAlmostEqualTo(conn1.Origin, tolerance)):
                log_message("Eletroduto existente encontrado entre {} e {} (ID: {})".format(conn1.Origin, conn2.Origin, conduit.Id), "INFO")
                return True
    return False

def find_conduit_type():
    """
    Encontra o tipo de eletroduto padrão no projeto.
    """
    # Preferível usar BuiltInParameter ou um diálogo de seleção se houver muitos tipos
    # Para este script, mantemos a busca por nome.
    conduit_types = FilteredElementCollector(doc).OfClass(ElementType).OfCategory(BuiltInCategory.OST_Conduit)
    for c_type in conduit_types:
        # Adapte os nomes de família e tipo conforme o seu template.
        # Exemplo: 'Eletroduto Flexivel Corrugado de PVC amarelo_Tigreflex' é o Type Name
        # e 'Conduite sem conexões' é o Family Name.
        if c_type.FamilyName == "Eletroduto Flexivel Corrugado de PVC amarelo_Tigreflex" and c_type.Name == "Eletroduto Flexivel Corrugado de PVC amarelo_Tigreflex":
            log_message("Tipo de eletroduto encontrado: {} - {}".format(c_type.FamilyName, c_type.Name), "DEBUG")
            return c_type
    log_message("Tipo de eletroduto 'Eletroduto Flexivel Corrugado de PVC amarelo_Tigreflex' não encontrado. Verifique seu template.", "ERROR")
    return None

def find_fitting_type_for_auto_creation():
    """
    Tenta encontrar um tipo de fitting que possa ser usado para criação automática.
    Geralmente, isso é feito pela API internamente.
    Este método é mais para garantir que um fitting Type exista, se necessário para uma criação manual.
    """
    fitting_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_ConduitFitting)
    for symb in fitting_symbols:
        # Exemplo: Procurar por uma curva padrão ou união.
        # A API do Revit muitas vezes pode inferir o fitting correto se os conectores são compatíveis.
        if symb.FamilyName == "Curva generica para eletroduto corrugado PVC Amarelo" and symb.Name == "Curva generica para eletroduto corrugado Amarelo":
            if not symb.IsActive:
                try:
                    symb.Activate()
                    doc.Regenerate() # Necessário após ativar FamilySymbol
                except Exception as e:
                    log_message("Erro ao ativar FamilySymbol {}: {}".format(symb.Name, e), "ERROR")
                    return None
            log_message("Tipo de curva fitting encontrado e ativado: {}".format(symb.Name), "DEBUG")
            return symb
    log_message("Tipo de curva fitting 'Curva generica para eletroduto corrugado Amarelo' não encontrado. Pode ser que a criação automática pela API não funcione para todos os casos.", "WARNING")
    return None

# --- Início da Execução do Script ---
setup_log()
log_message("Iniciando execução do script Ligar Caixas.")

try:
    # 1. Seleção dos Elementos
    forms.alert("Selecione a primeira caixa/elemento.")
    ref1 = uidoc.Selection.PickObject(ObjectType.Element, "Selecionar primeiro elemento")
    el1 = doc.GetElement(ref1)

    forms.alert("Selecione a segunda caixa/elemento.")
    ref2 = uidoc.Selection.PickObject(ObjectType.Element, "Selecionar segundo elemento")
    el2 = doc.GetElement(ref2)

    log_message("Elementos selecionados: '{}' (ID: {}) e '{}' (ID: {})".format(el1.Name, el1.Id, el2.Name, el2.Id))

    # 2. Encontrar Conectores
    # Tentativa de encontrar o conector mais próximo de cada elemento em relação ao outro elemento
    # Esta lógica é para encontrar o "melhor" par de conectores para a conexão.
    conn1 = get_closest_connector(el1, el2.Location.Point)
    conn2 = get_closest_connector(el2, el1.Location.Point)

    if not conn1 or not conn2:
        forms.alert("❌ Não foi possível encontrar conectores MEP válidos em um ou ambos os elementos selecionados.")
        log_message("Falha ao encontrar conectores válidos.", "ERROR")
        script.exit()
    
    # Validação adicional de conectores
    if conn1.Domain != Domain.DomainCableTrayConduit or conn2.Domain != Domain.DomainCableTrayConduit:
        forms.alert("❌ Os conectores selecionados não são de eletroduto/bandeja (DomainCableTrayConduit).")
        log_message("Domínio dos conectores incompatível.", "ERROR")
        script.exit()

    # 3. Verificação de Eletroduto Existente
    if conduit_exists_between(conn1, conn2):
        choice = forms.alert(
            "Já existe um eletroduto ou conexão muito próxima entre esses elementos.\nDeseja continuar e tentar criar mesmo assim?",
            options=["Sim, Continuar", "Não, Cancelar"]
        )
        if choice != "Sim, Continuar":
            log_message("Usuário cancelou a criação devido a eletroduto existente.", "INFO")
            script.exit()

    # 4. Encontrar Tipo de Eletroduto
    conduit_type = find_conduit_type()
    if not conduit_type:
        forms.alert("❌ Tipo de eletroduto necessário não encontrado no projeto. Verifique o console do PyRevit para detalhes.", title="Erro: Tipo de Eletroduto")
        script.exit()
    
    # A ativação do FamilySymbol para fittings não é estritamente necessária se a API cria o fitting.
    # Mas é uma boa prática garantir que ele está ativo se você fosse criar um FamilyInstance explicitamente.
    # fitting_type = find_fitting_type_for_auto_creation() 
    # if not fitting_type:
    #    forms.alert("Atenção: Tipo de curva fitting não encontrado. A criação automática de curva pela API pode ser afetada.", title="Aviso")


    level = active_view.GenLevel # Pega o nível da vista ativa
    if not level:
        # Fallback: tenta pegar um nível qualquer se a vista não tiver GenLevel
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        if levels:
            level = levels[0]
            log_message("Nível da vista ativa não encontrado, usando o primeiro nível disponível: {}".format(level.Name), "WARNING")
        else:
            forms.alert("❌ Não foi possível determinar o nível para criar o eletroduto.", title="Erro: Nível")
            log_message("Nenhum nível encontrado no projeto.", "ERROR")
            script.exit()

    # 5. Lógica de Conexão - Transação
    t = Transaction(doc, "LF Labs - Conectar Eletrodutos")
    try:
        t.Start()

        # Tentar conectar os conectores diretamente.
        # A API do Revit pode criar o eletroduto reto e/ou o fitting automaticamente.
        # É essencial que os conectores sejam compatíveis (domínio, diâmetro, etc.)
        log_message("Tentando conectar conectores: {} -> {}".format(conn1.Origin, conn2.Origin), "DEBUG")
        
        # Cria um eletroduto temporário para tentar a conexão.
        # A forma mais robusta é usar Conduit.Create direto e depois tentar conn.ConnectTo
        # ou deixar a API gerenciar a criação do segmento e fitting.
        
        # Opção 1: Criar um segmento reto e tentar conectar nas pontas
        # Calcula um ponto de destino temporário para o primeiro eletroduto
        # Usar um ponto intermediário pode ser útil para guiar a criação.
        # Se os conectores estiverem em alturas diferentes, o Revit criará um segmento vertical ou um fitting.
        
        # Define o ponto inicial e final para o segmento principal do eletroduto
        # Simplesmente conectando as origens dos conectores.
        # Este é o ponto onde o usuário espera que o eletroduto vá.
        start_pt = conn1.Origin
        end_pt = conn2.Origin

        # Criação do eletroduto principal
        new_conduit = Conduit.Create(doc, conduit_type.Id, start_pt, end_pt, level.Id)
        
        if new_conduit:
            log_message("Eletroduto principal (ID: {}) criado entre {} e {}".format(new_conduit.Id, start_pt, end_pt), "INFO")
            
            # Tentar aplicar diâmetro se o parâmetro existir
            param_diameter = new_conduit.LookupParameter("Diâmetro")
            if param_diameter and param_diameter.IsReadOnly == False:
                # Exemplo: 20 mm. Converta para as unidades internas do Revit (pés)
                # 20 mm = 20 / 304.8 pés
                try:
                    param_diameter.Set(20.0 / 304.8) # Diâmetro de 20mm
                    log_message("Diâmetro do eletroduto definido para 20mm.", "DEBUG")
                except Exception as e:
                    log_message("Não foi possível definir o diâmetro do eletroduto: {}".format(e), "WARNING")
            else:
                log_message("Parâmetro 'Diâmetro' não encontrado ou é somente leitura no eletroduto criado.", "WARNING")

            # Tentar conectar os conectores. A API pode estender o eletroduto recém-criado
            # ou criar um fitting se necessário e se os conectores forem compatíveis.
            # Percorre os conectores do *novo* eletroduto para tentar conectar com os elementos originais.
            
            new_conduit_connectors = get_mep_connectors(new_conduit)
            
            connected_successfully = False
            for new_c_from_conduit in new_conduit_connectors:
                # Tenta conectar o conector do eletroduto ao conector do primeiro elemento
                # Verifica se o conector do eletroduto ainda não está conectado (ao outro elemento)
                # e se é compatível (domínio, tipo de conexão)
                if new_c_from_conduit.Domain == conn1.Domain and \
                   new_c_from_conduit.ConnectorType == conn1.ConnectorType and \
                   not new_c_from_conduit.IsConnected: # Garante que o conector do eletroduto está livre
                    try:
                        new_c_from_conduit.ConnectTo(conn1)
                        log_message("Conectado conector do eletroduto ({} -> {}) ao elemento 1 ({}).".format(new_c_from_conduit.Origin, conn1.Origin, el1.Name), "DEBUG")
                        connected_successfully = True
                    except Exception as e:
                        log_message("Erro ao conectar eletroduto ao elemento 1 ({}): {}".format(el1.Name, e), "ERROR")
                        # Em caso de erro, pode ser que o conector já tenha sido usado ou seja incompatível.

                # Tenta conectar o conector do eletroduto ao conector do segundo elemento
                if new_c_from_conduit.Domain == conn2.Domain and \
                   new_c_from_conduit.ConnectorType == conn2.ConnectorType and \
                   not new_c_from_conduit.IsConnected: # Garante que o conector do eletroduto está livre
                    try:
                        new_c_from_conduit.ConnectTo(conn2)
                        log_message("Conectado conector do eletroduto ({} -> {}) ao elemento 2 ({}).".format(new_c_from_conduit.Origin, conn2.Origin, el2.Name), "DEBUG")
                        connected_successfully = True
                    except Exception as e:
                        log_message("Erro ao conectar eletroduto ao elemento 2 ({}): {}".format(el2.Name, e), "ERROR")
                        
            if not connected_successfully:
                 log_message("Nenhum conector do eletroduto pôde ser conectado aos elementos originais.", "WARNING")
                 forms.alert("⚠️ O eletroduto foi criado, mas a conexão automática com os elementos pode ter falhado. Verifique as extremidades.", title="Conexão Parcial")

        else:
            forms.alert("❌ Erro: Não foi possível criar o eletroduto principal.", title="Erro de Criação")
            log_message("Falha na criação do eletroduto principal.", "ERROR")
            t.RollBack()
            script.exit()

        t.Commit()
        forms.alert("✅ Eletroduto criado com sucesso e tentativa de conexão realizada!")
        log_message("Eletroduto criado e script finalizado com sucesso.", "INFO")

    except Exception as e:
        if t.Has  and t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
            log_message("Transação revertida devido a um erro.", "INFO")
        error_trace = traceback.format_exc()
        forms.alert("❌ Ocorreu um erro fatal durante a execução do script:\n{}".format(e), title="Erro Fatal")
        log_message("Erro fatal: {}".format(e), "ERROR")
        log_message("Traceback: {}".format(error_trace), "ERROR")
        script.exit()

except Exception as e:
    error_trace = traceback.format_exc()
    forms.alert("❌ Ocorreu um erro inesperado:\n{}".format(e), title="Erro Inesperado")
    log_message("Erro inesperado: {}".format(e), "ERROR")
    log_message("Traceback: {}".format(error_trace), "ERROR")
    script.exit()

finally:
    log_message("Finalizando script Ligar Caixas.", "INFO")