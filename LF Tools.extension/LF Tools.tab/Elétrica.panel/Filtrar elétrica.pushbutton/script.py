# -*- coding: utf-8 -*-
# Auditoria Elétrica Completa - Múltiplos Painéis

from pyrevit import revit, DB
from System.Collections.Generic import List
import sys
import os
from datetime import datetime

doc = revit.doc
uidoc = revit.uidoc


# ===============================================================
# Função para salvar erros no desktop
# ===============================================================
def salvar_erro(mensagem):
    try:
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo = os.path.join(desktop, "erro_auditoria_eletrica_{}.txt".format(timestamp))
        
        with open(arquivo, 'w') as f:
            f.write("ERRO NA AUDITORIA ELÉTRICA\n")
            f.write("=" * 50 + "\n")
            f.write("Data/Hora: {}\n".format(datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
            f.write("=" * 50 + "\n\n")
            f.write(mensagem)
        
        print("Erro salvo em: {}".format(arquivo))
    except:
        print("Não foi possível salvar o arquivo de erro")


# ===============================================================
# Função para processar um painel (mantém a lógica original)
# ===============================================================
def processar_painel(painel, dispositivos):
    try:
        pn = painel.LookupParameter("Panel Name")
        painel_nome = pn.AsString() if pn else painel.Name
        
        # Coletar TODOS os circuitos para ESTE painel
        todos_circuitos = DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        circuitos_derivados = []
        circuito_alimentador = None
        
        for c in todos_circuitos:
            param = c.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
            if param:
                nome = param.AsString()
                if nome and nome.upper().strip() == painel_nome.upper().strip():
                    circuitos_derivados.append(c)
            
            if c.BaseEquipment and c.BaseEquipment.Id == painel.Id:
                circuito_alimentador = c
        
        # Coletar dispositivos, conduítes e conexões
        elementos_do_sistema = set()
        luminarias_com_comando = []
        
        def adicionar_elementos(circ):
            try:
                for elid in circ.Elements:
                    inst = doc.GetElement(elid)
                    if not inst:
                        continue
                    
                    elementos_do_sistema.add(inst.Id)
                    
                    # Detectar luminárias com ID do comando
                    if inst.Category and inst.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_LightingFixtures):
                        param_cmd = inst.LookupParameter("ID do comando")
                        if param_cmd:
                            valor = param_cmd.AsString()
                            if valor:
                                luminarias_com_comando.append((inst, valor))
                    
                    # Conexões via MEPModel
                    try:
                        if hasattr(inst, "MEPModel") and inst.MEPModel:
                            for conn in inst.MEPModel.ConnectorManager.Connectors:
                                for ref in conn.AllRefs:
                                    ref_el = ref.Owner
                                    if ref_el:
                                        elementos_do_sistema.add(ref_el.Id)
                    except:
                        pass
            except:
                pass
        
        for c in circuitos_derivados:
            adicionar_elementos(c)
        
        if circuito_alimentador:
            adicionar_elementos(circuito_alimentador)
        
        # Incluir o próprio painel
        elementos_do_sistema.add(painel.Id)
        
        # Encontrar interruptores que compartilham o mesmo ID de comando
        interruptores_encontrados = set()
        
        if luminarias_com_comando:
            for lum, comando in luminarias_com_comando:
                for dev in dispositivos:
                    p = dev.LookupParameter("ID do comando")
                    if not p:
                        continue
                    if p.AsString() == comando:
                        interruptores_encontrados.add(dev.Id)
        
        # Criar conjunto de IDs para este painel
        ids_painel = set()
        ids_painel.add(painel.Id)
        
        for c in circuitos_derivados:
            ids_painel.add(c.Id)
        
        if circuito_alimentador:
            ids_painel.add(circuito_alimentador.Id)
        
        ids_painel.update(elementos_do_sistema)
        ids_painel.update(interruptores_encontrados)
        
        return ids_painel, painel_nome
        
    except Exception as e:
        return None, str(e)


# ===============================================================
# MAIN - Processar múltiplos painéis
# ===============================================================
try:
    # Seleção inicial
    sel_ids = uidoc.Selection.GetElementIds()
    if not sel_ids:
        print("Selecione um ou mais painéis.")
        sys.exit()
    
    elements = [doc.GetElement(eid) for eid in sel_ids]
    
    # Filtrar apenas painéis
    paineis = []
    for el in elements:
        if isinstance(el, DB.FamilyInstance):
            if el.Category and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                paineis.append(el)
    
    if not paineis:
        print("Nenhum painel encontrado na seleção.")
        sys.exit()
    
    # Coletar dispositivos uma vez (para verificação de ID do comando)
    dispositivos = DB.FilteredElementCollector(doc)\
        .OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Processar cada painel
    ids_finais = set(sel_ids)  # Começa com a seleção inicial
    paineis_processados = []
    erros = []
    
    for painel in paineis:
        resultado, info = processar_painel(painel, dispositivos)
        
        if resultado is not None:
            ids_finais.update(resultado)
            paineis_processados.append(info)
        else:
            pn = painel.LookupParameter("Panel Name")
            nome = pn.AsString() if pn else painel.Name
            erros.append("Painel {}: {}".format(nome, info))
    
    # Adicionar à seleção (ao invés de substituir)
    if ids_finais:
        ids_selecao = List[DB.ElementId](list(ids_finais))
        uidoc.Selection.SetElementIds(ids_selecao)
    
    # Salvar erros se houver
    if erros:
        mensagem_erro = "\n".join(erros)
        mensagem_erro += "\n\nPainéis processados com sucesso: {}\n".format(len(paineis_processados))
        mensagem_erro += "Painéis com erro: {}\n".format(len(erros))
        salvar_erro(mensagem_erro)

except Exception as e:
    # Salvar erro crítico
    import traceback
    mensagem_completa = "ERRO CRÍTICO:\n\n"
    mensagem_completa += str(e) + "\n\n"
    mensagem_completa += "TRACEBACK:\n"
    mensagem_completa += traceback.format_exc()
    salvar_erro(mensagem_completa)
    print("Erro crítico! Verifique o arquivo de erro no desktop.")