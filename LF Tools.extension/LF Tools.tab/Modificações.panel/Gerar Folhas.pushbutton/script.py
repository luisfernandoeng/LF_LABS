# -*- coding: utf-8 -*-
import clr
import datetime
import os
from System.Collections.Generic import List

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import UIApplication
from pyrevit import forms, script

# Acesso ao documento ativo usando pyRevit
uiapp = __revit__  # pyRevit injeta o UIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document if uidoc else None

# Caminho para logs no desktop
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")
LOG_PATH = os.path.join(DESKTOP_PATH, "renumeracao.log")

# ====== Funções Utilitárias ======

def log_error(message):
    """Registra mensagens em log no Desktop (compatível com IronPython)"""
    try:
        with open(LOG_PATH, "a") as log:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log.write(u"[{}] {}\n".format(timestamp, message))
    except Exception as e:
        forms.alert("Erro ao escrever no log: {}".format(str(e)))

def validate_document():
    """Valida se o documento é um projeto válido"""
    if not doc or doc.IsFamilyDocument or doc.IsDetached:
        log_error("Nenhum projeto válido ativo.")
        forms.alert("Abra um projeto RVT válido (não família ou template).")
        return False
    return True

def get_titleblock_type():
    """Retorna o primeiro tipo de carimbo (Title Block) disponível"""
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
    titleblock = collector.FirstElement()
    if not titleblock:
        log_error("Nenhum tipo de carimbo encontrado no projeto.")
        forms.alert("Nenhum tipo de carimbo encontrado. Verifique o template do projeto.")
        return None
    return titleblock

def gerar_codigo_folha(prefixo, disciplina, fase, numero, pavimento, revisao):
    """Função dedicada para montar o campo C-NOME-FOLHA automaticamente"""
    return "{}-{}-{}-{}-{}-R{:02d}".format(
        prefixo, disciplina, fase, str(numero).zfill(3), pavimento, revisao
    )

def set_param(element, param_name, value):
    """Define valor de parâmetro se existir"""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        try:
            param.Set(value)
            return True
        except Exception as e:
            log_error("Erro ao definir {} em {}: {}".format(param_name, element.Name, str(e)))
    return False

# ====== Função para Atualizar Revisão ======
def atualizar_revisao(folha, revisao_atual):
    """Função separada para incrementar e atualizar revisão da folha"""
    if not validate_document():
        return None

    prox_rev = revisao_atual + 1
    hoje = datetime.date.today().strftime("%d/%m/%y")

    t = Transaction(doc, "Atualizar Revisão")
    try:
        t.Start()
        log_error("Iniciando atualização de revisão para folha {}.".format(folha.Name))

        set_param(folha, "C-REVISÃO", str(prox_rev).zfill(2))
        set_param(folha, "C-REVISÃO-R{:02d}".format(prox_rev), str(prox_rev).zfill(2))
        set_param(folha, "C-DATA-R{:02d}".format(prox_rev), hoje)
        set_param(folha, "C-DESCRIÇÃO-R{:02d}".format(prox_rev), "Descrição padrão")  # Pode ser personalizada
        set_param(folha, "C-RESPONSÁVEL-R{:02d}".format(prox_rev), folha.LookupParameter("C-RESPONSÁVEL-R00").AsString())

        # Atualiza C-NOME-FOLHA após revisão
        numero = folha.LookupParameter("C-NÚMERO-DESENHO").AsString() or "001"
        pavimento = folha.LookupParameter("C-PAVIMENTO").AsString() or "TERREO"
        fase = folha.LookupParameter("C-FASE").AsString() or "PE"
        codigo = gerar_codigo_folha("V484", "ELE", fase, numero, pavimento, prox_rev)
        set_param(folha, "C-NOME-FOLHA", codigo)

        t.Commit()
        log_error("Atualização de revisão concluída para folha {}.".format(folha.Name))
        return prox_rev

    except Exception as e:
        log_error("Erro ao atualizar revisão: {}".format(str(e)))
        forms.alert("Erro ao atualizar revisão: {}".format(str(e)))
        return None
    finally:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
            log_error("Transação de atualização de revisão rolada de volta.")

# ====== Função para Criar Folha ======
def criar_folha():
    """Função separada para criar uma nova folha com base no carimbo disponível"""
    if not validate_document():
        return None

    titleblock = get_titleblock_type()
    if not titleblock:
        return None

    # Evita duplicação de número de folha
    existing_numbers = {s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()}
    base_number = 1
    while str(base_number).zfill(3) in existing_numbers:
        base_number += 1
    sheet_number = str(base_number).zfill(3)

    t = Transaction(doc, "Criar Nova Folha")
    try:
        t.Start()
        log_error("Iniciando criação de folha com número {}".format(sheet_number))

        new_sheet = ViewSheet.Create(doc, titleblock.Id)

        # Preenche parâmetros básicos
        hoje = datetime.date.today().strftime("%d/%m/%y")
        set_param(new_sheet, "C-REVISÃO", "00")
        set_param(new_sheet, "C-REVISÃO-R00", "00")
        set_param(new_sheet, "C-DATA-R00", hoje)
        set_param(new_sheet, "C-DESCRIÇÃO-R00", "EMISSÃO INICIAL - PE")
        set_param(new_sheet, "C-RESPONSÁVEL-R00", "Luís")
        set_param(new_sheet, "C-NÚMERO-DESENHO", sheet_number)
        set_param(new_sheet, "C-PROJETISTA", "Luís")
        set_param(new_sheet, "C-PAVIMENTO", "TERREO")
        set_param(new_sheet, "C-FASE", "PE")
        set_param(new_sheet, "C-ENDEREÇO-EMPREENDIMENTO", "Endereço Padrão")

        # Define o número nativo da folha
        new_sheet.SheetNumber = sheet_number

        # Gera código completo usando função separada
        codigo = gerar_codigo_folha("V484", "ELE", "PE", sheet_number, "TERREO", 0)
        set_param(new_sheet, "C-NOME-FOLHA", codigo)

        t.Commit()
        log_error("Criação de folha concluída com sucesso.")
        return new_sheet

    except Exception as e:
        log_error("Erro ao criar folha: {}".format(str(e)))
        forms.alert("Erro ao criar folha: {}".format(str(e)))
        return None
    finally:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
            log_error("Transação de criação de folha rolada de volta devido a erro.")

# ====== Função para Renumerar Folhas ======
def renumerar_folhas(folhas):
    """Função separada para renumerar folhas"""
    if not validate_document():
        return

    if not folhas:
        log_error("Nenhuma folha selecionada para renumeração.")
        forms.alert("Nenhuma folha selecionada.")
        return

    existing_numbers = {s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()}

    ordem_opcoes = ["C-PAVIMENTO", "C-NOME-FOLHA"]
    ordem = forms.SelectFromList.show(ordem_opcoes, title="Escolha o critério de ordenação", button_name="Ordenar")
    if not ordem:
        log_error("Nenhum critério de ordenação selecionado.")
        forms.alert("Nenhum critério de ordenação selecionado.")
        return

    # Ordena as folhas
    if ordem == "C-PAVIMENTO":
        folhas = sorted(folhas, key=lambda x: x.LookupParameter("C-PAVIMENTO").AsString() or "")
    else:
        folhas = sorted(folhas, key=lambda x: x.LookupParameter("C-NOME-FOLHA").AsString() or "")

    t = Transaction(doc, "Renumeração de Folhas")
    try:
        t.Start()
        log_error("Iniciando renumeração de {} folhas.".format(len(folhas)))

        for idx, folha in enumerate(folhas, start=1):
            novo_numero = str(idx).zfill(3)
            
            if novo_numero in existing_numbers and novo_numero != folha.SheetNumber:
                log_error("Número {} já está em uso para outra folha.".format(novo_numero))
                forms.alert("Número {} já está em uso. Pulando folha {}.".format(novo_numero, folha.Name))
                continue
            
            set_param(folha, "C-NÚMERO-DESENHO", novo_numero)
            
            old_number = folha.SheetNumber
            folha.SheetNumber = novo_numero
            existing_numbers.add(novo_numero)
            if old_number in existing_numbers:
                existing_numbers.remove(old_number)
            
            revisao = int(folha.LookupParameter("C-REVISÃO").AsString() or "0")
            codigo = gerar_codigo_folha(
                "V484", "ELE",
                folha.LookupParameter("C-FASE").AsString() or "PE",
                novo_numero,
                folha.LookupParameter("C-PAVIMENTO").AsString() or "TERREO",
                revisao
            )
            set_param(folha, "C-NOME-FOLHA", codigo)

        t.Commit()
        log_error("Renumeração concluída com sucesso.")
        forms.alert("Renumeração concluída com sucesso!")
    
    except Exception as e:
        log_error("Erro durante a renumeração: {}".format(str(e)))
        forms.alert("Erro durante a renumeração: {}".format(str(e)))
    finally:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
            log_error("Transação de renumeração rolada de volta devido a erro.")

# ====== Fluxo Principal ======
def main():
    if not doc:
        log_error("Nenhum documento ativo.")
        forms.alert("Abra um projeto RVT antes de rodar o script.")
        return

    log_error("Projeto ativo: {}. Iniciando script.".format(doc.Title))

    opcoes = ["Criar Nova Folha", "Renumeração de Folhas", "Atualizar Revisão"]
    opcao = forms.SelectFromList.show(opcoes, title="Escolha a Ação", button_name="Executar")

    if not opcao:
        log_error("Nenhuma ação escolhida.")
        return

    if opcao == "Criar Nova Folha":
        nova = criar_folha()
        if nova:
            forms.alert("Folha criada: {}".format(nova.LookupParameter("C-NOME-FOLHA").AsString()))
    elif opcao == "Renumeração de Folhas":
        selected = forms.select_sheets(title="Selecione as folhas")
        renumerar_folhas(selected)
    elif opcao == "Atualizar Revisão":
        selected = forms.select_sheets(title="Selecione as folhas para atualizar revisão")
        for folha in selected:
            revisao_atual = int(folha.LookupParameter("C-REVISÃO").AsString() or "0")
            atualizar_revisao(folha, revisao_atual)

if __name__ == "__main__":
    main()