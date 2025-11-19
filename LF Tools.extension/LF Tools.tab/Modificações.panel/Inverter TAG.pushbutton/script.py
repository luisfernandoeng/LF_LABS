# -*- coding: utf-8 -*-
# Script: Inverter_Tags.py (VERSÃO DIAGNÓSTICO)
# Autor: Pyrevit (Luís)
# Objetivo: Rodar um diagnóstico detalhado nas tags selecionadas para descobrir
#           por que o script falha em encontrar hosts editáveis.
# Compatível: pyRevit 4.8+ / IronPython 2.7 / Revit 2024–2025

from pyrevit import revit, DB, script, forms
from System import Guid
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit.forms import alert

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# --- Funções de Formulário ---
def ask_string(prompt, default="", title="Input"):
    """Função para pedir o nome do parâmetro."""
    return forms.ask_for_string(
        default=default,
        prompt=prompt,
        title=title
    )

# --- Funções Auxiliares (Link Element ID Corrigido) ---
def get_tagged_element(tag):
    """
    Obtém o elemento host (no documento atual ou linkado).
    Lógica de LinkElementId aprimorada para IronPython.
    """
    try:
        if hasattr(tag, "GetTaggedElementIds"):
            for element_id_in_list in tag.GetTaggedElementIds():
                
                # Verifica explicitamente o tipo de objeto em IronPython para LinkElementId
                is_link_element_id = element_id_in_list.GetType().Name == "LinkElementId"
                
                if is_link_element_id:
                    # 1. Host em Modelo Linkado
                    link_instance = doc.GetElement(element_id_in_list.LinkInstanceId)
                    if link_instance:
                        linked_doc = link_instance.GetLinkDocument()
                        if linked_doc:
                            linked_elem = linked_doc.GetElement(element_id_in_list.LinkedElementId)
                            if linked_elem:
                                return linked_elem, "Host Linkado"
                
                # 2. Host no Documento Atual (ElementId simples)
                elif element_id_in_list and element_id_in_list.IntegerValue != -1:
                    elem = doc.GetElement(element_id_in_list)
                    if elem:
                        return elem, "Host Local"

        # Fallback para métodos mais antigos
        elif hasattr(tag, "TaggedElementId"):
            eid = tag.TaggedElementId
            if eid and eid.IntegerValue != -1:
                return doc.GetElement(eid), "Host Local (Fallback)"

    except Exception as ex:
        return None, "ERRO API: {}".format(ex)

    return None, "Não Encontrado (Host Deletado?)"

# =================================================================
# FLUXO PRINCIPAL - DIAGNÓSTICO
# =================================================================

# 1) Obter tags selecionadas
selection_ids = uidoc.Selection.GetElementIds()
tags = [doc.GetElement(eid) for eid in selection_ids if doc.GetElement(eid) and isinstance(doc.GetElement(eid), DB.IndependentTag)]

if not tags:
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, 
            "Selecione as tags de identificação (IndependentTag) para diagnóstico. (Esc para cancelar)"
        )
        tags = [doc.GetElement(r.ElementId) for r in refs if isinstance(doc.GetElement(r.ElementId), DB.IndependentTag)]
    except Exception as ex:
        alert('Nenhuma tag selecionada ou seleção cancelada.\nDetalhe: {}'.format(ex))
        raise SystemExit

if not tags:
    alert('Nenhuma tag válida (IndependentTag) selecionada. Encerrando.')
    raise SystemExit

# 2) Perguntar o nome do parâmetro
param_name = ask_string(
    'Digite o nome do parâmetro para diagnóstico (Ex: Mark, Número do circuito):', 
    default='Mark',
    title='Diagnóstico de Parâmetros'
)
if not param_name:
    alert('Parâmetro não informado. Encerrando.')
    raise SystemExit
output.print_md("# --- RELATÓRIO DE DIAGNÓSTICO DO SCRIPT ---")
output.print_md("Parâmetro inspecionado: **{}**".format(param_name))
output.print_md("Tags encontradas: **{}**".format(len(tags)))
output.print_md("---")

# 3) Coletar dados de diagnóstico
debug_report = []
for tag in tags:
    host_data = get_tagged_element(tag)
    host = host_data[0]
    host_status = host_data[1]
    
    report_item = {
        'Tag ID': tag.Id,
        'Host ID': 'N/A',
        'Host Tipo': 'N/A',
        'Host Status': host_status,
        'Parâmetro Existe': 'NÃO',
        'EDITÁVEL (IsReadOnly)': 'N/A',
        'Valor Atual': 'N/A'
    }

    if host:
        report_item['Host ID'] = host.Id
        try:
            report_item['Host Tipo'] = revit.query.get_name(host) # Nome da família/elemento
        except Exception:
            report_item['Host Tipo'] = 'Erro ao obter nome'
        
        p = host.LookupParameter(param_name)
        
        if p:
            report_item['Parâmetro Existe'] = 'SIM'
            
            # PONTO CRÍTICO: Verifica se é somente leitura
            is_read_only = p.IsReadOnly
            report_item['EDITÁVEL (IsReadOnly)'] = 'NÃO' if is_read_only else '**SIM**'
            
            # Tenta ler o valor
            try:
                if p.StorageType == DB.StorageType.String:
                    report_item['Valor Atual'] = p.AsString() or 'Vazio'
                else:
                    report_item['Valor Atual'] = p.AsValueString() or 'Vazio'
            except Exception:
                report_item['Valor Atual'] = 'Erro ao ler valor'
        
    debug_report.append(report_item)

# 4) Gerar Relatório em Tabela
header = ["Tag ID", "Host ID", "Host Tipo", "Host Status", "Parâmetro Existe", "EDITÁVEL (IsReadOnly)", "Valor Atual"]
table_rows = []
for item in debug_report:
    table_rows.append([
        str(item['Tag ID']),
        str(item['Host ID']),
        str(item['Host Tipo']),
        str(item['Host Status']),
        item['Parâmetro Existe'],
        item['EDITÁVEL (IsReadOnly)'],
        item['Valor Atual']
    ])
    
# Formata em Markdown (para melhor visualização no pyRevit)
markdown_table = " | " + " | ".join(header) + " |\n"
markdown_table += " | " + " | ".join(["---"] * len(header)) + " |\n"
for row in table_rows:
    markdown_table += " | " + " | ".join(row) + " |\n"

output.print_md(markdown_table)
output.print_md("\n### **Análise:**")
output.print_md("Para que o script de inversão funcione, a coluna **EDITÁVEL (IsReadOnly)** deve ser **SIM**.")
output.print_md("Se for **NÃO** para todos os itens, você deve selecionar outro Parâmetro.")

# Fim do script, sem tentar transação.
raise SystemExit