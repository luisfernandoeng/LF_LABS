# -*- coding: utf-8 -*-
# LF Tools - Limpar Familias (Purge) v2.2
# Autor: Luís (com assistencia de Pyrevit)
# Data: 03/09/2025
# Descricao: Realiza um purge avancado em lote para arquivos de familia (.rfa),
#            simulando a ferramenta nativa "Excluir nao utilizados" do Revit.

import clr
import os
import datetime
import codecs

# --- Referencias da API do Revit ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    SaveAsOptions,
    WorksetConfiguration,
    WorksetConfigurationOption,
    ElementId,
    Family,
    Material, FillPatternElement, LinePatternElement, TextNoteType,
    DimensionType, ImportInstance, ImageType, Category
)
from System.Collections.Generic import List

# --- Referencias do pyrevit ---
from pyrevit import forms

# ==================== CONTEXTO REVIT ====================
# Metodo universal e mais compativel para obter os objetos do Revit
# __revit__ e uma variavel especial injetada pelo pyRevit que representa a UIApplication
uiapp = __revit__
app = uiapp.Application
doc = uiapp.ActiveUIDocument.Document


# ==================== CONFIGURACOES ====================
COMPACT_SAVE = True
DESKTOP_PATH = os.path.join(os.path.expanduser('~'), 'Desktop')
LOG_FILE_PATH = os.path.join(DESKTOP_PATH, 'LF_Tools_Purge_Log.txt')

# ==================== FUNCOES AUXILIARES ====================
def write_log(message, indent_level=0):
    """Escreve uma mensagem formatada no arquivo de log."""
    try:
        with codecs.open(LOG_FILE_PATH, 'a', 'utf-8') as log_file:
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            indent = '    ' * indent_level
            log_file.write('[{}] {}{}\n'.format(timestamp, indent, message))
    except Exception as ex:
        print('Erro ao escrever no log: {}'.format(ex))

# ==================== LOGICA PRINCIPAL DE PURGE ====================
def get_purgable_element_ids(doc_to_purge):
    """Coleta os IDs de todos os elementos que podem ser purgados."""
    purgable_ids = List[ElementId]()

    classes_to_purge = [
        Material, FillPatternElement, LinePatternElement, TextNoteType,
        DimensionType, ImportInstance, ImageType
    ]

    for cls in classes_to_purge:
        collector = FilteredElementCollector(doc_to_purge).OfClass(cls)
        for element_id in collector.ToElementIds():
            purgable_ids.Add(element_id)
            
    fam_collector = FilteredElementCollector(doc_to_purge).OfClass(Family)
    for fam in fam_collector:
        for symbol_id in fam.GetFamilySymbolIds():
            purgable_ids.Add(symbol_id)
            
    all_categories = doc_to_purge.Settings.Categories
    for cat in all_categories:
        if cat.SubCategories and cat.SubCategories.Size > 0:
            for sub_cat in cat.SubCategories:
                purgable_ids.Add(sub_cat.Id)

    return purgable_ids

def purge_unused_in_doc(fam_doc):
    """
    Executa rodadas de purga em um documento de familia ate que nada mais
    possa ser deletado. Retorna o total de itens deletados.
    """
    total_deleted_count = 0
    
    while True:
        deleted_in_this_round = 0
        ids_to_try_delete = get_purgable_element_ids(fam_doc)
        
        if not ids_to_try_delete or ids_to_try_delete.Count == 0:
            break

        t = Transaction(fam_doc, 'LF Tools - Rodada de Purge')
        t.Start()
        try:
            deleted_ids = fam_doc.Delete(ids_to_try_delete)
            deleted_in_this_round = deleted_ids.Count
            t.Commit()
        except Exception as e:
            t.RollBack()
            write_log("Erro durante transacao de purga: {}".format(e), 1)
            break

        total_deleted_count += deleted_in_this_round
        
        if deleted_in_this_round == 0:
            break
            
    return total_deleted_count

def process_family_file(rfa_path, output_dir):
    """
    Abre, purga, salva e fecha um unico arquivo de familia.
    Retorna (Sucesso, NomeDoArquivo, ItensDeletados, Mensagem).
    """
    fam_doc = None
    file_name = os.path.basename(rfa_path)
    
    try:
        fam_doc = app.OpenDocumentFile(rfa_path)
        
        if not fam_doc.IsFamilyDocument:
            return (False, file_name, 0, "Nao e um documento de familia.")

        deleted_count = purge_unused_in_doc(fam_doc)

        output_path = os.path.join(output_dir, file_name)
        save_options = SaveAsOptions()
        save_options.OverwriteExistingFile = True
        if COMPACT_SAVE:
            save_options.Compact = True
        
        fam_doc.SaveAs(output_path, save_options)
        
        return (True, file_name, deleted_count, "{} itens deletados".format(deleted_count))

    except Exception as e:
        error_message = str(e)
        return (False, file_name, 0, error_message)
    finally:
        if fam_doc:
            fam_doc.Close(False)

# ==================== SCRIPT PRINCIPAL ====================
def main():
    """Funcao principal que orquestra o processo."""
    source_folder = forms.pick_folder(title="Selecione a Pasta com as Familias .rfa")
    if not source_folder:
        return

    output_folder = os.path.join(source_folder, "familias_limpas")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    rfa_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.rfa')]
    if not rfa_files:
        forms.alert("Nenhum arquivo .rfa foi encontrado na pasta selecionada.", title="Aviso")
        return

    write_log("="*50)
    write_log("INICIO DA EXECUCAO - Limpeza de Familias v2.2")
    write_log("Pasta de Origem: {}".format(source_folder))
    write_log("Total de arquivos para processar: {}".format(len(rfa_files)))
    write_log("="*50)

    success_list, failure_list = [], []

    with forms.ProgressBar(title='Limpando Famílias ({total} total)...', cancellable=True) as pb:
        for i, file_name in enumerate(rfa_files):
            if pb.cancelled:
                write_log("Operacao cancelada pelo usuario.")
                break
            
            pb.update_progress(i, len(rfa_files))
            pb.title = 'Limpando Familias... ({}/{}) -> {}'.format(i + 1, len(rfa_files), file_name)

            full_path = os.path.join(source_folder, file_name)
            
            success, name, count, message = process_family_file(full_path, output_folder)

            if success:
                msg = "SUCESSO: {} ({}).".format(name, message)
                success_list.append(msg)
                write_log(msg)
            else:
                msg = "FALHA: {} -> {}".format(name, message)
                failure_list.append(msg)
                write_log(msg)

    
    summary_message = (
        u"Processo Finalizado!\n\n"
        u"Resultados salvos em:\n{}\n\n"
        u"{} famílias processadas com SUCESSO.\n"
        u"{} famílias FALHARAM.\n\n"
        u"Um log detalhado foi salvo em sua Área de Trabalho:\n{}"
    ).format(output_folder, len(success_list), len(failure_list), LOG_FILE_PATH)
    
    forms.alert(summary_message, title="LF Tools - Resumo da Limpeza")

if __name__ == '__main__':
    main()