# -*- coding: utf-8 -*-
from pyrevit import revit, script, forms
from Autodesk.Revit.DB import *
from System.Collections.Generic import List
import os
import datetime

doc = revit.doc
uidoc = revit.uidoc

log_path = os.path.join(os.environ["USERPROFILE"], "Desktop", "relatorio_familias_completo.txt")

def get_safe_param_val(param):
    try:
        if param.StorageType == StorageType.String:
            return param.AsString() or "<vazio>"
        elif param.StorageType == StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            if eid.IntegerValue == -1:
                return "None"
            try:
                el = doc.GetElement(eid)
                return el.Name if el else "ElementId: {}".format(eid.IntegerValue)
            except:
                return "ElementId: {}".format(eid.IntegerValue)
        else:
            return "<desconhecido>"
    except:
        return "<erro>"

def escrever_log(texto):
    with open(log_path, "a") as f:
        f.write(texto + "\n")

def analisar_parametros(elemento, tipo_el=None):
    """Analisa parâmetros de instância e tipo"""
    escrever_log("\n--- PARÂMETROS DE INSTÂNCIA ---")
    
    params_inst = {}
    for p in elemento.Parameters:
        try:
            nome = p.Definition.Name
            valor = get_safe_param_val(p)
            is_shared = p.IsShared
            is_readonly = p.IsReadOnly
            guid = p.GUID if is_shared else None
            
            info = "  {}: {}".format(nome, valor)
            if is_shared:
                info += " [COMPARTILHADO - GUID: {}]".format(guid)
            if is_readonly:
                info += " [SOMENTE LEITURA]"
            
            escrever_log(info)
            params_inst[nome] = valor
        except Exception as ex:
            escrever_log("  <erro ao ler parâmetro: {}>".format(str(ex)))
    
    # Parâmetros do tipo
    if tipo_el:
        escrever_log("\n--- PARÂMETROS DO TIPO ---")
        for p in tipo_el.Parameters:
            try:
                nome = p.Definition.Name
                if nome in params_inst:
                    continue  # Já listado
                
                valor = get_safe_param_val(p)
                is_shared = p.IsShared
                is_readonly = p.IsReadOnly
                guid = p.GUID if is_shared else None
                
                info = "  {}: {}".format(nome, valor)
                if is_shared:
                    info += " [COMPARTILHADO - GUID: {}]".format(guid)
                if is_readonly:
                    info += " [SOMENTE LEITURA]"
                
                escrever_log(info)
            except Exception as ex:
                escrever_log("  <erro ao ler parâmetro do tipo: {}>".format(str(ex)))

def analisar_conectores(elemento):
    """Analisa conectores MEP"""
    escrever_log("\n--- CONECTORES ---")
    
    try:
        # Verifica se é elemento MEP
        if isinstance(elemento, FamilyInstance):
            mep_model = elemento.MEPModel
            if mep_model:
                conn_manager = mep_model.ConnectorManager
                if conn_manager:
                    conectores = conn_manager.Connectors
                    if conectores.Size > 0:
                        for conn in conectores:
                            escrever_log("  Conector ID: {}".format(conn.Id))
                            escrever_log("    Tipo: {}".format(conn.ConnectorType))
                            escrever_log("    Domínio: {}".format(conn.Domain))
                            escrever_log("    Forma: {}".format(conn.Shape if hasattr(conn, 'Shape') else "N/A"))
                            
                            # Dimensões
                            try:
                                if conn.Shape == ConnectorProfileType.Round:
                                    escrever_log("    Diâmetro: {} mm".format(conn.Radius * 2 * 304.8))
                                elif conn.Shape == ConnectorProfileType.Rectangular:
                                    escrever_log("    Largura: {} mm".format(conn.Width * 304.8))
                                    escrever_log("    Altura: {} mm".format(conn.Height * 304.8))
                            except:
                                pass
                            
                            escrever_log("    Conectado: {}".format("Sim" if conn.IsConnected else "Não"))
                            escrever_log("")
                    else:
                        escrever_log("  Nenhum conector encontrado")
                else:
                    escrever_log("  Elemento não possui ConnectorManager")
            else:
                escrever_log("  Não é elemento MEP")
        else:
            escrever_log("  Tipo de elemento não suporta conectores")
    except Exception as ex:
        escrever_log("  <erro ao analisar conectores: {}>".format(str(ex)))

def analisar_familia(familia_instance):
    """Analisa informações da família"""
    escrever_log("\n--- INFORMAÇÕES DA FAMÍLIA ---")
    
    try:
        symbol = familia_instance.Symbol
        familia = symbol.Family
        
        escrever_log("Nome da Família: {}".format(familia.Name))
        escrever_log("ID da Família: {}".format(familia.Id))
        
        # Tipo de família
        if hasattr(familia, 'FamilyCategory'):
            escrever_log("Categoria da Família: {}".format(familia.FamilyCategory.Name if familia.FamilyCategory else "N/A"))
        
        # Verifica se é in-place
        is_inplace = familia.IsInPlace
        escrever_log("Família In-Place: {}".format("Sim" if is_inplace else "Não"))
        
        # Documento da família
        try:
            fam_doc = doc.EditFamily(familia)
            if fam_doc:
                escrever_log("Pode ser editada: Sim")
                fam_doc.Close(False)
        except:
            escrever_log("Pode ser editada: Não (ou família do sistema)")
        
        # Parâmetros compartilhados da família
        escrever_log("\n--- PARÂMETROS COMPARTILHADOS DA FAMÍLIA ---")
        fam_manager = familia.FamilyManager if hasattr(familia, 'FamilyManager') else None
        
        if fam_manager:
            shared_found = False
            for fam_param in fam_manager.Parameters:
                if fam_param.IsShared:
                    shared_found = True
                    escrever_log("  {}: [GUID: {}]".format(fam_param.Definition.Name, fam_param.GUID))
            
            if not shared_found:
                escrever_log("  Nenhum parâmetro compartilhado encontrado na família")
        else:
            escrever_log("  Não foi possível acessar FamilyManager")
        
    except Exception as ex:
        escrever_log("  <erro ao analisar família: {}>".format(str(ex)))

def analisar_geometria(elemento):
    """Analisa informações geométricas"""
    escrever_log("\n--- GEOMETRIA ---")
    
    try:
        options = Options()
        options.DetailLevel = ViewDetailLevel.Fine
        geom = elemento.get_Geometry(options)
        
        if geom:
            solid_count = 0
            face_count = 0
            volume_total = 0
            
            for geom_obj in geom:
                if isinstance(geom_obj, Solid):
                    if geom_obj.Volume > 0:
                        solid_count += 1
                        volume_total += geom_obj.Volume
                        face_count += geom_obj.Faces.Size
                elif isinstance(geom_obj, GeometryInstance):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    for inst_obj in inst_geom:
                        if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                            solid_count += 1
                            volume_total += inst_obj.Volume
                            face_count += inst_obj.Faces.Size
            
            escrever_log("Número de sólidos: {}".format(solid_count))
            escrever_log("Número de faces: {}".format(face_count))
            escrever_log("Volume total: {:.3f} m³".format(volume_total * 0.0283168))  # pés³ para m³
        else:
            escrever_log("Sem geometria disponível")
            
    except Exception as ex:
        escrever_log("  <erro ao analisar geometria: {}>".format(str(ex)))

def analisar_localizacao(elemento):
    """Analisa localização do elemento"""
    escrever_log("\n--- LOCALIZAÇÃO ---")
    
    try:
        location = elemento.Location
        
        if isinstance(location, LocationPoint):
            pt = location.Point
            escrever_log("Tipo de localização: Ponto")
            escrever_log("Coordenadas X: {:.3f} mm".format(pt.X * 304.8))
            escrever_log("Coordenadas Y: {:.3f} mm".format(pt.Y * 304.8))
            escrever_log("Coordenadas Z: {:.3f} mm".format(pt.Z * 304.8))
            
            if hasattr(location, 'Rotation'):
                escrever_log("Rotação: {:.2f}°".format(location.Rotation * 57.2958))  # rad para graus
                
        elif isinstance(location, LocationCurve):
            curve = location.Curve
            escrever_log("Tipo de localização: Curva")
            escrever_log("Comprimento: {:.3f} mm".format(curve.Length * 304.8))
            
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            escrever_log("Ponto inicial: ({:.2f}, {:.2f}, {:.2f})".format(
                start.X * 304.8, start.Y * 304.8, start.Z * 304.8))
            escrever_log("Ponto final: ({:.2f}, {:.2f}, {:.2f})".format(
                end.X * 304.8, end.Y * 304.8, end.Z * 304.8))
        else:
            escrever_log("Tipo de localização: Não disponível")
            
    except Exception as ex:
        escrever_log("  <erro ao analisar localização: {}>".format(str(ex)))

def analisar_workset(elemento):
    """Analisa informações de workset"""
    escrever_log("\n--- WORKSET ---")
    
    try:
        if doc.IsWorkshared:
            workset_id = elemento.WorksetId
            if workset_id != WorksetId.InvalidWorksetId:
                workset = doc.GetWorksetTable().GetWorkset(workset_id)
                escrever_log("Workset: {}".format(workset.Name))
                escrever_log("Workset ID: {}".format(workset_id.IntegerValue))
            else:
                escrever_log("Workset: Não atribuído")
        else:
            escrever_log("Documento não é colaborativo")
    except Exception as ex:
        escrever_log("  <erro ao analisar workset: {}>".format(str(ex)))

def analisar_fase(elemento):
    """Analisa informações de fase"""
    escrever_log("\n--- FASES ---")
    
    try:
        fase_criacao = doc.GetElement(elemento.CreatedPhaseId)
        fase_demolida = doc.GetElement(elemento.DemolishedPhaseId)
        
        escrever_log("Fase de Criação: {}".format(fase_criacao.Name if fase_criacao else "N/A"))
        escrever_log("Fase de Demolição: {}".format(fase_demolida.Name if fase_demolida else "Não demolido"))
    except Exception as ex:
        escrever_log("  <erro ao analisar fases: {}>".format(str(ex)))

# ==================== INÍCIO DO SCRIPT ====================

with open(log_path, "w") as f:
    f.write("=" * 80 + "\n")
    f.write("RELATORIO COMPLETO DE FAMILIAS E ELEMENTOS\n")
    f.write("Data/Hora: {}\n".format(datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
    f.write("Projeto: {}\n".format(doc.Title))
    f.write("=" * 80 + "\n\n")

sel_ids = uidoc.Selection.GetElementIds()
if not sel_ids:
    forms.alert("Selecione um ou mais elementos antes de rodar este script.")
    script.exit()

total = len(sel_ids)
escrever_log("Total de elementos selecionados: {}\n".format(total))
escrever_log("=" * 80 + "\n")

contador = 0
for elid in sel_ids:
    contador += 1
    el = doc.GetElement(elid)
    
    escrever_log("\n" + "=" * 80)
    escrever_log("ELEMENTO {}/{}".format(contador, total))
    escrever_log("=" * 80)
    
    escrever_log("\n--- INFORMAÇÕES BÁSICAS ---")
    escrever_log("ID do Elemento: {}".format(el.Id))
    escrever_log("Classe: {}".format(el.__class__.__name__))
    escrever_log("Categoria: {}".format(el.Category.Name if el.Category else "Sem categoria"))
    escrever_log("Nome: {}".format(el.Name if hasattr(el, 'Name') else "N/A"))
    
    # Tipo
    try:
        if isinstance(el, FamilyInstance):
            tipo = el.Symbol
            escrever_log("Tipo: {}".format(tipo.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
        elif hasattr(el, 'GetTypeId'):
            tipo_id = el.GetTypeId()
            if tipo_id != ElementId.InvalidElementId:
                tipo = doc.GetElement(tipo_id)
                escrever_log("Tipo: {}".format(tipo.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if tipo else "N/A"))
            else:
                tipo = None
                escrever_log("Tipo: N/A")
        else:
            tipo = None
            escrever_log("Tipo: N/A")
    except:
        tipo = None
        escrever_log("Tipo: <erro ao obter tipo>")
    
    # Análises detalhadas
    if isinstance(el, FamilyInstance):
        analisar_familia(el)
    
    analisar_parametros(el, tipo)
    analisar_conectores(el)
    analisar_geometria(el)
    analisar_localizacao(el)
    analisar_workset(el)
    analisar_fase(el)
    
    escrever_log("\n" + "=" * 80 + "\n")

escrever_log("\n\n" + "=" * 80)
escrever_log("FIM DO RELATÓRIO")
escrever_log("=" * 80)

forms.alert("✅ Relatório completo salvo em:\n{}".format(log_path), title="Sucesso")