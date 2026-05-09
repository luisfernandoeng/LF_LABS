# -*- coding: utf-8 -*-
from pyrevit import revit, script, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from System.Collections.Generic import List
import os
import re
import datetime

doc = revit.doc
uidoc = revit.uidoc

# Contexto de análise — trocado para link_doc quando inspecionando vínculo
analysis_doc = doc
link_context  = {"is_link": False, "name": None, "transform": None}

import System
desktop_path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
log_path = os.path.join(desktop_path, "relatorio_familias_completo.txt")

# Buffer para evitar I/O excessivo (performance)
log_buffer = []

def limpar_log_antigo():
    """Remove o relatório anterior antes de qualquer análise."""
    try:
        if os.path.exists(log_path):
            os.remove(log_path)
    except Exception as ex:
        forms.alert(
            u"Não foi possível apagar o relatório antigo:\n{}\n\n{}"
            .format(log_path, ex),
            title="Inspecionar Tipo"
        )
        script.exit()

def param_tem_valor_util(param):
    """Evita despejar parâmetros vazios/sem valor no relatório."""
    try:
        if not param or not param.Definition:
            return False
        nome = param.Definition.Name
        if not nome:
            return False
        try:
            if not param.HasValue:
                return False
        except:
            pass

        if param.StorageType == StorageType.String:
            valor = param.AsString()
            return bool(valor and str(valor).strip())
        if param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            return bool(eid and eid.IntegerValue != -1)
        return True
    except:
        return False

def formatar_parametro(param):
    nome = param.Definition.Name
    valor = get_safe_param_val(param)
    info = "  {}: {}".format(nome, valor)
    if param.IsShared:
        info += " [COMPARTILHADO - GUID: {}]".format(param.GUID)
    if param.IsReadOnly:
        info += " [SOMENTE LEITURA]"
    return info

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
                el = analysis_doc.GetElement(eid)
                return el.Name if el else "ElementId: {}".format(eid.IntegerValue)
            except:
                return "ElementId: {}".format(eid.IntegerValue)
        else:
            return "<desconhecido>"
    except:
        return "<erro>"

def get_param_val_display(param):
    """Retorna valor legível (AsValueString primeiro, depois raw)."""
    try:
        vs = param.AsValueString()
        if vs:
            return vs
    except:
        pass
    return str(get_safe_param_val(param))

def escrever_log(texto):
    log_buffer.append(texto)

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
    """Analisa informações da família (otimizado - sem abrir documento)"""
    escrever_log("\n--- INFORMAÇÕES DA FAMÍLIA ---")
    
    try:
        symbol = familia_instance.Symbol
        familia = symbol.Family
        
        escrever_log("Nome da Família: {}".format(familia.Name))
        escrever_log("ID da Família: {}".format(familia.Id))
        
        if hasattr(familia, 'FamilyCategory'):
            escrever_log("Categoria da Família: {}".format(familia.FamilyCategory.Name if familia.FamilyCategory else "N/A"))
        
        is_inplace = familia.IsInPlace
        escrever_log("Família In-Place: {}".format("Sim" if is_inplace else "Não"))
        escrever_log("Editável: {}".format("Não (família do sistema)" if familia.IsEditable == False else "Sim"))
        
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

            # Se for elemento de vínculo, exibe também as coordenadas no sistema do host
            if link_context["is_link"] and link_context["transform"]:
                try:
                    pt_host = link_context["transform"].OfPoint(pt)
                    escrever_log("--- Coordenadas no Host (após transform) ---")
                    escrever_log("  X (host): {:.3f} mm".format(pt_host.X * 304.8))
                    escrever_log("  Y (host): {:.3f} mm".format(pt_host.Y * 304.8))
                    escrever_log("  Z (host): {:.3f} mm".format(pt_host.Z * 304.8))
                except:
                    pass

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
        if analysis_doc.IsWorkshared:
            workset_id = elemento.WorksetId
            if workset_id != WorksetId.InvalidWorksetId:
                workset = analysis_doc.GetWorksetTable().GetWorkset(workset_id)
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
        fase_criacao = analysis_doc.GetElement(elemento.CreatedPhaseId)
        fase_demolida = analysis_doc.GetElement(elemento.DemolishedPhaseId)
        
        escrever_log("Fase de Criação: {}".format(fase_criacao.Name if fase_criacao else "N/A"))
        escrever_log("Fase de Demolição: {}".format(fase_demolida.Name if fase_demolida else "Não demolido"))
    except Exception as ex:
        escrever_log("  <erro ao analisar fases: {}>".format(str(ex)))


# ==================== SCHEDULE (TABELA DE QUANTIDADES) ====================

def analisar_schedule(schedule):
    """Analisa ViewSchedule: campos, filtros, ordenação e dados."""
    escrever_log("\n--- INFORMAÇÕES DA SCHEDULE ---")
    escrever_log("Nome: {}".format(schedule.Name))
    escrever_log("ID: {}".format(schedule.Id.IntegerValue))
    escrever_log("Classe: {}".format(schedule.__class__.__name__))

    try:
        escrever_log("É Template: {}".format("Sim" if schedule.IsTemplate else "Não"))
    except:
        pass

    schedule_def = None
    try:
        schedule_def = schedule.Definition
    except:
        escrever_log("  <erro ao obter Definition da Schedule>")
        return

    # --- Categoria ---
    try:
        cat_id = schedule_def.CategoryId
        if cat_id and cat_id.IntegerValue != -1:
            cat = Category.GetCategory(doc, cat_id)
            escrever_log("Categoria: {} (ID: {})".format(
                cat.Name if cat else "?", cat_id.IntegerValue))
        else:
            escrever_log("Categoria: Multicategoria")
    except:
        escrever_log("Categoria: <erro>")

    # --- Campos (Fields) ---
    escrever_log("\n--- CAMPOS (FIELDS) ---")
    try:
        field_count = schedule_def.GetFieldCount()
        escrever_log("Total de campos: {}".format(field_count))

        for i in range(field_count):
            field = schedule_def.GetField(i)
            fname = "(sem nome)"
            try:
                fname = field.GetName()
            except:
                pass

            hidden = ""
            try:
                if field.IsHidden:
                    hidden = " [OCULTO]"
            except:
                pass

            field_type = ""
            try:
                ft = field.FieldType
                field_type = " ({})".format(str(ft).replace("ScheduleFieldType.", ""))
            except:
                pass

            param_id_str = ""
            try:
                pid = field.ParameterId
                if pid and pid.IntegerValue > 0:
                    param_el = doc.GetElement(pid)
                    if param_el:
                        param_id_str = " [SharedParam: {}]".format(param_el.Name)
                elif pid and pid.IntegerValue < 0:
                    # BuiltInParameter
                    try:
                        bip = System.Enum.ToObject(BuiltInParameter, pid.IntegerValue)
                        param_id_str = " [BuiltIn: {}]".format(bip)
                    except:
                        param_id_str = " [BuiltIn ID: {}]".format(pid.IntegerValue)
            except:
                pass

            escrever_log("  [{}] {}{}{}{}".format(i, fname, field_type, hidden, param_id_str))
    except Exception as ex:
        escrever_log("  <erro ao ler campos: {}>".format(str(ex)))

    # --- Filtros ---
    escrever_log("\n--- FILTROS ---")
    try:
        filter_count = schedule_def.GetFilterCount()
        escrever_log("Total de filtros: {}".format(filter_count))
        for i in range(filter_count):
            sf = schedule_def.GetFilter(i)
            field_idx = sf.FieldId
            fname = "?"
            try:
                field = schedule_def.GetField(
                    schedule_def.GetFieldIndex(field_idx))
                fname = field.GetName()
            except:
                pass
            escrever_log("  Filtro {}: Campo='{}' Tipo={} Valor='{}'".format(
                i, fname,
                sf.FilterType,
                sf.GetStringValue() if hasattr(sf, 'GetStringValue') else "?"
            ))
    except Exception as ex:
        escrever_log("  <erro ao ler filtros: {}>".format(str(ex)))

    # --- Ordenação ---
    escrever_log("\n--- ORDENAÇÃO / AGRUPAMENTO ---")
    try:
        sort_count = schedule_def.GetSortGroupFieldCount()
        escrever_log("Campos de ordenação: {}".format(sort_count))
        for i in range(sort_count):
            sg = schedule_def.GetSortGroupField(i)
            field = schedule_def.GetField(
                schedule_def.GetFieldIndex(sg.FieldId))
            order = "Crescente" if sg.SortOrder == ScheduleSortOrder.Ascending else "Decrescente"
            escrever_log("  [{}] {} ({}){}".format(
                i, field.GetName(), order,
                " [AGRUPAR]" if sg.ShowHeader else ""))
    except Exception as ex:
        escrever_log("  <erro ao ler ordenação: {}>".format(str(ex)))

    # --- Dados (amostra de até 20 linhas) ---
    escrever_log("\n--- DADOS DA TABELA (amostra) ---")
    try:
        table_data = schedule.GetTableData()
        body = table_data.GetSectionData(SectionType.Body)
        num_rows = body.NumberOfRows
        num_cols = body.NumberOfColumns
        escrever_log("Linhas: {} | Colunas: {}".format(num_rows, num_cols))

        # Cabeçalhos (Header section)
        try:
            header_section = table_data.GetSectionData(SectionType.Header)
            header_rows = header_section.NumberOfRows
            if header_rows > 0:
                headers = []
                for c in range(header_section.NumberOfColumns):
                    try:
                        txt = schedule.GetCellText(SectionType.Header, 0, c)
                        headers.append(txt or "")
                    except:
                        headers.append("")
                if any(headers):
                    escrever_log("Cabeçalho: {}".format(" | ".join(headers)))
        except:
            pass

        max_rows = min(num_rows, 20)
        for r in range(max_rows):
            row_data = []
            for c in range(num_cols):
                try:
                    txt = schedule.GetCellText(SectionType.Body, r, c)
                    row_data.append(txt or "")
                except:
                    row_data.append("")
            escrever_log("  Linha {}: {}".format(r, " | ".join(row_data)))

        if num_rows > 20:
            escrever_log("  ... ({} linhas restantes omitidas)".format(
                num_rows - 20))

    except Exception as ex:
        escrever_log("  <erro ao ler dados: {}>".format(str(ex)))


# ==================== PANEL SCHEDULE (QUADRO DE CARGAS) ====================

def analisar_configuracoes_eletricas_doc():
    """Dumps ElectricalSetting do projeto: tensões e parâmetros relevantes."""
    escrever_log("\n--- CONFIGURAÇÕES ELÉTRICAS DO PROJETO (ElectricalSetting) ---")
    try:
        es = ElectricalSetting.GetElectricalSettings(doc)
        if not es:
            escrever_log("  ElectricalSetting não disponível neste documento.")
            return

        try:
            voltages = list(es.GetAllVoltageTypes())
            escrever_log("Tipos de Tensão definidos no projeto: {}".format(len(voltages)))
            for v in voltages:
                try:
                    escrever_log("  - {} (ID: {})".format(
                        v.Name if hasattr(v, 'Name') else "?", v.Id.IntegerValue))
                except:
                    pass
        except Exception as ex:
            escrever_log("  Tipos de Tensão: <erro: {}>".format(ex))

        try:
            demand_tables = list(es.GetDemandFactorTables())
            escrever_log("Tabelas de Fator de Demanda: {}".format(len(demand_tables)))
            for dt in demand_tables:
                try:
                    escrever_log("  - {}".format(dt.Name))
                except:
                    pass
        except:
            pass

        escrever_log("\nParâmetros do ElectricalSetting:")
        for p in es.Parameters:
            try:
                nome = p.Definition.Name
                if not nome:
                    continue
                valor = get_param_val_display(p)
                ro = " [RO]" if p.IsReadOnly else ""
                escrever_log("  {}: {}{}".format(nome, valor, ro))
            except:
                continue

    except Exception as ex:
        escrever_log("  <erro ao obter ElectricalSetting: {}>".format(ex))


def analisar_sistema_distribuicao_painel(panel):
    """Analisa o DistributionSysType do painel (tensão, fases, configuração)."""
    escrever_log("\n--- SISTEMA DE DISTRIBUIÇÃO DO PAINEL ---")
    try:
        dist_type_id = None

        # BIP varia por versão — busca por nome de parâmetro
        for nome_p in ["Sistema de distribuição", "Distribution System"]:
            p = panel.LookupParameter(nome_p)
            if p and p.StorageType == StorageType.ElementId:
                eid = p.AsElementId()
                if eid and eid != ElementId.InvalidElementId:
                    dist_type_id = eid
                    break

        # Fallback: varre params buscando ElementId de DistributionSysType
        if not dist_type_id:
            for p in panel.Parameters:
                try:
                    if p.StorageType != StorageType.ElementId:
                        continue
                    eid = p.AsElementId()
                    if not eid or eid == ElementId.InvalidElementId:
                        continue
                    el_cand = doc.GetElement(eid)
                    if el_cand and "DistributionSysType" in el_cand.__class__.__name__:
                        dist_type_id = eid
                        escrever_log("  Encontrado via parâmetro: {}".format(p.Definition.Name))
                        break
                except:
                    continue

        if not dist_type_id:
            escrever_log("  Sistema de distribuição não encontrado.")
            return

        dist_el = doc.GetElement(dist_type_id)
        if not dist_el:
            escrever_log("  ElementId {} não resolvido.".format(dist_type_id.IntegerValue))
            return

        escrever_log("Sistema de Distribuição: {} (ID: {})".format(
            dist_el.Name if hasattr(dist_el, 'Name') else "?",
            dist_el.Id.IntegerValue))
        escrever_log("Classe: {}".format(dist_el.__class__.__name__))

        for p in dist_el.Parameters:
            try:
                nome = p.Definition.Name
                if not nome:
                    continue
                valor = get_param_val_display(p)
                ro = " [RO]" if p.IsReadOnly else ""
                escrever_log("  {}: {}{}".format(nome, valor, ro))
            except:
                continue

    except Exception as ex:
        escrever_log("  <erro ao analisar sistema de distribuição: {}>".format(ex))


def diagnosticar_protecao_circuitos(systems):
    """Verifica por que RBS_ELEC_CIRCUIT_RATING_PARAM pode estar travado.

    Mostra para cada circuito: status RO/WRITE, carga calculada, rating exibido
    e o valor bruto em Amperes para facilitar diagnóstico do valor fixo em 16A.
    """
    escrever_log("\n--- DIAGNÓSTICO: PROTEÇÃO DO CIRCUITO (por que pode estar travada) ---")
    escrever_log("Legenda: [RO]=somente leitura (calculado pelo Revit) | [WRITE]=editável")

    bip_rating  = BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM
    bip_load    = BuiltInParameter.RBS_ELEC_APPARENT_LOAD
    bip_voltage = BuiltInParameter.RBS_ELEC_VOLTAGE
    bip_poles   = BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES
    bip_name    = BuiltInParameter.RBS_ELEC_CIRCUIT_NAME

    for sys in systems:
        cnum = sys.CircuitNumber or "?"
        try:
            p_rating = sys.get_Parameter(bip_rating)
            p_poles  = sys.get_Parameter(bip_poles)
            p_name   = sys.get_Parameter(bip_name)

            rating_val    = get_param_val_display(p_rating) if p_rating else "N/A"
            rating_status = "[RO]" if (p_rating and p_rating.IsReadOnly) else "[WRITE]"
            poles_val     = p_poles.AsInteger() if p_poles else 1
            nome_carga    = p_name.AsString()   if p_name  else "?"

            # Corrente aparente já em Amperes (AsDouble() = valor real em A)
            corrente_str = "?"
            corrente_val = 0.0
            for nome_c in ["Corrente aparente", "Apparent Current"]:
                p_c = sys.LookupParameter(nome_c)
                if p_c:
                    corrente_val = p_c.AsDouble()
                    corrente_str = p_c.AsValueString() or "{:.1f} A".format(corrente_val)
                    break

            # Carga e tensão — display value (já convertido)
            carga_str   = get_param_val_display(sys.get_Parameter(bip_load))    if sys.get_Parameter(bip_load)    else "?"
            tensao_str  = get_param_val_display(sys.get_Parameter(bip_voltage)) if sys.get_Parameter(bip_voltage) else "?"

            escrever_log("\n  Circuito {:>4} | Carga: {} | Tensão: {} | Pólos: {} | "
                         "Corrente: {} | Rating: {} {}".format(
                         cnum, carga_str, tensao_str, poles_val,
                         corrente_str, rating_val, rating_status))
            escrever_log("           Nome: {} | ID: {}".format(nome_carga, sys.Id.IntegerValue))

            # Alerta se proteção < corrente (dado inconsistente)
            try:
                rating_num = float(str(rating_val).replace(" A", "").replace(",", ".").strip())
                if corrente_val > 0 and rating_num < corrente_val * 0.9:
                    escrever_log("           >> ALERTA: Proteção ({}) menor que corrente ({}) — dado inconsistente!".format(
                        rating_val, corrente_str))
            except:
                pass

            for bip_extra_name, bip_extra_str in [
                ('Caixa da proteção (Frame)', 'RBS_ELEC_CIRCUIT_FRAME_PARAM'),
            ]:
                try:
                    bip_extra = getattr(BuiltInParameter, bip_extra_str, None)
                    if bip_extra is not None:
                        p_extra = sys.get_Parameter(bip_extra)
                        if p_extra:
                            escrever_log("           {}: {} {}".format(
                                bip_extra_name,
                                get_param_val_display(p_extra),
                                "[RO]" if p_extra.IsReadOnly else "[WRITE]"))
                except:
                    pass

        except Exception as ex:
            escrever_log("  Circuito {}: <erro: {}>".format(cnum, ex))

    escrever_log("\n>> CONCLUSÃO RÁPIDA:")
    escrever_log("   [WRITE] = parâmetro editável via API, mas a TABELA pode bloquear a célula")
    escrever_log("   pelo template (Gerenciar → Templates de Quadro de Cargas).")
    escrever_log("   A verificação de editabilidade de células aparece na seção")
    escrever_log("   'EDITABILIDADE DAS CÉLULAS' acima — procure a coluna 'In' ou 'Proteção'.")


def analisar_panel_schedule(panel_schedule_view):
    """Analisa PanelScheduleView: painel, circuitos e parâmetros editáveis."""
    escrever_log("\n--- INFORMAÇÕES DO QUADRO DE CARGAS ---")
    escrever_log("Nome da Tabela: {}".format(panel_schedule_view.Name))
    escrever_log("ID: {}".format(panel_schedule_view.Id.IntegerValue))
    escrever_log("Classe: {}".format(panel_schedule_view.__class__.__name__))

    # --- Configuração de Template ---
    try:
        from Autodesk.Revit.DB.Electrical import PanelScheduleTemplate
        template_id = panel_schedule_view.GetTemplate()
        if template_id and template_id.IntegerValue > 0:
            template = doc.GetElement(template_id)
            escrever_log("Template Associado: {} (ID: {})".format(template.Name, template.Id.IntegerValue))
        else:
            escrever_log("Template Associado: Nenhum / Personalizado")
    except Exception as ex:
        pass

    # --- Painel associado ---
    panel = None
    panel_id = None
    try:
        panel_id = panel_schedule_view.GetPanel()
    except:
        escrever_log("  <erro ao obter GetPanel()>")

    if panel_id and panel_id != ElementId.InvalidElementId:
        panel = doc.GetElement(panel_id)
    
    if not panel:
        escrever_log("Painel associado: NÃO ENCONTRADO")
        return

    panel_name = panel.Name
    for n in ["Nome do painel", "Panel Name", "Mark"]:
        p = panel.LookupParameter(n)
        if p and p.HasValue:
            panel_name = p.AsString()
            break

    escrever_log("Painel: {} (ID: {})".format(panel_name, panel.Id.IntegerValue))
    
    # --- Parâmetros de Instância e Tipo Específicos de Elétrica ---
    escrever_log("\n--- CONFIGURAÇÕES DE TIPO DO PAINEL (Crucial para Tabelas) ---")
    panel_type = doc.GetElement(panel.GetTypeId())
    if panel_type:
        try:
            nome_tipo = panel_type.Name
        except:
            p_tipo = panel_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            nome_tipo = p_tipo.AsString() if p_tipo else "N/A"
        escrever_log("Tipo: {}".format(nome_tipo))
        # Verificar Part Type (Tipo de Peça)
        part_type = panel_type.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
        if part_type:
            escrever_log("  Part Type (Comportamento): {}".format(part_type.AsValueString() or part_type.AsInteger()))
        
        # Verificar Max Poles e Panel Configuration
        max_poles = panel_type.get_Parameter(BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS)
        if max_poles:
            escrever_log("  Máx Pólos (Disjuntores): {}".format(max_poles.AsInteger()))
            
        for p in panel_type.Parameters:
            if p.IsShared or "Panel" in p.Definition.Name or "Schedule" in p.Definition.Name:
                escrever_log("  {}: {}".format(p.Definition.Name, get_param_val_display(p)))
    
    escrever_log("\n--- PARÂMETROS DE INSTÂNCIA DO PAINEL ---")
    # Mostrar configs especificas se encontrar
    try:
        panel_config = panel.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_CONFIGURATION)
        if panel_config:
            escrever_log(">> CONFIGURAÇÃO DO PAINEL (Colunas): {}".format(panel_config.AsValueString()))
    except AttributeError:
        pass
        
    for p in panel.Parameters:
        try:
            nome = p.Definition.Name
            valor = get_param_val_display(p)
            ro = " [RO]" if p.IsReadOnly else ""
            shared = " [COMPARTILHADO]" if p.IsShared else ""
            escrever_log("  {}: {}{}{}".format(nome, valor, ro, shared))
        except:
            continue

    # --- Circuitos ---
    escrever_log("\n--- CIRCUITOS CONECTADOS ---")
    systems = []
    try:
        if hasattr(panel, "MEPModel") and panel.MEPModel:
            systems = list(panel.MEPModel.GetAssignedElectricalSystems())
        else:
            systems = list(panel.GetAssignedElectricalSystems())
    except Exception as ex:
        escrever_log("  <erro ao obter circuitos: {}>".format(str(ex)))
        return

    def sort_key(s):
        try:
            digits = re.sub(r'\D', '', s.CircuitNumber or '')
            return int(digits) if digits else 9999
        except:
            return 9999
    systems.sort(key=sort_key)
    escrever_log("Total de circuitos válidos: {}".format(len(systems)))

    if not systems:
        escrever_log("  (nenhum circuito encontrado)")
        return

    # Dump detalhado de todos os circuitos
    escrever_log("\n--- DUMP DETALHADO DE TODOS OS CIRCUITOS ---")
    for sys in systems:
        cnum = sys.CircuitNumber or "?"
        escrever_log("\n>> CIRCUITO: {} (ID: {})".format(cnum, sys.Id.IntegerValue))
        
        # Parâmetros vitais Built-in primeiro
        bips = {
            'Nome (CIRCUIT_NAME)': BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
            'Tensão (VOLTAGE)': BuiltInParameter.RBS_ELEC_VOLTAGE,
            'Carga App (APPARENT_LOAD)': BuiltInParameter.RBS_ELEC_APPARENT_LOAD,
            'Polos (NUMBER_OF_POLES)': BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
            'Fio (WIRE_SIZE)': BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM,
        }
        for label, bip in bips.items():
            p = sys.get_Parameter(bip)
            if p: escrever_log("  - {}: {}".format(label, get_param_val_display(p)))
            
        # Demais parâmetros de instância
        for p in sys.Parameters:
            nome = p.Definition.Name
            # ignorar os que já mosrtamos ou nomes em branco
            if not nome: continue
            is_shared = " [COMPARTILHADO]" if p.IsShared else ""
            ro = " [RO]" if p.IsReadOnly else ""
            escrever_log("  - {}: {}{}{}".format(nome, get_param_val_display(p), ro, is_shared))

    # Seções da tabela do panel schedule (Completo)
    escrever_log("\n--- TABELA VISUAL (BODY DUMP COMPLETO) ---")
    try:
        table_data = panel_schedule_view.GetTableData()
        body = table_data.GetSectionData(SectionType.Body)
        nr = body.NumberOfRows
        nc = body.NumberOfColumns
        escrever_log("  Seção de Circuitos: {} linhas × {} colunas (Isso define se tem 1 ou 2 colunas visuais)".format(nr, nc))

        for r in range(nr):
            cells = []
            for c in range(nc):
                try:
                    txt = panel_schedule_view.GetCellText(SectionType.Body, r, c)
                    cells.append(txt or "_")
                except:
                    cells.append("_")
            escrever_log("    Linha {}: {}".format(r, " | ".join(cells)))
            
    except Exception as ex:
        escrever_log("  <erro ao ler estrutura visual da tabela: {}>".format(str(ex)))

    # --- Editabilidade das células da linha de circuito ---
    escrever_log("\n--- EDITABILIDADE DAS CÉLULAS (linha do 1º circuito) ---")
    escrever_log("Legenda: [EDIT]=clicável pelo usuário | [LOCKED]=bloqueado pelo template | [?]=API não suporta")
    try:
        _td = panel_schedule_view.GetTableData()
        _body = _td.GetSectionData(SectionType.Body)
        _nr = _body.NumberOfRows
        _nc = _body.NumberOfColumns

        # Cabeçalho das colunas (linha 1 normalmente)
        headers = []
        for _c in range(_nc):
            try:
                headers.append(panel_schedule_view.GetCellText(SectionType.Body, 1, _c) or "")
            except:
                headers.append("")

        # Primeira linha real de circuito
        for _r in range(_nr):
            _txt0 = panel_schedule_view.GetCellText(SectionType.Body, _r, 1) or ""
            if _txt0.strip() and _txt0.strip()[0].isdigit():
                escrever_log("Linha {} (Circuito {}):".format(_r, _txt0.strip()))
                for _c in range(_nc):
                    cell_txt = panel_schedule_view.GetCellText(SectionType.Body, _r, _c) or "_"
                    header   = headers[_c] if _c < len(headers) else ""

                    edit_str = "[?]"
                    try:
                        is_edit  = panel_schedule_view.IsCellEditable(SectionType.Body, _r, _c)
                        edit_str = "[EDIT]" if is_edit else "[LOCKED]"
                    except:
                        pass

                    param_str = ""
                    try:
                        pid = panel_schedule_view.GetCellParamId(SectionType.Body, _r, _c)
                        if pid and pid != ElementId.InvalidElementId:
                            try:
                                bip = System.Enum.ToObject(BuiltInParameter, pid.IntegerValue)
                                param_str = "-> BIP: {}".format(bip)
                            except:
                                param_str = "-> ParamID: {}".format(pid.IntegerValue)
                    except:
                        pass

                    escrever_log("  Col {:2} [{:20}] {:20} {} {}".format(
                        _c, header[:20], repr(cell_txt[:18]), edit_str, param_str))
                break
    except Exception as ex:
        escrever_log("  <erro ao verificar editabilidade: {}>".format(ex))

    # --- Diagnósticos focados no bloqueio da Proteção do Circuito ---
    analisar_sistema_distribuicao_painel(panel)
    analisar_configuracoes_eletricas_doc()
    if systems:
        diagnosticar_protecao_circuitos(systems)


# ==================== CIRCUITOS ELÉTRICOS ====================

def analisar_circuito(circuito):
    escrever_log("\n--- INFORMAÇÕES DO CIRCUITO ---")
    try:
        escrever_log("Nome da Carga: {}".format(circuito.LoadName or "?"))
        escrever_log("Painel: {}".format(circuito.PanelName or "?"))
        escrever_log("Polos: {}".format(circuito.PolesNumber))
        # Tensão e Carga via BIP para evitar unidades internas erradas
        p_volt = circuito.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
        p_load = circuito.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD)
        escrever_log("Tensão: {}".format(get_param_val_display(p_volt) if p_volt else "N/A"))
        escrever_log("Carga Aparente: {}".format(get_param_val_display(p_load) if p_load else "N/A"))
        # Corrente real
        for nome_c in ["Corrente aparente", "Apparent Current"]:
            p_c = circuito.LookupParameter(nome_c)
            if p_c:
                escrever_log("Corrente Aparente: {}".format(p_c.AsValueString() or "N/A"))
                break
        escrever_log("Rating (Proteção): {} A".format(circuito.Rating))
        comprimento_ft = circuito.Length
        escrever_log("Comprimento: {:.3f} m".format(comprimento_ft * 0.3048))
        if hasattr(circuito, 'WireSizeString'):
            escrever_log("Fiação: {}".format(circuito.WireSizeString))
        try:
            escrever_log("Classificação de Carga: {}".format(circuito.LoadClassification.Name))
        except:
            pass
    except Exception as ex:
        escrever_log("  <erro: {}>".format(str(ex)))

    escrever_log("\n--- BUILTINPARAMETERS ELÉTRICOS (RO/WRITE) ---")
    bips = [
        'RBS_ELEC_CIRCUIT_LENGTH', 'RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM',
        'RBS_ELEC_CIRCUIT_RATING_PARAM', 'RBS_ELEC_VOLTAGE_DROP_PARAM',
        'RBS_ELEC_VOLTAGE', 'RBS_ELEC_APPARENT_LOAD', 'RBS_ELEC_CIRCUIT_NAME'
    ]
    for bip_name in bips:
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is not None:
                p = circuito.get_Parameter(bip)
                if p:
                    valor = get_param_val_display(p)
                    ro = " [RO]" if p.IsReadOnly else " [WRITE]"
                    escrever_log("  {:35} {} {}".format(bip_name, ro, valor))
        except:
            pass

# ==================== DETECÇÃO DE TIPO DE ELEMENTO ====================

def analisar_parametros(elemento, tipo_el=None):
    """Versao enxuta: mostra somente parametros com valor util."""
    linhas_inst = []
    params_inst = set()

    try:
        parametros = elemento.Parameters
    except:
        parametros = []

    for p in parametros:
        try:
            nome = p.Definition.Name
            if nome:
                params_inst.add(nome)
            if param_tem_valor_util(p):
                linhas_inst.append(formatar_parametro(p))
        except Exception as ex:
            linhas_inst.append("  <erro ao ler parametro: {}>".format(str(ex)))

    if linhas_inst:
        escrever_log("\n--- PARAMETROS DE INSTANCIA (com valor) ---")
        for linha in linhas_inst:
            escrever_log(linha)

    if not tipo_el:
        return

    linhas_tipo = []
    try:
        parametros_tipo = tipo_el.Parameters
    except:
        parametros_tipo = []

    for p in parametros_tipo:
        try:
            nome = p.Definition.Name
            if nome in params_inst:
                continue
            if param_tem_valor_util(p):
                linhas_tipo.append(formatar_parametro(p))
        except Exception as ex:
            linhas_tipo.append("  <erro ao ler parametro do tipo: {}>".format(str(ex)))

    if linhas_tipo:
        escrever_log("\n--- PARAMETROS DO TIPO (com valor) ---")
        for linha in linhas_tipo:
            escrever_log(linha)


def analisar_conectores(elemento):
    """Versao enxuta: so emite a secao quando ha conectores."""
    linhas = []
    try:
        conn_manager = None
        if isinstance(elemento, FamilyInstance) and elemento.MEPModel:
            conn_manager = elemento.MEPModel.ConnectorManager
        elif hasattr(elemento, "ConnectorManager"):
            conn_manager = elemento.ConnectorManager

        if not conn_manager:
            return

        conectores = conn_manager.Connectors
        if not conectores or conectores.Size <= 0:
            return

        for conn in conectores:
            linhas.append("  Conector ID: {}".format(conn.Id))
            linhas.append("    Tipo: {}".format(conn.ConnectorType))
            linhas.append("    Dominio: {}".format(conn.Domain))
            try:
                linhas.append("    Forma: {}".format(conn.Shape))
                if conn.Shape == ConnectorProfileType.Round:
                    linhas.append("    Diametro: {:.1f} mm".format(conn.Radius * 2 * 304.8))
                elif conn.Shape == ConnectorProfileType.Rectangular:
                    linhas.append("    Largura: {:.1f} mm".format(conn.Width * 304.8))
                    linhas.append("    Altura: {:.1f} mm".format(conn.Height * 304.8))
            except:
                pass
            linhas.append("    Conectado: {}".format("Sim" if conn.IsConnected else "Nao"))
    except Exception as ex:
        linhas.append("  <erro ao analisar conectores: {}>".format(str(ex)))

    if linhas:
        escrever_log("\n--- CONECTORES ---")
        for linha in linhas:
            escrever_log(linha)


def analisar_geometria(elemento):
    """Versao enxuta: ignora geometria vazia."""
    try:
        options = Options()
        options.DetailLevel = ViewDetailLevel.Fine
        geom = elemento.get_Geometry(options)
        if not geom:
            return

        solid_count = 0
        face_count = 0
        volume_total = 0

        for geom_obj in geom:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
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

        if solid_count <= 0 and face_count <= 0:
            return

        escrever_log("\n--- GEOMETRIA ---")
        escrever_log("Numero de solidos: {}".format(solid_count))
        escrever_log("Numero de faces: {}".format(face_count))
        escrever_log("Volume total: {:.3f} m3".format(volume_total * 0.0283168))
    except Exception as ex:
        escrever_log("\n--- GEOMETRIA ---")
        escrever_log("  <erro ao analisar geometria: {}>".format(str(ex)))


def analisar_localizacao(elemento):
    """Versao enxuta: ignora elementos sem Location."""
    linhas = []
    try:
        if not hasattr(elemento, "Location"):
            return
        location = elemento.Location
        if not location:
            return
        if isinstance(location, LocationPoint):
            pt = location.Point
            linhas.append("Tipo de localizacao: Ponto")
            linhas.append("Coordenadas X: {:.3f} mm".format(pt.X * 304.8))
            linhas.append("Coordenadas Y: {:.3f} mm".format(pt.Y * 304.8))
            linhas.append("Coordenadas Z: {:.3f} mm".format(pt.Z * 304.8))

            if link_context["is_link"] and link_context["transform"]:
                try:
                    pt_host = link_context["transform"].OfPoint(pt)
                    linhas.append("--- Coordenadas no Host (apos transform) ---")
                    linhas.append("  X (host): {:.3f} mm".format(pt_host.X * 304.8))
                    linhas.append("  Y (host): {:.3f} mm".format(pt_host.Y * 304.8))
                    linhas.append("  Z (host): {:.3f} mm".format(pt_host.Z * 304.8))
                except:
                    pass

            if hasattr(location, 'Rotation'):
                linhas.append("Rotacao: {:.2f} deg".format(location.Rotation * 57.2958))
        elif isinstance(location, LocationCurve):
            curve = location.Curve
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            linhas.append("Tipo de localizacao: Curva")
            linhas.append("Comprimento: {:.3f} mm".format(curve.Length * 304.8))
            linhas.append("Ponto inicial: ({:.2f}, {:.2f}, {:.2f})".format(
                start.X * 304.8, start.Y * 304.8, start.Z * 304.8))
            linhas.append("Ponto final: ({:.2f}, {:.2f}, {:.2f})".format(
                end.X * 304.8, end.Y * 304.8, end.Z * 304.8))
    except Exception as ex:
        linhas.append("  <erro ao analisar localizacao: {}>".format(str(ex)))

    if linhas:
        escrever_log("\n--- LOCALIZACAO ---")
        for linha in linhas:
            escrever_log(linha)


def analisar_workset(elemento):
    """Versao enxuta: so mostra workset em arquivo colaborativo."""
    try:
        if not analysis_doc.IsWorkshared:
            return
        workset_id = elemento.WorksetId
        if workset_id == WorksetId.InvalidWorksetId:
            return
        workset = analysis_doc.GetWorksetTable().GetWorkset(workset_id)
        escrever_log("\n--- WORKSET ---")
        escrever_log("Workset: {}".format(workset.Name))
        escrever_log("Workset ID: {}".format(workset_id.IntegerValue))
    except Exception as ex:
        escrever_log("\n--- WORKSET ---")
        escrever_log("  <erro ao analisar workset: {}>".format(str(ex)))


def analisar_fase(elemento):
    """Versao enxuta: evita repetir 'nao demolido' para todo elemento."""
    linhas = []
    try:
        try:
            created_phase_id = elemento.CreatedPhaseId
        except:
            return
        try:
            demolished_phase_id = elemento.DemolishedPhaseId
        except:
            demolished_phase_id = ElementId.InvalidElementId

        fase_criacao = analysis_doc.GetElement(created_phase_id)
        fase_demolida = analysis_doc.GetElement(demolished_phase_id)
        if fase_criacao:
            linhas.append("Fase de Criacao: {}".format(fase_criacao.Name))
        if fase_demolida:
            linhas.append("Fase de Demolicao: {}".format(fase_demolida.Name))
    except Exception as ex:
        linhas.append("  <erro ao analisar fases: {}>".format(str(ex)))

    if linhas:
        escrever_log("\n--- FASES ---")
        for linha in linhas:
            escrever_log(linha)


def is_panel_schedule_view(el):
    """Verifica se o elemento é um PanelScheduleView."""
    try:
        class_name = el.__class__.__name__
        if "PanelScheduleView" in class_name:
            return True
        type_name = el.GetType().Name
        if "PanelScheduleView" in type_name:
            return True
    except:
        pass
    return False


def is_view_schedule(el):
    """Verifica se o elemento é um ViewSchedule (tabela de quantidades)."""
    try:
        if isinstance(el, ViewSchedule):
            return True
        class_name = el.__class__.__name__
        if class_name == "ViewSchedule":
            return True
    except:
        pass
    return False


def encontrar_panel_schedule_view(panel_id):
    """Busca o PanelScheduleView associado ao ID de um equipamento de painel físico."""
    try:
        for psv in FilteredElementCollector(doc).OfClass(PanelScheduleView):
            try:
                if psv.GetPanel() == panel_id:
                    return psv
            except:
                continue
    except:
        pass
    return None


# ==================== INÍCIO DO SCRIPT ====================

from Autodesk.Revit.UI.Selection import ObjectType as _ObjType
from Autodesk.Revit.DB import RevitLinkInstance as _RLI

limpar_log_antigo()

sel_ids = list(uidoc.Selection.GetElementIds())

# Detecta se a seleção contém apenas RevitLinkInstances (clique no vínculo inteiro)
# ou está vazia → oferece modo vínculo
selected_elems = [doc.GetElement(eid) for eid in sel_ids]
all_links = selected_elems and all(isinstance(e, _RLI) for e in selected_elems)

if not sel_ids or all_links:
    opcao = forms.CommandSwitchWindow.show(
        ["Elemento do Projeto Ativo", "Elemento do Vínculo"],
        message="O que deseja inspecionar?",
        title="Inspecionar Tipo"
    )
    if not opcao:
        script.exit()

    if "Vínculo" in opcao:
        try:
            ref = uidoc.Selection.PickObject(
                _ObjType.LinkedElement,
                "Clique no elemento do vínculo para inspecionar"
            )
        except Exception:
            script.exit()

        link_inst_picked = doc.GetElement(ref.ElementId)
        picked_link_doc  = link_inst_picked.GetLinkDocument()
        picked_transform = link_inst_picked.GetTotalTransform()

        if not picked_link_doc:
            forms.alert("Não foi possível acessar o documento do vínculo.")
            script.exit()

        # Troca contexto de análise para o link
        analysis_doc = picked_link_doc
        link_context["is_link"]    = True
        link_context["name"]       = picked_link_doc.Title
        link_context["transform"]  = picked_transform

        sel_ids   = [ref.LinkedElementId]
        _use_adoc = True
    else:
        # Modo ativo sem seleção prévia: pede para selecionar
        try:
            refs = uidoc.Selection.PickObjects(
                _ObjType.Element,
                "Selecione os elementos para inspecionar"
            )
            sel_ids = [r.ElementId for r in refs]
        except Exception:
            script.exit()
        _use_adoc = False
else:
    _use_adoc = False

if not sel_ids:
    forms.alert("Nenhum elemento selecionado.")
    script.exit()

escrever_log("=" * 80)
escrever_log("RELATORIO COMPLETO DE FAMILIAS E ELEMENTOS")
escrever_log("Data/Hora: {}".format(datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
escrever_log("Projeto: {}".format(doc.Title))
if link_context["is_link"]:
    escrever_log("Vínculo inspecionado: {}".format(link_context["name"]))
    is_id = link_context["transform"].IsIdentity
    escrever_log("Transform do vínculo: {}".format(
        "Identidade (mesma origem)" if is_id else "Com deslocamento/rotação"))
escrever_log("=" * 80)
escrever_log("")

# Resolve elementos — usa analysis_doc para link, doc para ativo
_edoc = analysis_doc if link_context["is_link"] else doc

novos_ids = []
for eid in sel_ids:
    if eid not in novos_ids:
        novos_ids.append(eid)
    try:
        el = _edoc.GetElement(eid)
        if hasattr(el, 'MEPModel') and el.MEPModel:
            sistemas = el.MEPModel.GetElectricalSystems()
            if sistemas:
                for sys in sistemas:
                    if sys.Id not in novos_ids:
                        novos_ids.append(sys.Id)
    except:
        pass

total = len(novos_ids)
escrever_log("Total de elementos (com circuitos associados): {}\n".format(total))
escrever_log("=" * 80 + "\n")

contador = 0
for elid in novos_ids:
    contador += 1
    el = _edoc.GetElement(elid)
    
    escrever_log("\n" + "=" * 80)
    escrever_log("ELEMENTO {}/{}".format(contador, total))
    escrever_log("=" * 80)
    
    # ===== DETECÇÃO ESPECIAL: PanelScheduleView =====
    if is_panel_schedule_view(el):
        escrever_log("\n--- INFORMAÇÕES BÁSICAS ---")
        escrever_log("ID do Elemento: {}".format(el.Id))
        escrever_log("Classe: {}".format(el.__class__.__name__))
        escrever_log("Categoria: Quadro de Cargas (PanelScheduleView)")
        escrever_log("Nome: {}".format(el.Name if hasattr(el, 'Name') else "N/A"))
        analisar_panel_schedule(el)
        analisar_parametros(el)
        escrever_log("\n" + "=" * 80 + "\n")
        continue

    # ===== DETECÇÃO ESPECIAL: ViewSchedule =====
    if is_view_schedule(el):
        escrever_log("\n--- INFORMAÇÕES BÁSICAS ---")
        escrever_log("ID do Elemento: {}".format(el.Id))
        escrever_log("Classe: {}".format(el.__class__.__name__))
        escrever_log("Categoria: Tabela de Quantidades (ViewSchedule)")
        escrever_log("Nome: {}".format(el.Name if hasattr(el, 'Name') else "N/A"))
        analisar_schedule(el)
        analisar_parametros(el)
        escrever_log("\n" + "=" * 80 + "\n")
        continue

    # ===== DETECÇÃO ESPECIAL: ElectricalSystem =====
    try:
        if isinstance(el, ElectricalSystem) or "ElectricalSystem" in el.__class__.__name__:
            escrever_log("\n--- INFORMAÇÕES BÁSICAS ---")
            escrever_log("ID do Elemento: {}".format(el.Id))
            escrever_log("Classe: {}".format(el.__class__.__name__))
            escrever_log("Categoria: Circuitos elétricos (ElectricalSystem)")
            escrever_log("Nome: {}".format(el.Name if hasattr(el, 'Name') else "N/A"))
            
            analisar_circuito(el)
            analisar_parametros(el, None)
            escrever_log("\n" + "=" * 80 + "\n")
            continue
    except:
        pass

    # ===== ELEMENTO NORMAL =====
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
                tipo = _edoc.GetElement(tipo_id)
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

    # Busca PanelScheduleView só para Equipamento elétrico (OST_ElectricalEquipment = -2001040)
    if isinstance(el, FamilyInstance):
        try:
            _cat_id = el.Category.Id.IntegerValue if el.Category else 0
        except:
            _cat_id = 0
        if _cat_id == -2001040:
            _psv = encontrar_panel_schedule_view(el.Id)
            if _psv:
                escrever_log("\n" + "=" * 60)
                escrever_log("QUADRO DE CARGAS (PanelScheduleView) ASSOCIADO A ESTE PAINEL")
                escrever_log("Nome: {} (ID: {})".format(_psv.Name, _psv.Id.IntegerValue))
                escrever_log("=" * 60)
                analisar_panel_schedule(_psv)

    escrever_log("\n" + "=" * 80 + "\n")

escrever_log("\n\n" + "=" * 80)
escrever_log("FIM DO RELATÓRIO")
escrever_log("=" * 80)

# Grava tudo de uma vez (performance)
import io
with io.open(log_path, "w", encoding="utf-8") as f:
    f.write(u"\n".join([unicode(line) if not isinstance(line, unicode) else line for line in log_buffer]))

forms.alert("✅ Relatório completo salvo em:\n{}".format(log_path), title="Sucesso")
