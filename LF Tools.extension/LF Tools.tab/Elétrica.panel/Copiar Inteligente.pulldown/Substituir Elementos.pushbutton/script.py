# coding: utf-8
"""Substituir Elementos Elétricos
Autor: LF Tools
Substitui famílias elétricas (X → Y) preservando:
  - Posição exata (XYZ, Rotação, Nível, Elevação do Ponto)
  - Circuito elétrico (AddToCircuit antes de deletar X)
  - Parâmetros compartilhados críticos de engenharia
"""

__title__ = "Substituir\nElementos"
__author__ = "LF Tools"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import math
import sys
import os

from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, ElementSet, ElementId,
    FamilyInstance, FamilySymbol, BuiltInCategory, BuiltInParameter,
    LocationPoint, StorageType, XYZ, Line, ConnectorType
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import InvalidOperationException
from pyrevit import forms, script, revit

try:
    # O lf_electrical_core está na pasta do outro plugin (LF Electrical.pulldown)
    core_lib_path = os.path.normpath(os.path.join(
        __file__, "..", "..", "..", "LF Electrical.pulldown", "lib"
    ))
    if core_lib_path not in sys.path:
        sys.path.append(core_lib_path)
        
    import lf_electrical_core
    from lf_electrical_core import dbg
except Exception as e:
    class DummyDbg:
        def section(self, *a, **k): pass
        def step(self, *a, **k): pass
        def warn(self, *a, **k): pass
        def fail(self, *a, **k): pass
        def ok(self, *a, **k): pass
        def elem_info(self, *a, **k): pass
        def exit(self, *a, **k): pass
    dbg = DummyDbg()
    print("Aviso: lf_electrical_core não carregado ({})".format(e))

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# ═══════════════════════════════════════════════════════════════════════
# PARÂMETROS CRÍTICOS — DICIONÁRIO RÍGIDO
# Chave: nome do parâmetro | Valor: GUID compartilhado (quando disponível)
# Estes são os parâmetros que DEVEM existir na família de destino.
# Se não existirem, o script registra aviso mas NÃO aborta (para robustez
# com famílias diferentes que não tenham todos os campos).
# ═══════════════════════════════════════════════════════════════════════

# GUIDs dos parâmetros compartilhados (extraídos do seu relatório)
SHARED_PARAM_GUIDS = {
    "Potência Aparente (VA)":     "44b19786-a579-4490-bcfe-f9ee378c8811",
    "Fator de Potência":          "55aa77f6-7eed-4763-97a7-6366566aacd8",
    "Tipo de Carga":              "5a93d187-b18c-4a94-bfa6-fd6c73516576",
    "N° de Fases":                "d1d0c4b4-47d8-45bc-a138-06f76d6f0beb",
    "Localização no Projeto":     "2b18c84e-bfb8-46a4-b375-8879cadca3a7",
    "Elevação do Ponto":          "b9a09243-1340-4d04-ac0d-880a915f4ba9",
    "Seção do Condutor Adotado":  "07d034ad-40ad-413f-a44b-9aecdee8be6a",
    "Tipo de Sistema":            "0f88be24-aabb-4568-ab56-6fc9e4f9aac7",
}

# Parâmetros adicionais buscados por nome (instância e tipo)
EXTRA_INSTANCE_PARAMS = [
    "Comentários",
    "Marca",
    "Código Planilha Custo",
]

# ═══════════════════════════════════════════════════════════════════════
# FILTRO DE SELEÇÃO — Apenas FamilyInstance elétrico
# ═══════════════════════════════════════════════════════════════════════

ELECTRICAL_CATS = {
    int(BuiltInCategory.OST_ElectricalFixtures),
    int(BuiltInCategory.OST_LightingFixtures),
    int(BuiltInCategory.OST_ElectricalEquipment),
    int(BuiltInCategory.OST_CommunicationDevices),
    int(BuiltInCategory.OST_DataDevices),
    int(BuiltInCategory.OST_FireAlarmDevices),
    int(BuiltInCategory.OST_SecurityDevices),
    int(BuiltInCategory.OST_NurseCallDevices),
    int(BuiltInCategory.OST_LightingDevices),
    int(BuiltInCategory.OST_TelephoneDevices),
}

class ElectricalInstanceFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if not isinstance(elem, FamilyInstance):
            return False
        if elem.Category and int(elem.Category.Id.IntegerValue) in ELECTRICAL_CATS:
            return True
        return False
    def AllowReference(self, ref, pos):
        return False


# ═══════════════════════════════════════════════════════════════════════
# UTILITÁRIOS DE PARÂMETRO
# ═══════════════════════════════════════════════════════════════════════

def read_param_value(param):
    """Lê o valor bruto de um Parameter do Revit."""
    if param is None or not param.HasValue:
        return None
    st = param.StorageType
    if st == StorageType.String:
        return param.AsString()
    elif st == StorageType.Double:
        return param.AsDouble()
    elif st == StorageType.Integer:
        return param.AsInteger()
    elif st == StorageType.ElementId:
        return param.AsElementId()
    return None


def set_param_value(param, value):
    """Grava valor em Parameter. Retorna True se sucesso."""
    if param is None or param.IsReadOnly or value is None:
        return False
    try:
        st = param.StorageType
        if st == StorageType.String:
            param.Set(str(value) if not isinstance(value, str) else value)
        elif st == StorageType.Double:
            param.Set(float(value))
        elif st == StorageType.Integer:
            param.Set(int(value))
        elif st == StorageType.ElementId:
            param.Set(value if isinstance(value, ElementId) else ElementId(int(value)))
        return True
    except Exception:
        return False


def transfer_shared_params(src_elem, dst_elem, logs):
    """
    Transfere os parâmetros compartilhados críticos (por GUID)
    de src_elem para dst_elem. Adiciona registros em logs[].
    """
    dbg.step("Transferindo shared params")
    for param_name, guid_str in SHARED_PARAM_GUIDS.items():
        # Busca por nome (LookupParameter é mais confiável via nome para shared params)
        src_p = src_elem.LookupParameter(param_name)
        if src_p is None:
            logs.append("  ⚠️ Origem sem parâmetro: '{}'".format(param_name))
            dbg.warn("Origem sem param '{}'".format(param_name))
            continue

        val = read_param_value(src_p)
        if val is None:
            logs.append("  ℹ️ Parâmetro '{}' vazio na origem — ignorado".format(param_name))
            continue

        dst_p = dst_elem.LookupParameter(param_name)
        if dst_p is None:
            logs.append("  ⚠️ Destino sem parâmetro: '{}' — família Y pode ser diferente".format(param_name))
            continue

        ok = set_param_value(dst_p, val)
        if ok:
            logs.append("  ✅ '{}' transferido".format(param_name))
            dbg.ok("Transferido: {} -> {}".format(param_name, val))
        else:
            logs.append("  ❌ Falha ao gravar '{}' (somente leitura?)".format(param_name))
            dbg.fail("Falha gravação param: {}".format(param_name))


def transfer_extra_params(src_elem, dst_elem, logs):
    """Transfere parâmetros adicionais por nome."""
    dbg.step("Transferindo parâmetros extras de instância")
    for pname in EXTRA_INSTANCE_PARAMS:
        src_p = src_elem.LookupParameter(pname)
        if src_p is None:
            continue
        val = read_param_value(src_p)
        if val is None:
            continue
        dst_p = dst_elem.LookupParameter(pname)
        if dst_p and not dst_p.IsReadOnly:
            if set_param_value(dst_p, val):
                dbg.ok("Extra param '{}' transferido".format(pname))


# ═══════════════════════════════════════════════════════════════════════
# MOTOR DE CIRCUITO E CONEXÕES
# ═══════════════════════════════════════════════════════════════════════

def get_physical_connections(elem):
    """
    Retorna lista com os conectores conectados fisicamente (conduítes, eletrocalhas).
    """
    connections = []
    try:
        cm = None
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            cm = elem.MEPModel.ConnectorManager
        elif hasattr(elem, 'ConnectorManager'):
            cm = elem.ConnectorManager
            
        if cm:
            for conn in cm.Connectors:
                if conn.IsConnected:
                    refs = []
                    for ref in conn.AllRefs:
                        try:
                            # Ignora conexão lógica (circuitos) ou do próprio elemento
                            if ref.ConnectorType == ConnectorType.Logical:
                                continue
                            if ref.Owner.Id != elem.Id:
                                refs.append(ref)
                        except Exception:
                            continue
                    
                    if refs:
                        connections.append({
                            'old_conn': conn,
                            'domain': conn.Domain,
                            'origin': conn.Origin,
                            'refs': refs
                        })
    except Exception as e:
        dbg.warn("Erro ao coletar conexões físicas: {}".format(e))
    return connections


def restore_physical_connections(new_elem, connections, logs):
    """
    Tenta reconectar os conduítes e eletrocalhas ao novo elemento.
    """
    if not connections:
        return
        
    try:
        cm = None
        if hasattr(new_elem, 'MEPModel') and new_elem.MEPModel:
            cm = new_elem.MEPModel.ConnectorManager
        elif hasattr(new_elem, 'ConnectorManager'):
            cm = new_elem.ConnectorManager
            
        if not cm:
            return
            
        # Converter ConnectorSet para lista Python
        new_conns = []
        for c in cm.Connectors:
            new_conns.append(c)
        
        for old_c in connections:
            best_match = None
            min_dist = float('inf')
            
            # Encontra o conector mais próximo do novo elemento que tenha o mesmo domínio
            for nc in new_conns:
                if nc.Domain == old_c['domain']:
                    dist = nc.Origin.DistanceTo(old_c['origin'])
                    if dist < min_dist:
                        min_dist = dist
                        best_match = nc
                        
            if best_match:
                new_conns.remove(best_match)
                
                # Desconecta os antigos e conecta ao novo
                for ref in old_c['refs']:
                    try:
                        # Tenta desconectar o antigo primeiro para liberar o ref
                        try:
                            old_c['old_conn'].DisconnectFrom(ref)
                        except Exception:
                            pass
                            
                        best_match.ConnectTo(ref)
                        logs.append("  🔗 Conexão com conduíte/eletrocalha restaurada")
                    except Exception as e:
                        logs.append("  ⚠️ Falha ao reconectar conduíte/eletrocalha: {}".format(e))
    except Exception as e:
        logs.append("  ⚠️ Erro ao restaurar conexões físicas: {}".format(e))


def find_circuit(elem):
    """
    Encontra o ElectricalSystem associado a um FamilyInstance.
    Tenta MEPModel → Connectores → busca global (fallback lento).
    Retorna o primeiro ElectricalSystem encontrado ou None.
    """
    # Método 1 - MEPModel.ElectricalSystems
    try:
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            mep = elem.MEPModel
            for method_name in ['GetElectricalSystems', 'GetAssignedElectricalSystems', 'ElectricalSystems']:
                if hasattr(mep, method_name):
                    res = getattr(mep, method_name)
                    systems = res() if callable(res) else res
                    if systems:
                        for sys in systems:
                            if isinstance(sys, ElectricalSystem):
                                return sys
    except Exception:
        pass

    # Método 2 - Conectores elétricos
    try:
        cm = None
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            cm = elem.MEPModel.ConnectorManager
        elif hasattr(elem, 'ConnectorManager'):
            cm = elem.ConnectorManager
        if cm:
            for conn in cm.Connectors:
                try:
                    from Autodesk.Revit.DB import Domain
                    if conn.Domain == Domain.DomainElectrical:
                        if hasattr(conn, 'MEPSystem') and conn.MEPSystem:
                            if isinstance(conn.MEPSystem, ElectricalSystem):
                                return conn.MEPSystem
                        for ref_conn in conn.AllRefs:
                            if isinstance(ref_conn.Owner, ElectricalSystem):
                                return ref_conn.Owner
                except Exception:
                    continue
    except Exception:
        pass

    # Método 3 - Busca global (fallback)
    try:
        for es in FilteredElementCollector(doc).OfClass(ElectricalSystem).ToElements():
            try:
                if es.Elements:
                    for member in es.Elements:
                        if member.Id == elem.Id:
                            return es
            except Exception:
                continue
    except Exception:
        pass

    return None


def add_to_circuit(circuit, new_elem):
    """Adiciona new_elem ao ElectricalSystem. Retorna True se bem-sucedido."""
    try:
        elem_set = ElementSet()
        elem_set.Insert(new_elem)
        circuit.AddToCircuit(elem_set)
        return True
    except Exception as ex:
        dbg.fail("Falha no AddToCircuit: {}".format(ex))
        return False


def remove_from_circuit(circuit, elem):
    """Remove elem do ElectricalSystem."""
    try:
        elem_set = ElementSet()
        elem_set.Insert(elem)
        circuit.RemoveFromCircuit(elem_set)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════════
# MOTOR DE SUBSTITUIÇÃO
# ═══════════════════════════════════════════════════════════════════════

def get_snapshot(elem):
    """
    Captura todos os dados geométricos e de contexto do elemento X.
    Retorna um dict com: xyz, level_id, rotation, elevation_offset, circuit
    """
    dbg.step("Capturando snapshot de {}".format(elem.Id))
    snap = {
        'xyz': None,
        'level_id': None,
        'rotation': 0.0,
        'elevation_offset': 0.0,
        'face_ref': None,
        'circuit': None,
        'elem_id': elem.Id,
    }

    # Localização
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        snap['xyz'] = loc.Point
        snap['rotation'] = loc.Rotation

    # Level
    try:
        snap['level_id'] = elem.LevelId
    except Exception:
        pass

    # Elevação do hospedeiro
    try:
        off_p = elem.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
        if off_p and off_p.HasValue:
            snap['elevation_offset'] = off_p.AsDouble()
    except Exception:
        pass

    # Circuito
    snap['circuit'] = find_circuit(elem)
    if snap['circuit']:
        dbg.ok("Circuito capturado: {}".format(snap['circuit'].Id))

    return snap


def create_replacement(snap, new_symbol, src_elem):
    """
    Cria a nova instância (Família Y) no mesmo local da Família X.
    Retorna o novo FamilyInstance ou levanta exceção.
    """
    dbg.step("Criando nova família")
    xyz = snap['xyz']
    level_id = snap['level_id']

    if xyz is None:
        raise ValueError("Não foi possível obter a localização do elemento original.")

    # Ativa o símbolo se necessário
    if not new_symbol.IsActive:
        dbg.step("Ativando símbolo destino")
        new_symbol.Activate()
        doc.Regenerate()

    # Criação
    if level_id and level_id != ElementId.InvalidElementId:
        level = doc.GetElement(level_id)
        from Autodesk.Revit.DB.Structure import StructuralType
        new_inst = doc.Create.NewFamilyInstance(
            xyz,
            new_symbol,
            level,
            StructuralType.NonStructural
        )
    else:
        from Autodesk.Revit.DB.Structure import StructuralType
        new_inst = doc.Create.NewFamilyInstance(
            xyz,
            new_symbol,
            StructuralType.NonStructural
        )

    doc.Regenerate()
    dbg.ok("Nova instância criada: ID {}".format(new_inst.Id))

    # Aplicar rotação (se diferente de 0)
    if abs(snap['rotation']) > 0.001:
        dbg.step("Aplicando rotação: {:.2f} rad".format(snap['rotation']))
        axis = Line.CreateBound(
            XYZ(xyz.X, xyz.Y, xyz.Z),
            XYZ(xyz.X, xyz.Y, xyz.Z + 1)
        )
        new_inst.Location.Rotate(axis, snap['rotation'])

    # Aplicar offset de elevação do hospedeiro
    try:
        if abs(snap['elevation_offset']) > 0.001:
            off_p = new_inst.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
            if off_p and not off_p.IsReadOnly:
                off_p.Set(snap['elevation_offset'])
                dbg.ok("Offset de elevação ({}) restaurado".format(snap['elevation_offset']))
    except Exception:
        pass

    return new_inst


def replace_element(src_elem, new_symbol, logs):
    """
    Substitui src_elem pela família new_symbol:
      1. Snapshot geométrico + circuito
      2. Cria nova instância no mesmo local
      3. Transfere parâmetros
      4. Adiciona ao circuito
      5. Deleta src_elem

    Retorna (True, novo_elem_id) ou (False, None).
    """
    dbg.section("Iniciando Substituição - Alvo: {}".format(src_elem.Id))
    
    # Fase A: Snapshot
    snap = get_snapshot(src_elem)
    circuit = snap['circuit']

    # Capturar conexões físicas antes de deletar
    phys_conns = get_physical_connections(src_elem)

    logs.append("📍 Localização: X={:.1f} Y={:.1f} Z={:.1f} | Rot: {:.2f}°".format(
        (snap['xyz'].X * 304.8) if snap['xyz'] else 0,
        (snap['xyz'].Y * 304.8) if snap['xyz'] else 0,
        (snap['xyz'].Z * 304.8) if snap['xyz'] else 0,
        math.degrees(snap['rotation'])
    ))

    if circuit:
        try:
            logs.append("⚡ Circuito: {} | Painel: {}".format(
                circuit.CircuitNumber,
                circuit.PanelId
            ))
        except Exception:
            logs.append("⚡ Circuito encontrado (sem número disponível)")
    else:
        logs.append("ℹ️ Elemento não está em nenhum circuito — será apenas reposicionado")

    # Fase B: Criar nova instância + transferir parâmetros
    new_inst = create_replacement(snap, new_symbol, src_elem)
    doc.Regenerate()

    # Parâmetros compartilhados
    transfer_shared_params(src_elem, new_inst, logs)
    # Parâmetros extras
    transfer_extra_params(src_elem, new_inst, logs)

    # Fase B.5: Restaurar conexões físicas
    if phys_conns:
        dbg.step("Restaurando conexões físicas")
        restore_physical_connections(new_inst, phys_conns, logs)

    # Fase C: Reconexão ao circuito
    if circuit:
        dbg.step("Restaurando circuito {}".format(circuit.Id))
        ok = add_to_circuit(circuit, new_inst)
        if ok:
            logs.append("✅ Adicionado ao circuito {}".format(circuit.CircuitNumber))
            dbg.ok("Conectado ao circuito")
        else:
            logs.append("⚠️ Não foi possível adicionar ao circuito automaticamente")
            dbg.fail("Falha reintegração ao circuito")

    # Deleta o elemento original
    src_id = src_elem.Id
    dbg.step("Deletando instância antiga {}".format(src_id))
    doc.Delete(src_id)
    logs.append("🗑️ Elemento original ({}) deletado".format(src_id.IntegerValue))

    return True, new_inst.Id


# ═══════════════════════════════════════════════════════════════════════
# SELEÇÃO DE FAMÍLIA/TIPO DESTINO
# ═══════════════════════════════════════════════════════════════════════

def choose_target_symbol(src_elements):
    """
    Mostra ao usuário a lista de FamilySymbols disponíveis no projeto
    para escolha do tipo de destino (Família Y).
    Retorna o FamilySymbol selecionado ou None.
    """
    dbg.section("Seleção de Família de Destino")

    # Coletamos categorias únicas dos elementos selecionados
    cats = set()
    for e in src_elements:
        if e.Category:
            cats.add(e.Category.Id.IntegerValue)

    # Lista de BICs para iterar
    bic_list = [
        ("OST_ElectricalFixtures", BuiltInCategory.OST_ElectricalFixtures),
        ("OST_LightingFixtures", BuiltInCategory.OST_LightingFixtures),
        ("OST_ElectricalEquipment", BuiltInCategory.OST_ElectricalEquipment),
        ("OST_CommunicationDevices", BuiltInCategory.OST_CommunicationDevices),
        ("OST_DataDevices", BuiltInCategory.OST_DataDevices),
        ("OST_FireAlarmDevices", BuiltInCategory.OST_FireAlarmDevices),
        ("OST_SecurityDevices", BuiltInCategory.OST_SecurityDevices),
        ("OST_LightingDevices", BuiltInCategory.OST_LightingDevices),
    ]

    # Coleta todos os FamilySymbols das mesmas categorias
    all_symbols = []
    for bic_name, bic in bic_list:
        try:
            cat_id = int(bic)
            if cat_id not in cats:
                continue
            symbols = (FilteredElementCollector(doc)
                       .OfClass(FamilySymbol)
                       .OfCategory(bic)
                       .ToElements())
            all_symbols.extend(symbols)
        except Exception:
            continue

    if not all_symbols:
        try:
            all_symbols = list(
                FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
            )
        except Exception:
            pass

    if not all_symbols:
        forms.alert("Nenhum tipo de familia encontrado no projeto.")
        dbg.fail("Sem FamilySymbols")
        return None

    # Monta dict para o seletor  {label: symbol}
    sym_dict = {}
    for sym in all_symbols:
        try:
            label = get_symbol_label(sym)
            sym_dict[label] = sym
        except Exception as e:
            dbg.warn("Falha ao ler symbol ID {}: {}".format(sym.Id.IntegerValue, e))
            continue

    if not sym_dict:
        forms.alert("Nao foi possivel listar as familias.")
        dbg.fail("sym_dict vazio")
        return None

    sorted_labels = sorted(sym_dict.keys())

    try:
        selected = forms.SelectFromList.show(
            sorted_labels,
            title="Escolha o Tipo de Destino (Familia Y)",
            button_name="Substituir por este tipo",
            multiselect=False
        )
    except Exception as ex:
        dbg.fail("Erro no SelectFromList: {}".format(ex))
        return None

    if not selected:
        dbg.warn("Seleção cancelada")
        return None

    dbg.ok("Selecionado: {}".format(selected))
    return sym_dict[selected]


# ═══════════════════════════════════════════════════════════════════════
# COLETA DE ELEMENTOS-FONTE
# ═══════════════════════════════════════════════════════════════════════

def get_source_elements():
    """
    Obtém os FamilyInstances a substituir.
    Usa a seleção atual se válida, caso contrário pede seleção interativa.
    Retorna lista de FamilyInstance ou [].
    """
    sel_filter = ElectricalInstanceFilter()

    # Seleção atual
    selected_ids = list(uidoc.Selection.GetElementIds())
    if selected_ids:
        valid = []
        for eid in selected_ids:
            el = doc.GetElement(eid)
            if el and sel_filter.AllowElement(el):
                valid.append(el)

        if valid:
            confirm = forms.alert(
                "{} elemento(s) pré-selecionado(s) serão substituídos.\n"
                "Deseja continuar com esses elementos?".format(len(valid)),
                yes=True, no=True, title="Confirmar Seleção"
            )
            if confirm:
                return valid

    # Seleção interativa
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            sel_filter,
            "Selecione os elementos a SUBSTITUIR (apenas dispositivos/luminárias elétricas)"
        )
        if not refs:
            return []
        return [doc.GetElement(r.ElementId) for r in refs]
    except Exception:
        return []


# ═══════════════════════════════════════════════════════════════════════
# MAIN E HELPERS
# ═══════════════════════════════════════════════════════════════════════

def get_symbol_label(sym):
    """Obtém o nome da Família e Tipo de forma robusta no IronPython."""
    # ── Nome da Família ──
    family_name = None
    try:
        p = sym.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        if p and p.HasValue:
            family_name = p.AsString()
    except Exception:
        pass
    if not family_name:
        try:
            p = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            if p and p.HasValue:
                family_name = p.AsString()
        except Exception:
            pass
    if not family_name:
        try:
            family_name = sym.Family.Name
        except Exception:
            pass
    if not family_name:
        family_name = "Desconhecido"

    # ── Nome do Tipo ──
    type_name = None
    try:
        p = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p and p.HasValue:
            type_name = p.AsString()
    except Exception:
        pass
    if not type_name:
        try:
            p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p and p.HasValue:
                type_name = p.AsString()
        except Exception:
            pass
    if not type_name:
        try:
            from Autodesk.Revit.DB import Element
            type_name = Element.Name.__get__(sym)
        except Exception:
            type_name = "ID {}".format(sym.Id.IntegerValue)

    family_name_u = unicode(family_name)
    type_name_u = unicode(type_name)
    return u"{} : {}".format(family_name_u, type_name_u)

def main():
    dbg.section("SCRIPT: Substituir Elementos Iniciado")
    output.print_md("## 🔄 Substituir Elementos Elétricos")
    output.print_md("---")

    # 1. Coletar elementos fonte
    src_elements = get_source_elements()
    if not src_elements:
        forms.alert("Nenhum elemento selecionado. Cancelando.")
        dbg.exit('main', 'Nenhum original escolhido')
        return

    output.print_md("**{}** elemento(s) selecionado(s) para substituição.".format(len(src_elements)))

    # Exibe primeiras famílias selecionadas
    family_names_src = set()
    for e in src_elements:
        try:
            if hasattr(e, 'Symbol') and e.Symbol:
                family_names_src.add(get_symbol_label(e.Symbol))
        except Exception:
            pass
    for n in sorted(family_names_src):
        output.print_md("  - Família X: **{}**".format(n))

    # 2. Escolher família de destino
    new_symbol = choose_target_symbol(src_elements)
    if not new_symbol:
        forms.alert("Nenhum tipo de destino selecionado. Cancelando.")
        dbg.exit('main', 'Nenhum destino escolhido')
        return

    target_label = get_symbol_label(new_symbol)
    output.print_md("**Tipo de destino (Família Y):** {}".format(target_label))
    output.print_md("---")

    # 3. Confirmação final
    confirm = forms.alert(
        "Substituir {} elemento(s) por:\n\n  ▶ {}\n\n"
        "Esta operação irá:\n"
        "  • Criar novo elemento no mesmo local\n"
        "  • Transferir parâmetros elétricos\n"
        "  • Manter o circuito existente\n"
        "  • Deletar o elemento original\n\n"
        "Continuar?".format(len(src_elements), target_label),
        yes=True, no=True, title="Confirmar Substituição"
    )
    if not confirm:
        output.print_md("❌ Operação cancelada pelo usuário.")
        dbg.exit('main', 'Cancelado no último aviso')
        return

    # 4. Execução em transação única
    success_count = 0
    fail_count    = 0
    all_logs      = []

    with Transaction(doc, "Substituir Elementos Elétricos") as t:
        t.Start()

        for idx, src_elem in enumerate(src_elements, start=1):
            elem_logs = []
            try:
                elem_name = "Elemento #{} (ID: {})".format(idx, src_elem.Id.IntegerValue)
                try:
                    elem_name = "{} (ID: {})".format(
                        get_symbol_label(src_elem.Symbol),
                        src_elem.Id.IntegerValue
                    )
                except Exception:
                    pass

                elem_logs.append("\n### {} → {}".format(elem_name, target_label))

                ok, new_id = replace_element(src_elem, new_symbol, elem_logs)
                if ok:
                    elem_logs.append("🆕 Novo elemento: ID {}".format(new_id.IntegerValue))
                    success_count += 1
                else:
                    fail_count += 1

            except Exception as ex:
                elem_logs.append("❌ **ERRO:** {}".format(str(ex)))
                dbg.fail("Erro Fatal Substituição: {}".format(str(ex)))
                fail_count += 1

            all_logs.extend(elem_logs)

        t.Commit()
        dbg.ok("Transação finalizada")

    # 5. Relatório
    output.print_md("---")
    output.print_md("## 📋 Relatório de Substituição")
    output.print_md("✅ **Sucesso:** {} | ❌ **Falha:** {}".format(success_count, fail_count))
    output.print_md("---")

    for log_line in all_logs:
        output.print_md(log_line)

    if success_count > 0:
        forms.toast(
            "✅ {} elemento(s) substituído(s) com sucesso!".format(success_count),
            title="Substituir Elementos"
        )
    dbg.section("FIM SCRIPT")


if __name__ == "__main__":
    main()
