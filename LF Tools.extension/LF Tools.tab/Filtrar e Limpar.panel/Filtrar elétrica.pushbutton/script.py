# -*- coding: utf-8 -*-
# Auditoria Elétrica Completa - Agora com "ID do comando" preservado

from pyrevit import revit, DB
from System.Collections.Generic import List
import sys

doc = revit.doc
uidoc = revit.uidoc


# ===============================================================
# 1) Seleção inicial (usuário deve selecionar o painel)
# ===============================================================

sel_ids = uidoc.Selection.GetElementIds()
if not sel_ids:
    print("Selecione um painel.")
    sys.exit()

elements = [doc.GetElement(eid) for eid in sel_ids]

painel = None
for el in elements:
    if isinstance(el, DB.FamilyInstance):
        if el.Category and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_ElectricalEquipment):
            painel = el
            break

if not painel:
    print("Nenhum painel encontrado.")
    sys.exit()

pn = painel.LookupParameter("Panel Name")
painel_nome = pn.AsString() if pn else painel.Name

print("Painel detectado:", painel_nome)


# ===============================================================
# 2) Coletar circuitos
# ===============================================================

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

print("Circuitos derivados encontrados:", len(circuitos_derivados))
print("Circuito alimentador encontrado:", "SIM" if circuito_alimentador else "NÃO")


# ===============================================================
# 3) Coletar dispositivos, conduítes e conexões
# ===============================================================

elementos_do_sistema = set()
luminarias_com_comando = []

def adicionar_elementos(circ):
    try:
        for elid in circ.Elements:
            inst = doc.GetElement(elid)
            if not inst:
                continue

            elementos_do_sistema.add(inst.Id)

            # NOVO — detectar luminárias com ID do comando
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


# Também incluir o próprio painel
elementos_do_sistema.add(painel.Id)

print("Elementos físicos coletados:", len(elementos_do_sistema))


# ===============================================================
# 4) NOVO — Encontrar interruptores que compartilham o mesmo ID de comando
# ===============================================================

interruptores_encontrados = set()

if luminarias_com_comando:
    print("\nVerificando ID do comando...")

    # coletar todos dispositivos elétricos
    dispositivos = DB.FilteredElementCollector(doc)\
        .OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures)\
        .WhereElementIsNotElementType()\
        .ToElements()

    for lum, comando in luminarias_com_comando:
        print("Luminária", lum.Id, "tem comando:", comando)

        for dev in dispositivos:
            p = dev.LookupParameter("ID do comando")
            if not p:
                continue
            if p.AsString() == comando:
                interruptores_encontrados.add(dev.Id)

    print("Interruptores encontrados:", len(interruptores_encontrados))


# ===============================================================
# 5) Criar seleção final
# ===============================================================

ids_finais = set(sel_ids)

# circuitos
for c in circuitos_derivados:
    ids_finais.add(c.Id)

if circuito_alimentador:
    ids_finais.add(circuito_alimentador.Id)

# elementos físicos
ids_finais.update(elementos_do_sistema)

# interruptores ligados por comando
ids_finais.update(interruptores_encontrados)

# aplicar no Revit
ids_selecao = List[DB.ElementId](list(ids_finais))
uidoc.Selection.SetElementIds(ids_selecao)

print("\n=== Seleção COMPLETA criada ===")
print("Total selecionado:", ids_selecao.Count)
print("=========================================")