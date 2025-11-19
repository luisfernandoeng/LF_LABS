# -*- coding: utf-8 -*-
import clr
import os
from datetime import datetime
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from pyrevit import forms, script

# Função para escrever no arquivo de log
def write_log(message):
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "debug_tags_mep.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(desktop_path, "a") as f:
        f.write("[{0}] {1}\n".format(timestamp, message))

# Acessa o documento ativo e a interface do usuário
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Verifica se o documento e a interface estão disponíveis
if not doc or not uidoc:
    error_msg = "Não foi possível acessar o documento ativo ou a interface do usuário."
    forms.alert(error_msg, title="Erro")
    write_log(error_msg)
    script.exit()

write_log("Script iniciado. Projeto: {0}, Vista ativa: {1}".format(doc.Title, doc.ActiveView.Name))

# Verifica o tipo de vista ativa
view_type = doc.ActiveView.ViewType
write_log("Tipo de vista ativa: {0}".format(view_type))
if view_type not in [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.DraftingView]:
    warn_msg = "Este script funciona melhor em vistas de planta ou vistas de desenho. Vista atual: {0}".format(view_type)
    forms.alert(warn_msg, title="Aviso")
    write_log(warn_msg)

# Pergunta modo de alinhamento
modo = forms.CommandSwitchWindow.show(
    ["Horizontal", "Vertical"],
    message="Escolha o modo de alinhamento das tags:"
)

if not modo:
    error_msg = "Operação cancelada pelo usuário."
    forms.alert(error_msg, title="Organizar Tags")
    write_log(error_msg)
    script.exit()

write_log("Modo de alinhamento selecionado: {0}".format(modo))

# Mapeia categorias disponíveis
categoria_map = {
    "Dispositivos Elétricos": BuiltInCategory.OST_ElectricalFixtureTags,
    "Conduits": BuiltInCategory.OST_ConduitTags,
    "Luminárias": BuiltInCategory.OST_LightingFixtureTags,
    "Equipamentos Elétricos": BuiltInCategory.OST_ElectricalEquipmentTags,
    "Dispositivos de Iluminação": BuiltInCategory.OST_LightingDeviceTags
}

# Usa todas as categorias do mapa por padrão (sem abrir janela)
categorias = list(categoria_map.keys())
write_log("Todas as categorias de tags selecionadas automaticamente.")

# Mapeia as categorias selecionadas para BuiltInCategory
selected_categories = [categoria_map[cat] for cat in categorias]
write_log("Categorias selecionadas: {0}".format(", ".join(categorias)))

# Verifica visibilidade das categorias na vista
for cat in selected_categories:
    category_id = ElementId(cat)
    if doc.ActiveView.GetCategoryHidden(category_id):
        warn_msg = "As tags da categoria {0} estão ocultas na vista ativa. Ative-as em 'Visibilidade/Gráficos'.".format(cat)
        forms.alert(warn_msg, title="Aviso")
        write_log(warn_msg)
        script.exit()

# Classe personalizada para implementar ISelectionFilter
class TagSelectionFilter(ISelectionFilter):
    def __init__(self, category_ids):
        self.category_ids = [int(cat_id) for cat_id in category_ids]
    def AllowElement(self, element):
        return element.Category.Id.IntegerValue in self.category_ids
    def AllowReference(self, reference, position):
        return True

# Seleciona tags
try:
    selection_filter = TagSelectionFilter(selected_categories)
    selected_refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        selection_filter,
        "Selecione as tags de {0} para alinhar".format(", ".join(categorias))
    )
    write_log("Elementos selecionados via PickObjects: {0}".format(len(selected_refs)))
except Exception as e:
    error_msg = "Erro durante a seleção com PickObjects: {0}".format(str(e))
    forms.alert("Nenhuma tag selecionada ou operação cancelada. Verifique o log para detalhes.", title="Organizar Tags")
    write_log(error_msg)
    script.exit()

# Converte referências em elementos
tags = [doc.GetElement(ref.ElementId) for ref in selected_refs]

# Verifica se tags foram selecionadas
if not tags:
    error_msg = "Nenhuma tag válida selecionada. Verifique se os elementos são tags de {0}.".format(", ".join(categorias))
    forms.alert(error_msg, title="Organizar Tags")
    write_log(error_msg)
    script.exit()

# Verifica se as tags são válidas (IndependentTag com TagHeadPosition)
valid_tags = []
invalid_info = []
for tag in tags:
    element_id = str(tag.Id)
    category_name = tag.Category.Name if tag.Category else "Desconhecida"
    category_id = str(tag.Category.Id) if tag.Category else "N/A"
    is_independent_tag = isinstance(tag, IndependentTag)
    has_leader = False
    has_tag_head = False
    tag_head_position = "Não acessível"
    try:
        if is_independent_tag:
            has_leader = tag.HasLeader
            tag_head_position = str(tag.TagHeadPosition)
            has_tag_head = True
    except Exception as e:
        tag_head_position = "Erro ao acessar TagHeadPosition: {0}".format(str(e))
    log_message = "Elemento ID {0} - Categoria: {1} (ID: {2}) - IndependentTag: {3} - HasLeader: {4} - TagHeadPosition: {5} ({6})".format(
        element_id, category_name, category_id, is_independent_tag, has_leader, has_tag_head, tag_head_position)
    write_log(log_message)
    if is_independent_tag and has_tag_head:
        valid_tags.append(tag)
    else:
        invalid_info.append(log_message)

if invalid_info:
    debug_msg = "Elementos inválidos selecionados (não são tags válidas):\n" + "\n".join(invalid_info)
    forms.alert(debug_msg, title="Depuração: Elementos Inválidos")
    write_log(debug_msg)

if not valid_tags:
    error_msg = "Nenhuma tag válida encontrada com TagHeadPosition. Certifique-se de selecionar tags de {0}.".format(", ".join(categorias))
    forms.alert(error_msg, title="Organizar Tags")
    write_log(error_msg)
    script.exit()

write_log("Tags válidas encontradas com TagHeadPosition: {0}".format(len(valid_tags)))

# Obter posições das tags válidas
posicoes = []
for tag in valid_tags:
    try:
        posicoes.append(tag.TagHeadPosition)
    except Exception as e:
        write_log("Erro ao obter TagHeadPosition para tag ID {0}: {1}".format(tag.Id, str(e)))

if not posicoes:
    error_msg = "Nenhuma posição de tag válida encontrada. Verifique o log para detalhes."
    forms.alert(error_msg, title="Organizar Tags")
    write_log(error_msg)
    script.exit()

write_log("Posições das tags válidas: {0}".format([str(pos) for pos in posicoes]))

# Calcular posição base
if modo == "Horizontal":
    y_base = posicoes[0].Y
    write_log("Alinhamento horizontal - Y base: {0}".format(y_base))
elif modo == "Vertical":
    x_base = posicoes[0].X
    write_log("Alinhamento vertical - X base: {0}".format(x_base))

# Inicia transação
with Transaction(doc, "Alinhar Tags de {0}".format(", ".join(categorias))) as trans:
    trans.Start()
    for tag in valid_tags:
        try:
            pos = tag.TagHeadPosition
            if modo == "Horizontal":
                novo_ponto = XYZ(pos.X, y_base, pos.Z)
            elif modo == "Vertical":
                novo_ponto = XYZ(x_base, pos.Y, pos.Z)
            tag.TagHeadPosition = novo_ponto
            write_log("Tag ID {0} movida para {1}".format(tag.Id, novo_ponto))
        except Exception as e:
            write_log("Erro ao mover tag ID {0}: {1}".format(tag.Id, str(e)))
    trans.Commit()

write_log("Tags alinhadas com sucesso.")
forms.alert("Tags alinhadas com sucesso!", title="Organizar Tags")