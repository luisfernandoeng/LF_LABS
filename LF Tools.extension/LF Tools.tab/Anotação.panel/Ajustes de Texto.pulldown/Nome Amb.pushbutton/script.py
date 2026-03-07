# -*- coding: utf-8 -*-
"""
Cria TextNotes com dados dos Rooms de um Revit Link
Correcoes: Filtro por Nivel e Anti-Sobreposicao
"""
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

# --- CONFIGURACOES ---
OFFSET_Y_FEET = 0.8  # Offset vertical do texto
TOLERANCE = 0.5      # Tolerancia de altura (em pes) para considerar que e o mesmo nivel
MIN_DIST_TEXT = 1.5  # Distancia minima (em pes) entre textos para evitar sobreposicao (aprox 45cm)

# --- FUNCOES AUXILIARES DE UNIDADES E MATEMATICA ---
def to_meters(feet):
    """Converte pes para metros"""
    return feet * 0.3048

def to_sq_meters(sq_feet):
    """Converte pes quadrados para metros quadrados"""
    return sq_feet * 0.092903

def is_close_point(pt1, points_list, min_dist):
    """Verifica se o ponto pt1 esta muito perto de algum ponto na lista"""
    for pt in points_list:
        if pt1.DistanceTo(pt) < min_dist:
            return True
    return False

# --- SELECAO DO LINK ---
def get_selected_link():
    """Obtem o Revit Link selecionado"""
    selection = uidoc.Selection
    selected_ids = selection.GetElementIds()
    
    if selected_ids.Count == 0:
        forms.alert("Selecione um Revit Link antes de executar.")
        return None
    
    first_id = list(selected_ids)[0]
    element = doc.GetElement(first_id)
    
    if not isinstance(element, DB.RevitLinkInstance):
        forms.alert("O elemento selecionado nao e um Revit Link.")
        return None
    
    return element

def get_rooms_from_link(link_instance):
    """Obtem todos os rooms do documento linkado"""
    link_doc = link_instance.GetLinkDocument()
    if link_doc is None:
        forms.alert("O link nao esta carregado.")
        return None, None
    
    collector = DB.FilteredElementCollector(link_doc)
    rooms = collector.OfCategory(DB.BuiltInCategory.OST_Rooms)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
    return rooms, link_doc

def get_default_text_note_type():
    """Busca o tipo de TextNote adequado"""
    collector = DB.FilteredElementCollector(doc)
    text_types = list(collector.OfClass(DB.TextNoteType).ToElements())
    
    if not text_types:
        return DB.ElementId.InvalidElementId

    target_specific = "RAO-ARQ-008-PLA-TER_TX-Arial-1"
    target_fallback = "Arial" 
    
    candidate_arial = None

    for text_type in text_types:
        try:
            p_name = text_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            name = p_name.AsString() if p_name else text_type.Name
            
            if not name:
                continue

            if name == target_specific:
                return text_type.Id
            
            if candidate_arial is None and target_fallback.lower() in name.lower():
                candidate_arial = text_type.Id

        except:
            continue

    if candidate_arial:
        return candidate_arial

    try:
        default_id = doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType)
        if default_id != DB.ElementId.InvalidElementId:
            return default_id
    except:
        pass

    return text_types[0].Id if text_types else DB.ElementId.InvalidElementId

# --- CRIACAO DAS NOTAS ---
def create_text_notes(link_instance, rooms, content_mode):
    """Cria as TextNotes baseado nos rooms do link"""
    transform = link_instance.GetTransform()
    text_type_id = get_default_text_note_type()
    
    if text_type_id == DB.ElementId.InvalidElementId:
        return 0, ["Nenhum tipo de Texto encontrado."]

    # Obter Elevacao da Vista Atual
    view_level = active_view.GenLevel
    if not view_level:
        return 0, ["A vista atual nao esta associada a um Nivel (Level)."]
    
    view_z = view_level.Elevation

    created_count = 0
    errors = []
    placed_points = []  # Lista para guardar onde ja colocamos texto
    
    with revit.Transaction("Criar Textos de Ambientes"):
        for room in rooms:
            try:
                if room.Location is None:
                    continue
                
                # Filtro de Nivel
                room_level = room.Level
                if not room_level:
                    continue
                
                # Calcula a cota final do room considerando a posicao do Link
                room_final_z = room_level.Elevation + transform.Origin.Z
                
                # Se a diferenca de altura for maior que a tolerancia, ignora
                if abs(room_final_z - view_z) > TOLERANCE:
                    continue

                # Obter Nome
                name_param = room.LookupParameter("Nome") or room.LookupParameter("Name")
                r_name = name_param.AsString().upper() if name_param and name_param.AsString() else "SEM NOME"
                
                final_text = r_name
                
                # Obter Area
                if content_mode in ["Nome + Area", "Completo (Nome+A+PD)"]:
                    area_m2 = to_sq_meters(room.Area)
                    final_text += "\nA: {:.2f}m²".format(area_m2)
                
                # Obter Pe Direito
                if content_mode == "Completo (Nome+A+PD)":
                    try:
                        height_feet = room.UnboundedHeight
                        pd_m = to_meters(height_feet)
                        
                        if pd_m <= 0.01:
                            offset_param = room.get_Parameter(DB.BuiltInParameter.ROOM_UPPER_OFFSET)
                            if offset_param:
                                pd_m = to_meters(offset_param.AsDouble())

                        final_text += "\nPD: {:.2f}m".format(pd_m)
                    except:
                        final_text += "\nPD: ?"

                # Posicionamento
                pt_link = room.Location.Point
                pt_host = transform.OfPoint(pt_link)
                pt_final = DB.XYZ(pt_host.X, pt_host.Y + OFFSET_Y_FEET, pt_host.Z)

                # Filtro Anti-Sobreposicao (2D apenas)
                pt_final_2d = DB.XYZ(pt_final.X, pt_final.Y, 0)
                
                is_overlapping = False
                for existing_pt in placed_points:
                    existing_pt_2d = DB.XYZ(existing_pt.X, existing_pt.Y, 0)
                    if pt_final_2d.DistanceTo(existing_pt_2d) < MIN_DIST_TEXT:
                        is_overlapping = True
                        break
                
                if is_overlapping:
                    continue

                # Criar TextNote
                DB.TextNote.Create(doc, active_view.Id, pt_final, final_text, text_type_id)
                placed_points.append(pt_final)
                created_count += 1
                
            except Exception as ex:
                errors.append("Erro no room: {}".format(str(ex)))
    
    return created_count, errors

# --- MAIN ---
def main():
    """Funcao principal"""
    if active_view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.AreaPlan]:
        forms.alert("Execute este script em uma vista de planta.")
        return
    
    link = get_selected_link()
    if not link:
        return
    
    rooms, link_doc = get_rooms_from_link(link)
    if not rooms:
        return
    
    options = ["Apenas Nome", "Nome + Area", "Completo (Nome+A+PD)"]
    selected_option = forms.CommandSwitchWindow.show(
        options,
        message="Selecione o formato da tag:"
    )
    
    if not selected_option:
        return 
        
    created, errors = create_text_notes(link, rooms, selected_option)
    
    msg = "Sucesso!\n{} notas criadas.".format(created)
    if errors:
        msg += "\n\n{} erros ocorreram.".format(len(errors))
        print("\n".join(errors))
        
    forms.alert(msg)

if __name__ == '__main__':
    main()