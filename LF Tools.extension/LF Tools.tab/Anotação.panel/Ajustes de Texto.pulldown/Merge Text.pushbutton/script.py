"""Merge Text Notes"""
import clr
from pyrevit import revit, DB, forms

# Pegar selecao atual
selection = revit.get_selection()
selected_elements = selection.elements

# Filtrar apenas o que e TextNote
text_notes = [x for x in selected_elements if isinstance(x, DB.TextNote)]

# Verificacao de seguranca
if len(text_notes) < 2:
    forms.alert("Selecione pelo menos 2 notas de texto.", exitscript=True)

# ORDENACAO:
# Organiza a lista baseada na coordenada Y (Altura)
# Do topo para baixo
text_notes.sort(key=lambda t: (-t.Coord.Y, t.Coord.X))

# O primeiro da lista sera o "Mestre"
master_note = text_notes[0]
remaining_notes = text_notes[1:]

# Iniciar Transacao
with revit.Transaction("Unir Textos"):
    for note in remaining_notes:
        # Pega o texto atual e adiciona uma quebra de linha (\r)
        master_note.Text = master_note.Text + "\r" + note.Text
        
        # Deleta a nota antiga
        revit.doc.Delete(note.Id)