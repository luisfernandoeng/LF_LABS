# -*- coding: utf-8 -*-
"""Ocultar abas do Revit.

Click normal: liga/desliga o perfil salvo.
Shift+Click: escolhe quais abas devem permanecer ocultas.
"""

__title__ = "Ocultar\nAbas"
__author__ = "Luis Fernando"

from pyrevit import forms, HOST_APP

from lf_ribbon_tabs import (
    apply_saved_profile,
    apply_visibility,
    get_hidden,
    get_tab_titles,
    save_hidden,
    set_active,
    is_active,
    install_persistent_hider,
    remove_saved_handlers,
)


class TabOption(forms.TemplateListItem):
    def __init__(self, title, hidden_set):
        forms.TemplateListItem.__init__(self, title)
        self.state = title in hidden_set

    @property
    def name(self):
        return self.item


def _selected_title(item):
    return getattr(item, "item", item)


def _is_shift_pressed():
    try:
        from System.Windows.Input import Keyboard, Key

        return (
            Keyboard.IsKeyDown(Key.LeftShift)
            or Keyboard.IsKeyDown(Key.RightShift)
        )
    except Exception:
        return False


def open_config(this_ext):
    old_hidden = get_hidden()
    tab_titles = get_tab_titles(exclude_ext=this_ext)

    if not tab_titles:
        forms.alert(u"Nenhuma aba encontrada.", title=u"Ocultar Abas")
        return

    selected = forms.SelectFromList.show(
        [TabOption(title, old_hidden) for title in tab_titles],
        title=u"Ocultar Abas - Configurar Perfil",
        button_name=u"Salvar Perfil",
        multiselect=True,
    )

    if selected is None:
        return

    new_hidden = set(_selected_title(item) for item in selected if item)
    save_hidden(new_hidden)
    set_active(bool(new_hidden))

    if new_hidden:
        install_persistent_hider(HOST_APP.uiapp, exclude_ext=this_ext)
    else:
        remove_saved_handlers(HOST_APP.uiapp)

    apply_visibility(
        to_hide=new_hidden,
        to_show=old_hidden - new_hidden,
        exclude_ext=this_ext,
    )

    if new_hidden:
        apply_saved_profile(exclude_ext=this_ext)
        forms.toast(u"Perfil salvo - {} aba(s) oculta(s).".format(len(new_hidden)))
    else:
        forms.toast(u"Perfil salvo - todas as abas visiveis.")


def toggle(this_ext):
    saved = get_hidden()

    if not saved:
        forms.toast(u"Nenhum perfil configurado - use Shift+Click para configurar.")
        return

    if is_active():
        set_active(False)
        remove_saved_handlers(HOST_APP.uiapp)
        apply_visibility(to_show=saved, exclude_ext=this_ext)
        forms.toast(u"Modo de ocultacao DESATIVADO. Abas restauradas.")
    else:
        set_active(True)
        install_persistent_hider(HOST_APP.uiapp, exclude_ext=this_ext)
        apply_saved_profile(exclude_ext=this_ext)
        forms.toast(u"Modo de ocultacao ATIVADO. {} aba(s) oculta(s).".format(len(saved)))


if __name__ == "__main__":
    extension_name = "LF Tools"

    if _is_shift_pressed():
        open_config(extension_name)
    else:
        toggle(extension_name)
