# -*- coding: utf-8 -*-
"""Ocultar Abas do Revit.

Click normal  → liga/desliga o perfil salvo (toggle).
Shift+Click   → abre configuração para escolher quais abas ocultar.
"""
__title__ = "Ocultar\nAbas"
__author__ = "Luís Fernando"

import clr
clr.AddReference('AdWindows')
import Autodesk.Windows as adWin

from pyrevit import forms, script

LF_HIDE_CONFIG_KEY = 'lf_hidden_tabs'


# ── Config ─────────────────────────────────────────────────────────────────

def _get_hidden():
    cfg = script.get_config(LF_HIDE_CONFIG_KEY)
    return set(cfg.get_option('hidden_tabs', []))


def _is_active():
    cfg = script.get_config(LF_HIDE_CONFIG_KEY)
    return cfg.get_option('is_active', False)


def _save_hidden(tab_names):
    cfg = script.get_config(LF_HIDE_CONFIG_KEY)
    cfg.hidden_tabs = list(tab_names)
    script.save_config()


def _set_active(state):
    cfg = script.get_config(LF_HIDE_CONFIG_KEY)
    cfg.is_active = state
    script.save_config()
    # Atualiza o ícone do botão para feedback visual
    script.toggle_icon(state)


# ── AdWindows helpers ───────────────────────────────────────────────────────

def _iter_user_tabs(exclude_ext):
    """Itera abas não-sistema, não-contextuais, não-LF."""
    for tab in adWin.ComponentManager.Ribbon.Tabs:
        title  = tab.Title or u""
        tab_id = tab.Id   or u""
        if not title:
            continue
        if exclude_ext in title:
            continue
        if "Modify" in tab_id:
            continue
        try:
            if tab.IsContextualTab:
                continue
        except Exception:
            continue
        yield tab


def _get_tab_titles(exclude_ext):
    return [tab.Title for tab in _iter_user_tabs(exclude_ext)]


def _apply_visibility(to_hide, to_show, exclude_ext):
    """Aplica delta de visibilidade — nunca toca em abas de sistema ou contextuais."""
    for tab in _iter_user_tabs(exclude_ext):
        title = tab.Title or u""
        if title in to_hide:
            tab.IsVisible = False
        elif title in to_show:
            tab.IsVisible = True


def _are_hidden_applied(saved_hidden, exclude_ext):
    """Retorna True se todas as abas do perfil estão atualmente ocultas."""
    if not saved_hidden:
        return False
    for tab in _iter_user_tabs(exclude_ext):
        if (tab.Title or u"") in saved_hidden and tab.IsVisible:
            return False
    return True


# ── UI ──────────────────────────────────────────────────────────────────────

class TabOption(forms.TemplateListItem):
    def __init__(self, title, hidden_set):
        super(TabOption, self).__init__(title)
        self.state = title in hidden_set

    @property
    def name(self):
        return self.item


def open_config(this_ext):
    """Shift+Click: abre diálogo para configurar quais abas ocultar."""
    old_hidden = _get_hidden()
    tab_titles = _get_tab_titles(exclude_ext=this_ext)

    if not tab_titles:
        forms.alert(u'Nenhuma aba encontrada.', title=u'Ocultar Abas')
        return

    selected = forms.SelectFromList.show(
        [TabOption(t, old_hidden) for t in tab_titles],
        title=u'Ocultar Abas — Configurar Perfil',
        button_name=u'Salvar Perfil',
        multiselect=True,
    )

    if selected is None:
        return

    new_hidden = set(str(t) for t in selected if t)
    _save_hidden(new_hidden)
    
    # Se salvou algo, ativa o modo de ocultação automaticamente
    if new_hidden:
        _set_active(True)
    else:
        _set_active(False)

    try:
        old_set = old_hidden
        _apply_visibility(
            to_hide=new_hidden - old_set,
            to_show=old_set - new_hidden,
            exclude_ext=this_ext,
        )
        if new_hidden:
            forms.toast(u'Perfil salvo — {} aba(s) oculta(s).'.format(len(new_hidden)))
        else:
            forms.toast(u'Perfil salvo — todas as abas visíveis.')
    except Exception as ex:
        forms.alert(
            u'Perfil salvo, mas não foi possível aplicar agora:\n{}\n\n'
            u'Reinicie o Revit para aplicar.'.format(ex),
            title=u'Ocultar Abas'
        )


def toggle(this_ext):
    """Click normal: liga/desliga o perfil salvo."""
    saved = _get_hidden()
    active = _is_active()

    if not saved:
        forms.toast(u'Nenhum perfil configurado — use Shift+Click para configurar.')
        return

    if active:
        # Modo ocultar está ATIVO → Desativar (Mostrar tudo)
        _set_active(False)
        _apply_visibility(to_hide=set(), to_show=saved, exclude_ext=this_ext)
        forms.toast(u'Modo de ocultação DESATIVADO (Abas restauradas).')
    else:
        # Modo ocultar está INATIVO → Ativar (Ocultar perfil)
        _set_active(True)
        _apply_visibility(to_hide=saved, to_show=set(), exclude_ext=this_ext)
        forms.toast(u'Modo de ocultação ATIVADO ({} aba(s) oculta(s)).'.format(len(saved)))


# ── Entry point ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    this_ext = script.get_extension_name()

    shift = False
    try:
        from System.Windows.Input import Keyboard, Key
        shift = Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)
    except Exception:
        pass

    if shift:
        open_config(this_ext)
    else:
        toggle(this_ext)
