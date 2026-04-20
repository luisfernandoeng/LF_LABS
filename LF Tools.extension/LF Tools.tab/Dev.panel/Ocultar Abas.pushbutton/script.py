# -*- coding: utf-8 -*-
"""Configura quais abas do Revit ficam ocultas ao iniciar.

Selecione as abas que deseja esconder e clique em Confirmar.
As alterações entram em vigor imediatamente e persistem nas próximas sessões.
"""
__title__ = "Ocultar\nAbas"
__author__ = "Luís Fernando"

from pyrevit import forms, script
from pyrevit.coreutils import ribbon

LF_HIDE_CONFIG_KEY = 'lf_hidden_tabs'


def _get_config():
    cfg = script.get_config(LF_HIDE_CONFIG_KEY)
    return cfg.get_option('hidden_tabs', [])


def _save_config(tab_names):
    cfg = script.get_config(LF_HIDE_CONFIG_KEY)
    cfg.hidden_tabs = tab_names
    script.save_config()


class TabOption(forms.TemplateListItem):
    def __init__(self, tab, currently_hidden):
        super(TabOption, self).__init__(tab)
        self.state = tab.name in currently_hidden

    @property
    def name(self):
        return self.item.name


def main():
    this_ext = script.get_extension_name()
    hidden_now = _get_config()

    all_tabs = [t for t in ribbon.get_current_ui() if this_ext not in t.name]

    selected = forms.SelectFromList.show(
        [TabOption(t, hidden_now) for t in all_tabs],
        title=u'Ocultar Abas do Revit',
        button_name=u'Aplicar',
        multiselect=True,
    )

    if selected is None:
        return

    new_hidden = [t.name for t in selected if t]
    _save_config(new_hidden)

    # Aplica imediatamente na sessão atual
    try:
        from pyrevit.runtime import types
        types.RibbonTabVisibilityUtils.StopHidingTabs()
        if new_hidden:
            types.RibbonTabVisibilityUtils.StartHidingTabs(new_hidden)
        forms.toast(
            u'{} aba(s) oculta(s). Persistirá nas próximas sessões.'.format(len(new_hidden))
            if new_hidden else u'Todas as abas visíveis.'
        )
    except Exception as ex:
        forms.alert(
            u'Configuração salva, mas não foi possível aplicar agora:\n{}\n\n'
            u'Reinicie o Revit para aplicar.'.format(ex),
            title=u'Ocultar Abas'
        )


if __name__ == '__main__':
    main()
