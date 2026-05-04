# -*- coding: utf-8 -*-
"""Persistent Revit ribbon tab visibility helpers for LF Tools."""

import clr

clr.AddReference("AdWindows")
import Autodesk.Windows as adWin

from pyrevit import script


CONFIG_KEY = "lf_hidden_tabs"
VIEW_HANDLER_KEY = "lf_tabs_visibility_view_handler"
IDLING_HANDLER_KEY = "lf_tabs_visibility_idling_handler"
DOC_HANDLER_KEY = "lf_tabs_visibility_doc_handler"
LEGACY_VIEW_HANDLER_KEY = "lf_tabs_visibility_handler"
_LOCAL_HANDLERS = {}


def _text(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)


def get_config():
    return script.get_config(CONFIG_KEY)


def _get_handler(uiapp, key):
    try:
        handler = getattr(uiapp, key)
        if handler:
            return handler
    except Exception:
        pass

    try:
        return script.get_envvar(key)
    except Exception:
        return _LOCAL_HANDLERS.get(key)


def _set_handler(uiapp, key, handler):
    _LOCAL_HANDLERS[key] = handler
    try:
        setattr(uiapp, key, handler)
    except Exception:
        pass
    try:
        script.set_envvar(key, handler)
    except Exception:
        pass


def get_hidden():
    return set(_text(tab) for tab in get_config().get_option("hidden_tabs", []))


def is_active():
    return bool(get_config().get_option("is_active", False))


def save_hidden(tab_names):
    get_config().hidden_tabs = sorted(_text(tab) for tab in tab_names if _text(tab))
    script.save_config()


def set_active(state):
    get_config().is_active = bool(state)
    script.save_config()
    try:
        script.toggle_icon(bool(state))
    except Exception:
        pass


def _extension_name(default=u"LF Tools"):
    try:
        return _text(script.get_extension_name())
    except Exception:
        return default


def _ribbon_tabs():
    ribbon = adWin.ComponentManager.Ribbon
    if ribbon is None:
        return []
    return list(ribbon.Tabs)


def is_managed_tab(tab, exclude_ext=None):
    title = _text(getattr(tab, "Title", None))
    tab_id = _text(getattr(tab, "Id", None))
    exclude_ext = _text(exclude_ext or _extension_name())

    if not title:
        return False
    if exclude_ext and exclude_ext in title:
        return False
    if "Modify" in tab_id:
        return False

    try:
        if tab.IsContextualTab:
            return False
    except Exception:
        pass

    return True


def iter_user_tabs(exclude_ext=None):
    for tab in _ribbon_tabs():
        if is_managed_tab(tab, exclude_ext=exclude_ext):
            yield tab


def get_tab_titles(exclude_ext=None):
    return [_text(tab.Title) for tab in iter_user_tabs(exclude_ext=exclude_ext)]


def apply_visibility(to_hide=None, to_show=None, exclude_ext=None):
    to_hide = set(_text(tab) for tab in (to_hide or []))
    to_show = set(_text(tab) for tab in (to_show or []))
    changed = 0

    for tab in iter_user_tabs(exclude_ext=exclude_ext):
        title = _text(tab.Title)
        try:
            if title in to_hide and tab.IsVisible:
                tab.IsVisible = False
                changed += 1
            elif title in to_show and not tab.IsVisible:
                tab.IsVisible = True
                changed += 1
        except Exception:
            pass

    return changed


def apply_saved_profile(exclude_ext=None):
    hidden = get_hidden()
    if not is_active() or not hidden:
        return 0
    return apply_visibility(to_hide=hidden, exclude_ext=exclude_ext)


def remove_saved_handlers(uiapp):
    legacy_view_handler = _get_handler(uiapp, LEGACY_VIEW_HANDLER_KEY)
    if legacy_view_handler:
        try:
            uiapp.ViewActivated -= legacy_view_handler
        except Exception:
            pass
        _set_handler(uiapp, LEGACY_VIEW_HANDLER_KEY, None)

    view_handler = _get_handler(uiapp, VIEW_HANDLER_KEY)
    if view_handler:
        try:
            uiapp.ViewActivated -= view_handler
        except Exception:
            pass
        _set_handler(uiapp, VIEW_HANDLER_KEY, None)

    idling_handler = _get_handler(uiapp, IDLING_HANDLER_KEY)
    if idling_handler:
        try:
            uiapp.Idling -= idling_handler
        except Exception:
            pass
        _set_handler(uiapp, IDLING_HANDLER_KEY, None)

    doc_handler = _get_handler(uiapp, DOC_HANDLER_KEY)
    if doc_handler:
        try:
            uiapp.Application.DocumentOpened -= doc_handler
        except Exception:
            pass
        _set_handler(uiapp, DOC_HANDLER_KEY, None)


def install_persistent_hider(uiapp, exclude_ext=None):
    """Register lightweight events that keep configured tabs hidden."""
    remove_saved_handlers(uiapp)

    if not is_active() or not get_hidden():
        return

    def reapply(sender=None, args=None):
        try:
            apply_saved_profile(exclude_ext=exclude_ext)
        except Exception:
            pass

    def on_idling(sender, args):
        reapply(sender, args)

    uiapp.ViewActivated += reapply
    uiapp.Idling += on_idling

    try:
        uiapp.Application.DocumentOpened += reapply
    except Exception:
        pass

    _set_handler(uiapp, VIEW_HANDLER_KEY, reapply)
    _set_handler(uiapp, IDLING_HANDLER_KEY, on_idling)
    _set_handler(uiapp, DOC_HANDLER_KEY, reapply)

    reapply()
