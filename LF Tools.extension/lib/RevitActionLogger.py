# -*- coding: utf-8 -*-
"""
RevitActionLogger — LF Tools
Monitora eventos da API do Revit e grava log estruturado em arquivo de texto.
"""
import System
import os
from datetime import datetime

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, ElementId, Domain
)
from Autodesk.Revit.UI.Events import *
from Autodesk.Revit.DB.Events import *

# ================================================================
# CONFIG
# ================================================================

desktop   = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
LOG_FILE  = os.path.join(desktop, "RevitActionLog.txt")

# Idling: intervalo mínimo entre checagens de seleção (segundos)
_IDLING_THROTTLE_S = 10

# Transações internas do Revit que não interessam no log
_IGNORED_TX = [
    "update", "calculations", "pre-edit", "regenerate",
    "selection", "modify element attributes", "auto-join",
    "sketch", "temporary", "highlight",
]

# AppDomain keys
_KEY_DOC_CHANGED  = "LF_ActionLogger_DocChanged"
_KEY_VIEW_ACT     = "LF_ActionLogger_ViewActivated"
_KEY_IDLING       = "LF_ActionLogger_Idling"
_KEY_DOC_OPENED   = "LF_ActionLogger_DocOpened"
_KEY_DOC_CLOSED   = "LF_ActionLogger_DocClosed"
_KEY_DOC_SAVED    = "LF_ActionLogger_DocSaved"
_KEY_DOC_SAVED_AS = "LF_ActionLogger_DocSavedAs"
_KEY_DOC_SYNCED   = "LF_ActionLogger_DocSynced"
_KEY_DIALOG       = "LF_ActionLogger_Dialog"
_KEY_FILE_EXP     = "LF_ActionLogger_FileExported"
_KEY_FILE_IMP     = "LF_ActionLogger_FileImported"
_KEY_IDLE_TIME    = "LF_ActionLogger_IdlingLastTime"
_KEY_SEL_COUNT    = "LF_ActionLogger_LastSelectionCount"

# ================================================================
# ESCRITA NO LOG
# ================================================================

def write_log(message):
    """Grava uma linha no log com timestamp, tolerante a falhas."""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a") as f:
            f.write("[{}] {}\n".format(timestamp, message))
    except Exception:
        pass

# ================================================================
# HELPERS DE ELEMENTO
# ================================================================

_ELEC_CATS = [
    int(BuiltInCategory.OST_ElectricalEquipment),
    int(BuiltInCategory.OST_ElectricalFixtures),
    int(BuiltInCategory.OST_LightingFixtures),
    int(BuiltInCategory.OST_LightingDevices),
    int(BuiltInCategory.OST_ElectricalCircuit),
    int(BuiltInCategory.OST_Wire),
]

def _is_electrical(elem):
    try:
        return elem and elem.Category and elem.Category.Id.IntegerValue in _ELEC_CATS
    except Exception:
        return False

def _is_tx_relevant(name):
    """Filtra transações internas do Revit que poluem o log."""
    n = (name or "").lower().strip()
    return not any(ign in n for ign in _IGNORED_TX)

def _elem_summary(elem):
    """Resumo de uma linha: tipo + nível. Sem dump de todos os parâmetros."""
    try:
        parts = []
        try:
            et = elem.Document.GetElement(elem.GetTypeId())
            if et and et.Name:
                parts.append("Tipo: " + et.Name)
        except Exception:
            pass
        for bip in (BuiltInParameter.FAMILY_LEVEL_PARAM,
                    BuiltInParameter.SCHEDULE_LEVEL_PARAM):
            try:
                p = elem.get_Parameter(bip)
                if p and p.HasValue:
                    v = p.AsValueString()
                    if v:
                        parts.append("Nivel: " + v)
                        break
            except Exception:
                pass
        return ("    " + " | ".join(parts)) if parts else ""
    except Exception:
        return ""

def _elec_context(elem):
    """Log rico para elementos elétricos: família, circuito, painel."""
    try:
        lines = []
        try:
            if hasattr(elem, "Symbol") and elem.Symbol:
                lines.append("    FAM/TIPO: {} | {}".format(
                    elem.Symbol.FamilyName or "", elem.Symbol.Name or ""))
        except Exception:
            pass

        for name in ("Panel", "Painel", "Load Name", "Nome da carga",
                     "Circuit Number", "Número do circuito", "Voltage", "Tensão"):
            try:
                p = elem.LookupParameter(name)
                if p and p.HasValue:
                    val = p.AsValueString() or p.AsString()
                    if not val:
                        try: val = str(p.AsDouble())
                        except Exception: pass
                    if val:
                        lines.append("    [{}] = {}".format(name, val))
            except Exception:
                pass

        try:
            if hasattr(elem, "MEPModel") and elem.MEPModel:
                systems = elem.MEPModel.ElectricalSystems
                if systems and systems.Count > 0:
                    for es in systems:
                        panel_name = ""
                        try:
                            pnl = elem.Document.GetElement(es.PanelId)
                            if pnl: panel_name = pnl.Name
                        except Exception:
                            pass
                        lines.append("    Sistema: Id={} Tipo={} Painel='{}'".format(
                            es.Id.IntegerValue, es.SystemType, panel_name))
        except Exception:
            pass

        return "\n".join(lines)
    except Exception:
        return ""

# ================================================================
# HANDLERS — DOCUMENTO
# ================================================================

def OnDocumentOpened(sender, args):
    try:
        doc  = args.GetDocument()
        path = doc.PathName if doc and doc.PathName else "Novo documento"
        write_log("=== DOCUMENTO ABERTO: '{}' ===".format(path))
    except Exception as e:
        write_log("ERRO OnDocumentOpened: " + str(e))

def OnDocumentClosed(sender, args):
    try:
        path = ""
        try: path = str(args.PathName)
        except Exception: pass
        write_log("=== DOCUMENTO FECHADO: '{}' ===".format(path))
    except Exception:
        pass

def OnDocumentSaved(sender, args):
    try:
        doc  = args.GetDocument()
        path = doc.PathName if doc and doc.PathName else ""
        write_log("=== DOCUMENTO SALVO: '{}' ===".format(path))
    except Exception:
        pass

def OnDocumentSavedAs(sender, args):
    try:
        doc  = args.GetDocument()
        path = doc.PathName if doc and doc.PathName else ""
        write_log("=== DOCUMENTO SALVO COMO: '{}' ===".format(path))
    except Exception:
        pass

def OnDocumentSynchronizedWithCentral(sender, args):
    try:
        doc  = args.GetDocument()
        path = doc.PathName if doc and doc.PathName else ""
        write_log("=== SINCRONIZADO COM CENTRAL: '{}' ===".format(path))
    except Exception:
        pass

def OnDocumentChanged(sender, args):
    try:
        tx_names = args.GetTransactionNames()
        tx_name  = tx_names[0] if tx_names else "Implicita"

        if not _is_tx_relevant(tx_name):
            return

        doc          = args.GetDocument()
        added_ids    = args.GetAddedElementIds()
        modified_ids = args.GetModifiedElementIds()
        deleted_ids  = args.GetDeletedElementIds()

        msgs = ["--- TRANSACAO: '{}' ---".format(tx_name)]

        if deleted_ids:
            msgs.append("  [DELETADO] {} elemento(s)".format(len(deleted_ids)))

        for e_id in added_ids:
            elem = doc.GetElement(e_id)
            if not elem: continue
            cat = elem.Category.Name if elem.Category else "Unknown"
            msgs.append("  [ADICIONADO] ID:{} | Cat:{} | Nome:{}".format(
                e_id.IntegerValue, cat, elem.Name))
            detail = _elec_context(elem) if _is_electrical(elem) else _elem_summary(elem)
            if detail: msgs.append(detail)

        for e_id in modified_ids:
            elem = doc.GetElement(e_id)
            if not elem: continue
            cat = elem.Category.Name if elem.Category else "Unknown"
            msgs.append("  [MODIFICADO] ID:{} | Cat:{} | Nome:{}".format(
                e_id.IntegerValue, cat, elem.Name))
            if _is_electrical(elem):
                detail = _elec_context(elem)
                if detail: msgs.append(detail)

        if len(msgs) > 1:
            write_log("\n".join(msgs))

    except Exception as e:
        write_log("ERRO OnDocumentChanged: " + str(e))

# ================================================================
# HANDLERS — ARQUIVOS E EXPORTAÇÃO
# ================================================================

def OnFileExported(sender, args):
    try:
        fmt  = ""
        path = ""
        try: fmt  = str(args.Format)
        except Exception: pass
        try: path = str(args.Path)
        except Exception: pass
        write_log("=== ARQUIVO EXPORTADO: {} -> '{}' ===".format(fmt, path))
    except Exception:
        pass

def OnFileImported(sender, args):
    try:
        fmt  = ""
        path = ""
        try: fmt  = str(args.Format)
        except Exception: pass
        try: path = str(args.Path)
        except Exception: pass
        write_log("=== ARQUIVO IMPORTADO: {} -> '{}' ===".format(fmt, path))
    except Exception:
        pass

# ================================================================
# HANDLERS — UI
# ================================================================

def OnViewActivated(sender, args):
    try:
        if args.CurrentActiveView:
            write_log("--- VISTA ATIVADA: '{}' ---".format(
                args.CurrentActiveView.Name))
    except Exception:
        pass

def OnDialogBoxShowing(sender, args):
    """Captura dialogs e avisos que o Revit exibe durante operações."""
    try:
        dialog_id = ""
        try: dialog_id = str(args.DialogId)
        except Exception: pass
        write_log("  [DIALOG] Id: '{}'".format(dialog_id))
    except Exception:
        pass

def OnIdling(sender, args):
    """
    Rastreia mudanças de seleção.
    Throttled: dispara no máximo uma vez a cada _IDLING_THROTTLE_S segundos.
    """
    try:
        domain   = System.AppDomain.CurrentDomain
        now      = System.DateTime.Now
        last     = domain.GetData(_KEY_IDLE_TIME)

        if last is not None:
            elapsed = (now - last).TotalSeconds
            if elapsed < _IDLING_THROTTLE_S:
                return

        domain.SetData(_KEY_IDLE_TIME, now)

        uidoc = sender.ActiveUIDocument
        if not uidoc:
            return

        count      = uidoc.Selection.GetElementIds().Count
        last_count = domain.GetData(_KEY_SEL_COUNT) or 0

        if count != last_count:
            write_log("  [SELECAO] {} elemento(s)".format(count))
            domain.SetData(_KEY_SEL_COUNT, count)

    except Exception:
        pass

# ================================================================
# ATTACH / DETACH GENÉRICO
# ================================================================

def _attach(obj, event_name, handler):
    try:
        getattr(obj, event_name).__iadd__(handler)
        return True
    except Exception:
        try:
            evt = getattr(obj, event_name)
            evt += handler
            return True
        except Exception:
            return False

def _detach(obj, event_name, handler):
    try:
        getattr(obj, event_name).__isub__(handler)
        return True
    except Exception:
        try:
            evt = getattr(obj, event_name)
            evt -= handler
            return True
        except Exception:
            return False

def _reg(domain, key, obj, event_name, handler_type, handler_fn):
    """Registra um handler e salva a referência no AppDomain."""
    try:
        h = System.EventHandler[handler_type](handler_fn)
        if _attach(obj, event_name, h):
            domain.SetData(key, h)
            return event_name
    except Exception:
        pass
    return None

def _unreg(domain, key, obj, event_name):
    """Remove um handler pelo AppDomain."""
    h = domain.GetData(key)
    if h:
        _detach(obj, event_name, h)
        domain.SetData(key, None)
        return True
    return False

# ================================================================
# START / STOP / STATUS
# ================================================================

def start_logger(uiapp):
    stop_logger(uiapp)

    app    = uiapp.Application
    domain = System.AppDomain.CurrentDomain
    ok     = []

    # — Documento —
    if _reg(domain, _KEY_DOC_CHANGED,  app,    "DocumentChanged",               DocumentChangedEventArgs,               OnDocumentChanged):               ok.append("DocumentChanged")
    if _reg(domain, _KEY_DOC_OPENED,   app,    "DocumentOpened",                DocumentOpenedEventArgs,                OnDocumentOpened):                ok.append("DocumentOpened")
    if _reg(domain, _KEY_DOC_CLOSED,   app,    "DocumentClosed",                DocumentClosedEventArgs,                OnDocumentClosed):                ok.append("DocumentClosed")
    if _reg(domain, _KEY_DOC_SAVED,    app,    "DocumentSaved",                 DocumentSavedEventArgs,                 OnDocumentSaved):                 ok.append("DocumentSaved")
    if _reg(domain, _KEY_DOC_SAVED_AS, app,    "DocumentSavedAs",               DocumentSavedAsEventArgs,               OnDocumentSavedAs):               ok.append("DocumentSavedAs")
    if _reg(domain, _KEY_DOC_SYNCED,   app,    "DocumentSynchronizedWithCentral", DocumentSynchronizedWithCentralEventArgs, OnDocumentSynchronizedWithCentral): ok.append("SynchronizedWithCentral")

    # — Arquivos —
    if _reg(domain, _KEY_FILE_EXP,     app,    "FileExported",                  FileExportedEventArgs,                  OnFileExported):                  ok.append("FileExported")
    if _reg(domain, _KEY_FILE_IMP,     app,    "FileImported",                  FileImportedEventArgs,                  OnFileImported):                  ok.append("FileImported")

    # — UI —
    if _reg(domain, _KEY_VIEW_ACT,     uiapp,  "ViewActivated",                 ViewActivatedEventArgs,                 OnViewActivated):                 ok.append("ViewActivated")
    if _reg(domain, _KEY_DIALOG,       uiapp,  "DialogBoxShowing",              DialogBoxShowingEventArgs,              OnDialogBoxShowing):              ok.append("DialogBoxShowing")
    if _reg(domain, _KEY_IDLING,       uiapp,  "Idling",                        IdlingEventArgs,                        OnIdling):
        domain.SetData(_KEY_IDLE_TIME,  None)
        domain.SetData(_KEY_SEL_COUNT, 0)
        ok.append("Idling")

    write_log("=== LOGGER INICIADO: {} ===".format(", ".join(ok)))
    return True


def stop_logger(uiapp):
    app    = uiapp.Application
    domain = System.AppDomain.CurrentDomain

    removed = []
    pairs = [
        (_KEY_DOC_CHANGED,  app,   "DocumentChanged"),
        (_KEY_DOC_OPENED,   app,   "DocumentOpened"),
        (_KEY_DOC_CLOSED,   app,   "DocumentClosed"),
        (_KEY_DOC_SAVED,    app,   "DocumentSaved"),
        (_KEY_DOC_SAVED_AS, app,   "DocumentSavedAs"),
        (_KEY_DOC_SYNCED,   app,   "DocumentSynchronizedWithCentral"),
        (_KEY_FILE_EXP,     app,   "FileExported"),
        (_KEY_FILE_IMP,     app,   "FileImported"),
        (_KEY_VIEW_ACT,     uiapp, "ViewActivated"),
        (_KEY_DIALOG,       uiapp, "DialogBoxShowing"),
        (_KEY_IDLING,       uiapp, "Idling"),
    ]

    for key, obj, event_name in pairs:
        if _unreg(domain, key, obj, event_name):
            removed.append(event_name)

    # Limpa estado auxiliar do Idling
    domain.SetData(_KEY_IDLE_TIME, None)
    domain.SetData(_KEY_SEL_COUNT, 0)

    if removed:
        write_log("=== LOGGER PARADO: {} ===".format(", ".join(removed)))

    return bool(removed)


def is_running():
    return System.AppDomain.CurrentDomain.GetData(_KEY_DOC_CHANGED) is not None
