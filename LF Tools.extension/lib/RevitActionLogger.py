# -*- coding: utf-8 -*-
"""
RevitActionLogger — LF Tools
Monitora eventos da API do Revit e grava log estruturado em arquivo de texto.
"""
import System
import os
import codecs
import time
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

# Noise controls. View/selection changes are expensive and usually do not
# explain a failing Revit operation, so keep them off unless debugging that.
_LOG_VIEW_CHANGES = False
_LOG_SELECTION_CHANGES = False
_LOG_DIALOGS = True
_LOG_ONLY_RELEVANT_CATEGORIES = True
_MAX_ELEMENTS_PER_TX = 25
_SUMMARY_SAMPLE_LIMIT = 100
_PERF_WARN_MS = 300

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
_KEY_LAST_ERROR   = "LF_ActionLogger_LastError"

# ================================================================
# ESCRITA NO LOG
# ================================================================

def write_log(message):
    """Grava uma linha no log com timestamp, tolerante a falhas."""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with codecs.open(LOG_FILE, "a", "utf-8") as f:
            f.write("[{}] {}\n".format(timestamp, message))
    except Exception as e:
        _remember_error("write_log: " + str(e))

def _remember_error(message):
    try:
        System.AppDomain.CurrentDomain.SetData(_KEY_LAST_ERROR, message)
    except Exception:
        pass

# ================================================================
# HELPERS DE ELEMENTO
# ================================================================

def _cat_ids(names):
    ids = []
    for name in names:
        try:
            ids.append(int(getattr(BuiltInCategory, name)))
        except Exception:
            pass
    return ids

_ELEC_CATS = _cat_ids([
    "OST_ElectricalEquipment",
    "OST_ElectricalFixtures",
    "OST_LightingFixtures",
    "OST_LightingDevices",
    "OST_ElectricalCircuit",
    "OST_Wire",
])

_RELEVANT_CATS = _cat_ids([
    "OST_ElectricalEquipment",
    "OST_ElectricalFixtures",
    "OST_LightingFixtures",
    "OST_LightingDevices",
    "OST_ElectricalCircuit",
    "OST_Wire",
    "OST_Conduit",
    "OST_ConduitFitting",
    "OST_CableTray",
    "OST_CableTrayFitting",
    "OST_CableTrayRun",
])

def _is_electrical(elem):
    try:
        return elem and elem.Category and elem.Category.Id.IntegerValue in _ELEC_CATS
    except Exception:
        return False

def _is_relevant(elem):
    if not _LOG_ONLY_RELEVANT_CATEGORIES:
        return True
    try:
        return elem and elem.Category and elem.Category.Id.IntegerValue in _RELEVANT_CATS
    except Exception:
        return False

def _category_name(elem):
    try:
        return elem.Category.Name if elem and elem.Category else "Unknown"
    except Exception:
        return "Unknown"

def _category_counts(doc, ids):
    counts = {}
    for idx, e_id in enumerate(ids):
        if idx >= _SUMMARY_SAMPLE_LIMIT:
            counts["..."] = counts.get("...", 0) + 1
            break
        try:
            elem = doc.GetElement(e_id)
            cat = _category_name(elem)
            counts[cat] = counts.get(cat, 0) + 1
        except Exception:
            counts["Unknown"] = counts.get("Unknown", 0) + 1
    return counts

def _format_counts(counts):
    try:
        return ", ".join(["{}:{}".format(k, counts[k]) for k in sorted(counts.keys())])
    except Exception:
        return ""

def _is_cabletray_related(elem):
    try:
        if not elem or not elem.Category:
            return False
        cat_name = (elem.Category.Name or "").lower()
        if "bandeja" in cat_name or "cable tray" in cat_name:
            return True
        bic = elem.Category.Id.IntegerValue
        return bic in [
            int(BuiltInCategory.OST_CableTray),
            int(BuiltInCategory.OST_CableTrayFitting),
        ]
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

def _xyz_text(pt):
    try:
        return "({:.4f}, {:.4f}, {:.4f})".format(pt.X, pt.Y, pt.Z)
    except Exception:
        return "(?)"

def _connector_manager(elem):
    try:
        return elem.ConnectorManager
    except Exception:
        pass
    try:
        if elem.MEPModel:
            return elem.MEPModel.ConnectorManager
    except Exception:
        pass
    return None

def _mep_curve_context(elem):
    """Log rico para bandejas/conexoes: curva, dimensoes e conectores."""
    try:
        lines = []

        try:
            loc = elem.Location
            if loc and hasattr(loc, "Curve") and loc.Curve:
                crv = loc.Curve
                lines.append("    Curva: {} -> {} | L={:.4f} ft".format(
                    _xyz_text(crv.GetEndPoint(0)),
                    _xyz_text(crv.GetEndPoint(1)),
                    crv.Length))
        except Exception:
            pass

        for label, bip in [
            ("Largura", BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM),
            ("Altura", BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM),
            ("Diametro", BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM),
            ("Comprimento", BuiltInParameter.CURVE_ELEM_LENGTH),
        ]:
            try:
                p = elem.get_Parameter(bip)
                if p and p.HasValue:
                    val = p.AsValueString()
                    if not val:
                        try:
                            val = "{:.1f} mm".format(p.AsDouble() * 304.8)
                        except Exception:
                            val = str(p.AsInteger())
                    lines.append("    {}: {}".format(label, val))
            except Exception:
                pass

        cm = _connector_manager(elem)
        if cm:
            try:
                conns = list(cm.Connectors)
                lines.append("    Conectores: {}".format(len(conns)))
                for idx, c in enumerate(conns[:12]):
                    try:
                        diam = ""
                        try:
                            if hasattr(c, "Radius"):
                                diam = " D={:.1f}mm".format(c.Radius * 2.0 * 304.8)
                        except Exception:
                            pass
                        try:
                            bz = c.CoordinateSystem.BasisZ
                            dir_txt = " Dir=({:.3f},{:.3f},{:.3f})".format(bz.X, bz.Y, bz.Z)
                        except Exception:
                            dir_txt = ""
                        refs = []
                        try:
                            for r in c.AllRefs:
                                try:
                                    if r.Owner and r.Owner.Id != elem.Id:
                                        refs.append(str(r.Owner.Id.IntegerValue))
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        ref_txt = " Refs=[{}]".format(",".join(refs)) if refs else ""
                        lines.append("      C{}: {} Shape={} Tipo={}{}{}{}".format(
                            idx,
                            _xyz_text(c.Origin),
                            c.Shape,
                            c.ConnectorType,
                            diam,
                            dir_txt,
                            ref_txt))
                    except Exception:
                        pass
            except Exception:
                pass

        return "\n".join(lines)
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
    t0 = time.time()
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

        skipped_added = 0
        skipped_modified = 0

        for idx, e_id in enumerate(added_ids):
            if idx >= _MAX_ELEMENTS_PER_TX:
                skipped_added += 1
                continue
            elem = doc.GetElement(e_id)
            if not elem: continue
            if not _is_relevant(elem):
                skipped_added += 1
                continue
            cat = _category_name(elem)
            msgs.append("  [ADICIONADO] ID:{} | Cat:{} | Nome:{}".format(
                e_id.IntegerValue, cat, elem.Name))
            if _is_cabletray_related(elem):
                detail = _mep_curve_context(elem) or _elem_summary(elem)
            else:
                detail = _elec_context(elem) if _is_electrical(elem) else _elem_summary(elem)
            if detail: msgs.append(detail)

        for idx, e_id in enumerate(modified_ids):
            if idx >= _MAX_ELEMENTS_PER_TX:
                skipped_modified += 1
                continue
            elem = doc.GetElement(e_id)
            if not elem: continue
            if not _is_relevant(elem):
                skipped_modified += 1
                continue
            cat = _category_name(elem)
            msgs.append("  [MODIFICADO] ID:{} | Cat:{} | Nome:{}".format(
                e_id.IntegerValue, cat, elem.Name))
            if _is_cabletray_related(elem):
                detail = _mep_curve_context(elem)
                if detail: msgs.append(detail)
            elif _is_electrical(elem):
                detail = _elec_context(elem)
                if detail: msgs.append(detail)

        if skipped_added:
            txt = _format_counts(_category_counts(doc, added_ids))
            msgs.append("  [ADICIONADO] {} item(ns) omitido(s) por filtro/limite{}".format(
                skipped_added, " | " + txt if txt else ""))
        if skipped_modified:
            txt = _format_counts(_category_counts(doc, modified_ids))
            msgs.append("  [MODIFICADO] {} item(ns) omitido(s) por filtro/limite{}".format(
                skipped_modified, " | " + txt if txt else ""))

        if len(msgs) > 1:
            write_log("\n".join(msgs))

        elapsed_ms = int((time.time() - t0) * 1000)
        if elapsed_ms >= _PERF_WARN_MS:
            write_log("  [PERF] OnDocumentChanged levou {} ms | tx='{}'".format(
                elapsed_ms, tx_name))

    except Exception as e:
        _remember_error("OnDocumentChanged: " + str(e))
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
    if not _LOG_VIEW_CHANGES:
        return
    try:
        if args.CurrentActiveView:
            write_log("--- VISTA ATIVADA: '{}' ---".format(
                args.CurrentActiveView.Name))
    except Exception:
        pass

def OnDialogBoxShowing(sender, args):
    if not _LOG_DIALOGS:
        return
    """Captura dialogs e avisos que o Revit exibe durante operações."""
    try:
        dialog_id = ""
        try: dialog_id = str(args.DialogId)
        except Exception: pass
        write_log("  [DIALOG] Id: '{}'".format(dialog_id))
    except Exception:
        pass

def OnIdling(sender, args):
    if not _LOG_SELECTION_CHANGES:
        return
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

    domain.SetData(_KEY_LAST_ERROR, None)
    write_log("=== LOGGER INICIADO: {} ===".format(", ".join(ok)))
    write_log("=== CONFIG: relevant_only={} max_elements={} view={} selection={} dialogs={} log='{}' ===".format(
        _LOG_ONLY_RELEVANT_CATEGORIES,
        _MAX_ELEMENTS_PER_TX,
        _LOG_VIEW_CHANGES,
        _LOG_SELECTION_CHANGES,
        _LOG_DIALOGS,
        LOG_FILE))
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
    domain.SetData(_KEY_LAST_ERROR, None)

    if removed:
        write_log("=== LOGGER PARADO: {} ===".format(", ".join(removed)))

    return bool(removed)


def is_running():
    return System.AppDomain.CurrentDomain.GetData(_KEY_DOC_CHANGED) is not None

def get_log_file():
    return LOG_FILE

def get_status_text():
    domain = System.AppDomain.CurrentDomain
    last_error = domain.GetData(_KEY_LAST_ERROR)
    return "\n".join([
        "Status: " + ("rodando" if is_running() else "parado"),
        "Log: " + LOG_FILE,
        "Filtro de categorias relevantes: " + str(_LOG_ONLY_RELEVANT_CATEGORIES),
        "Limite por transacao: " + str(_MAX_ELEMENTS_PER_TX),
        "Log de vista: " + str(_LOG_VIEW_CHANGES),
        "Log de selecao: " + str(_LOG_SELECTION_CHANGES),
        "Log de dialogs: " + str(_LOG_DIALOGS),
        "Ultimo erro interno: " + (str(last_error) if last_error else "nenhum"),
    ])
