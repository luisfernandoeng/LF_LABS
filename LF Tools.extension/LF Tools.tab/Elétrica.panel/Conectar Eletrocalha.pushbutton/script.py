# -*- coding: utf-8 -*-
"""
Conectar Eletrocalha - Roteamento sequencial automático entre elementos.

Shift-Click: Abre menu de configuração (Tipo, Tamanho, Elevação)
Normal: Seleção sequencial de elementos

A ferramenta criará trechos retos de eletrocalha (nível mantido constante)
e o próprio Revit tratará de colocar as curvas/junções automaticamente.
"""

__title__ = 'Conectar\nEletrocalha'
__author__ = 'Luis Fernando'

# ╚══════════════════════════════════════════════════════════════╝

# =====================================================================
#  IMPORTS
# =====================================================================
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import OperationCanceledException
import System
from System.Collections.Generic import List
from collections import OrderedDict
import traceback

from pyrevit import forms           # script importado de forma lazy em load/save_config
from lf_utils import DebugLogger, get_script_config, save_script_config
# Instância global — usar `dbg` em todo o script
dbg = DebugLogger(False)
# Referências globais — preenchidas no início de execute_connection()
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document if uidoc else None

# =====================================================================
#  HELPERS DE NOME
# =====================================================================
def __get_name__(obj):
    try: return obj.Name
    except: pass
    try: return Element.Name.GetValue(obj)
    except: pass
    try:
        p = obj.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.HasValue: return p.AsString()
    except: pass
    return ""

# =====================================================================
#  CONFIGURAÇÕES
# =====================================================================
def load_config():
    return get_script_config(__commandpath__, defaults={
        'cabletray_type': '',
        'default_width':  '200',
        'default_height': '100',
        'default_offset': '3.00',  # metros
        'use_connector': True,
        'debug_mode': False,
    })

def save_config(settings):
    save_script_config(__commandpath__, settings)

class SettingsWindow(forms.WPFWindow):
    def __init__(self, xaml_file, doc, settings):
        forms.WPFWindow.__init__(self, xaml_file)
        self._doc = doc
        self._settings = settings
        self._saved = False
        
        self.tb_width.Text = settings.get('default_width', '200')
        self.tb_height.Text = settings.get('default_height', '100')
        self.tb_offset.Text = settings.get('default_offset', '3.00')
        self.chk_debug.IsChecked = settings.get('debug_mode', False)
        
        if hasattr(self, 'chk_use_connector'):
            self.chk_use_connector.IsChecked = settings.get('use_connector', True)
        
        types = self._get_cabletray_types()
        self._populate_combo(self.cb_type, types, settings.get('cabletray_type', ''))
        
        self.btn_save.Click += self._on_save
        self.btn_cancel.Click += self._on_cancel

    def _get_cabletray_types(self):
        names = []
        for t in FilteredElementCollector(self._doc).OfClass(clr.GetClrType(CableTrayType)):
            n = __get_name__(t)
            if n:
                names.append(n)
        return ['(Padrão do Revit)'] + sorted(names, key=lambda x: x.lower())

    def _populate_combo(self, combo, items, selected):
        combo.Items.Clear()
        for item in items:
            combo.Items.Add(item)
        combo.SelectedItem = selected if selected in items else items[0]

    def _on_save(self, _sender, _args):
        s = self._settings
        
        sel = self.cb_type.SelectedItem
        s['cabletray_type'] = sel if sel else ''
        s['default_width'] = self.tb_width.Text.strip()
        s['default_height'] = self.tb_height.Text.strip()
        s['default_offset'] = self.tb_offset.Text.strip()
        try:
            if hasattr(self, 'chk_use_connector'):
                s['use_connector'] = bool(self.chk_use_connector.IsChecked)
        except:
            s['use_connector'] = True
        try:
            s['debug_mode'] = bool(self.chk_debug.IsChecked)
        except:
            s['debug_mode'] = False
            
        self._saved = True
        self.Close()

    def _on_cancel(self, _sender, _args):
        self.Close()

    def show(self):
        self.ShowDialog()
        return self._settings if self._saved else None


def show_settings():
    settings = load_config()
    import os
    xaml_path = os.path.join(os.path.dirname(__file__), 'settings.xaml')
    win = SettingsWindow(xaml_path, doc, settings)
    result = win.show()
    if result is not None:
        save_config(result)
        forms.toast("Configurações salvas!", title="Conectar Eletrocalha")

# =====================================================================
#  LÓGICA E UTILIDADES
# =====================================================================
def get_connectors(element):
    """Obtém conectores do domínio CableTray de um elemento."""
    connectors = []
    try:
        mgr = None
        if hasattr(element, "MEPModel") and element.MEPModel and element.MEPModel.ConnectorManager:
            mgr = element.MEPModel.ConnectorManager
        elif hasattr(element, "ConnectorManager") and element.ConnectorManager:
            mgr = element.ConnectorManager
        if mgr:
            for c in mgr.Connectors:
                if c.Domain == Domain.DomainCableTrayConduit:
                    # Filtra apenas conectores retangulares para ignorar eletrodutos (redondos)
                    try:
                        if c.Shape == ConnectorProfileType.Rectangular:
                            connectors.append(c)
                    except Exception:
                        connectors.append(c)
    except Exception:
        pass
    return connectors

def get_element_point(element, reference_point=None):
    """Tenta achar o ponto ideal de um elemento (conector mais próximo ou Location)."""
    conns = get_connectors(element)
    if conns:
        # Priorizar conectores que ainda não estão conectados a nada
        free_conns = [c for c in conns if not c.IsConnected]
        valid_conns = free_conns if free_conns else conns
        
        if reference_point:
            best_conn = min(valid_conns, key=lambda c: c.Origin.DistanceTo(reference_point))
            return best_conn.Origin, best_conn
        else:
            return valid_conns[0].Origin, valid_conns[0]
    try:
        loc = element.Location
        if isinstance(loc, LocationPoint):
            return loc.Point, None
        elif isinstance(loc, LocationCurve):
            return loc.Curve.Evaluate(0.5, True), None
    except Exception:
        pass
    try:
        bbox = element.get_BoundingBox(None)
        if bbox:
            return (bbox.Min + bbox.Max) / 2.0, None
    except Exception:
        pass
    return None, None

def get_default_cabletray_type():
    try:
        default_id = doc.GetDefaultElementTypeId(ElementTypeGroup.CableTrayType)
        if default_id != ElementId.InvalidElementId:
            return default_id
    except Exception:
        pass
    col = FilteredElementCollector(doc).OfClass(clr.GetClrType(CableTrayType))
    return col.FirstElementId()

# WarningSwallower vem de lf_utils.make_warning_swallower() — importado no topo

def get_connector_dimensions(conn):
    try:
        if conn.Shape == ConnectorProfileType.Rectangular:
            return conn.Width, conn.Height
        elif conn.Shape == ConnectorProfileType.Round:
            return conn.Radius * 2.0, conn.Radius * 2.0
        elif conn.Shape == ConnectorProfileType.Oval:
            return conn.Width, conn.Height
    except Exception:
        pass
    return None, None

# =====================================================================
#  EXECUÇÃO PRINCIPAL
# =====================================================================
def execute_connection():
    global uidoc, doc
    if not uidoc or not doc:
        forms.alert("Nenhum documento ativo encontrado.", title="Conectar Eletrocalha")
        return

    try:
        is_shift = __shiftclick__
    except NameError:
        is_shift = False

    if is_shift:
        show_settings()
        return

    dbg.section("Conectar Eletrocalha — Início")
    dbg.timer_start("total")

    settings = load_config()
    dbg.dump("settings", settings)

    # ── 1. Seleção de elementos ──────────────────────────────────────
    dbg.section("Fase 1: Seleção de Elementos")
    picked_elements = []
    
    selected_ids = uidoc.Selection.GetElementIds()
    if selected_ids:
        for eid in selected_ids:
            el = doc.GetElement(eid)
            if el:
                picked_elements.append(el)
    else:
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecione os dispositivos/eletrocalhas a conectar")
            for ref in refs:
                el = doc.GetElement(ref)
                if el:
                    picked_elements.append(el)
        except OperationCanceledException:
            pass

    if len(picked_elements) < 2:
        forms.alert(
            "Selecione pelo menos 2 elementos para formar uma rota.",
            title="Aviso"
        )
        return

    # Ordenar por proximidade para criar uma sequência lógica
    def sort_by_proximity(elements):
        if len(elements) <= 2:
            return elements
        
        pts = {}
        for el in elements:
            pt, _ = get_element_point(el)
            pts[el.Id] = pt if pt else XYZ.Zero
                
        # Achar o par mais distante para identificar as extremidades
        max_d = -1
        start_el = elements[0]
        for e1 in elements:
            for e2 in elements:
                if e1.Id != e2.Id:
                    d = pts[e1.Id].DistanceTo(pts[e2.Id])
                    if d > max_d:
                        max_d = d
                        start_el = e1
                        
        sorted_els = [start_el]
        remaining = [e for e in elements if e.Id != start_el.Id]
        
        current = start_el
        while remaining:
            closest = None
            min_d = float('inf')
            for r in remaining:
                d = pts[current.Id].DistanceTo(pts[r.Id])
                if d < min_d:
                    min_d = d
                    closest = r
            sorted_els.append(closest)
            remaining.remove(closest)
            current = closest
            
        return sorted_els

    picked_elements = sort_by_proximity(picked_elements)
    dbg.info("Ordem definida para {} elementos.".format(len(picked_elements)))

    # ── 2. Resolver tipo e parâmetros ────────────────────────────
    dbg.section("Fase 2: Parâmetros")

    pref_type_name = settings.get('cabletray_type', '')
    ct_type_id = None
    if pref_type_name and pref_type_name != "(Padrão do Revit)":
        for t in FilteredElementCollector(doc).OfClass(clr.GetClrType(CableTrayType)):
            if __get_name__(t) == pref_type_name:
                ct_type_id = t.Id
                break
        dbg.result(ct_type_id is not None,
                   "Tipo '{}' encontrado: Id={}".format(pref_type_name, ct_type_id))

    if not ct_type_id:
        ct_type_id = get_default_cabletray_type()
        dbg.debug("Tipo preferido não encontrado. Usando default: Id={}".format(ct_type_id))

    if not ct_type_id or ct_type_id == ElementId.InvalidElementId:
        forms.alert(
            "Nenhum tipo de eletrocalha encontrado no projeto.\n"
            "Carregue um tipo de eletrocalha antes de usar esta ferramenta.",
            title="Erro"
        )
        return

    def _parse_float(val, default_val):
        try:
            if val is None or str(val).strip() == '':
                return float(default_val)
            return float(str(val).replace(',', '.'))
        except Exception:
            return float(default_val)

    width_ft  = _parse_float(settings.get('default_width'), 200) / 304.8
    height_ft = _parse_float(settings.get('default_height'), 100) / 304.8
    offset_ft = _parse_float(settings.get('default_offset'), 3.00) / 0.3048

    dbg.debug("ct_type_id = {}".format(ct_type_id))
    dbg.debug("width_ft   = {:.6f}  ({} mm)".format(width_ft,  settings.get('default_width')))
    dbg.debug("height_ft  = {:.6f}  ({} mm)".format(height_ft, settings.get('default_height')))
    dbg.debug("offset_ft  = {:.6f}  ({} m)".format(offset_ft,  settings.get('default_offset')))

    # ── 3. Nível de referência ───────────────────────────────────
    base_level_id = picked_elements[0].LevelId
    if base_level_id == ElementId.InvalidElementId:
        view = doc.ActiveView
        if hasattr(view, "GenLevel") and view.GenLevel:
            base_level_id = view.GenLevel.Id
        else:
            base_level_id = FilteredElementCollector(doc).OfClass(Level).FirstElementId()
        dbg.debug("Elemento 0 sem LevelId. Usando nível da view: Id={}".format(base_level_id))

    base_level      = doc.GetElement(base_level_id)
    level_elevation = base_level.Elevation if base_level else 0.0
    dbg.debug("Nível: Id={}  elevation={:.4f} ft  Z_alvo={:.4f} ft".format(
        base_level_id, level_elevation, level_elevation + offset_ft))

    # ── 4. Transação ─────────────────────────────────────────────
    dbg.section("Fase 3: Criação das Eletrocalhas")
    t = Transaction(doc, "Conectar Eletrocalhas Inteligente")

    t.Start()

    try:
        drawn_trays    = []
        creation_errors = []
        true_z = level_elevation + offset_ft

        for i in range(len(picked_elements) - 1):
            el1 = picked_elements[i]
            el2 = picked_elements[i + 1]

            dbg.sub("Par {}/{}: Id={} → Id={}".format(
                i + 1, len(picked_elements) - 1, el1.Id, el2.Id))

            temp_pt_target, _ = get_element_point(el2)
            temp_pt_start, _  = get_element_point(el1)
            pt1, conn1 = get_element_point(el1, temp_pt_target)
            pt2, conn2 = get_element_point(el2, temp_pt_start)

            dbg.xyz("  pt1 (original)", pt1)
            dbg.xyz("  pt2 (original)", pt2)
            dbg.debug("  conn1={}  conn2={}".format(
                "OK" if conn1 else "None (Location fallback)",
                "OK" if conn2 else "None (Location fallback)"))

            if not pt1 or not pt2:
                msg = "Par {}: Coordenadas inválidas — pt1={} pt2={}".format(i + 1, pt1, pt2)
                dbg.error(msg)
                creation_errors.append(msg)
                continue

            # Usa Z real do conector se o elemento é MEP (CableTray, painel, etc.)
            # Só aplica offset configurado para elementos sem conector (ex: mobiliário, genéricos)
            z1 = pt1.Z if conn1 is not None else true_z
            z2 = pt2.Z if conn2 is not None else true_z
            route_pt1 = XYZ(pt1.X, pt1.Y, z1)
            route_pt2 = XYZ(pt2.X, pt2.Y, z2)
            dist = route_pt1.DistanceTo(route_pt2)

            dbg.xyz("  route_pt1 (Z final)", route_pt1)
            dbg.xyz("  route_pt2 (Z final)", route_pt2)
            dbg.debug("  z1={:.4f} ft ({}) z2={:.4f} ft ({})".format(
                z1, "conector" if conn1 else "offset cfg",
                z2, "conector" if conn2 else "offset cfg"))
            dbg.debug("  dist = {:.4f} ft  ({:.3f} m)".format(dist, dist * 0.3048))

            if dist < 0.1:
                msg = "Par {}: Pontos muito próximos ({:.4f} ft). Pulando.".format(i + 1, dist)
                dbg.debug(msg)
                creation_errors.append(msg)
                continue

            # ── Detectar tamanho das bandejas vizinhas ────────────────
            # Prioridade: conector 1 → conector 2 → parâmetros de el1 → el2 → configurações
            pair_w = None
            pair_h = None
            
            use_conn = settings.get('use_connector', True)

            if use_conn and conn1:
                cw, ch = get_connector_dimensions(conn1)
                if cw and ch:
                    pair_w, pair_h = cw, ch
                    
            if use_conn and not pair_w and conn2:
                cw, ch = get_connector_dimensions(conn2)
                if cw and ch:
                    pair_w, pair_h = cw, ch

            def _read_dim(el, bip):
                try:
                    p = el.get_Parameter(bip)
                    if p and p.HasValue and p.AsDouble() > 0:
                        return p.AsDouble()
                except Exception:
                    pass
                return None

            if not pair_w:
                pair_w = (_read_dim(el1, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                       or _read_dim(el2, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                       or width_ft)
                pair_h = (_read_dim(el1, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                       or _read_dim(el2, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                       or height_ft)
            dbg.debug("  Tamanho par: w={:.4f} ft  h={:.4f} ft".format(pair_w, pair_h))

            try:
                dbg.timer_start("CableTray.Create #{}".format(i))
                ctray = CableTray.Create(doc, ct_type_id, route_pt1, route_pt2, base_level_id)
                dbg.timer_end("CableTray.Create #{}".format(i))

                if ctray is None:
                    msg = "Par {}: CableTray.Create retornou None.".format(i + 1)
                    dbg.error(msg)
                    creation_errors.append(msg)
                    continue

                # ── Definir dimensões da nova bandeja ─────────────────
                p_w = ctray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                if p_w and not p_w.IsReadOnly:
                    p_w.Set(pair_w)
                    dbg.debug("  Width setado: {:.4f} ft".format(pair_w))
                else:
                    dbg.debug("  Width param ausente ou read-only.")

                p_h = ctray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                if p_h and not p_h.IsReadOnly:
                    p_h.Set(pair_h)
                    dbg.debug("  Height setado: {:.4f} ft".format(pair_h))
                else:
                    dbg.debug("  Height param ausente ou read-only.")

                # ── Conectar fisicamente aos elementos originais ───────
                # CableTray.Create posiciona geometricamente mas NÃO
                # estabelece a conexão lógica MEP — precisa de ConnectTo explícito.
                new_conns = get_connectors(ctray)
                dbg.debug("  Conectores da nova bandeja: {}".format(len(new_conns)))

                def _try_connect(new_c, orig_c, label):
                    if orig_c is None:
                        return
                    dist_c = new_c.Origin.DistanceTo(orig_c.Origin)
                    dbg.debug("  {} dist={:.4f} ft  new_connected={} orig_connected={}".format(
                        label, dist_c, new_c.IsConnected, orig_c.IsConnected))
                    if dist_c > 1.0:
                        dbg.debug("  {} distância alta ({:.4f} ft) — ConnectTo ignorado.".format(
                            label, dist_c))
                        return
                    if orig_c.IsConnected:
                        dbg.debug("  {} orig_c já está conectado — pulando.".format(label))
                        return
                    try:
                        new_c.ConnectTo(orig_c)
                        dbg.result(True, "  {} ConnectTo OK".format(label))
                    except Exception as ce:
                        dbg.debug("  {} ConnectTo falhou: {}".format(label, ce))

                if new_conns and len(new_conns) >= 2:
                    # Associar conector da nova bandeja ao conector mais próximo de cada extremidade
                    c_near_1 = min(new_conns, key=lambda c: c.Origin.DistanceTo(route_pt1))
                    c_near_2 = min(new_conns, key=lambda c: c.Origin.DistanceTo(route_pt2))
                    _try_connect(c_near_1, conn1, "conn→el1")
                    _try_connect(c_near_2, conn2, "conn→el2")
                else:
                    dbg.debug("  Nova bandeja com <2 conectores — conexão pulada.")

                drawn_trays.append(ctray)
                dbg.result(True, "Eletrocalha criada: Id={}".format(ctray.Id))

            except Exception as ex:
                msg = "Par {}: EXCEÇÃO em CableTray.Create — {}".format(i + 1, ex)
                dbg.error(msg)
                dbg.error(traceback.format_exc())
                creation_errors.append("{}: {}".format(msg, traceback.format_exc()))

        # ── 5. Conectar com elbows ────────────────────────────────
        if len(drawn_trays) > 1:
            dbg.section("Fase 4: Conexão com Elbows/Fittings")
            for i in range(len(drawn_trays) - 1):
                tray1 = drawn_trays[i]
                tray2 = drawn_trays[i + 1]

                c1_candidates = get_connectors(tray1)
                c2_candidates = get_connectors(tray2)

                dbg.debug("Trays {}/{}: connectors tray1={} tray2={}".format(
                    i + 1, len(drawn_trays) - 1,
                    len(c1_candidates), len(c2_candidates)))

                if not c1_candidates or not c2_candidates:
                    dbg.debug("Sem conectores disponíveis para par {}.".format(i + 1))
                    continue

                best_c1, best_c2, min_dist = None, None, float('inf')
                for c1 in c1_candidates:
                    if c1.IsConnected: continue
                    for c2 in c2_candidates:
                        if c2.IsConnected: continue
                        d = c1.Origin.DistanceTo(c2.Origin)
                        if d < min_dist:
                            min_dist, best_c1, best_c2 = d, c1, c2

                dbg.debug("  Melhor par de conectores: dist={:.4f} ft".format(min_dist))

                if best_c1 and best_c2 and min_dist < 2.0:
                    try:
                        doc.Create.NewElbowFitting(best_c1, best_c2)
                        dbg.result(True, "Elbow/fitting criado para par {}.".format(i + 1))
                    except Exception as ex:
                        dbg.debug("NewElbowFitting falhou ({}). Tentando ConnectTo...".format(ex))
                        try:
                            best_c1.ConnectTo(best_c2)
                            dbg.result(True, "ConnectTo OK para par {}.".format(i + 1))
                        except Exception as ex2:
                            dbg.error("ConnectTo também falhou: {}".format(ex2))
                else:
                    dbg.debug("Par {}: dist={:.4f} ft > 2.0 ft. Nenhuma curva criada.".format(
                        i + 1, min_dist))

        # ── Commit ────────────────────────────────────────────────
        t.Commit()
        dbg.section("Resultado")
        dbg.info("Transação commitada.")
        dbg.info("Eletrocalhas criadas: {}".format(len(drawn_trays)))
        if creation_errors:
            dbg.debug("Erros registrados durante a criação:")
            for idx, err in enumerate(creation_errors, 1):
                dbg.debug("  [{}] {}".format(idx, err))

        if not drawn_trays:
            err_detail = "\n\n".join(creation_errors) if creation_errors else "Sem detalhes."
            forms.alert(
                "Nenhuma eletrocalha foi criada.\n\nErros:\n{}".format(err_detail),
                title="Conectar Eletrocalha — Sem Resultado"
            )
        else:
            dbg.result(True, "{} eletrocalha(s) criada(s) com sucesso.".format(len(drawn_trays)))

    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        raise e

    dbg.timer_end("total")


def safe_execution():
    global dbg
    settings = load_config()
    dbg = DebugLogger(settings.get('debug_mode', False))
    
    dbg.section("Conectar Eletrocalha — BOOT")
    dbg.info("DEBUG_MODE = {}".format(settings.get('debug_mode', False)))
    try:
        execute_connection()
        dbg.section("Ferramenta Finalizada")
    except Exception as e:
        err_tb = traceback.format_exc()
        dbg.error("CRASH FATAL:\n{}".format(err_tb))
        if settings.get('debug_mode', False):
            forms.alert("CRASH FATAL (DEBUG):\n\n" + err_tb, title="Conectar Eletrocalha — Erro")
        else:
            forms.alert("Erro ao conectar eletrocalhas:\n" + str(e), title="Aviso")


if __name__ == "__main__":
    if __shiftclick__:
        show_settings()
    else:
        safe_execution()
