# coding: utf-8
"""
LF Electrical Core
Compartilhado entre Residencial, Industrial e Dados
"""

import os
import sys
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, script
import traceback
from collections import OrderedDict

# Queda de Tensão: inserir lib no path
lib_path = os.path.dirname(__file__)
if lib_path not in sys.path:
    sys.path.append(lib_path)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Letras proibidas na nomenclatura
SKIP_LETTERS = {'O', 'S'}
VALID_SWITCH_LETTERS = [chr(ord('a') + i) for i in range(26) if chr(ord('a') + i) not in ('o', 's')]

# ==================== SISTEMA DE DEBUG ====================

DEBUG_MODE = True  # <<< Mude para False quando tudo estiver estável

import System as _System
from datetime import datetime as _dt

_desktop = _System.Environment.GetFolderPath(_System.Environment.SpecialFolder.Desktop)
_LOG_FILE = os.path.join(_desktop, "LF_Electrical_Debug.txt")

class DebugLog(object):
    """Logger de debug que imprime no console do pyRevit e salva em arquivo."""

    _indent = 0

    @classmethod
    def _write(cls, level, msg):
        if not DEBUG_MODE:
            return
        prefix = "  " * cls._indent
        timestamp = _dt.now().strftime("%H:%M:%S.%f")[:-3]
        line = "[{}][{}] {}{}".format(timestamp, level, prefix, msg)
        # Console do pyRevit (aparece na janela de output)
        try:
            print(line)
        except Exception:
            pass
        # Arquivo no Desktop
        try:
            with open(_LOG_FILE, "a") as f:
                f.write(line + "\n")
        except Exception:
            pass

    @classmethod
    def section(cls, title):
        cls._write("====", "")
        cls._write("====", ">>> {} <<<".format(title))
        cls._write("====", "")

    @classmethod
    def enter(cls, func_name, **kwargs):
        details = ", ".join("{}={}".format(k, v) for k, v in kwargs.items()) if kwargs else ""
        cls._write("ENTR", "-> {}({})".format(func_name, details))
        cls._indent += 1

    @classmethod
    def exit(cls, func_name, result=None):
        cls._indent = max(0, cls._indent - 1)
        cls._write("EXIT", "<- {} => {}".format(func_name, result))

    @classmethod
    def step(cls, msg):
        cls._write("STEP", msg)

    @classmethod
    def ok(cls, msg):
        cls._write(" OK ", msg)

    @classmethod
    def warn(cls, msg):
        cls._write("WARN", "*** {} ***".format(msg))

    @classmethod
    def fail(cls, msg):
        cls._write("FAIL", "!!! {} !!!".format(msg))

    @classmethod
    def elem_info(cls, elem, label=""):
        """Loga informações chave de um elemento do Revit."""
        try:
            cat = elem.Category.Name if elem.Category else "SemCategoria"
            eid = elem.Id.IntegerValue
            fname = ""
            try:
                if hasattr(elem, 'Symbol') and elem.Symbol:
                    fname = elem.Symbol.FamilyName or ""
            except Exception:
                pass
            cls._write("ELEM", "{} Id={} Cat='{}' Fam='{}'".format(label, eid, cat, fname))
        except Exception as ex:
            cls._write("ELEM", "{} (erro ao inspecionar: {})".format(label, ex))

    @classmethod
    def param_check(cls, elem, param_name, found, value=None):
        if found:
            cls._write("PARM", "  [{}] = {}".format(param_name, value))
        else:
            cls._write("PARM", "  [{}] NAO ENCONTRADO".format(param_name))

    @classmethod
    def elem_full(cls, elem, label=""):
        """Relatorio completo: categoria, familia, conectores (com domain/IsConnected/MEPSystem/Poles/Volt),
        sub-componentes e parametros eletricos chave. Ativo apenas com DEBUG_MODE=True."""
        if not DEBUG_MODE:
            return
        try:
            cat = elem.Category.Name if elem.Category else "?"
            eid = elem.Id.IntegerValue
            fname, tname = "", ""
            try:
                if hasattr(elem, 'Symbol') and elem.Symbol:
                    fname = getattr(elem.Symbol, 'FamilyName', '') or ''
                    tname = getattr(elem.Symbol, 'Name', '') or ''
            except Exception:
                pass
            cls._write("DEBG", "{} Id={} Cat='{}' Fam='{}' Tipo='{}'".format(
                label, eid, cat, fname, tname))

            # ConnectorManager
            mgr = None
            try:
                if hasattr(elem, 'MEPModel') and elem.MEPModel:
                    mgr = getattr(elem.MEPModel, 'ConnectorManager', None)
                if not mgr and hasattr(elem, 'ConnectorManager'):
                    mgr = elem.ConnectorManager
            except Exception as ex:
                cls._write("DEBG", "  MEPModel/CM erro: {}".format(ex))

            if mgr:
                try:
                    conns = list(mgr.Connectors)
                except Exception:
                    conns = []
                cls._write("DEBG", "  Conectores totais: {}".format(len(conns)))
                for i, c in enumerate(conns):
                    try:
                        sys_obj = None
                        sys_id_str = "None"
                        try:
                            sys_obj = c.MEPSystem
                            sys_id_str = str(sys_obj.Id.IntegerValue) if sys_obj else "None"
                        except Exception:
                            sys_id_str = "ERRO"
                        poles_v, volt_v = "?", "?"
                        try: poles_v = int(c.Poles)
                        except Exception: pass
                        try: volt_v = "{}V".format(round(c.Voltage / 10.7639104167, 1))
                        except Exception: pass
                        cls._write("DEBG", "    [{i}] domain={d} connected={cn} MEPSystem={s} Poles={p} Volt={v}".format(
                            i=i, d=c.Domain, cn=c.IsConnected, s=sys_id_str, p=poles_v, v=volt_v))
                    except Exception as ex:
                        cls._write("DEBG", "    [{}] erro ao ler: {}".format(i, ex))
            else:
                cls._write("DEBG", "  Sem ConnectorManager")

            # Sub-componentes
            try:
                if hasattr(elem, 'GetSubComponentIds'):
                    sub_ids = elem.GetSubComponentIds()
                    if sub_ids and sub_ids.Count > 0:
                        cls._write("DEBG", "  Sub-componentes: {}".format(
                            [s.IntegerValue for s in sub_ids]))
                    else:
                        cls._write("DEBG", "  Sub-componentes: nenhum")
            except Exception:
                pass

            # Params eletricos chave
            for pname in ["N\xb0 de Fases", "N\xfamero de Fases", "Tens\xe3o (V)", "Tens\xe3o",
                          "Painel", "N\xfamero do circuito", "Sistema de distribui\xe7\xe3o"]:
                try:
                    p = elem.LookupParameter(pname)
                    if p and p.HasValue:
                        val = p.AsString() or p.AsValueString() or str(round(p.AsDouble(), 4))
                        ro = " [RO]" if p.IsReadOnly else ""
                        cls._write("DEBG", "  [{}]{} = {}".format(pname, ro, val))
                except Exception:
                    pass
        except Exception as ex:
            cls._write("DEBG", "{} (erro no relatorio: {})".format(label, ex))

dbg = DebugLog  # alias curto

# ==================== DIALOG HANDLER ====================

def _elec_circuit_dialog_handler(sender, args):
    """Auto-confirma o dialog de especificacao de circuito que aparece quando
    os params de tensao/fases sao somente-leitura na familia."""
    try:
        if hasattr(args, 'DialogId') and 'SpecifyCircuitInfo' in str(args.DialogId):
            dbg.step('AutoConfirm dialog: {}'.format(args.DialogId))
            args.OverrideResult(1)  # 1 = OK / Continuar
    except Exception:
        pass


class suppress_elec_dialog(object):
    """Context manager que assina DialogBoxShowing para auto-confirmar o dialog
    'Dialog_BuildingSystems_SpecifyCircuitInfo' durante ElectricalSystem.Create().

    Esse dialog aparece quando o Revit nao consegue determinar tensao/fases a
    partir dos parametros do elemento (params somente-leitura em familias de
    fabricante). Auto-confirmar com OK permite que a criacao prossiga usando
    o DistributionSysType do quadro para definir a tensao do circuito.

    Uso:
        with suppress_elec_dialog():
            circuit = ElectricalSystem.Create(doc, ids, PowerCircuit)
    """
    def __enter__(self):
        try:
            __revit__.DialogBoxShowing += _elec_circuit_dialog_handler
        except Exception as ex:
            dbg.warn('suppress_elec_dialog: falha ao subscrever: {}'.format(ex))
        return self

    def __exit__(self, *args):
        try:
            __revit__.DialogBoxShowing -= _elec_circuit_dialog_handler
        except Exception:
            pass


# Params de descricao em circuitos eletricos (PT-BR Revit)
CIRCUIT_DESC_PARAMS = [BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, "Coment\xe1rios", "Observa\xe7\xf5es", "Comments"]

# ==================== CONFIGURAÇÕES PERSISTENTES ====================

def load_config():
    config = script.get_config()
    defaults = {
        'session_panel_id': '',
        'dist_system': '',
        'naming_scheme': '',
        'circuit_prefix': '',
    }
    result = {}
    for key, default in defaults.items():
        try:
            result[key] = config.get_option(key, default)
        except Exception:
            result[key] = default
    return result

def save_config(settings):
    config = script.get_config()
    for key, value in settings.items():
        config.set_option(key, value)
    script.save_config()

def show_settings():
    settings = load_config()

    while True:
        current_dist = settings.get('dist_system', '') or '(nenhum)'
        current_naming = settings.get('naming_scheme', '') or '(nenhum)'
        current_prefix = settings.get('circuit_prefix', '') or '(nenhum)'

        opcoes = OrderedDict([
            ("1. Sistema de Distribuição Padrão: " + current_dist, "dist_system"),
            ("2. Nomenclatura de Circuito Padrão: " + current_naming, "naming_scheme"),
            ("3. Prefixo de Circuito Padrão: " + current_prefix, "circuit_prefix"),
            ("4. Salvar e Sair", "save"),
        ])

        escolha = forms.CommandSwitchWindow.show(
            opcoes.keys(),
            message="Configurações LF Electrical",
            title="Shift+Click - Configurações"
        )

        if not escolha or opcoes.get(escolha) == "save":
            save_config(settings)
            forms.toast("Configurações salvas!", title="LF Electrical")
            break

        key = opcoes[escolha]

        if key == 'dist_system':
            all_systems = []
            for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
                try:
                    name = s.Name
                except Exception:
                    p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    name = p.AsString() if p else "?"
                all_systems.append(name)

            if not all_systems:
                forms.alert("Nenhum sistema de distribuição encontrado no modelo.")
                continue

            chosen = forms.CommandSwitchWindow.show(sorted(all_systems), message="Selecione o sistema padrão:", title="Sistema de Distribuição")
            if chosen: settings['dist_system'] = chosen

        elif key == 'naming_scheme':
            all_schemes = []
            for s in FilteredElementCollector(doc).OfClass(CircuitNamingScheme).ToElements():
                all_schemes.append(s.Name)

            if not all_schemes:
                forms.alert("Nenhuma nomenclatura encontrada.")
                continue

            chosen = forms.CommandSwitchWindow.show(sorted(all_schemes), message="Selecione a nomenclatura padrão:", title="Nomenclatura")
            if chosen: settings['naming_scheme'] = chosen

        elif key == 'circuit_prefix':
            val = forms.ask_for_string(default=settings.get('circuit_prefix', ''), prompt="Prefixo (ex: QDL, QFL):", title="Prefixo de Circuito")
            if val is not None: settings['circuit_prefix'] = val


# ==================== FUNÇÕES AUXILIARES ====================

def get_panel_name(panel):
    try:
        for n in ["Nome do painel", "Panel Name", "Mark"]:
            p = panel.LookupParameter(n)
            if p and p.HasValue:
                return p.AsString()
        return panel.Name
    except Exception:
        return "Quadro"

def get_current_panel():
    dbg.enter('get_current_panel')
    settings = load_config()
    pid_str = settings.get('session_panel_id', '')
    if pid_str:
        try:
            panel = doc.GetElement(ElementId(int(pid_str)))
            if panel and panel.IsValidObject:
                dbg.ok('Quadro encontrado: Id={}'.format(pid_str))
                dbg.exit('get_current_panel', 'Id=' + pid_str)
                return panel
        except Exception:
            dbg.warn('Falha ao recuperar quadro Id={}'.format(pid_str))
    settings['session_panel_id'] = ''
    save_config(settings)
    dbg.warn('Nenhum quadro ativo')
    dbg.exit('get_current_panel', None)
    return None

def set_current_panel(panel_id):
    settings = load_config()
    settings['session_panel_id'] = str(panel_id.IntegerValue)
    save_config(settings)

def set_param(elem, names, value):
    found_readonly = []
    for n in names:
        try:
            p = elem.get_Parameter(n) if isinstance(n, BuiltInParameter) else elem.LookupParameter(n)
        except Exception:
            p = None
        if not p:
            continue
        if p.IsReadOnly:
            found_readonly.append(n if isinstance(n, str) else str(n))
            continue
        try:
            if isinstance(value, (int, float)):
                is_voltage = isinstance(n, str) and any(x in n.lower() for x in ["tens\xe3o", "voltage", "voltagem", "volts"])
                if is_voltage:
                    if p.SetValueString(str(value) + " V"):
                        dbg.ok('set_param [{}] = {} (ValueString+V)'.format(n, value))
                        return True
                    if p.SetValueString(str(value)):
                        dbg.ok('set_param [{}] = {} (ValueString)'.format(n, value))
                        return True
                    p.Set(float(value) * 10.7639104167)
                    dbg.ok('set_param [{}] = {} (raw float)'.format(n, value))
                    return True
                else:
                    p.Set(float(value))
            else:
                p.Set(str(value))
            dbg.ok('set_param [{}] = {}'.format(n, value))
            return True
        except Exception as ex:
            dbg.warn('set_param [{}] falhou: {}'.format(n, ex))
            continue
    if found_readonly:
        dbg.step('set_param somente leitura em {} (valor {})'.format(found_readonly, value))
    else:
        dbg.warn('set_param NENHUM param encontrado em {} para valor {}'.format(names, value))
    return False

def is_element_connected_to_panel(elem):
    try:
        for n in ["Painel", "Panel"]:
            p = elem.LookupParameter(n)
            if p and p.HasValue:
                val = p.AsString()
                if val and val.strip():
                    dbg.step('is_connected Id={} -> Painel="{}"'.format(elem.Id.IntegerValue, val))
                    return True
    except Exception: pass
    return False

def _find_matching_dist_sys(target_voltage):
    """Find ElementId of DistributionSysType whose voltage range covers target_voltage (Volts)."""
    v_internal = float(target_voltage) * 10.7639104167
    dbg.step('  Buscando sistema para {:.1f}V'.format(float(target_voltage)))

    voltage_bips = []
    for bip_name in ["RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_G_PARAM",
                     "RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_L_PARAM"]:
        try:
            voltage_bips.append(getattr(BuiltInParameter, bip_name))
        except Exception:
            dbg.step('  BuiltInParameter indisponivel: {}'.format(bip_name))
    if not voltage_bips:
        dbg.warn('  Parametros de tensao do sistema indisponiveis nesta versao do Revit')
        return None
    
    for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
        # RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_G_PARAM (Fase-Neutro)
        # RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_L_PARAM (Fase-Fase)
        for bip in voltage_bips:
            vp = s.get_Parameter(bip)
            if not vp or not vp.HasValue: continue
            
            vid = vp.AsElementId()
            if vid == ElementId.InvalidElementId: continue
            
            vtype = doc.GetElement(vid)
            if not vtype: continue
            
            min_p = vtype.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE_MIN_PARAM)
            max_p = vtype.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE_MAX_PARAM)
            if min_p and max_p:
                # Margem de +- 55 unidades internas (~5 Volts)
                if (min_p.AsDouble() - 55) <= v_internal <= (max_p.AsDouble() + 55):
                    dbg.ok('  Sistema compativel: {}'.format(s.Name))
                    return s.Id
    return None


def configure_element_for_voltage(elem, voltage, poles):
    """Force the MEP connector of elem to show target voltage and poles."""
    dbg.step('  Configurando elemento: {}V, {} polos'.format(voltage, poles))

    v_internal = float(voltage) * 10.7639104167

    # Method 1: BuiltInParameters on the element itself
    for bip, val in [(BuiltInParameter.RBS_ELEC_VOLTAGE, v_internal),
                     (BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES, float(poles))]:
        try:
            p = elem.get_Parameter(bip)
            if p and not p.IsReadOnly:
                p.Set(val)
                dbg.ok('    Param [{}] definido'.format(bip))
        except Exception: pass

    # Method 2: direct connector Voltage / Poles properties
    try:
        mgr = None
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            mgr = getattr(elem.MEPModel, 'ConnectorManager', None)
        if mgr:
            for c in mgr.Connectors:
                if c.Domain == Domain.DomainElectrical:
                    try:
                        c.Voltage = v_internal
                        c.Poles = poles
                        dbg.ok('    Conector elétrico configurado')
                    except Exception: pass
    except Exception: pass

    # Method 3: assign a DistributionSysType with the matching voltage range
    try:
        dist_id = _find_matching_dist_sys(voltage)
    except Exception as ex:
        dbg.warn('    _find_matching_dist_sys falhou: {}'.format(ex))
        dist_id = None
    if dist_id:
        for name in ["Sistema de distribui\xe7\xe3o", "Distribution System", "Sistema de Distribui\xe7\xe3o"]:
            try:
                p = elem.LookupParameter(name)
                if p and not p.IsReadOnly:
                    p.Set(dist_id)
                    dbg.ok('    Sistema de distribuição definido')
                    return
            except Exception: pass
    dbg.step('  configure_element_for_voltage concluido')


def ensure_element_is_free(elem):
    dbg.enter('ensure_element_is_free', Id=elem.Id.IntegerValue)
    try:
        systems_to_delete = set()

        try:
            if hasattr(elem, "MEPModel") and elem.MEPModel:
                for prop in ["AssignedElectricalSystems", "ElectricalSystems", "GetElectricalSystems"]:
                    try:
                        systems = None
                        if prop == "GetElectricalSystems":
                            if hasattr(elem.MEPModel, "GetElectricalSystems"):
                                systems = elem.MEPModel.GetElectricalSystems()
                        else:
                            systems = getattr(elem.MEPModel, prop, None)
                        if systems:
                            for sys_obj in systems:
                                if sys_obj:
                                    systems_to_delete.add(sys_obj.Id)
                                    dbg.step('Sistema por MEPModel.{} Id={}'.format(prop, sys_obj.Id.IntegerValue))
                    except Exception as ex:
                        dbg.warn('Falha ao ler MEPModel.{}: {}'.format(prop, ex))
        except Exception as ex:
            dbg.warn('Falha ao ler sistemas do MEPModel: {}'.format(ex))

        mgr = None
        if hasattr(elem, "MEPModel") and elem.MEPModel:
            mgr = getattr(elem.MEPModel, "ConnectorManager", None)
        if not mgr and hasattr(elem, "ConnectorManager"):
            mgr = elem.ConnectorManager
        if mgr:
            try:
                all_conns = list(mgr.Connectors)
            except Exception:
                all_conns = []
            dbg.step('Conectores: {}'.format(len(all_conns)))
            for i, c in enumerate(all_conns):
                try:
                    sys_obj = None
                    sys_id_str = "None"
                    try:
                        sys_obj = c.MEPSystem
                        if sys_obj:
                            sys_id_str = str(sys_obj.Id.IntegerValue)
                            systems_to_delete.add(sys_obj.Id)
                    except Exception:
                        sys_id_str = "ERRO"
                    dbg.step('  [{}] domain={} connected={} MEPSystem={}'.format(
                        i, c.Domain, c.IsConnected, sys_id_str))
                    
                    try:
                        if c.IsConnected and c.AllRefs:
                            for ref in c.AllRefs:
                                if ref.Owner and hasattr(ref.Owner, "Category") and ref.Owner.Category:
                                    if ref.Owner.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalCircuit):
                                        systems_to_delete.add(ref.Owner.Id)
                                        dbg.step('  Encontrado sistema via AllRefs Id={}'.format(ref.Owner.Id.IntegerValue))
                    except Exception as ex:
                        dbg.warn('  [{}] erro ao ler AllRefs: {}'.format(i, ex))

                except Exception as ex:
                    dbg.warn('  [{}] erro ao ler conector: {}'.format(i, ex))
        else:
            dbg.step('Sem ConnectorManager')

        if systems_to_delete:
            for sys_id in systems_to_delete:
                dbg.step('Deletando sistema Id={}'.format(sys_id.IntegerValue))
                try:
                    doc.Delete(sys_id)
                except Exception as del_ex:
                    dbg.warn('Falha ao deletar sistema {}: {}'.format(sys_id.IntegerValue, del_ex))
            dbg.ok('Removidos {} sistema(s)'.format(len(systems_to_delete)))
            dbg.exit('ensure_element_is_free', True)
            return True

        dbg.step('Nenhum sistema eletrico conectado - elemento livre')
    except Exception as ex:
        dbg.fail('ensure_element_is_free: {}'.format(ex))
    dbg.exit('ensure_element_is_free', False)
    return False

def get_element_poles(elem):
    """Lê o número de fases/polos reais do elemento (conector ou parâmetro).
    Retorna int ou None se não conseguir."""
    try:
        mgr = None
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            mgr = getattr(elem.MEPModel, 'ConnectorManager', None)
        if mgr:
            for c in mgr.Connectors:
                if c.Domain == Domain.DomainElectrical:
                    try:
                        v = int(c.Poles)
                        if v > 0:
                            return v
                    except Exception:
                        pass
    except Exception:
        pass
    for n in ["Número de Fases", "N\xb0 de Fases", "N\xba de Fases", "N\xfamero de polos", "Number of Poles", "Polos"]:
        try:
            p = elem.LookupParameter(n)
            if p and p.HasValue:
                v = p.AsInteger() if p.StorageType.ToString() == 'Integer' else int(p.AsDouble())
                if v > 0:
                    return v
        except Exception:
            pass
    return None


def get_family_name(elem):
    try:
        if hasattr(elem, 'Symbol') and elem.Symbol:
            return elem.Symbol.FamilyName or ""
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            fam_param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            if fam_param and fam_param.HasValue:
                return fam_param.AsString() or ""
    except Exception: pass
    return ""

def get_room_name(elem):
    eid = elem.Id.IntegerValue
    # Estratégia 1: Room via Phase (Revit PT-BR usa indexer Room[Phase])
    try:
        if hasattr(elem, "Room"):
            phase = None
            try:
                phase_id = elem.CreatedPhaseId
                if phase_id and phase_id != ElementId.InvalidElementId:
                    phase = elem.Document.GetElement(phase_id)
            except Exception:
                pass
            if phase:
                try:
                    rm = elem.Room[phase]
                    if rm:
                        name = rm.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
                        if name:
                            dbg.step('Room[Phase]: "{}" (Id={})'.format(name, eid))
                            return name
                except Exception:
                    pass
    except Exception:
        pass
    # Estratégia 2: Space via Phase
    try:
        if hasattr(elem, "Space"):
            phase = None
            try:
                phase_id = elem.CreatedPhaseId
                if phase_id and phase_id != ElementId.InvalidElementId:
                    phase = elem.Document.GetElement(phase_id)
            except Exception:
                pass
            if phase:
                try:
                    sp = elem.Space[phase]
                    if sp:
                        name = sp.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
                        if name:
                            dbg.step('Space[Phase]: "{}" (Id={})'.format(name, eid))
                            return name
                except Exception:
                    pass
    except Exception:
        pass
    # Estratégia 3: GetRoomAtPoint (fallback geométrico)
    # Para elementos hospedados em parede, Location.Point cai dentro da espessura da parede.
    # Geramos candidatos de ponto com offset nas direções da face (FacingOrientation) e
    # também nas 4 direções ortogonais — um deles estará dentro do ambiente.
    def _name_from_room_or_space(target_doc, pt):
        for getter in [target_doc.GetRoomAtPoint, target_doc.GetSpaceAtPoint]:
            try:
                obj = getter(pt)
                if obj:
                    p = obj.get_Parameter(BuiltInParameter.ROOM_NAME)
                    if p:
                        n = p.AsString()
                        if n:
                            return n
            except Exception:
                pass
        return None

    try:
        if hasattr(elem, "Location") and elem.Location and hasattr(elem.Location, "Point"):
            pt = elem.Location.Point

            # Candidatos: ponto original + offsets para atravessar a espessura da parede
            OFFSET = 0.5  # ~15 cm em pés
            candidates = [pt]
            try:
                fo = elem.FacingOrientation
                candidates.append(XYZ(pt.X + fo.X * OFFSET, pt.Y + fo.Y * OFFSET, pt.Z))
                candidates.append(XYZ(pt.X - fo.X * OFFSET, pt.Y - fo.Y * OFFSET, pt.Z))
            except Exception:
                pass
            for dx, dy in [(OFFSET, 0), (-OFFSET, 0), (0, OFFSET), (0, -OFFSET)]:
                candidates.append(XYZ(pt.X + dx, pt.Y + dy, pt.Z))

            for candidate in candidates:
                name = _name_from_room_or_space(elem.Document, candidate)
                if name:
                    dbg.step('Room por ponto: "{}" (Id={})'.format(name, eid))
                    return name

            # Buscar em Links (Modelos vinculados) com os mesmos candidatos
            for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements():
                link_doc = link.GetLinkDocument()
                if not link_doc:
                    continue
                try:
                    inv = link.GetTransform().Inverse
                    for candidate in candidates:
                        name = _name_from_room_or_space(link_doc, inv.OfPoint(candidate))
                        if name:
                            dbg.step('Room (Link) por ponto: "{}" (Id={})'.format(name, eid))
                            return name
                except Exception:
                    pass
    except Exception as ex:
        dbg.warn('get_room_name fallback erro: {} (Id={})'.format(ex, eid))
    dbg.warn('Sem room/space para Id={}'.format(eid))
    return ""

def get_valid_electrical_elements(elem, expected_domains=(Domain.DomainElectrical,)):
    dbg.enter('get_valid_electrical_elements', Id=elem.Id.IntegerValue)
    valid_pairs = []
    def _extract(e):
        try:
            mgr = None
            if hasattr(e, "MEPModel") and getattr(e, "MEPModel", None) and getattr(e.MEPModel, "ConnectorManager", None):
                mgr = e.MEPModel.ConnectorManager
            elif hasattr(e, "ConnectorManager") and getattr(e, "ConnectorManager", None):
                mgr = e.ConnectorManager
            if mgr:
                conns = [c for c in mgr.Connectors if c.Domain in expected_domains]
                if conns:
                    dbg.step('  Conectores validos: {} (Id={})'.format(len(conns), e.Id.IntegerValue))
                    return conns
                else:
                    dbg.step('  Nenhum conector no dominio esperado (Id={})'.format(e.Id.IntegerValue))
            else:
                dbg.step('  Sem ConnectorManager (Id={})'.format(e.Id.IntegerValue))
        except Exception as ex:
            dbg.warn('  _extract erro: {}'.format(ex))
        return []

    conns = _extract(elem)
    if conns: valid_pairs.append((elem.Id, conns))
        
    if hasattr(elem, "GetSubComponentIds"):
        sub_ids = elem.GetSubComponentIds()
        if sub_ids and sub_ids.Count > 0:
            dbg.step('Sub-componentes: {}'.format(sub_ids.Count))
            for sub_id in sub_ids:
                sub_elem = elem.Document.GetElement(sub_id)
                if sub_elem:
                    c_sub = _extract(sub_elem)
                    if c_sub: valid_pairs.append((sub_id, c_sub))
    dbg.exit('get_valid_electrical_elements', '{} par(es)'.format(len(valid_pairs)))
    return valid_pairs

class ConnectorDomainFilter(ISelectionFilter):
    def __init__(self, target_domains=None, allowed_categories=None):
        self.target_domains = target_domains
        # Se allowed_categories for passado, restringe a essas categorias.
        # Se None, usa o conjunto padrão de categorias elétricas.
        if allowed_categories is not None:
            self.allowed = set(allowed_categories)
        else:
            self.allowed = {
                int(BuiltInCategory.OST_ElectricalFixtures),
                int(BuiltInCategory.OST_ElectricalEquipment),
                int(BuiltInCategory.OST_LightingFixtures),
                int(BuiltInCategory.OST_LightingDevices),
                int(BuiltInCategory.OST_DataDevices),
                int(BuiltInCategory.OST_CommunicationDevices),
                int(BuiltInCategory.OST_TelephoneDevices),
            }

    def AllowElement(self, e):
        if not e.Category: return False
        return e.Category.Id.IntegerValue in self.allowed

    def AllowReference(self, ref, pos): return False

class CategoryFilter(ISelectionFilter):
    def __init__(self, cat_id): self.cat_id = cat_id
    def AllowElement(self, e): return e.Category and e.Category.Id.IntegerValue == int(self.cat_id)
    def AllowReference(self, ref, pos): return False

class ElectricalElementFilter(ISelectionFilter):
    ALLOWED_CATS = (
        int(BuiltInCategory.OST_LightingFixtures),
        int(BuiltInCategory.OST_ElectricalFixtures),
        int(BuiltInCategory.OST_ElectricalEquipment),
        int(BuiltInCategory.OST_LightingDevices),
        int(BuiltInCategory.OST_CommunicationDevices),
        int(BuiltInCategory.OST_DataDevices),
        int(BuiltInCategory.OST_TelephoneDevices),
        int(BuiltInCategory.OST_Conduit),
    )
    def AllowElement(self, e): return e.Category and e.Category.Id.IntegerValue in self.ALLOWED_CATS
    def AllowReference(self, ref, pos): return False

class PanelFilter(ISelectionFilter):
    def AllowElement(self, e): return e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment)
    def AllowReference(self, ref, pos): return False

def next_valid_letter(counter):
    while counter < 26:
        letter = chr(ord('A') + counter)
        if letter not in SKIP_LETTERS: return letter, counter
        counter += 1
    return 'A', 0

def _get_switch_label(counter):
    n = len(VALID_SWITCH_LETTERS)
    if counter < n: return VALID_SWITCH_LETTERS[counter]
    else:
        idx = counter - n
        letter = VALID_SWITCH_LETTERS[idx % n]
        repeats = 2 + (idx // n)
        return letter * repeats

def prompt_phase_voltage():
    opcoes = OrderedDict([
        ("1 Fase (Mono) - 127V",  {"poles": 1, "voltage": 127}),
        ("1 Fase (Mono) - 220V",  {"poles": 1, "voltage": 220}),
        ("2 Fases (Bifásico) - 220V", {"poles": 2, "voltage": 220}),
        ("2 Fases (Bifásico) - 380V", {"poles": 2, "voltage": 380}),
        ("3 Fases (Trifásico) - 220V", {"poles": 3, "voltage": 220}),
        ("3 Fases (Trifásico) - 380V", {"poles": 3, "voltage": 380})
    ])
    escolha = forms.CommandSwitchWindow.show(
        opcoes.keys(), message="Selecione as Fases e Tensão para este circuito:", title="Compatibilidade do Quadro"
    )
    if escolha: return opcoes[escolha]
    return None

def _get_compatible_systems(panel):
    dbg.enter('_get_compatible_systems', Id=panel.Id.IntegerValue)
    
    v_panel = None
    poles_panel = None
    
    # 1. Obter info do conector elétrico (que é o que o Revit usa na UI)
    try:
        if hasattr(panel, "MEPModel") and panel.MEPModel:
            cm = getattr(panel.MEPModel, "ConnectorManager", None)
            if cm:
                for c in cm.Connectors:
                    if c.Domain == Domain.DomainElectrical:
                        try: v_panel = c.Voltage
                        except Exception: pass
                        try: poles_panel = c.Poles
                        except Exception: pass
                        if v_panel is not None:
                            dbg.step('Conector encontrado: {}V, {} Polos'.format(round(v_panel/10.7639, 1), poles_panel))
                        break
    except Exception as ex:
        dbg.warn("Erro ao ler conector: {}".format(ex))

    # Obter todos os sistemas
    all_systems = {}
    for s in FilteredElementCollector(doc).OfClass(DistributionSysType).ToElements():
        n = ""
        try: n = s.Name
        except Exception:
            pp = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if pp: n = pp.AsString()
        if n: all_systems[n] = s

    if not all_systems:
        dbg.warn('Nenhum DistributionSysType no modelo')
        dbg.exit('_get_compatible_systems', '{}')
        return {}

    # Se nao conseguiu ler a tensao, ou for zero, a Familia no Revit tambem estara zerada ou incompleta
    if v_panel is None or v_panel <= 0:
        dbg.warn("Quadro não possui conector elétrico válido ou tensão é zero. Mostrando todos por segurança.")
        compatible = {k: v.Id for k, v in all_systems.items()}
        dbg.exit('_get_compatible_systems', '{} compativeis (Fallback sem conector)'.format(len(compatible)))
        return compatible
    
    compatible = {}
    for name, sys in all_systems.items():
        # Lógica de Fases
        try:
            phase_bip = getattr(BuiltInParameter, "RBS_ELEC_DISTRIBUTION_SYS_PHASE_PARAM")
        except Exception:
            phase_bip = None
        sys_phase_p = sys.get_Parameter(phase_bip) if phase_bip is not None else None
        sys_phase = sys_phase_p.AsInteger() if sys_phase_p else 2
        
        # Monofásico não aceita painel de 3 fases
        if sys_phase == 1 and poles_panel is not None and poles_panel >= 3:
            dbg.step('  [--] {} (Incompatível: Fases)'.format(name))
            continue
            
        # Lógica de Tensão
        v_ok = False
        voltage_params = []
        for bip_name in ["RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_L_PARAM",
                         "RBS_ELEC_DISTRIBUTION_SYS_VOLTAGE_L_G_PARAM"]:
            try:
                bip = getattr(BuiltInParameter, bip_name)
                voltage_params.append(sys.get_Parameter(bip))
            except Exception:
                pass
        
        for vparam in voltage_params:
            if vparam and vparam.AsElementId() != ElementId.InvalidElementId:
                vtype = doc.GetElement(vparam.AsElementId())
                if vtype:
                    min_p = vtype.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE_MIN_PARAM)
                    max_p = vtype.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE_MAX_PARAM)
                    if min_p and max_p:
                        v_min = min_p.AsDouble()
                        v_max = max_p.AsDouble()
                        # Margem flexível de +- 5 Volts (5 * 10.76 internamente ~ 55 unidades)
                        if (v_min - 55) <= v_panel <= (v_max + 55):
                            v_ok = True
                            break
                            
        if v_ok:
            compatible[name] = sys.Id
            dbg.step('  [OK] {}'.format(name))
        else:
            dbg.step('  [--] {} (Incompatível: Tensão não bate)'.format(name))

    # Fallback caso filtre 100% (evita bloquear o usuario)
    if not compatible:
        dbg.warn("Filtro algébrico zerou opções. Desativando filtro.")
        compatible = {k: v.Id for k, v in all_systems.items()}

    dbg.exit('_get_compatible_systems', '{} compativeis filtrados'.format(len(compatible)))
    return compatible

def configure_panel(panel):
    messages = []
    settings = load_config()

    new_name = forms.ask_for_string(default=get_panel_name(panel), prompt="Nome do Quadro (ex: QDL-01):", title="Configurar Quadro")
    if not new_name: return False, "Cancelado"

    with Transaction(doc, "Nome do Quadro") as t:
        t.Start()
        set_param(panel, ["Nome do painel", "Panel Name", "Mark"], new_name)
        t.Commit()
    messages.append("Nome: " + new_name)

    # Filtrar sistemas compativeis com o quadro
    sys_id = None
    compatible_systems = _get_compatible_systems(panel)

    default_sys = settings.get('dist_system', '')
    if compatible_systems:
        sys_names = sorted(compatible_systems.keys())
        if default_sys and default_sys in sys_names:
            sys_names.remove(default_sys)
            sys_names.insert(0, default_sys + " ★")
            compatible_systems[default_sys + " ★"] = compatible_systems[default_sys]
        chosen_sys = forms.CommandSwitchWindow.show(sys_names, message="Sistema de Distribuição para " + new_name + " ({} compatíveis):".format(len(sys_names)), title="Tipo de Quadro")
        if chosen_sys:
            sys_id = compatible_systems.get(chosen_sys)
            messages.append("Sistema: " + chosen_sys.replace(" ★", ""))
    else:
        forms.alert("Nenhum sistema de distribuição compatível com este quadro.")

    default_naming = settings.get('naming_scheme', '')
    nam_id = None
    all_namings = {}
    for s in FilteredElementCollector(doc).OfClass(CircuitNamingScheme).ToElements():
        all_namings[s.Name] = s.Id
    if all_namings:
        naming_names = sorted(all_namings.keys())
        if default_naming and default_naming in naming_names:
            naming_names.remove(default_naming)
            naming_names.insert(0, default_naming + " ★")
            all_namings[default_naming + " ★"] = all_namings[default_naming]
        chosen_naming = forms.CommandSwitchWindow.show(naming_names, message="Nomenclatura do Circuito:", title="Nomenclatura")
        if chosen_naming:
            nam_id = all_namings.get(chosen_naming)
            messages.append("Nomenclatura: " + chosen_naming.replace(" ★", ""))

    with Transaction(doc, "Sistema e Nomenclatura") as t:
        t.Start()
        p = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        if p and not p.IsReadOnly and sys_id:
            try: p.Set(sys_id)
            except Exception: messages.append("Sistema: FALHA")

        p = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")
        if p and not p.IsReadOnly and nam_id:
            p.Set(nam_id)

        prefix = settings.get('circuit_prefix', '')
        if prefix:
            set_param(panel, ["Prefixo do circuito", "Circuit Prefix"], prefix)

        t.Commit()
    return True, "\n".join(messages)

def select_and_configure_panel():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, PanelFilter(), "Selecione o QUADRO")
        panel = doc.GetElement(ref.ElementId)
        p_sys = panel.LookupParameter("Sistema de distribuição") or panel.LookupParameter("Distribution System")
        p_nam = panel.LookupParameter("Nomenclatura do circuito") or panel.LookupParameter("Circuit Naming")

        has_sys = p_sys and p_sys.HasValue and p_sys.AsElementId() != ElementId.InvalidElementId
        has_nam = p_nam and p_nam.HasValue and p_nam.AsElementId() != ElementId.InvalidElementId

        if has_sys and has_nam:
            set_current_panel(panel.Id)
            forms.alert("Quadro selecionado: " + get_panel_name(panel))
        else:
            success, msg = configure_panel(panel)
            if success:
                set_current_panel(panel.Id)
                forms.alert("Quadro Configurado!\n" + msg)
            else:
                forms.alert(msg)
    except Exception: pass

def call_queda_tensao():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, ElectricalElementFilter(), "Verificar Queda de Tensão")
    except Exception: return

    elem = doc.GetElement(ref.ElementId)
    is_conduit = (elem.Category and elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Conduit))

    comp_m, diam_mm = 0.0, 25.0
    if is_conduit:
        try: comp_m = round(elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 0.3048, 2)
        except Exception: pass
        try: diam_mm = round(elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsDouble() * 304.8, 2)
        except Exception: pass
    else:
        circuit = None
        for prop in ['AssignedElectricalSystems', 'ElectricalSystems']:
            if hasattr(elem, 'MEPModel') and elem.MEPModel and hasattr(elem.MEPModel, prop):
                syss = getattr(elem.MEPModel, prop)
                if syss and syss.Count > 0:
                    for es in syss:
                        circuit = es
                        break
            if circuit: break
            
        if not circuit:
            try:
                cm = elem.MEPModel.ConnectorManager
                if cm:
                    for c in cm.Connectors:
                        if c.Domain == Domain.DomainElectrical and c.MEPSystem and hasattr(c.MEPSystem, 'ElectricalSystemType'):
                            circuit = c.MEPSystem
                            break
            except Exception: pass

        if not circuit:
            forms.alert("Elemento não pertence a um circuito elétrico.")
            return

        try: comp_m = round(circuit.Length * 0.3048, 2)
        except Exception: pass

        try:
            conduits = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements()
            for cond in conduits:
                try:
                    cm = cond.ConnectorManager
                    if cm:
                        for c in cm.Connectors:
                            if c.IsConnected:
                                for ref_conn in c.AllRefs:
                                    owner = ref_conn.Owner
                                    if owner and owner.Id == elem.Id:
                                        p = cond.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                                        if p and p.HasValue:
                                            diam_mm = round(p.AsDouble() * 304.8, 2)
                                            raise StopIteration
                except StopIteration: break
                except Exception: continue
        except Exception: pass

        circ_info = ""
        try:
            load_name_p = circuit.LookupParameter("Nome da carga") or circuit.LookupParameter("Load Name")
            if load_name_p and load_name_p.HasValue: circ_info = load_name_p.AsString()
        except Exception: pass

        if circ_info:
            forms.toast("Circuito: {} | {}m | Ø{}mm".format(circ_info, comp_m, diam_mm), title="Queda de Tensão")

    data = {"length_m": comp_m, "diam_mm": diam_mm}
    try:
        from QuedaTensao.queda_tensao_ui import QuedaTensaoWindow
        xaml_path = os.path.join(lib_path, 'QuedaTensao', 'queda_tensao_window.xaml')
        win = QuedaTensaoWindow(xaml_path, doc, elem.Id, data)
        win.ShowDialog()
    except Exception as e:
        forms.alert("Erro ao abrir Queda de Tensão:\n" + str(e), title="Erro")
