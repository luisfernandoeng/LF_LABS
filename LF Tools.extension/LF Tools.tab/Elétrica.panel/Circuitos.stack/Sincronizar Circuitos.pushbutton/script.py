# -*- coding: utf-8 -*-
"""Sincronizar Circuitos
Selecione o quadro e o script dimensiona automaticamente
Disjuntor, Cabo e sincroniza a Distância de todos os circuitos.
"""
__title__ = "Sincronizar\nCircuitos"
__author__ = "LF Tools"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import re
import math

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# ================================================================
# TABELAS NBR 5410
# ================================================================

BREAKERS = [16, 20, 25, 32, 40, 50, 63, 70, 80, 100, 125, 160, 200, 225, 250, 300, 400]
SECTIONS = [1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240]
PANEL_PREFIXES = ("QD", "QDE", "QDG", "QGBT", "QDL", "QF", "QM", "QP")
PANEL_BREAKER_FACTOR = 1.15
PANEL_MIN_BREAKER = 32
PANEL_MIN_SECTION = 6.0
EPR_90_WIRE_TYPE_NAME = "[Cu/EPR-XLPE/0,6-1kV/90°]-Un-D-3Cc"

# Ampacidade (Iz) NBR 5410 - Método B1, 2 condutores, PVC 70°C
AMPACITIES = {
    1.5: 17.5, 2.5: 24.0, 4.0: 32.0, 6.0: 41.0, 10.0: 57.0, 16.0: 76.0,
    25.0: 101.0, 35.0: 125.0, 50.0: 151.0, 70.0: 192.0, 95.0: 232.0,
    120.0: 269.0, 150.0: 309.0, 185.0: 353.0, 240.0: 415.0
}

# Par Disjuntor → Cabo mínimo coordenado
PAIRED_SECTIONS = {
    16: 2.5, 20: 2.5, 25: 4.0, 32: 6.0, 40: 10.0,
    50: 10.0, 63: 16.0, 80: 25.0, 100: 35.0,
    125: 50.0, 160: 70.0, 200: 95.0, 225: 120.0, 250: 120.0,
    300: 150.0, 400: 240.0
}

# ================================================================
# HELPERS
# ================================================================

class PanelFilter(ISelectionFilter):
    def AllowElement(self, e):
        return (e.Category and
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment))
    def AllowReference(self, ref, pos): return False


class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        for f in fa.GetFailureMessages():
            if f.GetSeverity() == FailureSeverity.Warning:
                fa.DeleteWarning(f)
        return FailureProcessingResult.Continue


def get_panel_name(panel):
    try:
        for n in ["Nome do painel", "Panel Name", "Mark"]:
            p = panel.LookupParameter(n)
            if p and p.HasValue:
                return p.AsString()
        return panel.Name
    except:
        return "Quadro"


def set_parameter_value(p, value):
    if not p or p.IsReadOnly:
        return False
    if isinstance(value, ElementId):
        p.Set(value)
    elif p.StorageType == StorageType.String:
        p.Set(str(value))
    elif p.StorageType == StorageType.Integer:
        p.Set(int(value))
    elif p.StorageType == StorageType.ElementId:
        p.Set(value if isinstance(value, ElementId) else ElementId(int(value)))
    else:
        p.Set(float(value))
    return True


def write_param(elem, builtin=None, names=None, value=None):
    if builtin:
        try:
            if set_parameter_value(elem.get_Parameter(builtin), value):
                return True
        except:
            pass
    if names:
        for n in names:
            try:
                if set_parameter_value(elem.LookupParameter(n), value):
                    return True
            except:
                continue
    return False


def get_param_text(elem, builtin=None, names=None):
    params = []
    if builtin:
        try:
            params.append(elem.get_Parameter(builtin))
        except:
            pass
    if names:
        for n in names:
            try:
                params.append(elem.LookupParameter(n))
            except:
                pass

    for p in params:
        try:
            if p and p.HasValue:
                val = p.AsString() or p.AsValueString()
                if val:
                    return val
        except:
            continue
    return ""


def normalize_text(value):
    return re.sub(r"[^A-Z0-9]", "", (value or "").upper())


def get_standard_breaker(target):
    for b in BREAKERS:
        if b >= target:
            return b
    return BREAKERS[-1]


def get_next_standard_breaker(current):
    for b in BREAKERS:
        if b > current:
            return b
    return BREAKERS[-1]


def get_param_double(elem, builtin=None, names=None, default=0.0):
    params = []
    if builtin:
        try:
            params.append(elem.get_Parameter(builtin))
        except:
            pass
    if names:
        for n in names:
            try:
                params.append(elem.LookupParameter(n))
            except:
                pass

    for p in params:
        try:
            if p and p.HasValue:
                txt = p.AsValueString()
                if txt:
                    m = re.search(r"[-+]?\d+(?:[,.]\d+)?", txt)
                    if m:
                        return float(m.group(0).replace(",", "."))
                return float(p.AsDouble())
        except:
            continue
    return default


def get_circuit_label(circuit):
    return get_param_text(
        circuit,
        builtin=BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
        names=["Nome da carga", "Load Name"]
    )


def is_panel_feeder(circuit):
    circuit_number = normalize_text(circuit.CircuitNumber or "")
    load_name = normalize_text(get_circuit_label(circuit))
    prefix = normalize_text(get_param_text(circuit, names=["Préfixo Circuito", "Prefixo Circuito", "Circuit Prefix"]))

    names_to_check = [load_name, prefix, circuit_number]
    looks_like_panel = any(
        any(name.startswith(p) for p in PANEL_PREFIXES)
        for name in names_to_check if name
    )
    if not looks_like_panel:
        return False

    if load_name and circuit_number.startswith(load_name):
        return True
    if prefix and circuit_number.startswith(prefix):
        return True
    return False


def get_element_name(elem):
    try:
        return elem.Name
    except:
        try:
            return Element.Name.GetValue(elem)
        except:
            return ""


def find_epr_90_wire_type():
    exact = None
    fallback = None
    try:
        for wt in FilteredElementCollector(doc).OfClass(WireType):
            name = get_element_name(wt)
            if name == EPR_90_WIRE_TYPE_NAME:
                exact = wt
                break
            n = normalize_text(name)
            if "EPR" in n and "90" in n:
                fallback = wt
        return exact or fallback
    except:
        return None


def write_wire_type(circuit, wire_type):
    if not wire_type:
        return False
    try:
        return write_param(
            circuit,
            builtin=BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_TYPE_PARAM,
            names=["Tipo de fiação", "Wire Type"],
            value=wire_type.Id
        )
    except:
        return write_param(circuit, names=["Tipo de fiação", "Wire Type"], value=wire_type.Id)


def write_breaker(circuit, rating):
    debug = []

    p = None
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
    except:
        pass
    if not p:
        try:
            p = circuit.LookupParameter("Proteção do circuito")
        except:
            pass

    if p is None:
        debug.append("param nao encontrado")
        return False, 0.0, debug

    debug.append("storage={}".format(p.StorageType))
    debug.append("readOnly={}".format(p.IsReadOnly))

    try:
        val_before = p.AsDouble()
        debug.append("antes={:.1f}".format(val_before))
    except Exception as ex:
        val_before = 0.0
        debug.append("antes=err({})".format(ex))

    if p.IsReadOnly:
        return False, val_before, debug

    wrote = False
    try:
        if p.StorageType == StorageType.Double:
            result = p.Set(float(rating))
            wrote = True
            debug.append("Set(double {})={}".format(rating, result))
        elif p.StorageType == StorageType.Integer:
            result = p.Set(int(rating))
            wrote = True
            debug.append("Set(int {})={}".format(rating, result))
        elif p.StorageType == StorageType.String:
            result = p.Set(str(rating))
            wrote = True
            debug.append("Set(str {})={}".format(rating, result))
        else:
            debug.append("StorageType desconhecido: {}".format(p.StorageType))
    except Exception as ex:
        debug.append("Set excecao: {}".format(ex))

    try:
        val_after = p.AsDouble()
        debug.append("depois={:.1f}".format(val_after))
        actual = val_after
    except Exception as ex:
        debug.append("depois=err({})".format(ex))
        actual = 0.0

    return wrote, actual, debug


def calculate_vd(length_m, current, voltage, section_mm2, phases=1):
    if section_mm2 <= 0 or voltage <= 0:
        return 0.0
    factor = 2.0 if phases <= 2 else math.sqrt(3)
    vd_v = (factor * length_m * current) / (56.0 * section_mm2)
    return (vd_v * 100.0) / voltage


# ================================================================
# SELEÇÃO DO PAINEL
# ================================================================

def pick_panel():
    """Usa seleção atual ou pede para clicar no quadro."""
    for sid in list(uidoc.Selection.GetElementIds()):
        el = doc.GetElement(sid)
        if (el and el.Category and
                el.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment)):
            return el
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, PanelFilter(), "Selecione o QUADRO"
        )
        return doc.GetElement(ref.ElementId)
    except:
        return None


def get_circuits(panel):
    try:
        if hasattr(panel, "MEPModel") and panel.MEPModel:
            systems = list(panel.MEPModel.GetAssignedElectricalSystems())
        else:
            systems = list(panel.GetAssignedElectricalSystems())
    except:
        return []

    def sort_key(s):
        try:
            d = re.sub(r'\D', '', s.CircuitNumber or '')
            return int(d) if d else 9999
        except:
            return 9999

    systems.sort(key=sort_key)
    return systems


# ================================================================
# SINCRONIZAÇÃO
# ================================================================

def sync_circuits(circuits):
    """
    Para cada circuito:
      1. Distância: Length da API → metros (teto) → grava em L Considerado
      2. Disjuntor: próximo padrão ≥ I_corr (sem margem adicional)
      3. Cabo: par mínimo do disjuntor → sobe SEÇÃO enquanto VD > 3% ou Iz < In
    """
    ok, erros = [], []
    epr_90_wire_type = find_epr_90_wire_type()

    with Transaction(doc, "Sincronizar Circuitos") as t:
        t.Start()
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(opts)

        for c in circuits:
            num = c.CircuitNumber or "?"
            updated = []
            err = []
            panel_feeder = is_panel_feeder(c)
            if panel_feeder:
                updated.append("Alimentador")

            # 1. DISTÂNCIA
            dist_m = 0.0
            try:
                raw_ft = c.Length
                if raw_ft > 0:
                    dist_m = float(math.ceil(raw_ft * 0.3048))
                    if write_param(c, names=["L Considerado", "L Considerado (m)"], value=dist_m):
                        updated.append("Dist: {}m".format(int(dist_m)))
            except Exception as ex:
                err.append("Dist: " + str(ex))

            # 2. CORRENTE E TENSÃO
            try:
                i_proj = (c.ApparentLoad / c.Voltage) if c.Voltage > 0 else 0.0
                v_nominal = get_param_double(
                    c,
                    builtin=BuiltInParameter.RBS_ELEC_VOLTAGE,
                    names=["Tensão", "TensÃ£o", "Voltage"],
                    default=127.0
                )
                poles = c.PolesNumber

                fca, fct = 1.0, 1.0
                try:
                    pf = c.LookupParameter("FCA")
                    if pf and pf.HasValue:
                        fca = float(str(pf.AsValueString() or pf.AsDouble()).replace(",", "."))
                    pf = c.LookupParameter("FCT")
                    if pf and pf.HasValue:
                        fct = float(str(pf.AsValueString() or pf.AsDouble()).replace(",", "."))
                except:
                    pass
                if fca <= 0: fca = 1.0
                if fct <= 0: fct = 1.0

                i_corr = i_proj / (fca * fct)
            except Exception as ex:
                err.append("Corrente: " + str(ex))
                i_proj, v_nominal, poles, i_corr = 0.0, 127.0, 1, 0.0

            # 3. DIMENSIONAMENTO
            try:
                rating = get_standard_breaker(i_corr)
                if panel_feeder:
                    rating = get_standard_breaker(i_corr * PANEL_BREAKER_FACTOR)
                    rating = max(rating, get_next_standard_breaker(get_standard_breaker(i_corr)))
                    rating = max(rating, PANEL_MIN_BREAKER)
                section = PAIRED_SECTIONS.get(rating, 2.5)
                if panel_feeder:
                    section = max(section, PANEL_MIN_SECTION)

                for _ in range(len(SECTIONS)):
                    vd = calculate_vd(dist_m, i_proj, v_nominal, section, phases=poles)
                    iz = AMPACITIES.get(section, 0.0) * fca * fct
                    if (vd <= 3.0 or dist_m <= 0) and iz >= rating:
                        break
                    idx = next((i for i, s in enumerate(SECTIONS) if s == section), -1)
                    if idx == -1 or idx + 1 >= len(SECTIONS):
                        break
                    section = SECTIONS[idx + 1]

                breaker_written, actual_rating, breaker_debug = write_breaker(c, rating)
                debug_str = " [{}]".format(", ".join(breaker_debug)) if breaker_debug else ""
                if breaker_written:
                    if actual_rating and abs(actual_rating - rating) > 0.1:
                        updated.append("Disj pedido: {}A (ficou {}A){}".format(
                            int(rating), int(round(actual_rating)), debug_str))
                    else:
                        updated.append("Disj: {}A".format(int(rating)))
                else:
                    err.append("Disjuntor: nao gravado{}".format(debug_str))
                if write_param(c, names=["Seção do Condutor Adotado (mm²)", "Condutor Adotado"], value=section):
                    updated.append("Cabo: {}mm²".format(section))
                if panel_feeder:
                    if write_wire_type(c, epr_90_wire_type):
                        updated.append("Tipo: EPR 90")
                    elif epr_90_wire_type:
                        err.append("Tipo de fiacao: parametro bloqueado")
                    else:
                        err.append("Tipo de fiacao: EPR 90 nao encontrado no projeto")
            except Exception as ex:
                err.append("Dimensionamento: " + str(ex))

            if updated:
                line = "| {} | {} |".format(num, ", ".join(updated))
                if err:
                    line += " ⚠️ " + "; ".join(err)
                ok.append(line)
            elif err:
                erros.append("| {} | ❌ {} |".format(num, "; ".join(err)))

        t.Commit()

    return ok, erros


# ================================================================
# MAIN
# ================================================================

def main():
    panel = pick_panel()
    if not panel:
        return

    circuits = get_circuits(panel)
    if not circuits:
        forms.alert("Nenhum circuito encontrado no quadro '{}'.".format(get_panel_name(panel)))
        return

    panel_name = get_panel_name(panel)

    confirm = forms.alert(
        "Quadro: {}\n{} circuito(s) serão sincronizados.\n\nContinuar?".format(
            panel_name, len(circuits)),
        title="Sincronizar Circuitos",
        yes=True, no=True
    )
    if not confirm:
        return

    ok_lines, err_lines = sync_circuits(circuits)

    output.print_md("## ⚡ Sincronização — {}".format(panel_name))
    output.print_md("| Circuito | Resultado |")
    output.print_md("|---|---|")
    for line in ok_lines:
        output.print_md(line)
    for line in err_lines:
        output.print_md(line)

    output.print_md("\n✅ **{} sincronizado(s)** | ❌ **{} erro(s)**".format(
        len(ok_lines), len(err_lines)))

    forms.toast("✅ {} circuito(s) sincronizado(s)".format(len(ok_lines)))


if __name__ == '__main__':
    main()
