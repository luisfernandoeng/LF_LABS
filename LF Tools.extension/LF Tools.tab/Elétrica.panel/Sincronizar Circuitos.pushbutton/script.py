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


def write_param(elem, builtin=None, names=None, value=None):
    if builtin:
        try:
            p = elem.get_Parameter(builtin)
            if p and not p.IsReadOnly:
                p.Set(float(value)) if not isinstance(value, str) else p.Set(value)
                return True
        except:
            pass
    if names:
        for n in names:
            try:
                p = elem.LookupParameter(n)
                if p and not p.IsReadOnly:
                    p.Set(float(value)) if not isinstance(value, str) else p.Set(value)
                    return True
            except:
                continue
    return False


def get_standard_breaker(target):
    for b in BREAKERS:
        if b >= target:
            return b
    return BREAKERS[-1]


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

    with Transaction(doc, "Sincronizar Circuitos") as t:
        t.Start()
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(opts)

        for c in circuits:
            num = c.CircuitNumber or "?"
            updated = []
            err = []

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
                p_volt = c.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
                v_nominal = p_volt.AsDouble() if p_volt else 127.0
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
                section = PAIRED_SECTIONS.get(rating, 2.5)

                for _ in range(len(SECTIONS)):
                    vd = calculate_vd(dist_m, i_proj, v_nominal, section, phases=poles)
                    iz = AMPACITIES.get(section, 0.0) * fca * fct
                    if (vd <= 3.0 or dist_m <= 0) and iz >= rating:
                        break
                    idx = next((i for i, s in enumerate(SECTIONS) if s == section), -1)
                    if idx == -1 or idx + 1 >= len(SECTIONS):
                        break
                    section = SECTIONS[idx + 1]

                if write_param(c, builtin=BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM, value=rating):
                    updated.append("Disj: {}A".format(int(rating)))
                if write_param(c, names=["Seção do Condutor Adotado (mm²)", "Condutor Adotado"], value=section):
                    updated.append("Cabo: {}mm²".format(section))
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
