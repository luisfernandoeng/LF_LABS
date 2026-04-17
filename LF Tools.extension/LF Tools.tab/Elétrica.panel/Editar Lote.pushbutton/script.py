# -*- coding: utf-8 -*-
"""Editar Circuitos em Lote
Autor: LF Tools
Edita Nome da Carga, Distância e Auto-Dimensiona Cabo+Disjuntor
de múltiplos circuitos de um painel, de uma vez.
"""

__title__ = "Editar\nLote"
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
from pyrevit import forms, script, revit

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# ================================================================
# FILTRO DE SELEÇÃO
# ================================================================

class PanelFilter(ISelectionFilter):
    def AllowElement(self, e):
        return (e.Category and
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment))
    def AllowReference(self, ref, pos):
        return False


class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        for f in fa.GetFailureMessages():
            if f.GetSeverity() == FailureSeverity.Warning:
                fa.DeleteWarning(f)
        return FailureProcessingResult.Continue


# ================================================================
# FUNÇÕES AUXILIARES
# ================================================================

def get_panel_name(panel):
    try:
        for n in ["Nome do painel", "Panel Name", "Mark"]:
            p = panel.LookupParameter(n)
            if p and p.HasValue:
                return p.AsString()
        return panel.Name
    except:
        return "Quadro"


def read_param_string(elem, builtin=None, names=None):
    """Lê parâmetro como texto legível."""
    if builtin:
        try:
            p = elem.get_Parameter(builtin)
            if p and p.HasValue:
                vs = p.AsValueString()
                if vs:
                    return vs
                if p.StorageType == StorageType.String:
                    return p.AsString() or ""
                if p.StorageType == StorageType.Double:
                    return str(round(p.AsDouble(), 2))
                if p.StorageType == StorageType.Integer:
                    return str(p.AsInteger())
        except:
            pass
    if names:
        for n in names:
            try:
                p = elem.LookupParameter(n)
                if p and p.HasValue:
                    vs = p.AsValueString()
                    if vs:
                        return vs
                    if p.StorageType == StorageType.String:
                        return p.AsString() or ""
                    if p.StorageType == StorageType.Double:
                        return str(round(p.AsDouble(), 2))
                    if p.StorageType == StorageType.Integer:
                        return str(p.AsInteger())
            except:
                continue
    return ""


# Tabela de disjuntores e seções comerciais (Brasil)
BREAKERS = [16, 20, 25, 32, 40, 50, 63, 70, 80, 100, 125, 160, 200, 225, 250, 300, 400]
SECTIONS = [1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240]

# Ampacidade (Iz) NBR 5410 - Método B1 (Eletroduto embutido), 2 condutores carregados, PVC 70°C
AMPACITIES = {
    1.5: 17.5, 2.5: 24.0, 4.0: 32.0, 6.0: 41.0, 10.0: 57.0, 16.0: 76.0,
    25.0: 101.0, 35.0: 125.0, 50.0: 151.0, 70.0: 192.0, 95.0: 232.0,
    120.0: 269.0, 150.0: 309.0, 185.0: 353.0, 240.0: 415.0
}

# Mapeamento Padrão: Disjuntor -> Seção do Cabo (mm2) para Coordenação
PAIRED_SECTIONS = {
    16: 2.5, 20: 2.5, 25: 4.0, 32: 6.0, 40: 10.0, 
    50: 10.0, 63: 16.0, 80: 25.0, 100: 35.0,
    125: 50.0, 160: 70.0, 200: 95.0, 225: 120.0, 250: 120.0,
    300: 150.0, 400: 240.0
}

def get_standard_breaker(target):
    """Retorna o próximo disjuntor padrão >= target."""
    for b in BREAKERS:
        if b >= target:
            return b
    return BREAKERS[-1]

def calculate_vd(length_m, current, voltage, section_mm2, phases=1):
    """Calcula queda de tensão percentual simplificada (Cobre)."""
    if section_mm2 <= 0 or voltage <= 0: return 0.0
    # Fator: 2 para Monofásico/Bifásico (Fase-Neutro/Fase-Fase), sqrt(3) para Trifásico
    factor = 2.0 if phases <= 2 else math.sqrt(3)
    # rho cobre ~ 1/56
    vd_v = (factor * length_m * current) / (56.0 * section_mm2)
    return (vd_v * 100.0) / voltage

def get_next_section(current_s):
    """Retorna a próxima seção comercial de cabo."""
    for s in SECTIONS:
        if s > current_s:
            return s
    return current_s

def read_param_double(elem, builtin=None, names=None):
    """Lê parâmetro como double (unidades internas)."""
    if builtin:
        try:
            p = elem.get_Parameter(builtin)
            if p and p.HasValue and p.StorageType == StorageType.Double:
                return p.AsDouble()
        except:
            pass
    if names:
        for n in names:
            try:
                p = elem.LookupParameter(n)
                if p and p.HasValue and p.StorageType == StorageType.Double:
                    return p.AsDouble()
            except:
                continue
    return 0.0


def write_param(elem, builtin=None, names=None, value=None):
    """Tenta gravar parâmetro. Retorna True se gravou."""
    if builtin:
        try:
            p = elem.get_Parameter(builtin)
            if p and not p.IsReadOnly:
                if isinstance(value, str):
                    p.Set(value)
                else:
                    p.Set(float(value))
                return True
        except:
            pass
    if names:
        for n in names:
            try:
                p = elem.LookupParameter(n)
                if p and not p.IsReadOnly:
                    if isinstance(value, str):
                        p.Set(value)
                    else:
                        p.Set(float(value))
                    return True
            except:
                continue
    return False


# ================================================================
# NOMES DE PARÂMETRO DA TABELA DO USUÁRIO
# ================================================================

DIST_PARAM_NAMES = [
    "L Considerado", "L Considerado (m)", 
    "Comprimento do circuito", "Circuit Length",
    "Distância", "Distance", "Comprimento"
]


def find_distance_param(circuit):
    """Encontra o parâmetro de distância editável do circuito.
    Retorna tupla (Parameter, nome) ou (None, None)."""
    # Primeiro tenta built-in
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH)
        if p and not p.IsReadOnly:
            return p, "Comprimento do circuito (BuiltIn)"
    except:
        pass

    # Tenta nomes conhecidos
    for n in DIST_PARAM_NAMES:
        try:
            p = circuit.LookupParameter(n)
            if p and not p.IsReadOnly:
                return p, n
        except:
            continue

    # Busca ampla: qualquer param editável com "distân" ou "compri" ou "length" no nome
    for p in circuit.Parameters:
        try:
            nm = p.Definition.Name.lower()
            if p.IsReadOnly:
                continue
            if p.StorageType != StorageType.Double:
                continue
            if any(kw in nm for kw in ["distân", "distan", "compri", "length"]):
                return p, p.Definition.Name
        except:
            continue

    return None, None


def read_distance_m(circuit):
    """Lê distância do circuito em metros."""
    # Tenta built-in
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH)
        if p and p.HasValue:
            return p.AsDouble() * 0.3048
    except:
        pass

    # Tenta nomes conhecidos
    for n in DIST_PARAM_NAMES:
        try:
            p = circuit.LookupParameter(n)
            if p and p.HasValue and p.StorageType == StorageType.Double:
                val = p.AsDouble()
                # Se val > 100, provavelmente em pés → converte
                # Se val < 100, pode ser metros ou pés
                # Heurística: Revit armazena em pés internamente
                return val * 0.3048
        except:
            continue

    return 0.0


def read_voltage(circuit):
    """Lê tensão do circuito em Volts."""
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
        if p and p.HasValue:
            vs = p.AsValueString()
            if vs:
                num = re.sub(r'[^\d.,]', '', vs).replace(',', '.')
                return float(num) if num else 220.0
            return p.AsDouble()
    except:
        pass
    return 220.0


def read_load_va(circuit):
    """Lê carga aparente do circuito em VA."""
    try:
        p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD)
        if p and p.HasValue:
            vs = p.AsValueString()
            if vs:
                num = re.sub(r'[^\d.,]', '', vs).replace(',', '.')
                return float(num) if num else 0.0
            return p.AsDouble()
    except:
        pass
    return 0.0


# ================================================================
# SELEÇÃO DE PAINEL + CIRCUITOS
# ================================================================

def select_panel():
    """Seleciona painel e retorna (panel, [circuits])."""
    # Verifica seleção atual
    selected_ids = list(uidoc.Selection.GetElementIds())
    panel = None
    for sid in selected_ids:
        el = doc.GetElement(sid)
        if (el and el.Category and
                el.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment)):
            panel = el
            break

    if not panel:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, PanelFilter(),
                "Selecione o QUADRO"
            )
            panel = doc.GetElement(ref.ElementId)
        except:
            return None, []

    if not panel:
        return None, []

    # Obter circuitos
    systems = []
    try:
        if hasattr(panel, "MEPModel") and panel.MEPModel:
            systems = list(panel.MEPModel.GetAssignedElectricalSystems())
        else:
            systems = list(panel.GetAssignedElectricalSystems())
    except:
        forms.alert("Não foi possível obter os circuitos do painel.")
        return panel, []

    # Ordenar por número
    def sort_key(s):
        try:
            digits = re.sub(r'\D', '', s.CircuitNumber or '')
            return int(digits) if digits else 9999
        except:
            return 9999

    systems.sort(key=sort_key)
    return panel, systems


# ================================================================
# LISTAR CIRCUITOS
# ================================================================

def list_circuits(circuits):
    """Mostra resumo dos circuitos no console."""
    output.print_md("### 📋 {} Circuitos".format(len(circuits)))
    output.print_md(
        "| # | Nome da Carga | Tensão | Carga (VA) | Fio | Disj | Distância |")
    output.print_md(
        "|---|---|---|---|---|---|---|")

    for c in circuits:
        try:
            num = c.CircuitNumber or "?"
            name = read_param_string(c, BuiltInParameter.RBS_ELEC_CIRCUIT_NAME) or "-"
            voltage = read_param_string(c, BuiltInParameter.RBS_ELEC_VOLTAGE) or "-"
            load = read_param_string(c, BuiltInParameter.RBS_ELEC_APPARENT_LOAD) or "-"
            wire = read_param_string(
                c, builtin=None, names=["Seção do Condutor Adotado (mm²)", "Fio", "Wire Size"]) or "-"
            rating = read_param_string(
                c, builtin=None, names=["Proteção do circuito", "Classificação", "Rating"]) or "-"
            dist = "{:.1f}m".format(read_distance_m(c))
            output.print_md(
                "| {} | {} | {} | {} | {} | {} | {} |".format(
                    num, name, voltage, load, wire, rating, dist))
        except:
            continue


# ================================================================
# FILTRAR CIRCUITOS
# ================================================================

def filter_circuits(circuits):
    """Menu para filtrar circuitos."""
    if not circuits:
        return []

    action = forms.CommandSwitchWindow.show(
        ['✅ Todos ({})'.format(len(circuits)),
         '🔍 Filtrar por Nome',
         '📊 Faixa numérica (ex: 1-10)'],
        message="Quais circuitos editar?",
        title="Filtrar"
    )
    if not action:
        return []

    if 'Todos' in action:
        return list(circuits)

    elif 'Nome' in action:
        text = forms.ask_for_string(
            prompt="Filtrar circuitos que contenham no nome:",
            title="Filtro"
        )
        if not text:
            return []
        result = []
        for c in circuits:
            name = read_param_string(
                c, BuiltInParameter.RBS_ELEC_CIRCUIT_NAME) or ""
            if text.lower() in name.lower():
                result.append(c)
        if not result:
            forms.alert("Nenhum circuito com '{}' no nome.".format(text))
        return result

    elif 'Faixa' in action or 'aixa' in action:
        faixa = forms.ask_for_string(
            prompt="Faixa (ex: S1-S35 ou 1-10):",
            title="Faixa"
        )
        if not faixa:
            return []
        parts = re.split(r'[-–]', faixa.strip())
        if len(parts) != 2:
            forms.alert("Use formato: S1-S35 ou 1-10")
            return []
        try:
            start_n = int(re.sub(r'\D', '', parts[0].strip()) or '0')
            end_n = int(re.sub(r'\D', '', parts[1].strip()) or '0')
        except:
            forms.alert("Valores inválidos.")
            return []

        result = []
        for c in circuits:
            try:
                d = re.sub(r'\D', '', c.CircuitNumber or '')
                n = int(d) if d else 9999
                if start_n <= n <= end_n:
                    result.append(c)
            except:
                continue
        if not result:
            forms.alert("Nenhum circuito na faixa {}-{}.".format(start_n, end_n))
        return result

    return []


# ================================================================
# AÇÃO: EDITAR NOME DA CARGA
# ================================================================

def action_edit_load_name(circuits):
    """Edita Nome da Carga em lote com template {n}."""
    template = forms.ask_for_string(
        prompt=(
            "Template para Nome da Carga:\n"
            "Use {n} para numerar automaticamente (1,2,3...)\n"
            "Sem {n} → mesmo nome para todos\n\n"
            "Exemplo: Aspirador de Pó Baia {n}"
        ),
        title="Nome da Carga em Lote",
        default="Aspirador de Pó Baia {n}"
    )
    if template is None:
        return 0

    start = 1
    if '{n}' in template:
        start_str = forms.ask_for_string(
            prompt="Começar numeração em:",
            title="Início", default="1"
        )
        try:
            start = int(start_str)
        except:
            start = 1

    count = 0
    with Transaction(doc, "Editar Nome da Carga em Lote") as t:
        t.Start()
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(opts)

        for i, c in enumerate(circuits):
            name = template.replace('{n}', str(start + i))
            ok = write_param(
                c,
                builtin=BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
                names=["Nome da carga", "Load Name"],
                value=name
            )
            if ok:
                count += 1
        t.Commit()
    return count


# ================================================================
# AÇÃO: EDITAR DISTÂNCIA
# ================================================================

def action_edit_distance(circuits):
    """Edita distância em lote."""
    dist_str = forms.ask_for_string(
        prompt="Nova distância (metros) para {} circuito(s):".format(
            len(circuits)),
        title="Distância em Lote",
        default="25"
    )
    if not dist_str:
        return 0
    try:
        dist_m = float(dist_str.replace(',', '.'))
    except:
        forms.alert("Valor inválido.")
        return 0

    dist_ft = dist_m / 0.3048  # Revit usa pés internamente

    # Detecta qual parâmetro de distância usar a partir do 1º circuito
    param_obj, param_name = find_distance_param(circuits[0])

    if not param_obj:
        output.print_md("### ⚠️ Parâmetro de distância editável não encontrado")
        output.print_md("Parâmetros disponíveis no circuito:")
        for p in circuits[0].Parameters:
            try:
                ro = " 🔒" if p.IsReadOnly else " ✏️"
                val = p.AsValueString() or p.AsString() or ""
                output.print_md("- **{}**{} = {}".format(
                    p.Definition.Name, ro, val))
            except:
                continue
        return 0

    output.print_md("ℹ️ Usando parâmetro: **{}**".format(param_name))

    count = 0
    with Transaction(doc, "Editar Distância em Lote") as t:
        t.Start()
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(opts)

        for c in circuits:
            p_obj, _ = find_distance_param(c)
            if p_obj and not p_obj.IsReadOnly:
                try:
                    p_obj.Set(dist_ft)
                    count += 1
                except:
                    pass
        t.Commit()
    return count


# ================================================================

def action_sync_parameters(circuits):
    """
    Sincroniza e CORRIGE dimensionamento:
    1. Distância: Comprimento API -> metros arredondados -> L Considerado
    2. Disjuntor: (I_projeto / FCA / FCT) + 10A -> Próximo padrão
    3. Cabo: Extrai bitola Revit -> Verifica Queda de Tensão -> Aumenta se > 3%
    """
    import math
    import re
    count = 0
    synced_data = []

    with Transaction(doc, "Dimensionamento e Sincronização") as t:
        t.Start()
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(opts)

        for c in circuits:
            num = c.CircuitNumber or "?"
            updated_params = []
            errs = []
            
            # --- 1. DISTÂNCIA ---
            dist_m = 0.0
            try:
                raw_ft = c.Length
                if raw_ft > 0:
                    dist_m = float(math.ceil(raw_ft * 0.3048))
                    if write_param(c, builtin=None, names=["L Considerado", "L Considerado (m)"], value=dist_m):
                        updated_params.append("Dist: {}m".format(int(dist_m)))
            except Exception as ex:
                errs.append("Dist: " + str(ex))

            # --- 2. DADOS ELÉTRICOS BÁSICOS ---
            try:
                # Corrente de projeto (Calculada pela API)
                # No relatório do usuário: ApparentLoad / Voltage = I real
                if c.Voltage > 0:
                    i_proj = c.ApparentLoad / c.Voltage
                else:
                    # Fallback via BIP se for 0
                    p_curr = c.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_CURRENT_PARAM)
                    i_proj = p_curr.AsDouble() if p_curr else 0.0
                
                # Tensão nominal (para Queda de Tensão)
                # Buscamos o valor "legível" (127, 220) em vez do raw da API que vem escalado (2368)
                p_volt = c.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
                v_nominal = p_volt.AsDouble() if p_volt else 127.0
                
                # Polos (PolesNumber)
                poles = c.PolesNumber
                
                # Fatores de Correção
                fca_val = 1.0
                fct_val = 1.0
                try:
                    p_fca = c.LookupParameter("FCA")
                    if p_fca and p_fca.HasValue: fca_val = float(str(p_fca.AsValueString() or p_fca.AsDouble()).replace(",", "."))
                    p_fct = c.LookupParameter("FCT")
                    if p_fct and p_fct.HasValue: fct_val = float(str(p_fct.AsValueString() or p_fct.AsDouble()).replace(",", "."))
                except: pass
                
                if fca_val <= 0: fca_val = 1.0
                if fct_val <= 0: fct_val = 1.0
                
                i_corr = i_proj / (fca_val * fct_val)
                
            except Exception as ex:
                errs.append("Basicos: " + str(ex))
                i_proj, v_nominal, poles, i_corr = 0.0, 127.0, 1, 0.0
                
            # --- 3. DIMENSIONAMENTO (In > Ip_corr + 10 E Par Cabo/Disjuntor) ---
            try:
                # 3.1 Cálculo de Disjuntor Inicial
                target_rating = i_corr + 10.0
                final_rating = get_standard_breaker(target_rating)
                
                # 3.2 Cabo Inicial (Par do Disjuntor)
                section = PAIRED_SECTIONS.get(final_rating, 2.5)
                
                # 3.3 Loop de Queda de Tensão (máx 3% - Sobe Disjuntor E Cabo em Par)
                tries = 0
                while tries < 10:
                    vd = calculate_vd(dist_m, i_proj, v_nominal, section, phases=poles)
                    # Verifica também Iz_corr (coordenação NBR 5410)
                    iz_base = AMPACITIES.get(section, 0.0)
                    iz_corr = iz_base * fca_val * fct_val
                    
                    if (vd <= 3.0 or dist_m <= 0) and final_rating <= iz_corr:
                        break
                    
                    # Se falhou, aumenta o PAR
                    next_idx = -1
                    for idx, b in enumerate(BREAKERS):
                        if b == final_rating:
                            next_idx = idx + 1
                            break
                    
                    if next_idx >= len(BREAKERS) or next_idx == -1:
                        break # Limite atingido
                    
                    final_rating = BREAKERS[next_idx]
                    section = PAIRED_SECTIONS.get(final_rating, section)
                    tries += 1
                
                # Grava Disjuntor Final
                if write_param(c, builtin=BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM, value=final_rating):
                    updated_params.append("Disj: {}A".format(int(final_rating)))
                
                # Grava Cabo Final
                if write_param(c, builtin=None, names=["Seção do Condutor Adotado (mm²)", "Condutor Adotado"], value=section):
                    updated_params.append("Cabo: {}mm²".format(section))
                    
            except Exception as ex:
                errs.append("Dimensionamento: " + str(ex))
                
            if updated_params:
                count += 1
                info = "   - **Circuito {}**: {}".format(num, ", ".join(updated_params))
                if errs: info += " (Erros parciais: {})".format("; ".join(errs))
                synced_data.append(info)
            elif errs:
                synced_data.append("   - **❌ Circuito {}**: Erros: {}".format(num, "; ".join(errs)))

        t.Commit()

    if count > 0:
        output.print_md("### ✅ Sincronização e Dimensionamento")
        output.print_md("Foram processados **{}** circuitos:".format(count))
        for line in synced_data:
            output.print_md(line)
    else:
        # Tenta listar erros mesmo se não atualizou nada
        if synced_data:
            output.print_md("### ⚠️ Tentativa de Sincronização Falhou")
            for line in synced_data: output.print_md(line)
        else:
            forms.alert("Nenhum parâmetro pôde ser sincronizado ou dimensionado.")

    return count


# ================================================================
# AÇÃO: DIAGNOSTICAR PARÂMETROS
# ================================================================

def action_diagnostics(circuits):
    """Lista todos os parâmetros editáveis de um circuito para debug."""
    if not circuits:
        return

    c = circuits[0]
    num = c.CircuitNumber or "?"
    output.print_md("### 🔧 Diagnóstico do Circuito {}".format(num))
    output.print_md(
        "| Parâmetro | Valor | Editável | Tipo |")
    output.print_md(
        "|---|---|---|---|")

    for p in c.Parameters:
        try:
            name = p.Definition.Name
            ro = "🔒 Não" if p.IsReadOnly else "✏️ **Sim**"
            val = p.AsValueString() or p.AsString() or ""
            st = str(p.StorageType).replace("StorageType.", "")
            output.print_md("| {} | {} | {} | {} |".format(
                name, val, ro, st))
        except:
            continue


# ================================================================
# MENU PRINCIPAL
# ================================================================

def main():
    panel, circuits = select_panel()
    if not panel:
        return

    if not circuits:
        forms.alert("Nenhum circuito no painel '{}'.".format(
            get_panel_name(panel)))
        return

    panel_name = get_panel_name(panel)

    while True:
        action = forms.CommandSwitchWindow.show(
            [
                '📋 Listar Circuitos ({})'.format(len(circuits)),
                '✏️ Editar Nome da Carga em Lote',
                '📏 Editar Distância em Lote',
                '⚡ Sincronizar Distância e Cabo Adotado',
                '🔧 Diagnóstico de Parâmetros',
                '🔄 Trocar Quadro',
                '◀ Sair',
            ],
            message="Quadro: {} | {} circuito(s)".format(
                panel_name, len(circuits)),
            title="Editar Circuitos em Lote"
        )

        if not action or 'Sair' in action:
            break

        if 'Listar' in action:
            list_circuits(circuits)

        elif 'Nome' in action:
            filtered = filter_circuits(circuits)
            if filtered:
                count = action_edit_load_name(filtered)
                forms.toast(
                    "✅ {} nome(s) atualizado(s)".format(count))

        elif 'Editar Dist' in action:
            filtered = filter_circuits(circuits)
            if filtered:
                count = action_edit_distance(filtered)
                if count > 0:
                    forms.toast(
                        "✅ {} distância(s) atualizada(s)".format(count))

        elif 'Sincronizar' in action:
            filtered = filter_circuits(circuits)
            if filtered:
                count = action_sync_parameters(filtered)
                forms.toast(
                    "✅ {} circuito(s) sincronizado(s)".format(count))

        elif 'Diagnóstico' in action or 'Diagn' in action:
            action_diagnostics(circuits)

        elif 'Trocar' in action:
            panel, circuits = select_panel()
            if panel:
                panel_name = get_panel_name(panel)
            else:
                break


if __name__ == '__main__':
    main()
