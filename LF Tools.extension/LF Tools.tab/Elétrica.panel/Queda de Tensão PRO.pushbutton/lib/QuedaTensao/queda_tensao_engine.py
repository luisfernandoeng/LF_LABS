# -*- coding: utf-8 -*-
import math

# =========================================================================
# TABELA 1: CAPACIDADE DE CONDUÇÃO (Ampacidade)
# NBR 5410 Tabela 36 (Método B1 - Eletroduto embutido - Cobre - PVC 70C)
# Formato: 'Bitola': (Ampacidade A, Diâmetro total externo em mm)
# =========================================================================
BITOLAS_CAPACIDADE = {
    '1.5': (15.5, 3.0),
    '2.5': (21.0, 3.4),
    '4.0': (28.0, 4.0),
    '6.0': (36.0, 4.6),
    '10':  (50.0, 5.8),
    '16':  (68.0, 7.0),
    '25':  (89.0, 8.6),
    '35':  (111.0, 10.0),
    '50':  (134.0, 11.5),
    '70':  (171.0, 13.5),
    '95':  (207.0, 15.5),
    '120': (239.0, 17.5),
    '150': (271.0, 19.5),
    '185': (309.0, 21.5),
    '240': (364.0, 24.5),
}
# Ordered list of keys for step up/down logic
ORDERED_BITOLAS = ['1.5', '2.5', '4.0', '6.0', '10', '16', '25', '35', '50', '70', '95', '120', '150', '185', '240']

# =========================================================================
# TABELA 2: DIÂMETROS DE ELETRODUTOS
# Polegadas Nominal: Diâmetro Interno Estimado em mm
# =========================================================================
ELETRODUTOS_DIAM_INTERNO = {
    '16': 12.0,   # 1/2"
    '20': 16.0,   # 3/4"
    '25': 21.0,   # 1"
    '32': 27.0,   # 1.1/4"
    '40': 35.0,   # 1.1/2"
    '50': 44.0,   # 2"
    '60': 53.0,   # 2.1/2"
    '75': 69.0,   # 3"
    '85': 78.0,   # 3.1/2"
    '100': 91.0,  # 4"
}
ORDERED_ELETRODUTOS = ['16', '20', '25', '32', '40', '50', '60', '75', '85', '100']

# =========================================================================
# TABELA 3: FATORES DE CORREÇÃO / AGRUPAMENTO NBR 5410
# =========================================================================
FATOR_AGRUPAMENTO = {
    1: 1.00,
    2: 0.80,
    3: 0.70,
    4: 0.65,
    5: 0.60,
    6: 0.57,
    7: 0.54,
    8: 0.52,
    9: 0.50,
}

# =========================================================================
# TABELA 4: RESISTIVIDADE E REATÂNCIA (Cobre - Padrão C=58)
# =========================================================================
RESISTIVIDADE_COBRE_70C = 0.0217 # Ω.mm²/m (Padrao NBR 5410 a 70C)
RESISTIVIDADE_COBRE_BASE = 1.0 / 58.0 # Ω.mm²/m (Aprox 0.01724 - Sugestão Imagem)

# Reatância indutiva aproximada em Ω/km
# (Se quiser ignorar reatância para simplificar como na foto, basta usar X=0)
REATANCIA = {
    '1.5': 0.140,
    '2.5': 0.130,
    '4.0': 0.120,
    '6.0': 0.110,
    '10':  0.100,
    '16':  0.095,
    '25':  0.090,
    '35':  0.085,
    '50':  0.080,
    '70':  0.075,
    '95':  0.072,
    '120': 0.070,
    '150': 0.068,
    '185': 0.066,
    '240': 0.064,
}

# =========================================================================
# MOTOR E FUNÇÕES CORE
# =========================================================================

class QuedaTensaoEngine(object):

    @staticmethod
    def _encontrar_eletroduto_adequado(num_condutores, diam_cabo_mm):
        """Retorna o eletroduto ideal do mercado para o número de cabos e bitola atuais"""
        # Limite permitido
        if num_condutores == 1:
            limite = 0.53
        elif num_condutores == 2:
            limite = 0.31
        else:
            limite = 0.40

        area_condutor = math.pi * ((float(diam_cabo_mm) / 2.0) ** 2)
        area_ocupada = num_condutores * area_condutor

        for eletroduto_dn in ORDERED_ELETRODUTOS:
            diam_interno = ELETRODUTOS_DIAM_INTERNO[eletroduto_dn]
            area_eletroduto = math.pi * ((float(diam_interno) / 2.0) ** 2)
            
            ocupacao_percent = (area_ocupada / area_eletroduto)
            if ocupacao_percent <= limite:
                return eletroduto_dn, round(ocupacao_percent * 100, 1), round(limite * 100, 1)
        
        # Caso ultrapasse o máximo (100mm)
        return "100", 100.0, round(limite*100, 1)

    @staticmethod
    def estimar_eletroduto_proximo(diam_atual_mm):
        """Mapeia um diâmetro customizado do modelo nativo pro mais próximo da tabela pra testes de folga"""
        closest = None
        min_diff = float('inf')
        
        for elet in ORDERED_ELETRODUTOS:
            diff = abs(float(elet) - float(diam_atual_mm))
            if diff < min_diff:
                min_diff = diff
                closest = elet
        return closest

    @staticmethod
    def calcular_dimensionamento(inputs):
        """ 
        inputs dict expected:
        - diam_eletroduto_mm (float)
        - comprimento_m (float)
        - tipo_carga ('Iluminação', 'TUG', 'TUE', 'Motor')
        - potencia_w (float ou None)
        - corrente_a (float ou None)
        - tensao_v (int)
        - fp (float)
        - tipo_circuito ('mono', 'bi', 'tri')
        - num_circuitos_agrup (int)
        - bitola_manual (str ou None)
        """
        
        erros = []
        dicas = []

        # 1. CALCULA A CORRENTE
        if inputs.get('corrente_a'):
            corrente = float(inputs['corrente_a'])
        else:
            p = float(inputs['potencia_w'])
            v = float(inputs['tensao_v'])
            fp = float(inputs['fp'])
            if inputs['tipo_circuito'] in ['mono', 'bi']:
                corrente = p / (v * fp)
            else:
                corrente = p / (1.732 * v * fp)
        
        # 2. SELECIONA A BITOLA (Por Ampacidade Corrigida)
        agrup = int(inputs['num_circuitos_agrup'])
        if agrup > 9: agrup = 9
        if agrup < 1: agrup = 1
        
        fca = FATOR_AGRUPAMENTO[agrup]
        corrente_corrigida = corrente / fca

        bitola_recomendada = None
        bitola_manual = inputs.get('bitola_manual')

        if bitola_manual and bitola_manual in BITOLAS_CAPACIDADE:
            bitola_recomendada = bitola_manual
            amp_manual = BITOLAS_CAPACIDADE[bitola_manual][0]
            if amp_manual < corrente_corrigida:
                erros.append("⚠️ Bitola Manual ({}mm²) insuficiente para Corrente Corrigida ({}A). (Capacidade: {}A)".format(bitola_manual, round(corrente_corrigida,1), amp_manual))
        else:
            # Automático
            for b in ORDERED_BITOLAS:
                amp, diam = BITOLAS_CAPACIDADE[b]
                if amp >= corrente_corrigida:
                    carga = inputs['tipo_carga']
                    if carga == 'Iluminação' and float(b) < 1.5: continue
                    elif carga in ['TUG', 'TUE'] and float(b) < 2.5: continue
                    bitola_recomendada = b
                    break
        
        if not bitola_recomendada:
            erros.append("Corrente ({}A) muito alta para as seções padronizadas.".format(round(corrente_corrigida,1)))
            return {
                "ok": False, 
                "bitola": None,
                "erros": erros,
                "dicas": dicas,
                "corrente": round(corrente, 1),
                "corrente_fca": round(corrente_corrigida, 1),
                "fca": fca
            }

        # 3. CONDUTORES E OCUPAÇÃO
        num_condutores = inputs.get('num_fios', 3)

        diam_cabo = BITOLAS_CAPACIDADE[bitola_recomendada][1]
        elet_ideal, ocup_ideal, limite = QuedaTensaoEngine._encontrar_eletroduto_adequado(num_condutores, diam_cabo)
        elet_atual_mapper = QuedaTensaoEngine.estimar_eletroduto_proximo(inputs['diam_eletroduto_mm'])

        if float(elet_ideal) > float(elet_atual_mapper):
            erros.append("⚠️ Tubulação ATUAL (Ø {}mm) não comporta a bitola {}mm² (Ocupação: {}%)".format(elet_atual_mapper, bitola_recomendada, round(ocup_ideal,1)))
        
        # 4. CALCULO DA QUEDA DE TENSÃO
        c = 58.0 # Cobre
        tensao = float(inputs['tensao_v'])
        comp = float(inputs['comprimento_m'])
        s_mm2 = float(bitola_recomendada)
        fator = 1.732 if inputs['tipo_circuito'] == 'tri' else 2.0

        # Formula da imagem: Δv (V) = (I * L * Fator) / (C * S)
        delta_v_absoluto = (corrente * comp * fator) / (c * s_mm2)
        queda_percentual = (delta_v_absoluto / tensao) * 100.0

        # Distância máxima (4%)
        if corrente > 0:
            L_max = (0.04 * tensao * c * s_mm2) / (corrente * fator)
        else:
            L_max = 9999

        motivo_bitola = "Ampacidade p/ Corrente Corrigida"
        if bitola_manual:
            motivo_bitola = "Seleção Manual do Usuário"
            
        # Fallback se não for manual
        if not bitola_manual and queda_percentual > 4.0:
            idx_atual = ORDERED_BITOLAS.index(bitola_recomendada)
            for idx in range(idx_atual+1, len(ORDERED_BITOLAS)):
                b_test = ORDERED_BITOLAS[idx]
                s_test = float(b_test)
                dv_t = (corrente * comp * fator) / (c * s_test)
                q_t = (dv_t / tensao) * 100.0
                
                if q_t <= 4.0 or idx == len(ORDERED_BITOLAS)-1:
                    dicas.append("ℹ️ Bitola elevada para {}mm² para atender limite de 4% de queda.".format(b_test))
                    bitola_recomendada = b_test
                    queda_percentual = q_t
                    motivo_bitola = "Limite de Queda de Tensão (4%)"
                    
                    # Recalcula ocupação com a nova bitola
                    diam_cabo = BITOLAS_CAPACIDADE[bitola_recomendada][1]
                    elet_ideal, ocup_ideal, limite = QuedaTensaoEngine._encontrar_eletroduto_adequado(num_condutores, diam_cabo)
                    break

        queda_ok = queda_percentual <= 4.0
        if not queda_ok:
            erros.append("❌ Queda de Tensão crítica: {}% (Máximo NBR 4%)".format(round(queda_percentual,2)))

        return {
            "ok": len(erros) == 0,
            "bitola": bitola_recomendada,
            "motivo_bitola": motivo_bitola,
            "corrente": round(corrente, 1),
            "corrente_fca": round(corrente_corrigida, 1),
            "fca": fca,
            "num_condutores": num_condutores,
            "ocupacao": round(ocup_ideal, 1),
            "limite_ocupacao": limite,
            "queda_tensao": round(queda_percentual, 2),
            "queda_ok": queda_ok,
            "dist_maxima": round(L_max, 1),
            "eletroduto_recomendado": elet_ideal,
            "eletroduto_atualizado_para_check": elet_atual_mapper,
            "erros": erros,
            "dicas": dicas
        }
