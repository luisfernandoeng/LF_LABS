# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationFramework')
from System.Windows import Window, MessageBox

from pyrevit import forms, DB, revit
from QuedaTensao.queda_tensao_engine import QuedaTensaoEngine

class QuedaTensaoWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, doc, element_id, base_data):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.doc = doc
        self.element_id = element_id
        
        # Guardar valores de engenharia na instância
        self.diam_mm = base_data.get('diam_mm', 0.0)
        self.comp_m = base_data.get('length_m', 0.0)

        # Preenche os Texts inoperacionais
        self.TxtDiametro.Text = "Ø {} mm".format(self.diam_mm)
        self.TxtComprimento.Text = "{} m".format(self.comp_m)

        self.resultado_atual = None
        
        # Eventos para "Live Update" (Software de Teste)
        self.InputPotencia.TextChanged += self.on_calc_click
        self.InputCorrente.TextChanged += self.on_calc_click
        self.InputFP.TextChanged += self.on_calc_click
        self.InputAgrupamento.TextChanged += self.on_calc_click
        
        self.ComboCarga.SelectionChanged += self.on_calc_click
        self.ComboTensao.SelectionChanged += self.on_calc_click
        self.ComboCircuito.SelectionChanged += self.on_calc_click
        self.ComboBitola.SelectionChanged += self.on_calc_click

    def _safe_get_combo_text(self, combo):
        try:
            if combo.SelectedItem:
                if hasattr(combo.SelectedItem, 'Content'):
                    return str(combo.SelectedItem.Content)
                return str(combo.SelectedItem)
            return combo.Text
        except:
            return ""
    
    def on_calc_click(self, sender, args):
        try:
            # 1. Parsing the Inputs
            carga_str = self._safe_get_combo_text(self.ComboCarga)
            if "TUG" in carga_str: carga_str = "TUG"
            elif "TUE" in carga_str: carga_str = "TUE"

            circuito_str = self._safe_get_combo_text(self.ComboCircuito)
            tipo_circuito = 'mono'
            if 'Bi' in circuito_str: tipo_circuito = 'bi'
            elif 'Tri' in circuito_str: tipo_circuito = 'tri'

            circ_idx = self.ComboCircuito.SelectedIndex
            # 0: 1F+N+T, 1: 2F+T, 2: 2F+N+T, 3: 3F+T, 4: 3F+N+T
            
            tipo_circuito = 'mono' if circ_idx <= 0 else ('bi' if circ_idx <= 2 else 'tri')
            
            # Número total de fios para ocupação (Manual override)
            wires_map = {0: 3, 1: 3, 2: 4, 3: 4, 4: 5}
            total_wires = wires_map.get(circ_idx, 3)

            v_str = self._safe_get_combo_text(self.ComboTensao)
            v_val = 220
            if '127' in v_str: v_val = 127
            elif '380' in v_str: v_val = 380

            pot_txt = self.InputPotencia.Text.replace(',', '.')
            potencia_f = float(pot_txt) if pot_txt else 0.0

            curr_txt = self.InputCorrente.Text.replace(',', '.')
            corrente_f = float(curr_txt) if curr_txt else None

            # Validação silenciosa para live update
            if not potencia_f and not corrente_f:
                self.ResStatusGlobal.Text = "Aguardando Potência ou Corrente..."
                return
            
            fp_txt = self.InputFP.Text.replace(',', '.')
            fp = float(fp_txt) if fp_txt else 1.0

            agrup_txt = self.InputAgrupamento.Text
            agrup = int(agrup_txt) if agrup_txt else 1

            bit_manual_str = self._safe_get_combo_text(self.ComboBitola)
            bit_manual = None
            if bit_manual_str and bit_manual_str != "Automático":
                bit_manual = bit_manual_str

            inp_dict = {
                'diam_eletroduto_mm': self.diam_mm,
                'comprimento_m': self.comp_m,
                'tipo_carga': carga_str,
                'potencia_w': potencia_f,
                'corrente_a': corrente_f,
                'tensao_v': v_val,
                'fp': fp,
                'tipo_circuito': tipo_circuito,
                'num_fios': total_wires,
                'num_circuitos_agrup': agrup,
                'bitola_manual': bit_manual
            }

            # Aciona Engine de Cálculo
            res = QuedaTensaoEngine.calcular_dimensionamento(inp_dict)
            self.resultado_atual = res

            # Renderiza Resultados
            self._render_results(res)
            
            # Libera Botão de Aplicar
            self.BtnAplicar.IsEnabled = res.get('ok', False)

        except Exception as e:
            # Em vez de MessageBox, usa o status label pra não interromper o fluxo de teste
            self.ResStatusGlobal.Text = "Erro: Parâmetros incompletos"
            from System.Windows.Media import BrushConverter
            self.ResStatusGlobal.Foreground = BrushConverter().ConvertFromString("#ff5252")
    
    def _render_results(self, res):
        # UI Setup - Resetting display
        self.ResTuboDir.Text = ""
        self.ResStatusGlobal.Text = ""
        
        if not res['bitola']:
            self.ResBitola.Text = "❌ Erro Crítico: Seleção Impossível."
            self.ResCorrente.Text = ""
            self.ResOcupacao.Text = ""
            self.ResQueda.Text = ""
            self.ResDistMax.Text = ""
            self.ResStatusGlobal.Text = "\n".join(res['erros'])
            from System.Windows.Media import BrushConverter
            self.ResStatusGlobal.Foreground = BrushConverter().ConvertFromString("#ff5252")
            return

        # 1. Bitola e Motivo
        self.ResBitola.Text = "✅ Bitola: {} mm²".format(res['bitola'])
        
        corrente_info = "Corrente: {}A".format(res['corrente'])
        if res.get('fca', 1.0) < 1.0:
            corrente_info += " (Agrupamento FCA {}: Corrigida p/ {}A)".format(res['fca'], res['corrente_fca'])
        self.ResCorrente.Text = corrente_info
        
        # 2. Ocupação Conduíte
        tubo = res['eletroduto_recomendado']
        tubo_atual = res['eletroduto_atualizado_para_check']
        
        self.ResOcupacao.Text = "Condutores: {} | Ocupação do Tubo Ø {}: {}% (Max {}%)".format(
            res['num_condutores'], tubo, res['ocupacao'], int(res['limite_ocupacao'])
        )

        dir_msgs = []
        dir_msgs.append("🔍 Seleção Baseada em: {}".format(res.get('motivo_bitola', 'N/A')))

        if float(tubo) > float(tubo_atual):
            dir_msgs.append("⚠️ Tubulação ATUAL ({}mm) não comporta os cabos calculados!".format(tubo_atual))
        elif float(tubo) < float(tubo_atual):
            dir_msgs.append("ℹ️ O Eletroduto ATUAL ({}mm) possui folga confortável.".format(tubo_atual))
        else:
            dir_msgs.append("✅ Tubulação ATUAL ({}mm) está ideal para norma.".format(tubo_atual))
        
        for d in res['dicas']: dir_msgs.append("💡 " + d)
        for e in res['erros']: dir_msgs.append("⚠️ " + e)

        self.ResTuboDir.Text = "\n".join(dir_msgs)

        # 3. Drop Tensao e Distância
        q_ok = "✅" if res['queda_ok'] else "❌ Excede 4%"
        self.ResQueda.Text = "Queda de Tensão ({}m): {}% {}".format(self.comp_m, res['queda_tensao'], q_ok)
        
        # Feedback de Distância como solicitado
        dist_msg = "📏 Com o cabo de {}mm², você pode chegar a até {} metros nesta carga.".format(res['bitola'], res['dist_maxima'])
        self.ResDistMax.Text = dist_msg
        
        # Global Status
        if res['ok'] and res['queda_ok']:
            self.ResStatusGlobal.Text = "✅ PROJETO CONFORME NBR 5410"
            self.ResStatusGlobal.Foreground = self.FindResource("AccentColor")
        else:
            self.ResStatusGlobal.Text = "❌ PROJETO COM RESTRIÇÕES/ALERTAS"
            from System.Windows.Media import BrushConverter
            self.ResStatusGlobal.Foreground = BrushConverter().ConvertFromString("#ff5252")

    def on_apply_click(self, sender, args):
        if not self.resultado_atual:
            return
            
        res = self.resultado_atual
        t = DB.Transaction(self.doc, "Aplicar Info Queda de Tensao")
        t.Start()
        
        try:
            el = self.doc.GetElement(self.element_id)
            memo = "Dimensionado via Queda de Tensão PRO:\n"
            memo += "Bitola: {}mm²\n".format(res['bitola'])
            memo += "Queda(Atual): {}%\n".format(res['queda_tensao'])
            memo += "Ocupacao do Tubo: {}%\n".format(res['ocupacao'])
            if res['ok'] and res['queda_ok']:
                memo += "Status: Conforme"
            else:
                memo += "Status: Atencao Plena"

            comments_param = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if comments_param and not comments_param.IsReadOnly:
                comments_param.Set(memo)
            
            t.Commit()
            forms.alert("Dimensionamento aplicado com sucesso nos Comentários!", title="Sucesso")
            self.Close()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Erro Gravando no Revit: " + str(e), "Transaction Error")
