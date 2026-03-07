# Ideias e Melhorias para o LF Tools

Este documento compila as ideias de novos plugins e propostas de atualização para as ferramentas já existentes na sua aba **LF Tools**. Use este material como guia para as próximas implementações.

---

## 🌟 1. Novos Plugins Propostos

### 1.1. Auto-Cotas (Dimensionamento Automático)
* **O que faz:** Gera cotas alinhadas automaticamente para múltiplos elementos de uma categoria selecionada de uma só vez. Por exemplo, permite que o usuário trace uma "linha guia" e o plugin cota todas as paredes, portas ou eixos que cruzam essa linha, ou cota automaticamente o contorno de um ambiente.
* **Por que é útil:** O detalhamento é a parte mais demorada do Revit. Um plugin que automatiza a colocação de cotas economiza horas de cliques manuais e reduz erros, sendo um diferencial enorme para a produtividade de qualquer escritório.

### 1.2. Copiar/Transferir Configurações Avançadas (Transfer Settings Prop)
* **O que faz:** Vai muito além do "Transfer Project Standards" nativo do Revit. Permite que o usuário abra uma interface e escolha exatamente *quais* Filtros de Vista, *quais* View Templates ou *quais* Regras de Navegador ele quer copiar de um modelo (linkado ou aberto na sessão) para o modelo atual. Permite selecionar por meio de checkboxes a transferência de:
  - Filtros de Vista (selecionados por nome)
  - Modelos de Vista (View Templates)
  - Padrões de Preenchimento (Fill Patterns) e Materiais específicos
  - Parâmetros Compartilhados do Projeto
* **Por que é útil:** Atualmente, a transferência nativa do Revit copia tudo de uma vez. Ter uma ferramenta que permita selecionar "cirurgicamente" apenas 2 filtros e 1 modelo de vista seria uma salvação para coordenadores de BIM que gerenciam templates.

### 1.3. Alinhador de Tags e Textos (Tag/Text Aligner)
* **O que faz:** O usuário seleciona múltiplas tags, textos ou cotas e escolhe alinhar (Top, Bottom, Left, Right) ou distribuir uniformemente o espaço entre eles.
* **Por que é útil:** Facilita muito o detalhamento da prancha, mantendo os desenhos com aparência profissional e limpa, poupando tempo de alinhamento manual.

### 1.4. Renumerador Sequencial por Clique (Click-Numbering)
* **O que faz:** Ferramenta interativa onde o usuário define um prefixo/sufixo e valor inicial, e clica na tela. Cada elemento clicado (estacas, vagas, portas) recebe o próximo número da sequência.
* **Por que é útil:** Ideal para elementos não lineares ou organizados de forma arbitrária que não podem ser renumerados por um simples "sort", mas que precisam de uma ordem lógica baseada no projeto.

### 1.5. Modelador de Rodapés/Acabamentos por Ambiente (Room to Finishes)
* **O que faz:** Lê o perímetro de um ambiente selecionado, subtrai vãos de porta e cria automaticamente paredes finas (ou sweeps) de revestimento, rodapé ou gesso.
* **Por que é útil:** Automatiza uma das tarefas de modelagem mais repetitivas em detalhamento de interiores.

### 1.6. Resolvedor de Avisos (Smart Warnings Manager)
* **O que faz:** Uma janela (WPF) que agrupa os avisos (*Warnings*) do Revit por importância ou categoria, permitindo isolar elementos em 3D e aplicando resoluções automáticas (ex: acionar o Overkill para *overlapping elements*).
* **Por que é útil:** Transforma a interface terrível e confusa de erros nativa do Revit em um dashboard gerencial tratável e amigável para o Coordenador BIM.

### 1.7. Criador de Vistas por Ambiente (Room Views Automator)
* **O que faz:** Com um clique, gera as vistas necessárias para o ambiente (1 Planta Ampliada, 1 Forro, 4 Elevações Internas), ajusta os Crop Regions correspondentes e aplica os View Templates corretos.
* **Por que é útil:** "Ferramenta Mágica" que poupa dias e dias de trabalho tedioso criando e recortando elevações de interiores e paginações de parede.

### 1.8. Deep Purge (Limpeza Profunda)
* **O que faz:** Limpa os "restos" que o Purge nativo não pega: Padrões de Linha inúteis, View Templates não aplicados, Filtros de Vista órfãos, Viewports sem uso, Tags/CADs escondidos.
* **Por que é útil:** Essencial para entrega de modelo "As-Built" ou "LOD 400", otimizando consideravelmente o tamanho do arquivo.

---

## 🚀 2. Melhorias para os Plugins Existentes

### 2.1. Inverter Anotação (Suporte a Plantas Rotacionadas)
**O Problema Atual:**
O script atual usa vetores absolutos globais (`XYZ(0, 1, 0)` e `XYZ(1, 0, 0)`) para decidir o que é "horizontal" e "vertical". Quando a vista do usuário está rotacionada (por um Scope Box ou rotação do Norte do Projeto na folha), o plugin inverte a anotação na diagonal em vez de seguir o "papel", causando confusão.

**A Solução Técnica:**
Em vez de usar `XYZ` fixos, os eixos de espelhamento devem ser lidos diretamente a partir da geometria da vista ativa no momento da execução:
* Horizontal (Normal à direção Cima da tela): `active_view.UpDirection`
* Vertical (Normal à direção Direita da tela): `active_view.RightDirection`
Com essa pequena mudança na matemática, a anotação vai "tombar" perfeitamente seguindo os olhos do usuário, independentemente de como a planta foi rotacionada!

### 2.2. Renomear+ (Renomeação Dinâmica + Case Converter)
**A Melhoria 1: Parâmetros Dinâmicos**
* **O que é:** Implementar o mesmo sistema de "Tags" que já existe no *Gerar Folhas*. Ao renomear um lote de elementos (como portas ou eixos), o usuário poderá compor o novo nome puxando valores de outros parâmetros daquele próprio elemento usando chaves.
* **Exemplo de Uso:** Um prefixo como `PT-{Nível}-`. O plugin leria o parâmetro "Nível" de cada porta isoladamente e renomearia para `PT-Térreo-01`, `PT-Pavimento 1-02`, etc. Isso transforma a ferramenta em um renomeador paramétrico ultra-poderoso, concorrendo com os melhores plugins pagos do mercado.

**A Melhoria 2: Case Converter (Maiúsculas e Minúsculas)**
* **O que é:** Adicionar botões rápidos na janela WPF para formatar o lote inteiro para: TUDO MAIÚSCULO, tudo minúsculo, ou Iniciais Maiúsculas `.upper()`, `.lower()`, `.title()`. 
* **Exemplo de Uso:** Padronizar nomes de Vistas e Modelos de Vista criados por diferentes projetistas para que o navegador fique organizado.

### 2.3. Gerar Folhas (Perfis de Exportação / Presets)
**O Problema Atual:**
Se o usuário exporta "Versão Prefeitura" em PDF e depois "Versão Executivo" em DWG (com padrões de nomes diferentes), ele precisa reconfigurar as marcações da interface toda vez.

**A Solução Técnica:**
Adicionar um campo no topo (ComboBox) chamado "Perfil:". O usuário configura tudo (padrão de nome, destino, formato) e clica em "Salvar Perfil". O plugin salva esses dados num arquivo JSON na própria pasta da extensão (`AppData\Roaming\pyRevit\Extensions`). No próximo uso, basta ele selecionar "Perfil Executivo" e todos os checkboxes, padrões de texto e caminhos de pasta se preenchem sozinhos, ficando disponível como um Preset global da máquina para qualquer projeto.

### 2.4. Overkill (Preview Analysis de Segurança)
**O Problema Atual:**
Rodar um plugin que apaga duplicatas sempre gera um "frio na barriga" no usuário, que fica com medo de o algoritmo ter apagado algo importante que estava levemente sobreposto por engano.

**A Solução Técnica:**
Após clicar em "Analisar", em vez de abrir um *forms.alert* direto, abrir uma pequena lista (DataGrid) listando as duplicatas encontradas, separadas por Categoria. Cada item teria um checkbox. O usuário bate o olho, vê que foram "2 Paredes" e "10 Textos", e clica em "Confirmar Exclusão". Isso dá controle e confiança total na ferramenta, transformando-a de "arriscada" numa de auditoria de qualidade.
