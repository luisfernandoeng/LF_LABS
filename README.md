<p align="center">
  <strong>LF Tools</strong><br>
  Suite de Engenharia &amp; Automação BIM para Autodesk Revit
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Revit-2020–2025-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/pyRevit-4.8+-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/IronPython-2.7-green?style=flat-square" />
  <img src="https://img.shields.io/badge/UI-WPF%20%2F%20XAML-blueviolet?style=flat-square" />
</p>

---

**LF Tools** é uma extensão de alta performance para **Autodesk Revit**, construída sobre o ecossistema **pyRevit**.  
O objetivo é eliminar tarefas repetitivas e dar superpoderes ao engenheiro que trabalha com BIM — especialmente em projetos de **instalações elétricas**, **documentação** e **gerenciamento de dados**.

Toda a interface utiliza **WPF nativo (XAML)** com um design system próprio (Light Theme v2) e componentes padronizados, garantindo uma experiência consistente e profissional em cada ferramenta.

---

## Índice

- [Painel Anotação](#-painel-anotação)
- [Painel Automatizar](#-painel-automatizar)
- [Painel Dados e Filtros](#-painel-dados-e-filtros)
- [Painel Elétrica](#-painel-elétrica)
- [Painel Manutenção](#-painel-manutenção)
- [Painel AutoSave](#-painel-autosave)
- [Painel Dev](#-painel-dev)
- [Tech Stack](#-tech-stack)
- [Instalação](#️-instalação)
- [Contribuições](#-contribuições)

---

## ✒️ Painel Anotação

Ferramentas para refinamento e precisão em anotações técnicas.

| Ferramenta | Descrição |
|---|---|
| **Inverter Anotação** | Corrige tags e textos que ficaram espelhados após operações de `Mirror`. |
| **Merge Text** | Consolida notas de texto fragmentadas em uma única, preservando a ordem espacial (X → Y). |
| **Nome Amb** | Extrai dados de ambientes (Rooms) de arquivos vinculados e gera tags automáticas no modelo ativo. |
| **Renumerar** | Renumeração sequencial inteligente baseada na ordem de clique (click-order). |
| **Renomear+** | Renomeação em massa de Vistas, Folhas e Tabelas com preview em tempo real e padrões customizáveis. |

---

## 🤖 Painel Automatizar

Automação de processos complexos e geração de conteúdo.

| Ferramenta | Descrição |
|---|---|
| **Auto-Cotas** | Dimensionamento automático inteligente. Detecta o eixo central de elementos (conduítes, tubos, paredes) e gera cotas alinhadas em relação a referências (eixos, paredes). Interface WPF com seleção de modos. |
| **Gerar Folhas** | Criação massiva de folhas com alinhamento automático de viewports, configuração de padrões de exportação e interface de nível 2 com multi-select e preview. |
| **Smart Crop** | Ajusta o *Crop Region* das vistas ao limite exato da geometria visível, otimizando o desempenho gráfico. Inclui cópia de Crop Region entre vistas. |

---

## 📊 Painel Dados e Filtros

Manipulação granular de informações e auditoria de modelos.

| Ferramenta | Descrição |
|---|---|
| **Filtrar Avançado** | Query builder com lógica booleana (Igual, Contém, Maior que, etc.) para filtrar categorias simples ou múltiplas, com interface rica e presets salvos. |
| **Inspecionar Tipo** | Diagnóstico de `BuiltInParameters`, IDs de categoria e dados internos de famílias (MEP / Arquitetura). |
| **Soma Dist** | Totalizador métrico para elementos lineares — Eletrodutos, Eletrocalhas, Tubos e Linhas. |
| **To Excel** | Sincronização bidirecional de alta performance entre Revit e Excel (exporta e importa), sem dependência de drivers COM. Usa `xlsxwriter` embarcado. Preview de dados com seleção de abas. |

---

## ⚡ Painel Elétrica

O maior módulo da extensão — ferramentas especializadas para engenharia de instalações elétricas.

### Roteamento & Conexão

| Ferramenta | Descrição |
|---|---|
| **Conectar Eletroduto** | Conexão automática de conduit runs entre pontos elétricos, com configuração de tipo de eletroduto, ângulo de curva (90°/45°), modo de roteamento (auto/manual) e detecção de interferências (clash avoidance). Interface com settings persistentes. |
| **Conectar Eletrocalha** | Conexão automática de cable tray runs com configuração de largura, altura e tipo de eletrocalha. |

### Circuitos & Quadros

| Ferramenta | Descrição |
|---|---|
| **LF Electrical** | Central de automação de circuitos com três modos: **Residencial** (circuitos TUG/TUE, nomeação alfabética de disjuntores), **Industrial** e **Dados** (quadros de cabeamento). Configura quadros (fases/distribuição) e cria circuitos agrupados automaticamente. |
| **Gerenciar Circuito** | Adiciona/Remove elementos de circuitos existentes com highlight visual em vermelho. Suporta desconexão via conectores e deleção de circuitos inteiros. |
| **Editar Lote** | Edição em massa de circuitos de um quadro — renomeia cargas com template `{n}`, define distância em lote e auto-dimensiona cabo + disjuntor conforme NBR 5410 (Método B1, PVC 70 °C). |
| **Sincronizar Circuitos** | Lê `Length` da API, converte pés → metros, calcula disjuntor e cabo coordenados (In ≥ Ip / FCA·FCT, VD ≤ 3 %) e grava tudo em uma transação. |
| **Filtrar Elétrica** | Rastreamento recursivo de topologia de rede — seleciona todos os dispositivos conectados a um quadro. |
| **Transferir Circuitos** | Transfere circuitos entre modelos vinculados. |

### Interoperabilidade & Coordenação

| Ferramenta | Descrição |
|---|---|
| **Pontos por Vínculo** | Copia famílias elétricas de um modelo de arquitetura vinculado para o projeto MEP, sincronizando fases e posições. Inclui classificação de cargas e perfis de configuração persistentes. |
| **Substituir Elementos** | Substitui famílias em massa preservando conexões elétricas existentes. |
| **Transfer Settings** | Transferência cirúrgica de Filtros, Modelos de Vista e Padrões de Preenchimento entre arquivos abertos ou links. |

### Utilitários

| Ferramenta | Descrição |
|---|---|
| **Analisar Geometria** | Gera um relatório JSON com coordenadas (X, Y em mm) dos elementos selecionados — projetado para alimentar prompts de IA com dados geométricos. |
| **Gerar Diagrama** | Gera diagrama unifilar automático a partir do quadro elétrico selecionado, criando uma Vista de Desenho com a representação dos circuitos. |

---

## 🧹 Painel Manutenção

Higiene e performance do banco de dados do Revit.

| Ferramenta | Descrição |
|---|---|
| **Overkill** | Identifica e remove elementos duplicados/sobrepostos. Realiza purge seletivo de vistas não utilizadas. |
| **SmartSelectSimilar** | Seleção inteligente de elementos similares com filtros avançados por parâmetro — vai além do `Select Similar` nativo do Revit. |

---

## 💾 Painel AutoSave

| Ferramenta | Descrição |
|---|---|
| **Smart AutoSave** | Sistema de auto-salvamento inteligente que roda em background via `DispatcherTimer`. Inicializa automaticamente no startup do Revit (se habilitado). Configurável via interface WPF. |

---

## 🛠️ Painel Dev

Ferramentas de desenvolvimento e customização do ambiente.

| Ferramenta | Descrição |
|---|---|
| **Action Logger** | Sistema de log de ações do usuário no Revit. Permite **Start**, **Stop** e **View Log** das sessões. Útil para auditoria de produtividade e debug de workflows. |
| **Ocultar Abas** | Oculta/exibe abas do Ribbon do Revit no startup — reduz poluição visual para focar nas ferramentas necessárias. Configuração persistente via `pyrevit script config`. |

---

## 💻 Tech Stack

| Camada | Tecnologia |
|---|---|
| **Core** | Autodesk Revit API (2020–2025) |
| **Linguagem** | Python (IronPython 2.7) |
| **Framework** | pyRevit 4.8+ |
| **Interface** | XAML / WPF — Design System "Light Theme v2" com paleta, componentes e guidelines padronizados |
| **Excel** | `xlsxwriter` embarcado (sem COM) |
| **Normas** | NBR 5410 — Tabelas de ampacidade, seções comerciais e coordenação disjuntor–cabo |

---

## ⚙️ Instalação

1. Instale o [pyRevit](https://github.com/pyrevitlabs/pyRevit) (v4.8 ou superior).
2. Baixe ou clone a pasta `LF Tools.extension`.
3. Mova para o diretório de extensões do pyRevit:
   ```
   %appdata%\pyRevit\Extensions\
   ```
4. Reinicie o Revit ou execute `Reload` no pyRevit.

A aba **LF Tools** aparecerá automaticamente no Ribbon.

---

## 🤝 Contribuições

Pull requests e issues são bem-vindos!  
Se você tem uma ideia para melhorar o fluxo de trabalho BIM, sinta-se à vontade para colaborar.

---

## 📧 Contato

**Luís Fernando** — [lufe.machado@gmail.com](mailto:lufe.machado@gmail.com)

---

> *Ferramentas desenvolvidas por engenheiros, para engenheiros. Focadas em precisão e alto volume de dados.* 🚀
