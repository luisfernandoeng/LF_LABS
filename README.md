**LF Tools** é uma suite de engenharia e automação de alta performance desenvolvida para **Autodesk Revit**, focada em maximizar a produtividade e eliminar tarefas repetitivas em fluxos de trabalho BIM.

Desenvolvida sobre o ecossistema **pyRevit** utilizando **IronPython 2.7** e interfaces nativas **WPF (Windows Presentation Foundation)**, esta extensão oferece ferramentas robustas para manipulação de dados, documentação, interoperabilidade e gerenciamento de modelos.

---

## 🔥 Módulos e Funcionalidades

### ✒️ Painel Anotação (Documentation & Text)

Ferramentas para refinamento estético e precisão em anotações técnicas.

- **Ajustes de Texto (Pulldown)**:
    - **Inverter Anotação**: Corrige tags e textos que ficaram invertidos após operações de `Mirror`.
    - **Merge Text**: Consolida notas de texto fragmentadas preservando a ordem espacial (X/Y).
    - **Nome Amb (Linked Rooms)**: Extrai dados de Ambientes de arquivos vinculados e gera tags automáticas.
- **Renumerar**: Algoritmo de renumeração sequencial inteligente baseado na ordem de seleção (Click-order).
- **Renomear+**: Motor de renomeação em massa para Vistas, Folhas e Tabelas com preview em tempo real.

---

### 🤖 Painel Automatizar (Efficiency & Workflow)

Automação de processos complexos e geração de conteúdo.

- **Auto-Cotas**: Dimensionamento automático inteligente. Detecta o eixo central de elementos (conduítes, tubos, paredes) e gera cotas alinhadas em relação a referências (eixos, paredes).
- **Gerar Folhas**: Criação massiva de folhas com alinhamento automático de viewports e configuração de padrões de exportação.
- **Smart Crop**: Ajusta o *Crop Region* das vistas ao limite exato da geometria visível, otimizando o desempenho gráfico.

---

### 📊 Painel Dados e Filtros (Data & Query)

Manipulação granular de informações e auditoria de modelos.

- **Filtrar Avançado**: Query builder com lógica booleana (Igual, Contém, Maior que) para categorias simples ou múltiplas.
- **Inspecionar Tipo**: Diagnóstico de `BuiltInParameters`, IDs de categoria e dados internos de famílias (MEP/Arquitetura).
- **Soma Dist**: Totalizador métrico para elementos lineares (Eletrodutos, Tubos e Linhas).
- **To Excel (High-Performance)**: Sincronização bidirecional ultra-rápida entre Revit e Excel sem dependência de drivers COM.
- **Transfer Settings**: Transferência cirúrgica de Filtros, Modelos de Vista e Padrões entre arquivos abertos ou links.

---

### ⚡ Painel Elétrica (MEP Systems)

Utilitários especializados para engenharia de instalações elétricas.

- **Filtrar Elétrica (Circuit Tracer)**: Rastreamento recursivo de topologia de rede, selecionando todos os dispositivos conectados a um quadro.
- **LF Electrical**: Central de automação de circuitos. Configura quadros (fases/distribuição), cria circuitos agrupados (TUG/TUE) e nomeia interruptores em sequência alfabética.

---

### 🧹 Painel Manutenção (Audit & Optimization)

Higiene e performance do banco de dados do Revit.

- **Overkill (Model Cleanup)**: Identifica e remove elementos duplicados/sobrepostos e realiza o purge seletivo de vistas não utilizadas.

---

## 💻 Tech Stack

- **Core**: Autodesk Revit API (2020-2025)
- **Language**: Python (IronPython 2.7)
- **Framework**: pyRevit v4.8+
- **UI/UX**: XAML / WPF com estilização customizada.

---

## ⚙️ Instalação

1.  Baixe a pasta `LF Tools.extension`.
2.  Mova para a pasta de extensões do pyRevit:
    `%appdata%\pyRevit\Extensions\`
3.  Reinicie o Revit ou use o comando `Reload`.

> _"Ferramentas desenvolvidas por engenheiros, para engenheiros. Focadas em precisão e alto volume de dados."_ :rocket:

### 🤝 Contribuições

Pull requests e issues são bem-vindos! Se você tem uma ideia para melhorar o fluxo de trabalho BIM, sinta-se à vontade para colaborar.

### 📧 Contato

**Luís Fernando** - [lufe.machado@gmail.com]
