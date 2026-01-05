# LF Tools - Revit Automation Suite 🚀

**LF Tools** é uma suite de engenharia e automação de alta performance desenvolvida para **Autodesk Revit**, focada em maximizar a produtividade e eliminar tarefas repetitivas em fluxos de trabalho BIM.

Desenvolvida sobre o ecossistema **pyRevit** utilizando **IronPython 2.7** e interfaces nativas **WPF (Windows Presentation Foundation)**, esta extensão oferece ferramentas robustas para manipulação de dados, documentação interoperabilidade e gerenciamento de modelos.

---

## 🔥 Módulos e Funcionalidades

### 🛠️ Painel Modificações (Data & Batch Processing)

Ferramentas para manipulação massiva de dados e parâmetros.

#### **1. Renumerar (Smart Renumbering)**
Algoritmo de renumeração sequencial inteligente.
- **Fluxo Híbrido**: Permite seleção contínua de elementos no modelo sem fechar a interface.
- **Ordenação Dinâmica**: Numera baseando-se na ordem de seleção do usuário (click-order).
- **Customização**: Suporte a prefixos, sufixos e *padding* (zeros à esquerda).

#### **2. Renomear+ (Advanced Batch Renamer)**
Motor de renomeação em massa com suporte a regras complexas.
- **Find & Replace**: Substituição de strings em parâmetros de Vistas, Folhas e Tabelas.
- **Numeração Sequencial**: Re-indexação de folhas e vistas.
- **Preview Real-time**: Visualização das alterações antes da aplicação no banco de dados do Revit.

#### **3. To Excel (High-Performance IO)**
Sincronização bidirecional de dados entre Revit e Excel sem dependência de drivers COM.
- **Performance O(1)**: Otimizado para grandes volumes de dados usando bibliotecas nativas (`xlsxwriter`/`xlrd`).
- **Relatórios de Integridade**: Feedback detalhado sobre células modificadas, ignoradas (imutáveis) ou erros de tipo.
- **Aplicações**: Edição em massa de Tabelas de Quantidades e Quadros de Cargas.

#### **4. Gerar Folhas (Sheet Automation)**
Automação de documentação técnica.
- **Batch Creation**: Geração automática de múltiplas folhas baseada em vistas selecionadas.
- **DWG Auto-Setup**: Configuração automática de padrões de exportação (AIA Layers, True Colors) se inexistentes no projeto.
- **Alinhamento Inteligente**: Centralização automática de Viewports no Title Block.

#### **5. Inspecionar Tipo (Type Inspector)**
Ferramenta de diagnóstico rápido de elementos.
- **Introspecção**: Revela parâmetros ocultos (BuiltInParameters), IDs de categoria e dados de conectores MEP.
- **Debug Tool**: Essencial para coordenadores BIM identificarem inconsistências em famílias.

#### **6. Inverter Anotação (Mirror Fix)**
Correção automática da orientação de anotações.
- **Algoritmo**: Detecta e corrige anotações de texto e tags que ficaram invertidas/espelhadas após operações de `Mirror` no modelo.

#### **7. Merge Text (Text Consolidation)**
Consolidação de notas de texto fragmentadas.
- **Algoritmo Espacial**: Unifica múltiplas notas de texto selecionadas em uma única entidade mestre.
- **Ordenação Y/X**: Preserva a ordem de leitura baseada nas coordenadas espaciais dos elementos originais.

#### **8. Nome Amb (Linked Room Tagging)**
Anotação automatizada baseada em vínculos (Revit Links).
- **Data Extraction**: Lê dados de Ambientes (Rooms) de arquivos vinculados (impossível com tags nativas de anotação genérica).
- **Collision Avoidance**: Algoritmo que evita sobreposição de textos em plantas densas.
- **Multi-Parâmetros**: Extrai Nome, Área e Pé Direito (Unbounded Height).

#### **9. Soma Dist (Route Totalizer)**
Totalizador métrico para elementos lineares.
- **Cálculo de Rede**: Soma o comprimento total de Eletrodutos, Tubos ou Linhas selecionadas.
- **Aplicações**: Estimativa rápida de cabeamento e tubulação.

---

### 🔍 Painel Filtrar e Limpar (Audit & Optimization)

Ferramentas para auditoria, limpeza e seleção precisa de elementos.

#### **10. Filtro Avançado (Query Builder)**
Seleção baseada em regras lógicas, similar aos Filtros de Vista, mas para seleção ativa.
- **Lógica Booleana**: Suporte a operadores (Igual, Contém, Diferente, Maior que).
- **Multi-Categoria**: Permite filtrar elementos de categorias distintas simultaneamente.

#### **11. Overkill (Model Cleanup)**
Ferramenta de saneamento do modelo.
- **Deduplicação**: Identifica e remove elementos geometricamente idênticos sobrepostos (clash zero).
- **Limpeza de Vistas**: Purge seletivo de vistas e folhas não utilizadas.

#### **12. Smart Crop (Viewport Optimization)**
Ajuste algorítmico de Viewports.
- **Bounding Box Analysis**: Redefine o Crop Region da vista para o limite exato da geometria visível.
- **Benefício**: Reduz o processamento gráfico da vista e otimiza o espaço em prancha.

---

### ⚡ Painel Elétrica (MEP Systems)

Utilitários específicos para projetos de instalações elétricas.

#### **13. Filtrar Elétrica (Circuit Tracer)**
Rastreamento inteligente de sistemas elétricos.
- **Topologia de Rede**: Seleciona um Painel e identifica recursivamente todos os dispositivos e circuitos conectados a ele.
- **Copy/Monitor Aux**: Facilita a cópia de pavimentos inteiros garantindo que a integridade do circuito seja mantida na seleção.

---

## 💻 Tech Stack

- **Core**: Autodesk Revit API 2024
- **Language**: Python (IronPython 2.7)
- **Framework**: pyRevit v4.8+
- **UI/UX**: WPF (Xaml) com Estilização "Dark Mode" Customizada via ResourceDictionaries.

---

## ⚙️ Instalação

1.  Baixe a pasta `LF Tools.extension`.
2.  Mova para a pasta de extensões do pyRevit:
    `%appdata%\pyRevit\Extensions\`
3.  Reinicie o Revit.

> _"Ferramentas desenvolvidas por engenheiros, para engenheiros. Focadas em alto volume de dados e precisão."_ :rocket:

### 🤝 Contribuições

Curtiu? Acha que pode melhorar algo? Se quiser contribuir, por favor, **envie um *pull request*** ou **abra uma *issue*** para melhorias e correções. Todo *feedback* é bem-vindo!

### 📧 Contato

Para dúvidas, sugestões ou só para mandar um "e aí", me envie um e-mail: **[lufe.machado@gmail.com]**
