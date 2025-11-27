## 🛠️ LF Tools Extension

Opa! Se você trabalha com projetos de **engenharia, arquitetura ou automação residencial**, sabe que tem um monte de tarefa repetitiva que só atrapalha o fluxo. Esta extensão nasceu justamente para isso: **facilitar minha vida e, agora, a sua também!**

Aqui eu junto um conjunto de *plugins* que desenvolvi para **automatizar e integrar processos específicos** no meu dia a dia. Chega de perder tempo com cliques desnecessários!

### ✨ O que essa extensão faz?

* **Automação na veia:** Plugins variados que acabam com as tarefas mais chatas e repetitivas.
* **Integração:** Conexão fácil com outras ferramentas e APIs que eu uso.
* **Fluxo de Trabalho Personalizado:** Você ganha mais liberdade para focar no que realmente importa no seu projeto.

---

### 🚀 Plugins Inclusos (Por enquanto)

Dá uma olhada no que já está rodando por aqui:

#### 1. Filtro Avançado

Ele te ajuda a **filtrar elementos** no projeto sem precisar de seleção prévia.

* Funciona com **múltiplos parâmetros** no mesmo filtro.
* Você filtra elementos com caracteristicas específicas.
* A lógica é igualzinha aos filtros de vista do Revit: você pode filtrar por **"igual a", "contém", "diferente de"**, etc. É só usar a criatividade!

#### 2. Filtrar Elétrica

Um dos que eu mais **amo**!

* Você seleciona o **quadro** primeiro, depois roda o plugin.
* Ele **seleciona todos os circuitos** ligados naquele quadro.
* Eu uso ele para copiar elementos de um pavimento para outro sem perder o circuito.
* *Obs.:* Por enquanto, os interruptores perdem o `Switch ID`, mas **já estou de olho para resolver isso!**

#### 3. Overkill

Esse é fácil: é o **Overkill do CAD**, mas no Revit!

* Você seleciona o que quer "limpar".
* Diz se quer **deletar os duplicados** ou **apenas selecioná-los** para saber onde estão.
* *Atenção:* Por enquanto, tem poucas categorias, mas vou colocando mais conforme a **necessidade aparecer!**

#### 4. Gerar Folhas

**Esse deu trabalho! e vai ser o queridinho de muita gente** É um gerador automático de folhas que salva a pátria na hora de entregar o projeto.

* Faz o **PDF e DWG de várias folhas de uma vez**.
* Ele pega o nome do arquivo a partir de um parâmetro seu (eu uso o `NOME-FOLHA`, que é o padrão da construtora).
* Na hora de salvar, o DWG **já sai sem aquelas vistas anexadas**, gerando um arquivo único e limpo.
* **Configuração é simples:** Você escolhe a pasta de saída, marca as folhas que quer na primeira aba e ajusta as opções de PDF/DWG na segunda.

#### 5. Inspecionar Tipo

Basicamente, um **detetive de elementos**.

* Quer saber **o que cada elemento é**? Quais **parâmetros** ele tem?
* É só selecionar uma tomada, por exemplo, e ele te diz qual o nome, se tem conector elétrico, e todos os parâmetros internos.

#### 6. Inverter Anotação

Sabe quando você usa o `mirror` e aquelas anotações genéricas **insistem em ficar espelhadas/invertidas**?

* Você seleciona as anotações caprichosas e ele **espelha todas de uma vez**, resolvendo o problema rapidinho.

#### 7. Renomear+

Esse é para quem precisa de **edição de texto em massa** nos parâmetros!

* **Exemplo:** Trocar o nome de vários elementos ou re-numerar folhas seguindo um padrão (tipo `UN-01`, `UN-02`, etc.).
* Tem um texto que você tem que substituir em varios elementos, procurar e substituir por aqui.
* *Em progresso:* Estou tentando implementar **expressões regulares (`regex`) para o plugin**, mas ainda sem sucesso. Quem quiser testar, sinta-se à vontade!

#### 8. Renumerar

Mais focado em **numeração sequencial** de elementos.

* Você seleciona os elementos que quer numerar (exemplo: preencher o parâmetro **"marca"**).
* Ele pede onde você quer **começar** (do 1, do 10, do 20) e segue a ordem: `01, 02, 03`, etc.
* **Importante:** A numeração é feita na **ordem em que você clicou/selecionou**.

#### 9. Soma Dist

Simples e direto!

* Precisa saber a **distância total** de um trecho de eletroduto?
* Você seleciona os elementos e ele te retorna **a contagem/distância total**.

#### 10. To Excel

Simples e direto!

* Precisa alterar tabelas ou parametros no revit em massa?
* Você seleciona as tabelas que quer alaterar e manda elas pro excel e depois importa de volta
---

### ⚙️ Como a mágica acontece?

Cada plugin é um arquivo específico que contém os *scripts* e configurações para rodar. 
Eles são carregados e usados via plataforma compatível (se precisar de detalhes de como carregar na sua plataforma, me avisa!).

### 📥 Como Instalar (Para usuários **pyRevit**)

**Pré-requisito:** Você precisa ter o **![pyRevit](https://github.com/pyrevitlabs/pyRevit/releases)** instalado.

1.  **Baixe ou Clone:** Clone o repositório ou baixe o arquivo ZIP da pasta principal `LF Tools.extension`.
2.  **Acesse a pasta de extensões:**
    * Abra o menu **Executar** do Windows (`Win` + `R`).
    * Digite `%appdata%` e pressione **Enter**.
    * Navegue até a pasta `...\pyRevit\Extensions`.
    * *(O caminho completo deve ser algo como: `C:\Users\[SeuUsuario]\AppData\Roaming\pyRevit\Extensions`)*
3.  **Mova a pasta:** Copie a pasta `LF Tools.extension` e cole dentro da pasta `Extensions`.
4.  **Reinicie o Revit:** Feche e abra o Revit (ou a aba pyRevit) para que a extensão seja carregada. Pronto!

### ⌨️ Como Usar

1.  Abra a ferramenta/interface correspondente na sua plataforma.
2.  Carregue o plugin desejado (`Filtro Avançado`, `Gerar Folhas`, etc.).
3.  Configure as opções que ele pedir (se houver).
4.  Execute e veja a mágica acontecer!

### 🤝 Contribuições

Curtiu? Acha que pode melhorar algo? Se quiser contribuir, por favor, **envie um *pull request*** ou **abra uma *issue*** para melhorias e correções. Todo *feedback* é bem-vindo!

### 📧 Contato

Para dúvidas, sugestões ou só para mandar um "e aí", me envie um e-mail: **[lufe.machado@gmail.com]**
