# Reconciliação Independente da Base

Projeto novo e independente de qualquer painel de conciliação existente.
Nasceu da constatação de que ler a aba CHECK como fonte de verdade repete,
sem verificar, qualquer erro de preenchimento que já esteja nela — inclusive
número digitado à mão e fórmula apontando pra célula errada.

## Princípio

A CHECK deixa de ser fonte e vira **resultado conferido**:

1. Recalcula **Folha** somando a `FOPAG 2026` diretamente, por conta —
   reaproveitando a própria composição que a fórmula da CHECK já descreve
   (`SUMIFS(...)`), sem depender do valor que ela cacheia.
2. Recalcula **Contábil** seguindo a composição da própria fórmula da CHECK
   quando ela é uma referência direta e resolve por completo — o que inclui
   ajustes de reclassificação manual do RAZÃO que a fórmula já descreve
   (ex.: `BALANCETE!J160 + RAZÃO!G700 + RAZÃO!G702`), lidos direto da célula
   de origem, não do cache. Quando a fórmula está quebrada ou o valor é
   digitado, cai no fallback: busca a conta direto no BALANCETE por código
   (Débito − Crédito) — não segue a referência da CHECK, então uma
   referência quebrada não contamina o recálculo.
3. Cruza os três números (declarado, Folha recalculada, Contábil
   recalculado) e separa dois tipos de alerta, hoje inexistentes num painel
   que só lê a CHECK:
   - **Financeiro** — Folha diverge de Contábil de verdade, nas fontes primárias.
   - **Preenchimento** — a CHECK diverge da fonte (valor digitado ou fórmula
     quebrada), independente de o número final bater por coincidência.
4. Produz um **Score de Confiabilidade da Base**, separado do risco
   financeiro: responde "o resumo que eu li está certo?", não "o fechamento
   está certo?".
5. Lista **contas fantasmas** — contas presentes no BALANCETE, na mesma
   família contábil que a CHECK já cobre, sem linha correspondente nela.

Ver a especificação completa (achados, arquitetura, casos de teste) no
memorando técnico produzido antes deste projeto.

## Funcionalidades operacionais

- **Drill-down por rubrica**: clicar numa linha da tabela de contas abre a
  composição — quais colunas da FOPAG 2026 formam o valor de Folha, e (se
  houver ajuste do RAZÃO) quais lançamentos específicos compõem o Contábil,
  com valor individual de cada um.
- **Custo por centro de custo**: soma o custo total da FOPAG 2026 agrupado
  por Centro de Custo, direto da fonte (não passa pela CHECK), com filtro
  por competência. Coluna de total e coluna de mês são detectadas pelo
  cabeçalho da planilha, não por posição fixa.
- **Exportação CSV**: exporta a tabela de contas respeitando os filtros
  ativos (competência, tipo de alerta, severidade).
- **Dicionário de contas**: upload opcional de um CSV
  (`código;severidade;responsável`) que enriquece a tabela com severidade e
  responsável sugerido, e habilita filtro por severidade.

## Como rodar

Sem build, sem dependências de runtime. Abra `index.html` num navegador
(Chrome ou Edge) — ou sirva a pasta com qualquer servidor estático:

```bash
npx serve .
# ou
python3 -m http.server 8080
```

Carregue um arquivo único de fechamento (XLSX/XLSM/XLS) contendo, no
mínimo, a aba `CHECK`. `FOPAG 2026` e `BALANCETE MM.AAAA` são detectados
automaticamente pelo nome da aba.

## Testes

```bash
npm test
```

Roda `test/engine.test.js` — testes de unidade com fixtures sintéticas que
reproduzem, minimizados, os achados reais que motivaram este projeto: valor
digitado (padrão observado em Janeiro/Fevereiro de um fechamento real),
fórmula apontando pra célula vazia (padrão observado em Junho) e fórmula
com ajuste do RAZÃO embutido (padrão observado no 13º Salário). O arquivo
real usado para validar o motor durante o desenvolvimento não é versionado
aqui — só as fixtures sintéticas que replicam o mesmo formato de erro.

## Arquitetura

```
index.html   — upload, dashboard, filtros, drill-down, exportação
app.js       — glue: lê o XLSX via SheetJS, detecta abas/colunas por
               cabeçalho (não por posição fixa), monta o WorkbookAdapter,
               orquestra o engine e agrega custo por centro de custo
engine.js    — motor puro (sem dependência de biblioteca de planilha):
               parser de fórmula, classificador de proveniência de célula,
               recalculadores (Folha via FOPAG, Contábil via composição da
               fórmula ou fallback por código no BALANCETE), reconciliação
               3 vias, score de confiabilidade, cobertura de contas
vendor/      — SheetJS (xlsx.full.min.js), local — sem CDN externo
test/        — testes de unidade do engine.js (fixtures sintéticas)
```

`engine.js` recebe um **WorkbookAdapter** (`sheetNames`, `cell(sheet,col,row)`,
`usedRowCount(sheet)`) em vez de depender do SheetJS diretamente — isso é o
que permite testar o motor em Node com fixtures simples, e reusar a mesma
lógica no navegador em cima do SheetJS de verdade.

## Limitações conhecidas (v0.2)

- O recálculo de Contábil só segue a composição da fórmula da CHECK quando
  ela é uma referência direta (`='ABA'!CEL+...`) e resolve por completo. Se
  o Contábil vier de outro tipo de fórmula (ex.: `SUMIFS` também do lado
  Contábil), cai direto no fallback por código de conta — não foi testado
  contra esse padrão.
- Detecção de blocos de mês na CHECK assume o layout 3-colunas
  (FOPAG | CONTÁBIL | DIFERENÇA) com um cabeçalho de rótulos comum a todos
  os blocos — não foi testado contra outros layouts de CHECK.
- Cobertura de contas fantasmas compara pelo prefixo de 3 segmentos do
  código de conta mais comum entre as contas já mapeadas na CHECK; em
  planos de contas com nomenclatura muito diferente isso pode não
  segmentar bem.
- Custo por centro de custo depende de achar, pelo cabeçalho, uma coluna de
  mês, uma de "Centro de Custo" e uma de total na mesma aba da FOPAG — se a
  planilha nomear essas colunas de forma muito diferente, a seção
  simplesmente não aparece (falha silenciosa, não trava o resto do painel).
- Dicionário de contas é um CSV simples (sem suporte a aspas/escapes
  complexos) — funciona bem pra um `código;severidade;responsável` direto,
  não é um parser CSV completo.
