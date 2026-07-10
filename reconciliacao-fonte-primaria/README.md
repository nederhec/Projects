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
2. Recalcula **Contábil** buscando a conta direto no `BALANCETE` do mês
   (Débito − Crédito), pelo código de conta — não segue a referência da
   CHECK, então uma fórmula quebrada não contamina o recálculo.
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
reproduzem, minimizados, os dois achados reais que motivaram este projeto:
valor digitado (padrão observado em Janeiro/Fevereiro de um fechamento real)
e fórmula apontando pra célula vazia (padrão observado em Junho). O arquivo
real usado para validar o motor durante o desenvolvimento não é versionado
aqui — só as fixtures sintéticas que replicam o mesmo formato de erro.

## Arquitetura

```
index.html   — upload, dashboard, filtros (por competência / tipo de alerta)
app.js       — glue: lê o XLSX via SheetJS, detecta abas e blocos de mês na
               CHECK, monta o WorkbookAdapter e orquestra o engine
engine.js    — motor puro (sem dependência de biblioteca de planilha):
               parser de fórmula, classificador de proveniência de célula,
               recalculadores, reconciliação 3 vias, score, cobertura
vendor/      — SheetJS (xlsx.full.min.js), local — sem CDN externo
test/        — testes de unidade do engine.js (fixtures sintéticas)
```

`engine.js` recebe um **WorkbookAdapter** (`sheetNames`, `cell(sheet,col,row)`,
`usedRowCount(sheet)`) em vez de depender do SheetJS diretamente — isso é o
que permite testar o motor em Node com fixtures simples, e reusar a mesma
lógica no navegador em cima do SheetJS de verdade.

## Limitações conhecidas (v0.1)

- O recálculo de Contábil soma Débito − Crédito do BALANCETE por código de
  conta; ainda não incorpora ajustes de reclassificação lançados
  manualmente no RAZÃO (alguns fechamentos somam células específicas do
  RAZÃO por cima do BALANCETE — ver fórmula real da conta "13º Salário" no
  arquivo original). Enquanto isso não entra, contas com esse padrão podem
  aparecer com um alerta financeiro que, na prática, é uma reclassificação
  já esperada — vale revisar, não é necessariamente erro.
- Detecção de blocos de mês na CHECK assume o layout 3-colunas
  (FOPAG | CONTÁBIL | DIFERENÇA) com um cabeçalho de rótulos comum a todos
  os blocos — não foi testado contra outros layouts de CHECK.
- Cobertura de contas fantasmas compara pelo prefixo de 3 segmentos do
  código de conta mais comum entre as contas já mapeadas na CHECK; em
  planos de contas com nomenclatura muito diferente isso pode não
  segmentar bem.
