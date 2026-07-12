# Reconciliação Independente da Base

Projeto novo e independente de qualquer painel de conciliação existente.
Nasceu da constatação de que ler a aba CHECK como fonte de verdade repete,
sem verificar, qualquer erro de preenchimento que já esteja nela — inclusive
número digitado à mão e fórmula apontando pra célula errada.

## Segurança — histórico

O leitor de planilha era originalmente SheetJS (`xlsx`) 0.18.5, a mesma
versão do painel original — última publicada no npm (mar/2022), com duas
vulnerabilidades conhecidas de severidade alta sem correção disponível por
canal alcançável deste ambiente (a SheetJS passou a distribuir correções só
pelo CDN próprio, bloqueado pela política de rede daqui; um espelho
comunitário no npm foi encontrado mas descartado por ser republicação de
terceiro não verificada, não a fonte oficial).

**Resolvido**: o projeto trocou de biblioteca — agora usa
[ExcelJS](https://github.com/exceljs/exceljs) (`vendor/exceljs.min.js`),
ativamente mantida, sem as vulnerabilidades acima, publicada oficialmente no
npm. `engine.js` não foi afetado (já era desacoplado de qualquer biblioteca
de planilha via `WorkbookAdapter`); só `app.js` (monta o adapter sobre o
ExcelJS em vez do SheetJS) e o `<script>` do `index.html` mudaram. Migração
validada: os 25 testes automatizados passam, e o dashboard no navegador
produz exatamente os mesmos números (financeiro/preenchimento/sem-fonte/
fantasmas) antes e depois da troca, tanto no arquivo real do cliente quanto
na fixture sintética.

**Achado residual, verificado como não-explorável no nosso uso**: o
`npm audit` acusa uma vulnerabilidade moderada numa dependência transitiva
do ExcelJS (`uuid`, [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq)
— falta de checagem de limites do buffer nas funções `v3`/`v5`/`v6` quando
um buffer é passado manualmente). Conferido no código-fonte do ExcelJS: ele
só chama `uuid.v4()`, sem argumentos — as funções e o padrão de uso que a
vulnerabilidade exige não são exercitados por este projeto. Um
`npm audit fix --force` "resolveria" isso fazendo downgrade do ExcelJS pra
3.4.0 (versão antiga, não mantida) — pior, não melhor. Não aplicado.

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
   família e profundidade contábil (nº de segmentos do código) que a CHECK
   já cobre, sem linha correspondente nela. A profundidade importa: um
   código-pai/subtotal do plano de contas (ex.: `3.1.01.003`, que soma as
   próprias contas-filha `3.1.01.003.003` + `3.1.01.003.007` + ...) não é
   uma conta faltando, é o mesmo dinheiro em outra granularidade — sem
   filtrar por profundidade, ele aparecia na lista como se fosse.

Ver a especificação completa (achados, arquitetura, casos de teste) no
memorando técnico produzido antes deste projeto.

## Funcionalidades operacionais

- **Drill-down por rubrica**: clicar numa linha da tabela de contas abre a
  composição — quais colunas da FOPAG 2026 formam o valor de Folha, e (se
  houver ajuste do RAZÃO) quais lançamentos específicos compõem o Contábil,
  com valor individual de cada um.
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

## Hospedagem — Cloudflare Pages + Basic Auth

Pra rodar além da máquina local (segundo parecer de outra pessoa da equipe,
por exemplo), o projeto suporta deploy no Cloudflare Pages com Basic Auth
obrigatório na frente — nada é servido sem autenticação, nem os arquivos
estáticos (`app.js`, `engine.js`, `vendor/`).

**Como funciona**: `functions/_middleware.js` roda como Cloudflare Pages
Function antes de qualquer requisição. Sem `Authorization: Basic` válido,
responde `401` (`WWW-Authenticate`) e não chega a servir nada. Se as
variáveis de ambiente não estiverem configuradas, responde `500` em vez de
liberar o acesso — falha fechada, não aberta.

**Configurar (uma vez)**:
1. No Cloudflare Pages, criar um projeto apontando pra este repositório.
   Como o repo tem vários projetos independentes, definir **Root directory**
   = `reconciliacao-fonte-primaria`.
2. **Build command**: vazio (nenhum build). **Build output directory**: `/`.
3. Em *Settings → Environment variables*, adicionar `BASIC_AUTH_USER` e
   `BASIC_AUTH_PASS` como **Secret** (não como texto plano, e nunca no
   repositório). Repetir pros ambientes de Production e Preview.
4. Todo push no branch conectado dispara deploy automático.

**Testar localmente** (opcional, requer `npx wrangler`):

```bash
cd reconciliacao-fonte-primaria
npx wrangler pages dev . --port 8788 -b BASIC_AUTH_USER=teste -b BASIC_AUTH_PASS=segredo123
```

Importante rodar com `.` como diretório *e* de dentro da pasta do projeto —
`wrangler pages dev reconciliacao-fonte-primaria` (a partir da raiz do
repo) não detecta `functions/` e serve tudo sem autenticação nenhuma
(confirmado testando as duas formas).

Trocar o usuário/senha depois é só editar as variáveis no painel — não
precisa de novo deploy nem alterar código.

## Testes

Duas camadas, nenhuma delas depende do arquivo real do cliente (não
versionado aqui):

```bash
npm test              # unidade — zero dependências, roda em qualquer lugar
npm run test:integration   # integração — precisa de `npm install` antes (devDependency exceljs)
npm run test:all      # as duas
```

CI: `.github/workflows/reconciliacao-fonte-primaria-tests.yml` roda
`npm run test:all` automaticamente em todo PR e push no `main` que toque
este diretório — nenhuma mudança chega no `main` sem passar pelas 25
verificações.

**Unidade** (`test/engine.test.js`) — fixtures em memória, escritas à mão,
testando cada função do `engine.js` isoladamente: valor digitado (padrão
Jan/Fev do arquivo real), fórmula apontando pra célula vazia (padrão
Junho), fórmula com ajuste do RAZÃO embutido (padrão 13º Salário),
hierarquia pai/filho na cobertura de contas.

**Integração** (`test/integration.test.js`) — carrega um `.xlsx` de verdade
via ExcelJS (mesma biblioteca que o `app.js` usa no navegador):
`test/fixtures/fechamento-ficticio-teste.xlsx`, dado 100%
fictício (gerado por `test/fixtures/build-synthetic-fixture.js`,
`npm run fixtures:build` regenera). Não é uma cópia do arquivo real — usa
uma família de conta diferente (`4.2.01.*`), cabeçalho de aba fora da linha
1, nome de aba BALANCETE com separador alternativo (hífen em vez de ponto),
e uma divergência financeira real plantada de propósito — pra provar que a
lógica generaliza em vez de estar decorada pra um arquivo só. Essa segunda
camada já achou um bug real: o rótulo de rubrica no drill-down assumia
cabeçalho na linha 1 da FOPAG e mostrava letra de coluna ("G") em vez do
nome da rubrica ("ENCARGOS") sempre que o cabeçalho estava em outra linha —
incluindo no arquivo real do cliente, que tem cabeçalho na linha 3.
Corrigido: `findColumnLabel` agora procura a primeira célula de texto nas
primeiras linhas em vez de assumir posição fixa.

A validação de detecção de abas/blocos/colunas do `app.js` (que é
DOM-coupled, não roda em Node) foi feita manualmente num navegador real via
Playwright contra essa mesma fixture — confirmado: dashboard renderiza,
KPIs batem com o esperado, cobertura de contas exclui a conta-pai
corretamente.

## Arquitetura

```
index.html   — upload, dashboard, filtros, drill-down, exportação
app.js       — glue: lê o XLSX via ExcelJS, detecta abas/colunas por
               cabeçalho (não por posição fixa), monta o WorkbookAdapter
               e orquestra o engine
engine.js    — motor puro (sem dependência de biblioteca de planilha):
               parser de fórmula, classificador de proveniência de célula,
               recalculadores (Folha via FOPAG, Contábil via composição da
               fórmula ou fallback por código no BALANCETE), reconciliação
               3 vias, score de confiabilidade, cobertura de contas
vendor/      — ExcelJS (exceljs.min.js), local — sem CDN externo
test/        — unidade (fixtures em memória) + integração (fixture .xlsx
               fictícia real, ver seção Testes)
```

`engine.js` recebe um **WorkbookAdapter** (`sheetNames`, `cell(sheet,col,row)`,
`usedRowCount(sheet)`) em vez de depender de biblioteca de planilha
nenhuma diretamente — isso é o que permite testar o motor em Node com
fixtures simples, reusar a mesma lógica no navegador em cima do ExcelJS de
verdade, e foi o que tornou a troca de SheetJS pra ExcelJS um problema
isolado em `app.js`, sem tocar `engine.js`.

## Investigado e descartado — Gerencial × Contábil, Rescisão, Provisões

Três eixos adicionais foram avaliados e **não entraram**, com evidência do
motivo — documentado aqui pra não serem retentados sem essa base:

- **Gerencial × Contábil** (`Base fechamento` vs. total Contábil): o total
  mensal da `Base fechamento` roda ~2× o total da família contábil de
  pessoal (ex.: Abril: R$3,17M na Base fechamento contra ~R$1,57–2,3M na
  família `3.1.01.*` do BALANCETE, dependendo de como essa família é
  somada — ver próximo item). O gap não tem explicação recuperável só do
  arquivo; pode ser metodologia de custo pleno/rateio que a Base fechamento
  aplica e a contabilidade não. Comparar os dois como se devessem bater
  seria fabricar uma divergência (ou uma "conciliação") sem base.
- **Custo total (FOPAG) × Contábil (família completa do BALANCETE)**: esta
  era a alternativa mais defensável — comparar a coluna de custo total da
  FOPAG 2026 contra a soma de *todas* as contas `3.1.01.*` do BALANCETE, não
  só as que a CHECK lista. Uma primeira tentativa pareceu bater quase
  perfeitamente (diferença de R$939 em ~R$1,57M) — mas era coincidência de
  dois erros se cancelando: código de teste ignorando células do BALANCETE
  armazenadas como texto, e a soma por prefixo contando conta-pai e
  conta-filha ao mesmo tempo (dobrando parte do valor). Corrigidos os dois
  problemas, a soma correta da família completa (só contas-folha, texto
  parseado) fica em ~R$2,31M contra ~R$1,57M da FOPAG — uma diferença de
  ~47% sem explicação disponível no arquivo. Não foi implementado.
- **Cross-check de Rescisão** (`Rescisoes` vs. colunas de rescisão da FOPAG):
  o arquivo tem só 1 evento de rescisão em Abril, insuficiente pra validar
  um de-para de rubricas. E o de-para não é direto: `Custo Total` da aba
  `Rescisoes` inclui encargos do empregador que as colunas de rescisão da
  FOPAG (pagas ao empregado) não cobrem — só uma das quatro colunas testadas
  bateu exatamente (Salário Bruto = Saldo Salário Rescisão).
- **Cross-check de Provisões** (`CONSULTA DE PROVISÕES DE FÉRIAS` /
  `CONSULTA PROVISÃO DE 13º SALÁRIO` vs. FOPAG): essas duas abas são
  **saldo acumulado** por referência mensal (quanto já foi provisionado até
  aquele mês), enquanto as colunas de provisão da FOPAG 2026 são o
  **movimento daquele mês** (ex.: 1/12 do período). São conceitos
  diferentes — comparar diretamente dá uma diferença de ~5× que não
  significa nada (confirmado com números reais: Abril, 13º salário,
  R$515 mil de saldo acumulado contra R$96 mil de movimento do mês).

Pra qualquer um desses virar uma verificação de verdade, precisa da
definição de negócio de quem monta a `Base fechamento`/`Rescisoes` — o que
exatamente cada total representa e como ele deveria reconciliar (ou não)
com a contabilidade. Sem isso, apresentar um número "batendo" ou "não
batendo" seria exatamente o tipo de falsa confiança que este projeto existe
pra eliminar.

## Limitações conhecidas (v0.2)

- O recálculo de Contábil só segue a composição da fórmula da CHECK quando
  ela é uma referência direta (`='ABA'!CEL+...`) e resolve por completo. Se
  o Contábil vier de outro tipo de fórmula (ex.: `SUMIFS` também do lado
  Contábil), cai direto no fallback por código de conta — não foi testado
  contra esse padrão.
- Detecção de blocos de mês na CHECK assume o layout 3-colunas
  (FOPAG | CONTÁBIL | DIFERENÇA) com um cabeçalho de rótulos comum a todos
  os blocos — validado contra dois arquivos com esse layout (real e
  fixture sintética), mas não contra um layout genuinamente diferente
  (ex.: 2 ou 4 colunas por bloco).
- Cobertura de contas fantasmas compara pelo prefixo de 3 segmentos e pela
  profundidade (nº de segmentos) mais comuns entre as contas já mapeadas na
  CHECK — validado contra duas famílias de conta diferentes (`3.1.01.*` no
  arquivo real, `4.2.01.*` na fixture), mas em planos de contas com
  nomenclatura muito diferente (nº de segmentos variável entre contas
  legítimas, por exemplo) isso pode não segmentar bem.
- Dicionário de contas é um CSV simples (sem suporte a aspas/escapes
  complexos) — funciona bem pra um `código;severidade;responsável` direto,
  não é um parser CSV completo.
