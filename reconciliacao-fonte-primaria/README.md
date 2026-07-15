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
- **Colunas da tabela**: Competência, Conta, Descrição, Valor Folha, Valor
  Contábil, Diferença Recalculada, Alertas, Tipo. Os valores de Folha e
  Contábil são os recalculados direto da fonte primária (FOPAG/BALANCETE),
  não o que a CHECK declara. "Tipo" classifica a causa: `Diferença de
  valor`, `Falta lançamento contábil` (Folha tem valor, Contábil não),
  `Falta lançamento na Folha` (o inverso), `Sem divergência` ou `Não
  verificável`. O filtro de alertas tem uma opção "Todas as Divergências" —
  mostra tudo, exceto contas com alerta "ok" **e** tipo "Sem divergência" ao
  mesmo tempo (ou seja, mantém qualquer conta com algum alerta, mesmo que o
  tipo seja "Sem divergência").
- **Resumo financeiro**: 5 cards interativos acima da tabela — Total da
  Folha (RH), Total Contábil, Divergências Folha, Divergências Contábil e
  Impacto Total (soma das diferenças financeiras confirmadas). Tem seletor
  de competência próprio — os 5 números recalculam pro mês escolhido (ou
  todas, combinado). Clicar num card filtra a tabela abaixo pelas contas
  correspondentes (competência incluída) e rola até ela.
- **Exportação CSV**: exporta a tabela de contas respeitando os filtros
  ativos (competência, tipo de alerta), com as mesmas colunas de valor/tipo
  da tabela, mais a nota sobre a origem do Contábil e o selo de previsão —
  pra a planilha ficar utilizável sozinha, sem precisar abrir a ferramenta
  de novo.
- **Marca de previsão**: competências cujo BALANCETE tem "PREV" no nome
  (mês ainda não fechado, ex. "PREV. Balancete 052026") ganham um selo
  "previsão" no resumo, no filtro e em cada linha da tabela — sem isso, um
  alerta financeiro em cima de dado projetado parece divergência de livro
  fechado, quando não é.
- **Rascunho de e-mail**: botão "Gerar e-mail" na composição de cada conta
  monta um rascunho direcionado a quem vai ajustar a divergência — formato
  "CONTA/RUBRICA → VALORES → DIAGNÓSTICO", adaptado do painel de
  conciliação antigo (`painel-conciliacao-folha`) para
  o modelo de dados deste projeto. O DIAGNÓSTICO tem o Tipo (mesma
  classificação da tabela) e um parágrafo em linguagem direta apontando qual
  lado está errado/faltando e por quanto (ex.: "Valor de R$ X lançado em
  Folha (RH), porém AUSENTE na contabilidade"), mais uma linha "Evidência no
  razão" com a contagem de lançamentos e os totais de débito/crédito da
  conta no razão auxiliar daquele mês. Essa evidência é uma varredura
  **independente** do razão — `engine.lookupContaNoRazao` procura a conta
  direto na aba de RAZÃO (formato de blocos "Conta: N - X.X.XX.XXX.XXX ..."
  até "Total conta:"), sem depender em nada do que a fórmula da CHECK
  referencia. Nada é enviado automaticamente — é só um rascunho pra revisão
  manual, com botões de copiar e abrir no cliente de e-mail (`mailto:`).
- **Painel de Custos e Conciliação**: segunda aba do dashboard (`Painel de
  Custos`), implementada em `painel-custos.js`. Fica **bloqueada** (botão
  desabilitado, com cadeado) só até um arquivo ser carregado e lido — assim
  que o dashboard aparece (`habilitarPainelCustos()` em `app.js`, chamada no
  fim de `renderDashboard()`), a aba já fica disponível, **mesmo que a
  competência mais recente ainda tenha divergências em aberto**: nem sempre
  o mês vai estar 100% conferido, e o painel serve justamente pra
  acompanhar custo e cobertura da conciliação enquanto isso, não só depois.
  - **Atualização em tempo real conforme divergências são sanadas**: como o
    app não tem edição manual de conta — o único jeito de sanar uma
    divergência é corrigir a fonte primária (FOPAG/BALANCETE) e recarregar o
    arquivo —, `renderDashboard()` chama `renderPainelCustosSeVisivel()` toda
    vez que um arquivo é processado. Se o Painel de Custos já estiver aberto
    nesse momento, ele se reconstrói sozinho com os números atualizados, sem
    precisar trocar de aba pra "forçar" o refresh (mesmo helper usado pelo
    toggle de tema, que também precisa redesenhar os gráficos ao vivo).
  - **KPIs** (cards no topo, competência mais recente): Custo Contábil,
    Variação do Custo (mês a mês, R$ e %), % Conciliado, Custo de Pessoal
    (Salário + Encargos + Benefícios), Diferença Líquida, Diferença
    Absoluta e Divergências no Upload (com pill de quantas ainda estão
    Pendentes). Tudo calculado em cima dos mesmos `state.resultados` que a
    tabela de Conciliação usa — nada é recalculado aqui, só agregado por
    competência.
  - **Gráficos** (Chart.js, `vendor/chart.umd.min.js`, vendorizado
    localmente via npm — CDNs como jsDelivr/unpkg são bloqueados pela
    política de rede deste ambiente): evolução do Custo Contábil vs. Custo
    de Pessoal, composição do Custo de Pessoal por categoria (empilhado),
    % Conciliado ao longo do tempo, variação mensal do custo contábil
    (barras verde/vermelho pelo sinal), Diferença Líquida por competência,
    exposição (Diferença Absoluta), divergências identificadas vs.
    pendentes, e Rescisões por competência. Diferente do antigo relatório
    em PDF, os gráficos desenham direto num `<canvas>` visível na página —
    sem exportação, sem canvas fora da tela, sem a bagunça de embutir PNG/JPEG
    num arquivo (essa aba não gera PDF).
  - **Custo de Pessoal, Encargos, Benefícios**: não há nenhuma aba no arquivo
    real do cliente com esse total já pronto (conferido: nem o BALANCETE, nem
    RESUMO FECHAMENTO FOPAG, nem a aba ENCARGOS trazem essa quebra) — por
    isso a classificação é uma **heurística por palavra-chave** na descrição
    da conta (Encargos: INSS/FGTS/PIS/GRRF; Benefícios: vale/assistência
    médica/seguro de vida/plano de saúde/auxílio; demais contas: Salário).
    O critério fica descrito no próprio painel pra quem for usar os números
    conferir antes.
  - **Rescisões**: conta pessoas distintas (por NOME) na aba "Rescisoes" do
    arquivo, agrupadas pelo mês da coluna MÊS — mostra "—" (não 0) quando essa
    aba não existe no arquivo carregado, pra não parecer "zero rescisões"
    quando na verdade é "sem como verificar".
  - **% Conciliado**: percentual de contas da competência com Folha **e**
    Contábil confirmados direto na fonte primária — mede cobertura da
    conciliação, não confiabilidade da CHECK (esse é o Score de
    Confiabilidade, KPI diferente já existente no dashboard).
  - **Divergências no Upload vs. Pendentes**: a contagem de divergências no
    momento em que uma competência é vista pela primeira vez neste
    navegador é gravada como "baseline" e nunca mais sobrescrita
    automaticamente — ajustes feitos depois só reduzem o indicador de
    Pendentes, não o de "no Upload". Isso deixa visível quanto trabalho de
    conciliação já foi feito desde a entrega original do arquivo. O
    armazenamento é `localStorage`, chave
    `reconciliacao-fonte-primaria:baseline-divergencias:{AAAA-MM}`. **Limite
    conhecido**: a chave é só pela competência (mês/ano), não pelo cliente
    ou arquivo — trocar de arquivo de um cliente pra outro sem resetar pode
    misturar baselines de competências com o mesmo rótulo. O botão "Zerar
    baseline de divergências" no próprio painel limpa todas as chaves desse
    prefixo, pra usar ao trocar de cliente.

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

## Hospedagem — Cloudflare Workers + Basic Auth

Pra rodar além da máquina local (segundo parecer de outra pessoa da equipe,
por exemplo), o projeto suporta deploy como Cloudflare Worker, com suporte
opcional a Basic Auth na frente.

**Estado atual do deploy em produção**: o Basic Auth está **desativado**
(`assets.run_worker_first: false` em `wrangler.jsonc`, por decisão
explícita do usuário) — a URL pública serve os arquivos estáticos direto,
sem exigir login. Como o processamento do XLSX é 100% client-side (nada é
enviado nem armazenado no servidor), o risco é alguém com a URL acessar a
ferramenta, não vazamento de dados de terceiros. A URL de produção tem um
prefixo aleatório (não é indexada nem previsível), mas não deixa de ser
pública. Para reativar a autenticação, ver "Reativar Basic Auth" abaixo.

Nota: o painel da Cloudflare vem migrando de "Pages" pra "Workers" como
modelo padrão pra sites estáticos com Git integration — por isso o projeto
usa `worker.js` + `wrangler.jsonc` (modelo Workers), não mais
`functions/_middleware.js` (modelo Pages, descontinuado aqui).

**Como funciona (quando ativado)**: `worker.js` roda antes de qualquer
asset estático ser servido — mas só se `wrangler.jsonc` tiver
`assets.run_worker_first: true`. **Isso é obrigatório pra ativar**: o
padrão do Workers é servir os arquivos estáticos direto, sem rodar o
script (o oposto do antigo Pages Functions). Sem essa opção, o Basic Auth
simplesmente não roda pra nenhuma requisição que bata num arquivo
estático — vira decoração, nada fica protegido de verdade. Sem
`Authorization: Basic` válido, responde `401` (`WWW-Authenticate`). Se as
variáveis de ambiente não estiverem configuradas, responde `500` em vez de
liberar o acesso — falha fechada.

**Reativar Basic Auth**:
1. No painel da Cloudflare, **Compute (Workers)** → **Create application**
   → **Connect to Git** (não a opção antiga "Pages", que pode nem aparecer
   mais dependendo da conta) → selecionar o repositório.
2. Nas configurações de build, se houver campo **Root directory** (ou
   **Path**, dependendo da versão do wizard), definir
   `reconciliacao-fonte-primaria` (o repo tem vários projetos
   independentes). O nome do Worker no painel **precisa bater** com o
   campo `"name"` de `wrangler.jsonc` (`apura-folha-app`), senão o build
   falha.
3. Configurar `BASIC_AUTH_USER` e `BASIC_AUTH_PASS` como **Secret** do
   Worker (runtime, nunca no repositório) — no painel usado neste projeto
   isso não ficou em "Bindings" nem em "Build → Variables and secrets"
   (esse último é só pra build); procure a seção de variáveis de runtime em
   Settings.
4. Editar `wrangler.jsonc` e voltar `assets.run_worker_first` para `true`,
   commitar e dar push (dispara redeploy automático via
   `npx wrangler deploy`, já configurado pelo próprio painel).

**Testar localmente** (opcional, requer `npx wrangler`):

```bash
cd reconciliacao-fonte-primaria
npx wrangler dev --var BASIC_AUTH_USER:teste --var BASIC_AUTH_PASS:segredo123
```

Achado ao testar: com `assets.directory` apontando pra raiz do projeto
(onde o próprio `wrangler dev` também guarda seu estado local em
`.wrangler/`), o servidor entra num loop de reload — a escrita do próprio
estado é detectada como mudança de asset. Não afeta o deploy de verdade
(que usa `wrangler deploy`, não `dev`), só o teste local; contornado com
`--persist-to <pasta fora do projeto>` ao testar.

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
index.html       — upload, dashboard, abas (Conciliação / Painel de
                    Custos), filtros, drill-down, exportação
app.js            — glue: lê o XLSX via ExcelJS, detecta abas/colunas por
                    cabeçalho (não por posição fixa), monta o
                    WorkbookAdapter, orquestra o engine e controla o
                    troca de abas / desbloqueio do Painel de Custos
engine.js         — motor puro (sem dependência de biblioteca de planilha):
                    parser de fórmula, classificador de proveniência de
                    célula, recalculadores (Folha via FOPAG, Contábil via
                    composição da fórmula ou fallback por código no
                    BALANCETE), reconciliação 3 vias, score de
                    confiabilidade, cobertura de contas
painel-custos.js  — aba de KPIs e gráficos por competência, calculados em
                    cima de `state.resultados` (nada recalculado); mantém
                    o baseline de divergências no localStorage
vendor/           — ExcelJS (exceljs.min.js) e Chart.js (chart.umd.min.js),
                    locais — sem CDN externo
test/             — unidade (fixtures em memória) + integração (fixture
                    .xlsx fictícia real, ver seção Testes)
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

- **Financeiro/tesouraria** (o que efetivamente saiu do banco): investigado
  numa cópia mais completa do arquivo do mesmo cliente. A única coluna que
  parecia ser isso, `PAGAMENTO LIQUIDO` na aba `RESUMO FECHAMENTO FOPAG`, é
  uma referência a um workbook externo desvinculado (`#VALUE!` em toda
  célula) — e mesmo se resolvesse, somaria a coluna errada (`FOPAG
  2026!$AE`, que é "INSS Acerto", não líquido de pagamento). Varrida a aba
  FOPAG 2026 inteira (101 colunas) e não existe nenhuma coluna de valor
  líquido/pago. Não há fonte de Financeiro neste arquivo — o motor
  continua reconciliando só RH/DP × Contábil, não as 3 pernas.

## Achados reais de um arquivo mais completo do mesmo cliente

Uma segunda cópia do arquivo do cliente (mesma empresa, mas com todos os 12
meses no CHECK, 101 colunas na FOPAG 2026 e 6 pares BALANCETE/RAZÃO em vez
dos 4 originais) revelou 3 problemas reais no motor — nenhum causava dado
errado no arquivo original só por coincidência de estrutura, mas todos
causariam silenciosamente:

- **Coluna do BALANCETE fixa (`lookupContaNoBalancete`)**: o fallback por
  código de conta assumia código sempre na coluna B, Débito sempre H,
  Crédito sempre I. No arquivo maior, Janeiro/Fevereiro têm o código na
  coluna A (sem coluna "Código" sequencial) e Débito/Crédito em E/F; Março
  em diante ganham a coluna "Código" e empurram tudo uma coluna à direita.
  Corrigido: `findBalanceteColunas` acha as três pelo texto do cabeçalho
  ("Classificação"/"Débito"/"Crédito"), com o valor fixo antigo como
  fallback só se a varredura não achar nada.
- **Nomenclatura de aba BALANCETE sem separador**: `findBalanceteMap`
  exigia um separador (`.`/`-`/`/`) entre mês e ano — "BALANCETE 04.2026",
  "Balancete 03-2026". O arquivo maior usa "Balancete 012026" (mês e ano
  colados, sem separador nenhum), que não batia com o regex — resultado:
  **zero balancetes detectados**, e toda verificação de Contábil caía
  silenciosamente em "sem fonte". Corrigido: separador agora é opcional
  (`\d{2}` e `\d{4}` são largura fixa, então isso não introduz ambiguidade
  entre os três formatos).
- **SUMIFS com referência de coluna inteira**: o parser de fórmula exigia
  número de linha na faixa (`$E$3:$E$100`). O arquivo maior usa referência
  de coluna inteira (`'Fopag 2026'!$E:$E`), que não batia com o regex —
  resultado: toda fórmula de FOPAG desse arquivo caía no ramo genérico
  "intra" do parser (por não ser reconhecida como SUMIFS) e ficava
  "sem-verificação", apesar de ter SUMIFS de verdade. Esse foi o mais
  amplo dos três: sozinho já derrubava a verificação de Folha do arquivo
  inteiro (96 de 96 contas×competência foram para "sem fonte").
  Corrigido: número de linha na faixa agora é opcional.

Antes dos três: "Sem fonte para verificar" = 96 (praticamente tudo). Depois:
0 — e 52 divergências financeiras reais passaram a aparecer (antes
escondidas atrás do "sem fonte"). Confirmado manualmente que os valores
recalculados batem com o que a própria CHECK declara nos casos onde a
CHECK está correta (ex.: SALARIOS E ORDENADOS de Janeiro, R$2.294,03 nos
dois lados). Os três têm teste de regressão automatizado (`engine.test.js`
e `integration.test.js`, com uma fixture com layout de BALANCETE
deliberadamente deslocado) — verificado que os testes falham sem a
correção e passam com ela, não só que passam.

**Verificação independente do razão auxiliar (`lookupContaNoRazao`)**:
construída depois, ao notar que a linha "Evidência no razão" do e-mail só
aparecia quando a própria fórmula da CHECK compunha BALANCETE + RAZÃO
explicitamente — uma dependência estreita do que a CHECK referencia, o
oposto do princípio do projeto. Agora `lookupContaNoRazao` varre a aba de
RAZÃO de verdade, direto pelo código da conta, formato de blocos ("Conta:
N - X.X.XX.XXX.XXX Nome" ... lançamentos ... "Total conta:", com os totais
de Débito/Crédito já calculados na própria linha de fechamento). Achado
real ao validar contra o arquivo do cliente: a paginação do razão original
(exportado de um PDF) repete a linha de cabeçalho de coluna
("Data | Histórico | ... | Débito | Crédito | ...") no meio do bloco de uma
conta quando um lançamento atravessa uma quebra de página — essa linha tem
texto ("Débito"/"Crédito") nas colunas de valor, não número, e sem filtrar
por `Number.isFinite` ela conta como um lançamento a mais (39 lançamentos
reais viravam 40). Teste de regressão em `engine.test.js` verificado (falha
sem o filtro, passa com ele). Cobertura no arquivo real: 94 das 96
contas×competência têm pelo menos um lançamento no razão do mês.

## Limitações conhecidas (v0.3)

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
