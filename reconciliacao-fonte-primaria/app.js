/* ==========================================================================
   app.js — camada de UI sobre engine.js
   Lê o arquivo único de fechamento, detecta CHECK / FOPAG / BALANCETE por
   mês, monta o WorkbookAdapter sobre o ExcelJS e roda a reconciliação
   independente para cada conta x competência encontrada na CHECK.
   ========================================================================== */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const engine = window.ReconEngine;

  const state = {
    adapter: null, workbook: null, resultados: [], confiabilidade: null,
    cobertura: { contasFantasmas: [] }, competencias: [], competenciaDatas: new Map()
  };

  // ------------------------------------------------------------ utilidades

  function normalize(s) {
    return String(s ?? '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
  }
  function colIdxToLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  const money = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
  const pct = (n) => (n === null || n === undefined ? '—' : `${n}%`);

  // -------------------------------------------------- adapter sobre ExcelJS

  /** Empacota o Workbook do ExcelJS num objeto com `.SheetNames` (array de
   *  string) e `.Sheets` (nome -> worksheet) — mesma forma que o resto deste
   *  arquivo já espera, pra manter o diff pequeno na troca de biblioteca. */
  function normalizeWorkbook(exceljsWorkbook) {
    const sheets = {};
    exceljsWorkbook.worksheets.forEach((ws) => { sheets[ws.name] = ws; });
    return { SheetNames: exceljsWorkbook.worksheets.map((ws) => ws.name), Sheets: sheets };
  }

  function makeAdapter(workbook) {
    return {
      sheetNames: workbook.SheetNames,
      cell(sheet, col, row) {
        const ws = workbook.Sheets[sheet];
        if (!ws) return undefined;
        const c = ws.getCell(`${col}${row}`);
        if (!c) return undefined;
        const formula = c.formula || undefined;
        const value = formula ? c.result : c.value;
        if (value === undefined || value === null || value === '') return undefined;
        return { value, formula };
      },
      usedRowCount(sheet) {
        const ws = workbook.Sheets[sheet];
        return ws ? ws.rowCount || 0 : 0;
      }
    };
  }

  /** Acha a coluna cujo cabeçalho (linha 1) contém um dos apelidos, dentro
   *  das primeiras colunas de uma aba — usado pra não depender de posição
   *  fixa de coluna em planilhas de clientes diferentes. */
  /** Varre por ordem de prioridade dos apelidos (não por posição de coluna)
   *  — "custo total contabil" precisa ganhar de "custo fechamento" mesmo a
   *  coluna dele vindo depois na planilha; testar todas as colunas pro
   *  primeiro apelido antes de cair pro segundo evita pegar a coluna errada
   *  só porque ela aparece antes fisicamente. */
  function findColumnByHeader(adapter, sheet, aliases, headerRow, maxCol) {
    for (const alias of aliases) {
      const a = normalize(alias);
      for (let c = 1; c <= (maxCol || 150); c++) {
        const col = colIdxToLetter(c);
        const cell = adapter.cell(sheet, col, headerRow);
        if (cell && normalize(cell.value).includes(a)) return col;
      }
    }
    return null;
  }

  // --------------------------------------------------- detecção de abas

  function findSheet(workbook, aliases) {
    const norm = aliases.map(normalize);
    return workbook.SheetNames.find((name) => {
      const n = normalize(name);
      return norm.some((a) => n.includes(a));
    }) || null;
  }

  function findBalanceteMap(workbook) {
    const map = {};
    for (const name of workbook.SheetNames) {
      // Separador entre mês e ano é opcional — achado real: o arquivo mais
      // completo do cliente usa "Balancete 012026" (mês+ano colados, sem
      // separador nenhum), enquanto o arquivo original usa "BALANCETE
      // 04.2026" (ponto) e a fixture sintética testa hífen. \d{2}/\d{4} são
      // largura fixa, então tornar o separador opcional não introduz
      // ambiguidade em nenhum dos três formatos.
      const m = name.match(/balancete\s*(\d{2})[.\-/]?(\d{4})/i);
      if (m) map[`${m[2]}-${m[1]}`] = name;
    }
    return map;
  }

  // ------------------------------------------ detecção de blocos na CHECK

  /** Acha a linha de cabeçalho a partir de um rótulo-âncora, sem assumir
   *  posição fixa — planilhas reais têm o cabeçalho em linhas diferentes
   *  (a CHECK deste arquivo tem "FOPAG" na linha 3, a FOPAG 2026 tem "Centro
   *  de Custo" na linha 3 também, mas nada garante que sempre serão iguais). */
  function findHeaderRow(adapter, sheet, anchors, maxRow, maxCol) {
    const norm = anchors.map(normalize);
    for (let row = 1; row <= (maxRow || 8); row++) {
      for (let c = 1; c <= (maxCol || 80); c++) {
        const cell = adapter.cell(sheet, colIdxToLetter(c), row);
        if (cell && norm.includes(normalize(cell.value))) return row;
      }
    }
    return null;
  }

  function detectBlocks(adapter, sheet, headerRow) {
    const blocks = [];
    for (let c = 1; c <= 200; c++) {
      const col = colIdxToLetter(c);
      const cell = adapter.cell(sheet, col, headerRow);
      if (!cell || normalize(cell.value) !== 'fopag') continue;
      const contabilCol = colIdxToLetter(c + 1);
      const diferencaCol = colIdxToLetter(c + 2);
      const mesCell = adapter.cell(sheet, col, headerRow - 2);
      const dataCell = adapter.cell(sheet, col, headerRow - 1);
      // Blocos-modelo pra meses futuros sem dado nenhum (comum em templates
      // anuais) não têm uma data real na linha da competência — sem isso não
      // dá pra distinguir "mês sem movimento" de "coluna nunca usada", então
      // o bloco é descartado em vez de virar uma competência fantasma.
      if (!dataCell || !(dataCell.value instanceof Date)) continue;
      blocks.push({
        fopagCol: col,
        contabilCol,
        diferencaCol,
        competencia: mesCell && typeof mesCell.value === 'string' ? mesCell.value : dataCell.value.toISOString().slice(0, 7),
        data: dataCell.value
      });
    }
    return blocks;
  }

  function collectAccountRows(adapter, sheet, block, firstDataRow, maxRows) {
    const rows = [];
    let vazias = 0;
    for (let r = firstDataRow; r < firstDataRow + maxRows; r++) {
      const desc = adapter.cell(sheet, 'A', r);
      if (!desc) { vazias++; if (vazias > 15) break; continue; }
      vazias = 0;
      const texto = String(desc.value).trim();
      if (!texto || /^total/i.test(texto)) continue;
      const codigo = engine.parseAccountCode(texto);
      if (!codigo) continue;
      rows.push({ row: r, descricao: texto, codigo });
    }
    return rows;
  }

  /** Deriva o prefixo de família de contas (3 primeiros segmentos, ex.:
   *  "3.1.01") mais comum entre os códigos já mapeados na CHECK, pra limitar
   *  a busca de contas fantasmas ao mesmo grupo (evita listar todo o plano
   *  de contas de resultado como se fosse ausência de folha). */
  function detectAccountFamilyPrefix(codigosNaCheck) {
    const contagem = new Map();
    codigosNaCheck.forEach((codigo) => {
      const partes = codigo.split('.').slice(0, 3).join('.');
      contagem.set(partes, (contagem.get(partes) || 0) + 1);
    });
    let melhor = '3.';
    let max = 0;
    contagem.forEach((n, prefixo) => { if (n > max) { max = n; melhor = `${prefixo}.`; } });
    return melhor;
  }

  /** Profundidade (nº de segmentos) mais comum entre os códigos já mapeados
   *  na CHECK — usada pra não deixar conta-pai/subtotal do BALANCETE (menos
   *  segmentos que uma conta-folha, ex.: "3.1.01.003" vs "3.1.01.003.003")
   *  aparecer como se fosse uma conta faltando na cobertura. */
  function detectAccountDepth(codigosNaCheck) {
    const contagem = new Map();
    codigosNaCheck.forEach((codigo) => {
      const n = codigo.split('.').length;
      contagem.set(n, (contagem.get(n) || 0) + 1);
    });
    let melhor = null, max = 0;
    contagem.forEach((n, profundidade) => { if (n > max) { max = n; melhor = profundidade; } });
    return melhor;
  }

  // -------------------------------------------------------- orquestração

  function processWorkbook(workbook, fileName) {
    const adapter = makeAdapter(workbook);
    const checkSheet = findSheet(workbook, ['check']);
    if (!checkSheet) {
      showError('Não foi possível localizar a aba CHECK no arquivo. Sem ela, não há como ancorar as contas — confira o nome da aba.');
      return;
    }
    const headerRow = findHeaderRow(adapter, checkSheet, ['fopag']);
    if (headerRow === null) {
      showError(`Aba "${checkSheet}" encontrada, mas o cabeçalho FOPAG/CONTABIL/DIFERENÇA não foi localizado nas primeiras linhas.`);
      return;
    }
    const blocks = detectBlocks(adapter, checkSheet, headerRow);
    const balanceteMap = findBalanceteMap(workbook);

    const resultados = [];
    const codigosNaCheck = new Set();
    const competenciasAtivas = [];
    for (const block of blocks) {
      const contas = collectAccountRows(adapter, checkSheet, block, headerRow + 2, 300);
      // Planilhas anuais costumam ter a linha de data pré-preenchida em todos
      // os 12 meses mesmo sem nenhum dado lançado ainda — só entra como
      // competência de verdade se algum valor de Folha/Contábil existir.
      const temDado = contas.some((conta) =>
        adapter.cell(checkSheet, block.fopagCol, conta.row) || adapter.cell(checkSheet, block.contabilCol, conta.row));
      if (!temDado) continue;
      competenciasAtivas.push(block.competencia);
      let balanceteSheet = null;
      if (block.data instanceof Date) {
        const key = `${block.data.getFullYear()}-${String(block.data.getMonth() + 1).padStart(2, '0')}`;
        balanceteSheet = balanceteMap[key] || null;
      }
      block.balanceteSheet = balanceteSheet;
      for (const conta of contas) {
        codigosNaCheck.add(conta.codigo);
        const fopagCell = adapter.cell(checkSheet, block.fopagCol, conta.row);
        const contabilCell = adapter.cell(checkSheet, block.contabilCol, conta.row);
        const diferencaCell = adapter.cell(checkSheet, block.diferencaCol, conta.row);
        resultados.push(engine.reconcileConta({
          adapter, checkSheet, contaCodigo: conta.codigo, descricao: conta.descricao,
          competencia: block.competencia, fopagCell, contabilCell, diferencaCell, balanceteSheet
        }));
      }
    }

    // A cobertura só faz sentido dentro da "família" de contas que a CHECK já
    // cobre (ex.: 3.1.01 = pessoal) — comparar contra todo o grupo de
    // resultado (3.*) traria centenas de contas não-folha como falso positivo.
    // `profundidade` evita listar conta-pai/subtotal (menos segmentos) como
    // se fosse uma conta faltando quando ela só agrega contas-folha que a
    // CHECK já cobre individualmente (achado real neste arquivo).
    const prefixo = detectAccountFamilyPrefix(codigosNaCheck);
    const profundidade = detectAccountDepth(codigosNaCheck);
    const cobertura = { contasFantasmas: [] };
    const vistos = new Set();
    Object.values(balanceteMap).forEach((sheet) => {
      const r = engine.computeCobertura(adapter, sheet, codigosNaCheck, { prefixo, profundidade });
      r.contasFantasmas.forEach((f) => { if (!vistos.has(f.codigo)) { vistos.add(f.codigo); cobertura.contasFantasmas.push(f); } });
    });

    const blocksAtivos = blocks.filter((b) => competenciasAtivas.includes(b.competencia));

    state.adapter = adapter;
    state.workbook = workbook;
    state.resultados = resultados;
    state.confiabilidade = engine.computeConfiabilidade(resultados);
    state.cobertura = cobertura;
    state.competencias = [...new Set(competenciasAtivas)];
    state.competenciaDatas = new Map(blocksAtivos.map((b) => [b.competencia, b.data]));

    renderMeta(fileName, checkSheet, blocksAtivos, balanceteMap);
    renderDashboard();
  }

  // ------------------------------------------------------------- render

  function showError(msg) {
    const el = $('avisos');
    el.hidden = false;
    el.innerHTML = `<div class="aviso aviso-erro">${escapeHtml(msg)}</div>`;
    $('dashboard').hidden = true;
  }
  function escapeHtml(v) {
    return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function renderMeta(fileName, checkSheet, blocks, balanceteMap) {
    $('avisos').hidden = true;
    $('meta-arquivo').textContent = fileName;
    $('meta-check').textContent = checkSheet;
    $('meta-competencias').textContent = blocks
      .map((b) => b.competencia + (b.balanceteSheet && /prev/i.test(b.balanceteSheet) ? ' (previsão)' : ''))
      .join(', ') || '—';
    $('meta-balancetes').textContent = Object.values(balanceteMap).join(', ') || 'nenhum encontrado';
  }

  const STATUS_MAP = {
    'formula-viva': ['ok', 'fórmula viva'],
    'intra': ['ok', 'fórmula viva'],
    'formula-quebrada': ['critical', 'fórmula quebrada'],
    'valor-digitado': ['warn', 'valor digitado'],
    'verificado': ['ok', 'verificado'],
    'sem-fonte': ['neutral', 'sem fonte'],
    'sem-verificacao': ['neutral', 'sem verificação'],
    'conta-nao-encontrada': ['warn', 'conta não achada']
  };
  const ALERTA_MAP = {
    financeiro: ['critical', 'Financeiro'],
    preenchimento: ['flag', 'Preenchimento'],
    'financeiro-nao-verificado': ['warn', 'Financeiro (não verificado)']
  };

  function badge(status) {
    const [cls, label] = STATUS_MAP[status] || ['neutral', status || '—'];
    return `<span class="badge badge-${cls}">${label}</span>`;
  }

  function statusLabel(status) {
    return (STATUS_MAP[status] || [null, status || '—'])[1];
  }

  function alertBadge(tipo) {
    const [cls, label] = ALERTA_MAP[tipo] || ['neutral', tipo];
    return `<span class="badge badge-${cls}">${label}</span>`;
  }

  function alertLabel(tipo) {
    return (ALERTA_MAP[tipo] || [null, tipo])[1];
  }

  function renderDashboard() {
    $('dashboard').hidden = false;

    const c = state.confiabilidade;
    $('score-confiabilidade').textContent = pct(c.score);
    $('score-confiabilidade-detail').textContent =
      `${c.celulasVivas} célula(s) viva(s) · ${c.celulasQuebradas} quebrada(s) · ${c.celulasDigitadas} digitada(s)`;

    const alertasFinanceiro = state.resultados.filter((r) => r.alertas.includes('financeiro'));
    const alertasPreenchimento = state.resultados.filter((r) => r.alertas.includes('preenchimento'));
    const semFonte = state.resultados.filter((r) => r.recalculado.statusContabil === 'sem-fonte' || r.recalculado.statusFopag === 'sem-verificacao');

    $('kpi-financeiro').textContent = alertasFinanceiro.length;
    $('kpi-preenchimento').textContent = alertasPreenchimento.length;
    $('kpi-sem-fonte').textContent = semFonte.length;
    $('kpi-fantasmas').textContent = state.cobertura.contasFantasmas.length;

    populateFiltroCompetencia();
    renderTabela();
    renderFantasmas();
  }

  /** A competência é "previsão" (ainda não fechada) quando o BALANCETE que
   *  a alimenta tem "PREV" no nome — achado real: o arquivo mais completo
   *  do cliente tem meses futuros como "PREV. Balancete 052026", e sem
   *  marcar isso na tela um alerta financeiro em cima de previsão parece
   *  divergência de livro fechado, quando não é. */
  function linhaEhPrevisao(r) {
    return Boolean(r.recalculado.fonteContabil && /prev/i.test(r.recalculado.fonteContabil));
  }

  function competenciaEhPrevisao(competencia) {
    const r = state.resultados.find((x) => x.competencia === competencia);
    return Boolean(r && linhaEhPrevisao(r));
  }

  function populateFiltroCompetencia() {
    const sel = $('filtro-competencia');
    sel.innerHTML = '<option value="">Todas as competências</option>' +
      state.competencias.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}${competenciaEhPrevisao(c) ? ' (previsão)' : ''}</option>`).join('');
  }

  function currentFilters() {
    return {
      competencia: $('filtro-competencia').value,
      alerta: $('filtro-alerta').value,
      severidade: $('filtro-severidade') ? $('filtro-severidade').value : ''
    };
  }

  function filteredResultados() {
    const { competencia, alerta, severidade } = currentFilters();
    let linhas = state.resultados;
    if (competencia) linhas = linhas.filter((r) => r.competencia === competencia);
    if (alerta === 'financeiro') linhas = linhas.filter((r) => r.alertas.includes('financeiro') || r.alertas.includes('financeiro-nao-verificado'));
    if (alerta === 'preenchimento') linhas = linhas.filter((r) => r.alertas.includes('preenchimento'));
    if (alerta === 'limpas') linhas = linhas.filter((r) => r.alertas.length === 0);
    if (severidade) linhas = linhas.filter((r) => (dictEntry(r.codigo) || {}).severidade === severidade);
    return linhas;
  }

  function dictEntry(codigo) {
    return state.dicionario ? state.dicionario.get(codigo) : null;
  }

  function renderTabela() {
    const linhas = filteredResultados();
    const host = $('tabela-contas');
    if (!linhas.length) { host.innerHTML = '<p class="vazio">Nenhuma conta nesse filtro.</p>'; return; }

    const comDicionario = Boolean(state.dicionario && state.dicionario.size);
    const colspan = comDicionario ? 11 : 9;
    const rowsHtml = linhas.map((r) => {
      const key = `${r.competencia}__${r.codigo}`;
      const dict = dictEntry(r.codigo);
      const previsao = linhaEhPrevisao(r);
      const extraCols = comDicionario
        ? `<td>${dict && dict.severidade ? `<span class="badge badge-${sevClass(dict.severidade)}">${escapeHtml(dict.severidade)}</span>` : '—'}</td><td>${escapeHtml((dict && dict.responsavel) || '—')}</td>`
        : '';
      return `
      <tr class="row-conta ${r.alertas.includes('financeiro') ? 'row-financeiro' : ''} ${r.alertas.includes('preenchimento') ? 'row-preenchimento' : ''}" data-key="${escapeHtml(key)}" tabindex="0" role="button" aria-expanded="false">
        <td><span class="expand-caret">▸</span></td>
        <td>${escapeHtml(r.competencia)}${previsao ? ' <span class="badge badge-neutral" title="Mês ainda não fechado — BALANCETE de previsão, não contábil final">previsão</span>' : ''}</td>
        <td class="mono">${escapeHtml(r.codigo)}</td>
        <td>${escapeHtml(r.descricao)}</td>
        <td class="num">${money.format(r.declarado.diferenca)}</td>
        <td class="num">${r.recalculado.diferenca === null ? '—' : money.format(r.recalculado.diferenca)}</td>
        <td>${badge(r.proveniencia.fopag)}</td>
        <td>${badge(r.proveniencia.contabil)}</td>
        <td><div class="alertas-cell">${r.alertas.length ? r.alertas.map(alertBadge).join('') : '<span class="badge badge-ok">ok</span>'}${r.recalculado.ajusteRazao ? '<span class="badge badge-flag">ajuste RAZÃO</span>' : ''}</div></td>
        ${extraCols}
      </tr>
      <tr class="row-detalhe" data-detail-for="${escapeHtml(key)}" hidden><td colspan="${colspan}">${renderDetalheConta(r)}</td></tr>`;
    }).join('');

    host.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th></th><th>Competência</th><th>Conta</th><th>Descrição</th>
            <th class="num">Diferença (CHECK)</th><th class="num">Diferença (recalculada)</th>
            <th>Proveniência FOPAG</th><th>Proveniência Contábil</th><th>Alertas</th>
            ${comDicionario ? '<th>Severidade</th><th>Responsável</th>' : ''}
          </tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>`;

    host.querySelectorAll('.row-conta').forEach((row) => {
      const toggle = () => {
        const key = row.dataset.key;
        const detail = host.querySelector(`.row-detalhe[data-detail-for="${CSS.escape(key)}"]`);
        const expanded = row.getAttribute('aria-expanded') === 'true';
        row.setAttribute('aria-expanded', String(!expanded));
        row.querySelector('.expand-caret').textContent = expanded ? '▸' : '▾';
        if (detail) detail.hidden = expanded;
      };
      row.addEventListener('click', toggle);
      row.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); } });
    });

    host.querySelectorAll('.btn-gerar-email').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const r = linhas.find((x) => `${x.competencia}__${x.codigo}` === btn.dataset.key);
        if (r) gerarEmail(r);
      });
    });
  }

  /** Nota (texto puro) sobre a origem do Contábil recalculado — reutilizada
   *  tanto no HTML do drill-down quanto no rascunho de e-mail. */
  function contabilNota(r) {
    if (r.recalculado.ajusteRazao) {
      const a = r.recalculado.ajusteRazao;
      return `Contábil = BALANCETE + ${a.termos.length} lançamento(s) do RAZÃO que a própria CHECK já reclassifica, totalizando ${money.format(a.valor)}.`;
    }
    if (r.recalculado.origemContabil === 'balancete-por-codigo') {
      return 'Contábil recalculado direto no BALANCETE pelo código de conta (a fórmula da CHECK não pôde ser seguida com confiança).';
    }
    if (r.recalculado.statusContabil === 'sem-fonte') {
      return 'Sem BALANCETE disponível para esta competência — não há como verificar o Contábil.';
    }
    return 'Contábil confirmado direto pela composição da própria fórmula da CHECK.';
  }

  function renderDetalheConta(r) {
    const key = `${r.competencia}__${r.codigo}`;
    const fopagTermos = r.recalculado.fopagTermos || [];
    const fopagHtml = fopagTermos.length
      ? `<table class="detail-table">
          <thead><tr><th>Rubrica (FOPAG 2026)</th><th class="num">Valor</th></tr></thead>
          <tbody>${fopagTermos.map((t) => `<tr><td>${escapeHtml(t.rubrica)}</td><td class="num">${money.format(t.value)}</td></tr>`).join('')}</tbody>
         </table>`
      : '<p class="vazio">Sem composição verificável (FOPAG sem fórmula reconhecida nessa célula).</p>';

    let contabilHtml = `<p class="detail-note">${escapeHtml(contabilNota(r))}</p>`;
    if (r.recalculado.ajusteRazao) {
      const a = r.recalculado.ajusteRazao;
      contabilHtml += `
        <table class="detail-table">
          <thead><tr><th>Origem</th><th class="num">Valor</th></tr></thead>
          <tbody>${a.termos.map((t) => `<tr><td class="mono">${escapeHtml(t.sheet)}!${escapeHtml(t.coluna || t.col)}${t.row}</td><td class="num">${money.format(t.value)}</td></tr>`).join('')}</tbody>
        </table>`;
    }

    return `<div class="detail-grid">
      <div><h4>Composição da Folha</h4>${fopagHtml}</div>
      <div><h4>Composição do Contábil</h4>${contabilHtml}</div>
    </div>
    <div class="email-cta">
      <button type="button" class="btn-ghost btn-small btn-gerar-email" data-key="${escapeHtml(key)}">Gerar e-mail</button>
    </div>`;
  }

  /** Frase objetiva sobre QUAL lado (Folha ou Contábil) tem valor
   *  confirmado direto na fonte primária e qual não tem — aponta de onde
   *  vem o erro (falta de dado de um lado) em vez de deixar isso implícito
   *  nos rótulos técnicos de proveniência. */
  function origemDivergencia(r) {
    const temFopag = r.recalculado.statusFopag === 'verificado';
    const temContabil = r.recalculado.statusContabil === 'verificado';
    const vFopag = temFopag ? money.format(r.recalculado.fopag) : null;
    const vContabil = temContabil ? money.format(r.recalculado.contabil) : null;

    if (temFopag && temContabil) {
      return `Os dois lados têm valor confirmado direto na fonte — Folha (FOPAG): ${vFopag}; Contábil (BALANCETE): ${vContabil}. A diferença é real, não é falta de dado de nenhum dos dois lados.`;
    }
    if (temFopag && !temContabil) {
      const motivo = r.recalculado.statusContabil === 'conta-nao-encontrada'
        ? `a conta ${r.codigo} não foi encontrada no BALANCETE desta competência`
        : 'o BALANCETE desta competência não pôde ser localizado/verificado';
      return `O erro está do lado do CONTÁBIL: a Folha (FOPAG) tem valor confirmado (${vFopag}), mas ${motivo}.`;
    }
    if (!temFopag && temContabil) {
      return `O erro está do lado da FOLHA: o Contábil (BALANCETE) tem valor confirmado (${vContabil}), mas a fórmula da CHECK para a Folha (FOPAG) não pôde ser recalculada com confiança nessa célula.`;
    }
    return `Nem Folha (FOPAG) nem Contábil (BALANCETE) puderam ser confirmados direto na fonte para a conta ${r.codigo} — não dá pra apontar de qual lado vem o erro sem revisão manual da planilha.`;
  }

  /** Monta um rascunho de e-mail (assunto + corpo) com o detalhamento da
   *  divergência de uma conta, pra agilizar a validação com o responsável.
   *  Nada é enviado — só preenche o painel para revisão manual. */
  function gerarEmailTexto(r) {
    const dict = dictEntry(r.codigo);
    const subject = `Divergência de conciliação — ${r.competencia} — ${r.codigo} — ${r.descricao}`;
    const fopagTermos = r.recalculado.fopagTermos || [];
    const body = [
      'Prezado(a) [nome],',
      '',
      `Na reconciliação da competência ${r.competencia} foi identificada a divergência abaixo, apurada direto das fontes primárias (FOPAG/BALANCETE). Segue o detalhamento para agilizar a validação.`,
      '',
      '1) CONTA',
      `   • Competência: ${r.competencia}`,
      `   • Conta: ${r.codigo} — ${r.descricao}`,
      '',
      '2) VALORES',
      `   • Diferença declarada na CHECK: ${money.format(r.declarado.diferenca)}`,
      `   • Diferença recalculada (Folha × Contábil na fonte): ${r.recalculado.diferenca === null ? '—' : money.format(r.recalculado.diferenca)}`,
      `   • Folha (FOPAG): ${r.recalculado.statusFopag === 'verificado' ? money.format(r.recalculado.fopag) : `— (${statusLabel(r.recalculado.statusFopag)})`}`,
      `   • Contábil (BALANCETE): ${r.recalculado.statusContabil === 'verificado' ? money.format(r.recalculado.contabil) : `— (${statusLabel(r.recalculado.statusContabil)})`}`,
      '',
      '3) DE ONDE VEM O ERRO',
      `   • ${origemDivergencia(r)}`,
      `   • ${contabilNota(r)}`,
      '',
      '4) DIAGNÓSTICO',
      `   • Alertas: ${r.alertas.length ? r.alertas.map(alertLabel).join(', ') : 'nenhum'}`,
      fopagTermos.length ? '   • Composição da Folha (FOPAG):' : '   • Sem composição verificável (FOPAG sem fórmula reconhecida nessa célula).',
      ...fopagTermos.map((t) => `     - ${t.rubrica}: ${money.format(t.value)}`),
      '',
      '5) RESPONSÁVEL',
      `   • Severidade: ${(dict && dict.severidade) || 'A definir'}`,
      `   • Responsável: ${(dict && dict.responsavel) || 'A definir'}`,
      '',
      'Solicito a validação da causa e o retorno com a tratativa para fechamento da competência.',
      '',
      'Atenciosamente,',
      '[Seu nome]'
    ].filter((line) => line !== undefined).join('\n');
    return { subject, body };
  }

  function gerarEmail(r) {
    const { subject, body } = gerarEmailTexto(r);
    $('email-title').textContent = subject;
    $('email-text').value = body;
    $('email-placeholder').classList.add('is-hidden');
    $('email-panel').classList.remove('is-hidden');
    $('email-panel').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function sevClass(severidade) {
    const n = normalize(severidade);
    if (n.includes('alt')) return 'critical';
    if (n.includes('med')) return 'warn';
    if (n.includes('baix')) return 'ok';
    return 'neutral';
  }

  // ------------------------------------------------- dicionário de contas

  /** Tokeniza um CSV respeitando campos entre aspas (aspas escapadas como
   *  "", delimitador ou quebra de linha dentro do campo) — sem isso, um
   *  responsável como "Silva, João" (delimitador vírgula) ou com ; no nome
   *  (delimitador ponto-e-vírgula) corromperia as colunas seguintes. */
  function parseCsvRows(text, delim) {
    const rows = [];
    let row = [], field = '', inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else { inQuotes = false; }
        } else {
          field += c;
        }
      } else if (c === '"') {
        inQuotes = true;
      } else if (c === delim) {
        row.push(field); field = '';
      } else if (c === '\n') {
        row.push(field); field = ''; rows.push(row); row = [];
      } else if (c === '\r') {
        // ignora — o \n seguinte fecha a linha
      } else {
        field += c;
      }
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function parseDicionarioCsv(text) {
    const map = new Map();
    const semBom = text.replace(/^\uFEFF/, '');
    if (!semBom.trim()) return map;
    const primeiraLinha = semBom.slice(0, semBom.search(/\r?\n/) === -1 ? semBom.length : semBom.search(/\r?\n/));
    const delim = primeiraLinha.includes(';') ? ';' : ',';
    const rows = parseCsvRows(semBom, delim).filter((r) => r.some((c) => c.trim()));
    if (!rows.length) return map;
    let start = 0;
    if (normalize(rows[0][0]).includes('codigo')) start = 1;
    for (let i = start; i < rows.length; i++) {
      const codigo = (rows[i][0] || '').trim();
      if (!codigo) continue;
      map.set(codigo, { severidade: (rows[i][1] || '').trim(), responsavel: (rows[i][2] || '').trim() });
    }
    return map;
  }

  async function handleDicionario(file) {
    if (!file) return;
    const text = await file.text();
    state.dicionario = parseDicionarioCsv(text);
    $('nome-dicionario').textContent = `${file.name} (${state.dicionario.size} conta(s) mapeada(s))`;
    populateFiltroSeveridade();
    if (!$('dashboard').hidden) renderTabela();
  }

  function populateFiltroSeveridade() {
    const sel = $('filtro-severidade');
    if (!sel || !state.dicionario) return;
    const severidades = new Set();
    state.dicionario.forEach((v) => { if (v.severidade) severidades.add(v.severidade); });
    sel.innerHTML = '<option value="">Todas as severidades</option>' +
      [...severidades].sort().map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
  }

  // ------------------------------------------------------------ exportação

  function toCsvValue(v) {
    const s = String(v ?? '');
    return /[",;\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }

  function numCsv(v) {
    return v === null || v === undefined ? '' : v.toFixed(2).replace('.', ',');
  }

  function exportarCsv() {
    const linhas = filteredResultados();
    if (!linhas.length) return;
    const header = [
      'Competência', 'Previsão', 'Conta', 'Descrição',
      'Diferença (CHECK)', 'Diferença (recalculada)',
      'Folha (FOPAG)', 'Status Folha', 'Contábil (BALANCETE)', 'Status Contábil',
      'Alertas', 'Nota do Contábil', 'Severidade', 'Responsável'
    ];
    const rows = linhas.map((r) => {
      const dict = dictEntry(r.codigo) || {};
      return [
        r.competencia,
        linhaEhPrevisao(r) ? 'Sim' : 'Não',
        r.codigo, r.descricao,
        numCsv(r.declarado.diferenca),
        numCsv(r.recalculado.diferenca),
        numCsv(r.recalculado.statusFopag === 'verificado' ? r.recalculado.fopag : null),
        statusLabel(r.recalculado.statusFopag),
        numCsv(r.recalculado.statusContabil === 'verificado' ? r.recalculado.contabil : null),
        statusLabel(r.recalculado.statusContabil),
        r.alertas.length ? r.alertas.map(alertLabel).join(', ') : 'Nenhum',
        contabilNota(r),
        dict.severidade || '', dict.responsavel || ''
      ];
    });
    const csv = [header, ...rows].map((row) => row.map(toCsvValue).join(';')).join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `reconciliacao-independente-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function renderFantasmas() {
    const host = $('lista-fantasmas');
    const items = state.cobertura.contasFantasmas;
    if (!items.length) { host.innerHTML = '<p class="vazio">Nenhuma conta do BALANCETE ficou de fora da CHECK.</p>'; return; }
    host.innerHTML = `<ul class="fantasma-list">${items.map((f) =>
      `<li><span class="mono">${escapeHtml(f.codigo)}</span> — presente em <em>${escapeHtml(f.sheet)}</em> (linha ${f.row}), sem linha correspondente na CHECK</li>`
    ).join('')}</ul>`;
  }

  // ------------------------------------------------------------- eventos

  // Alterna e lembra o tema (claro/escuro) de uma sessão pra outra. O tema
  // inicial já foi aplicado por um script inline no <head> (evita flash);
  // aqui só cuidamos do clique e, enquanto o usuário não escolher nada
  // explicitamente, seguimos mudanças do tema do sistema operacional.
  function initTema() {
    const KEY = 'reconciliacao-fonte-primaria:tema';
    const btn = $('btn-tema');
    if (!btn) return;

    function aplicar(tema) {
      document.documentElement.dataset.theme = tema;
      btn.setAttribute('aria-pressed', tema === 'dark' ? 'true' : 'false');
    }
    aplicar(document.documentElement.dataset.theme || 'light');

    btn.addEventListener('click', () => {
      const proximo = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
      aplicar(proximo);
      try { localStorage.setItem(KEY, proximo); } catch (e) { /* sem storage disponível, tudo bem */ }
    });

    let temEscolhaSalva = false;
    try { temEscolhaSalva = localStorage.getItem(KEY) !== null; } catch (e) {}
    if (!temEscolhaSalva && window.matchMedia) {
      const mq = window.matchMedia('(prefers-color-scheme: dark)');
      const onMudancaSistema = (e) => {
        let aindaSemEscolha = true;
        try { aindaSemEscolha = localStorage.getItem(KEY) === null; } catch (err) {}
        if (aindaSemEscolha) aplicar(e.matches ? 'dark' : 'light');
      };
      if (mq.addEventListener) mq.addEventListener('change', onMudancaSistema);
      else if (mq.addListener) mq.addListener(onMudancaSistema);
    }
  }

  function setCarregando(ativo) {
    const el = $('carregando');
    if (el) el.hidden = !ativo;
    const input = $('input-arquivo');
    if (input) input.disabled = ativo;
  }

  async function handleFile(file) {
    if (!file) return;
    $('nome-arquivo').textContent = file.name;
    setCarregando(true);
    // Cede o controle pro navegador pintar o indicador de carregamento antes
    // do trabalho síncrono pesado (parse do XLSX inteiro, recálculo de todas
    // as contas) — sem isso, um arquivo grande trava a thread e o spinner
    // nunca chega a aparecer.
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    try {
      const buf = await file.arrayBuffer();
      const exceljsWorkbook = new window.ExcelJS.Workbook();
      await exceljsWorkbook.xlsx.load(buf);
      const workbook = normalizeWorkbook(exceljsWorkbook);
      processWorkbook(workbook, file.name);
    } catch (err) {
      showError(`Não foi possível ler o arquivo: ${err.message}`);
    } finally {
      setCarregando(false);
    }
  }

  function init() {
    initTema();
    $('input-arquivo').addEventListener('change', (e) => handleFile(e.target.files[0]));
    const dropzone = $('dropzone');
    ['dragenter', 'dragover'].forEach((evt) => dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('is-dragover'); }));
    ['dragleave', 'drop'].forEach((evt) => dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('is-dragover'); }));
    dropzone.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));
    $('filtro-competencia').addEventListener('change', renderTabela);
    $('filtro-alerta').addEventListener('change', renderTabela);
    $('btn-exportar-csv').addEventListener('click', exportarCsv);
    $('filtro-severidade').addEventListener('change', renderTabela);
    $('input-dicionario').addEventListener('change', (e) => handleDicionario(e.target.files[0]));

    $('btn-copy-email').addEventListener('click', async (e) => {
      const btn = e.currentTarget;
      const value = $('email-text').value;
      if (!value) return;
      try { await navigator.clipboard.writeText(value); }
      catch (_) { $('email-text').select(); document.execCommand('copy'); }
      const original = btn.textContent;
      btn.textContent = 'Copiado!';
      setTimeout(() => { btn.textContent = original; }, 1500);
    });
    $('btn-mailto').addEventListener('click', () => {
      const subject = $('email-title').textContent;
      const body = $('email-text').value;
      window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
