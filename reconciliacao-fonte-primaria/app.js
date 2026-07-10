/* ==========================================================================
   app.js — camada de UI sobre engine.js
   Lê o arquivo único de fechamento, detecta CHECK / FOPAG / BALANCETE por
   mês, monta o WorkbookAdapter sobre o SheetJS e roda a reconciliação
   independente para cada conta x competência encontrada na CHECK.
   ========================================================================== */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const engine = window.ReconEngine;

  const state = { adapter: null, workbook: null, resultados: [], confiabilidade: null, cobertura: { contasFantasmas: [] }, competencias: [] };

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

  // -------------------------------------------------- adapter sobre SheetJS

  function makeAdapter(workbook) {
    return {
      sheetNames: workbook.SheetNames,
      cell(sheet, col, row) {
        const ws = workbook.Sheets[sheet];
        if (!ws) return undefined;
        const c = ws[`${col}${row}`];
        if (!c || c.v === undefined || c.v === null || c.v === '') return undefined;
        return { value: c.v, formula: c.f };
      },
      usedRowCount(sheet) {
        const ws = workbook.Sheets[sheet];
        if (!ws || !ws['!ref']) return 0;
        return window.XLSX.utils.decode_range(ws['!ref']).e.r + 1;
      }
    };
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
      const m = name.match(/balancete\s*(\d{2})[.\-/](\d{4})/i);
      if (m) map[`${m[2]}-${m[1]}`] = name;
    }
    return map;
  }

  // ------------------------------------------ detecção de blocos na CHECK

  /** Acha a linha de cabeçalho (FOPAG | CONTABIL | DIFERENÇA) dentro das
   *  primeiras linhas da CHECK, sem assumir posição fixa. */
  function findHeaderRow(adapter, sheet) {
    for (let row = 1; row <= 8; row++) {
      for (let c = 1; c <= 80; c++) {
        const cell = adapter.cell(sheet, colIdxToLetter(c), row);
        if (cell && normalize(cell.value) === 'fopag') return row;
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

  // -------------------------------------------------------- orquestração

  function processWorkbook(workbook, fileName) {
    const adapter = makeAdapter(workbook);
    const checkSheet = findSheet(workbook, ['check']);
    if (!checkSheet) {
      showError('Não foi possível localizar a aba CHECK no arquivo. Sem ela, não há como ancorar as contas — confira o nome da aba.');
      return;
    }
    const headerRow = findHeaderRow(adapter, checkSheet);
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
    const prefixo = detectAccountFamilyPrefix(codigosNaCheck);
    const cobertura = { contasFantasmas: [] };
    const vistos = new Set();
    Object.values(balanceteMap).forEach((sheet) => {
      const r = engine.computeCobertura(adapter, sheet, codigosNaCheck, { prefixo });
      r.contasFantasmas.forEach((f) => { if (!vistos.has(f.codigo)) { vistos.add(f.codigo); cobertura.contasFantasmas.push(f); } });
    });

    state.adapter = adapter;
    state.workbook = workbook;
    state.resultados = resultados;
    state.confiabilidade = engine.computeConfiabilidade(resultados);
    state.cobertura = cobertura;
    state.competencias = [...new Set(competenciasAtivas)];

    renderMeta(fileName, checkSheet, blocks.filter((b) => competenciasAtivas.includes(b.competencia)), balanceteMap);
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
    $('meta-competencias').textContent = blocks.map((b) => b.competencia).join(', ') || '—';
    $('meta-balancetes').textContent = Object.values(balanceteMap).join(', ') || 'nenhum encontrado';
  }

  function badge(status) {
    const map = {
      'formula-viva': ['ok', 'fórmula viva'],
      'intra': ['ok', 'fórmula viva'],
      'formula-quebrada': ['critical', 'fórmula quebrada'],
      'valor-digitado': ['warn', 'valor digitado'],
      'verificado': ['ok', 'verificado'],
      'sem-fonte': ['neutral', 'sem fonte'],
      'sem-verificacao': ['neutral', 'sem verificação'],
      'conta-nao-encontrada': ['warn', 'conta não achada']
    };
    const [cls, label] = map[status] || ['neutral', status || '—'];
    return `<span class="badge badge-${cls}">${label}</span>`;
  }

  function alertBadge(tipo) {
    const map = {
      financeiro: ['critical', 'Financeiro'],
      preenchimento: ['flag', 'Preenchimento'],
      'financeiro-nao-verificado': ['warn', 'Financeiro (não verificado)']
    };
    const [cls, label] = map[tipo] || ['neutral', tipo];
    return `<span class="badge badge-${cls}">${label}</span>`;
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

  function populateFiltroCompetencia() {
    const sel = $('filtro-competencia');
    sel.innerHTML = '<option value="">Todas as competências</option>' +
      state.competencias.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  }

  function currentFilters() {
    return {
      competencia: $('filtro-competencia').value,
      alerta: $('filtro-alerta').value
    };
  }

  function renderTabela() {
    const { competencia, alerta } = currentFilters();
    let linhas = state.resultados;
    if (competencia) linhas = linhas.filter((r) => r.competencia === competencia);
    if (alerta === 'financeiro') linhas = linhas.filter((r) => r.alertas.includes('financeiro') || r.alertas.includes('financeiro-nao-verificado'));
    if (alerta === 'preenchimento') linhas = linhas.filter((r) => r.alertas.includes('preenchimento'));
    if (alerta === 'limpas') linhas = linhas.filter((r) => r.alertas.length === 0);

    const host = $('tabela-contas');
    if (!linhas.length) { host.innerHTML = '<p class="vazio">Nenhuma conta nesse filtro.</p>'; return; }

    const rowsHtml = linhas.map((r) => `
      <tr class="${r.alertas.includes('financeiro') ? 'row-financeiro' : ''} ${r.alertas.includes('preenchimento') ? 'row-preenchimento' : ''}">
        <td>${escapeHtml(r.competencia)}</td>
        <td class="mono">${escapeHtml(r.codigo)}</td>
        <td>${escapeHtml(r.descricao)}</td>
        <td class="num">${money.format(r.declarado.diferenca)}</td>
        <td class="num">${r.recalculado.diferenca === null ? '—' : money.format(r.recalculado.diferenca)}</td>
        <td>${badge(r.proveniencia.fopag)}</td>
        <td>${badge(r.proveniencia.contabil)}</td>
        <td>${r.alertas.length ? r.alertas.map(alertBadge).join(' ') : '<span class="badge badge-ok">ok</span>'}</td>
      </tr>`).join('');

    host.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Competência</th><th>Conta</th><th>Descrição</th>
            <th class="num">Diferença (CHECK)</th><th class="num">Diferença (recalculada)</th>
            <th>Proveniência FOPAG</th><th>Proveniência Contábil</th><th>Alertas</th>
          </tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>`;
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

  async function handleFile(file) {
    if (!file) return;
    $('nome-arquivo').textContent = file.name;
    try {
      const buf = await file.arrayBuffer();
      const workbook = window.XLSX.read(buf, { type: 'array', cellFormula: true, cellDates: true });
      processWorkbook(workbook, file.name);
    } catch (err) {
      showError(`Não foi possível ler o arquivo: ${err.message}`);
    }
  }

  function init() {
    $('input-arquivo').addEventListener('change', (e) => handleFile(e.target.files[0]));
    const dropzone = $('dropzone');
    ['dragenter', 'dragover'].forEach((evt) => dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('is-dragover'); }));
    ['dragleave', 'drop'].forEach((evt) => dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('is-dragover'); }));
    dropzone.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));
    $('filtro-competencia').addEventListener('change', renderTabela);
    $('filtro-alerta').addEventListener('change', renderTabela);
  }

  document.addEventListener('DOMContentLoaded', init);
})();
