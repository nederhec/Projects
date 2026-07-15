/* ==========================================================================
   painel-custos.js — Painel de Custos e Conciliação (nova aba)
   Consome os mesmos resultados já reconciliados pela tabela de "Contas por
   competência" (state.resultados) — não recalcula nada por conta própria.
   Fica disponível assim que um arquivo é lido (ver habilitarPainelCustos em
   app.js) — não exige a competência sem divergências, já que isso nem
   sempre vai acontecer.
   Gráficos via Chart.js (vendor/chart.umd.min.js), desenhados ao vivo na
   página — substituiu a versão anterior em PDF.
   ========================================================================== */
(function () {
  'use strict';

  const BASELINE_PREFIX = 'reconciliacao-fonte-primaria:baseline-divergencias:';
  const charts = [];

  // -------------------------------------------------------------- formatação

  function moneyAbrev(v, money) {
    const sinal = v < 0 ? '-' : '';
    const abs = Math.abs(v);
    if (abs >= 1_000_000) return `${sinal}R$ ${(abs / 1_000_000).toFixed(2).replace('.', ',')} mi`;
    if (abs >= 1_000) return `${sinal}R$ ${Math.round(abs / 1_000).toLocaleString('pt-BR')} mil`;
    return money.format(v);
  }

  function pctFmt(n) {
    return `${n >= 0 ? '+' : ''}${n.toFixed(1).replace('.', ',')}%`;
  }

  // ------------------------------------------------------- classificação

  /** Mesma heurística por palavra-chave documentada no README — sem fonte
   *  pronta no arquivo com a quebra Salário/Encargos/Benefícios. */
  function classificarPessoal(descricao, normalize) {
    const d = normalize(descricao);
    if (/\binss\b|\bfgts\b|\bpis\b|\bgrrf\b|encargo/.test(d)) return 'Encargos';
    if (/\bvale\b|assistenc|seguro de vida|plano de saude|auxilio|beneficio/.test(d)) return 'Benefícios';
    return 'Salário';
  }

  function acharColunaPorCabecalho(adapter, sheet, aliases, headerRow, colIdxToLetter, normalize, maxCol) {
    for (const alias of aliases) {
      for (let c = 1; c <= (maxCol || 40); c++) {
        const col = colIdxToLetter(c);
        const cell = adapter.cell(sheet, col, headerRow);
        if (cell && normalize(cell.value).includes(alias)) return col;
      }
    }
    return null;
  }

  /** Conta pessoas distintas (por NOME) na aba de Rescisões, agrupadas pela
   *  competência (mês/ano) da coluna MÊS. Sem essa aba, devolve Map vazio —
   *  o KPI mostra "—", não 0, pra não parecer "zero rescisões" quando na
   *  verdade é "sem como verificar". */
  function contarRescisoesPorCompetencia(state, normalize, colIdxToLetter) {
    const resultado = new Map();
    const workbook = state.workbook;
    const adapter = state.adapter;
    if (!workbook || !adapter) return resultado;
    const sheetName = workbook.SheetNames.find((n) => normalize(n).includes('rescis'));
    if (!sheetName) return resultado;

    let headerRow = null;
    for (let r = 1; r <= 6 && !headerRow; r++) {
      for (let c = 1; c <= 30; c++) {
        const cell = adapter.cell(sheetName, colIdxToLetter(c), r);
        if (cell && normalize(cell.value) === 'nome') { headerRow = r; break; }
      }
    }
    if (!headerRow) return resultado;

    const nomeCol = acharColunaPorCabecalho(adapter, sheetName, ['nome'], headerRow, colIdxToLetter, normalize);
    const mesCol = acharColunaPorCabecalho(adapter, sheetName, ['mes'], headerRow, colIdxToLetter, normalize);
    if (!nomeCol || !mesCol) return resultado;

    const nomesPorChaveMes = new Map();
    const total = adapter.usedRowCount(sheetName) || 0;
    for (let r = headerRow + 1; r <= total; r++) {
      const nomeCell = adapter.cell(sheetName, nomeCol, r);
      const mesCell = adapter.cell(sheetName, mesCol, r);
      if (!nomeCell || !mesCell || !(mesCell.value instanceof Date)) continue;
      const chave = `${mesCell.value.getFullYear()}-${mesCell.value.getMonth()}`;
      if (!nomesPorChaveMes.has(chave)) nomesPorChaveMes.set(chave, new Set());
      nomesPorChaveMes.get(chave).add(normalize(nomeCell.value));
    }

    state.competenciaDatas.forEach((data, label) => {
      const chave = `${data.getFullYear()}-${data.getMonth()}`;
      if (nomesPorChaveMes.has(chave)) resultado.set(label, nomesPorChaveMes.get(chave).size);
    });
    return resultado;
  }

  // ------------------------------------------------------ baseline (localStorage)

  function chaveBaseline(data) {
    return `${BASELINE_PREFIX}${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, '0')}`;
  }

  function lerOuGravarBaseline(chave, valorAtual) {
    try {
      const existente = localStorage.getItem(chave);
      if (existente !== null) return Number(existente);
      localStorage.setItem(chave, String(valorAtual));
      return valorAtual;
    } catch (e) {
      return valorAtual;
    }
  }

  function zerarBaselines() {
    try {
      Object.keys(localStorage)
        .filter((k) => k.startsWith(BASELINE_PREFIX))
        .forEach((k) => localStorage.removeItem(k));
    } catch (e) { /* sem storage disponível, tudo bem */ }
  }

  // ------------------------------------------------------------------ KPIs

  function computeKpisMes(state, competencia, tipoDivergencia, normalize, rescisoesPorCompetencia) {
    const linhas = state.resultados.filter((r) => r.competencia === competencia);
    let custoContabil = 0;
    let diferencaLiquida = 0;
    let diferencaAbsoluta = 0;
    let conciliadas = 0;
    let divergenciasAtuais = 0;
    const pessoal = { Salário: 0, Encargos: 0, Benefícios: 0 };

    linhas.forEach((r) => {
      const fopagOk = r.recalculado.statusFopag === 'verificado';
      const contabilOk = r.recalculado.statusContabil === 'verificado';
      if (contabilOk) {
        custoContabil += r.recalculado.contabil;
        pessoal[classificarPessoal(r.descricao, normalize)] += r.recalculado.contabil;
      }
      if (fopagOk && contabilOk) conciliadas++;
      if (r.recalculado.diferenca !== null) diferencaLiquida += r.recalculado.diferenca;
      if (r.alertas.includes('financeiro')) diferencaAbsoluta += Math.abs(r.recalculado.diferenca);
      if (r.alertas.length > 0 || tipoDivergencia(r) !== 'Sem divergência') divergenciasAtuais++;
    });

    const data = state.competenciaDatas.get(competencia);
    const divergenciasNoUpload = data ? lerOuGravarBaseline(chaveBaseline(data), divergenciasAtuais) : divergenciasAtuais;

    return {
      competencia,
      custoContabil, diferencaLiquida, diferencaAbsoluta,
      percConciliado: linhas.length ? (conciliadas / linhas.length) * 100 : 0,
      totalContas: linhas.length,
      pessoal,
      rescisoes: rescisoesPorCompetencia.has(competencia) ? rescisoesPorCompetencia.get(competencia) : null,
      divergenciasNoUpload,
      divergenciasPendentes: divergenciasAtuais
    };
  }

  // -------------------------------------------------------------- gráficos

  function corTema(varName, fallback) {
    const v = getComputedStyle(document.documentElement).getPropertyValue(varName).trim();
    return v || fallback;
  }

  function destruirGraficos() {
    charts.forEach((c) => c.destroy());
    charts.length = 0;
  }

  function criarGrafico(canvasId, config) {
    const el = document.getElementById(canvasId);
    if (!el) return null;
    const chart = new window.Chart(el.getContext('2d'), config);
    charts.push(chart);
    return chart;
  }

  const OPCOES_BASE = {
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: { legend: { display: false } }
  };

  function corTexto() { return corTema('--text-2', '#45526E'); }
  function corGrade() { return corTema('--border', '#E4D8BE'); }

  function escalasPadrao(extra) {
    const cor = corTexto(), grade = corGrade();
    return Object.assign({
      x: { ticks: { color: cor, font: { size: 11 } }, grid: { display: false } },
      y: { ticks: { color: cor, font: { size: 11 } }, grid: { color: grade } }
    }, extra || {});
  }

  // ------------------------------------------------------------------ render

  function render(ctx) {
    const { state, tipoDivergencia, money, normalize, colIdxToLetter, escapeHtml } = ctx;
    destruirGraficos();

    const competencias = [...state.competencias].sort((a, b) => state.competenciaDatas.get(a) - state.competenciaDatas.get(b));
    const rescisoesPorCompetencia = contarRescisoesPorCompetencia(state, normalize, colIdxToLetter);
    const kpisPorMes = competencias.map((c) => computeKpisMes(state, c, tipoDivergencia, normalize, rescisoesPorCompetencia));

    const host = document.getElementById('custos-conteudo');
    if (!kpisPorMes.length) { host.innerHTML = '<p class="vazio">Sem dados suficientes pra montar o painel.</p>'; return; }

    const ultima = kpisPorMes[kpisPorMes.length - 1];
    const penultima = kpisPorMes.length > 1 ? kpisPorMes[kpisPorMes.length - 2] : null;
    const variacaoValor = penultima ? ultima.custoContabil - penultima.custoContabil : null;
    const variacaoPct = penultima && penultima.custoContabil ? (variacaoValor / penultima.custoContabil) * 100 : null;
    const custoPessoalTotal = ultima.pessoal.Salário + ultima.pessoal.Encargos + ultima.pessoal.Benefícios;
    const pctDoCusto = ultima.custoContabil ? (ultima.diferencaAbsoluta / Math.abs(ultima.custoContabil)) * 100 : 0;
    const pendentesMelhorou = ultima.divergenciasPendentes <= ultima.divergenciasNoUpload;

    host.innerHTML = `
      <section class="card">
        <h2>Painel de Custos e Conciliação</h2>
        <p class="card-note">Competência de referência: <strong>${escapeHtml(ultima.competencia)}</strong> — evolução calculada em cima das mesmas contas recalculadas da aba Conciliação, nada refeito aqui.</p>
        <div class="exec-summary">
          O painel separa <strong>custo</strong>, <strong>cobertura da conciliação</strong> e <strong>exposição a divergências</strong>.
          Custo de Pessoal, Encargos e Benefícios são classificados por palavra-chave na descrição da conta (sem fonte pronta com essa quebra no arquivo) — confira antes de usar os números.
          <strong>Divergências no Upload</strong> é a contagem congelada neste navegador na primeira vez que cada competência foi vista; ajustes posteriores só reduzem o indicador de <strong>Pendentes</strong>.
          <button type="button" class="btn-ghost btn-small" id="btn-zerar-baseline" style="margin-left:8px;">Zerar baseline de divergências</button>
        </div>
        <div class="kpi-exec-grid">
          ${kpiExecCard('Custo Contábil', moneyAbrev(ultima.custoContabil, money), [
            penultima ? { texto: `Mês anterior ${moneyAbrev(penultima.custoContabil, money)}` } : null
          ])}
          ${kpiExecCard('Variação do Custo', variacaoValor === null ? '—' : moneyAbrev(variacaoValor, money), [
            variacaoPct === null ? null : { texto: `Variação % ${pctFmt(variacaoPct)}`, tom: variacaoValor >= 0 ? 'up' : 'down' }
          ])}
          ${kpiExecCard('% Conciliado', `${ultima.percConciliado.toFixed(1).replace('.', ',')}%`, [
            { texto: `${ultima.totalContas} conta(s) na competência` }
          ])}
          ${kpiExecCard('Custo de Pessoal', moneyAbrev(custoPessoalTotal, money), [
            { texto: 'Salário + Encargos + Benefícios' }
          ])}
          ${kpiExecCard('Diferença Líquida', moneyAbrev(ultima.diferencaLiquida, money), [
            { texto: 'Saldo entre Folha e Contábil', tom: ultima.diferencaLiquida >= 0 ? 'up' : 'down' }
          ])}
          ${kpiExecCard('Diferença Absoluta', moneyAbrev(ultima.diferencaAbsoluta, money), [
            { texto: `% do custo ${pctDoCusto.toFixed(1).replace('.', ',')}%` }
          ])}
          ${kpiExecCard('Divergências no Upload', String(ultima.divergenciasNoUpload), [
            { texto: `Pendentes ${ultima.divergenciasPendentes}`, tom: pendentesMelhorou ? 'up' : 'down' }
          ])}
        </div>
      </section>

      <section class="card">
        <div class="chart-box">
          <h3>Custo contábil e custo de pessoal por competência</h3>
          <p class="chart-note">Ambos recalculados direto da fonte primária — Custo de Pessoal é a soma de Salário + Encargos + Benefícios classificados por palavra-chave.</p>
          <div class="chart-canvas-wrap"><canvas id="chart-custo-linha"></canvas></div>
        </div>
      </section>

      <section class="card">
        <div class="chart-grid">
          <div class="chart-box">
            <h3>Composição do custo de pessoal</h3>
            <p class="chart-note">Salário, encargos e benefícios empilhados por competência.</p>
            <div class="chart-canvas-wrap"><canvas id="chart-composicao"></canvas></div>
          </div>
          <div class="chart-box">
            <h3>% conciliado por competência</h3>
            <p class="chart-note">Contas com Folha e Contábil confirmados direto na fonte, sobre o total da competência.</p>
            <div class="chart-canvas-wrap"><canvas id="chart-conciliado"></canvas></div>
          </div>
        </div>
      </section>

      <section class="card">
        <div class="chart-grid">
          <div class="chart-box">
            <h3>Variação mensal do custo contábil</h3>
            <p class="chart-note">Mudança em R$ frente ao mês anterior.</p>
            <div class="chart-canvas-wrap"><canvas id="chart-variacao"></canvas></div>
          </div>
          <div class="chart-box">
            <h3>Diferença líquida entre folha e contabilidade</h3>
            <p class="chart-note">Saldo assinado (Folha − Contábil) por competência.</p>
            <div class="chart-canvas-wrap"><canvas id="chart-dif-liquida"></canvas></div>
          </div>
        </div>
      </section>

      <section class="card">
        <div class="chart-grid">
          <div class="chart-box">
            <h3>Exposição total das divergências</h3>
            <p class="chart-note">Soma das diferenças financeiras confirmadas (Diferença Absoluta) por competência.</p>
            <div class="chart-canvas-wrap"><canvas id="chart-exposicao"></canvas></div>
          </div>
          <div class="chart-box">
            <h3>Divergências identificadas e pendentes</h3>
            <p class="chart-note">Identificadas = baseline congelado no upload; Pendentes = contagem atual.</p>
            <div class="chart-canvas-wrap"><canvas id="chart-pendencias"></canvas></div>
          </div>
        </div>
      </section>

      <section class="card">
        <div class="chart-box">
          <h3>Rescisões por competência</h3>
          <p class="chart-note">Pessoas distintas na aba "Rescisoes" do arquivo, por mês. Competências sem essa aba não aparecem na contagem.</p>
          <div class="chart-canvas-wrap"><canvas id="chart-rescisoes"></canvas></div>
        </div>
      </section>

      <section class="card">
        <h2>Detalhamento mensal dos KPIs</h2>
        <p class="card-note">Mais recente primeiro.</p>
        <div id="tabela-custos"></div>
      </section>`;

    const btnZerar = document.getElementById('btn-zerar-baseline');
    if (btnZerar) {
      btnZerar.addEventListener('click', () => {
        zerarBaselines();
        render(ctx);
      });
    }

    renderGraficos(kpisPorMes, money);
    renderTabelaCustos(kpisPorMes, money);
  }

  function kpiExecCard(label, valor, pills) {
    const pillsHtml = (pills || []).filter(Boolean).map((p) =>
      `<span class="pill${p.tom === 'up' ? ' is-up' : p.tom === 'down' ? ' is-down' : ''}">${p.texto}</span>`
    ).join('');
    return `<div class="kpi-exec">
      <div class="label-row">
        <span class="label">${label}</span>
        <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 16v-5M12 8h.01"/></svg>
      </div>
      <div class="value">${valor}</div>
      ${pillsHtml ? `<div class="pill-row">${pillsHtml}</div>` : ''}
    </div>`;
  }

  function renderGraficos(kpisPorMes, money) {
    const labels = kpisPorMes.map((k) => k.competencia);
    const corAccent = corTema('--accent-strong', '#9C7418');
    const corFlag = corTema('--flag', '#2E6660');
    const corWarn = corTema('--warn', '#B15C1D');
    const corOk = corTema('--ok', '#4E7D52');
    const corCritical = corTema('--critical', '#A23B2E');
    const corNeutral = corTema('--neutral', '#6B6455');

    criarGrafico('chart-custo-linha', {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'Custo Contábil', data: kpisPorMes.map((k) => k.custoContabil), borderColor: corAccent, backgroundColor: `${corAccent}22`, fill: false, tension: 0.3, pointRadius: 3 },
          { label: 'Custo de Pessoal', data: kpisPorMes.map((k) => k.pessoal.Salário + k.pessoal.Encargos + k.pessoal.Benefícios), borderColor: corFlag, backgroundColor: `${corFlag}22`, fill: false, tension: 0.3, pointRadius: 3 }
        ]
      },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { legend: { display: true, position: 'bottom', labels: { color: corTexto() } }, tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${money.format(c.parsed.y)}` } } },
        scales: escalasPadrao({ y: { ticks: { color: corTexto(), callback: (v) => moneyAbrev(v, money) }, grid: { color: corGrade() } } })
      })
    });

    criarGrafico('chart-composicao', {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Salário', data: kpisPorMes.map((k) => k.pessoal.Salário), backgroundColor: corAccent },
          { label: 'Encargos', data: kpisPorMes.map((k) => k.pessoal.Encargos), backgroundColor: corFlag },
          { label: 'Benefícios', data: kpisPorMes.map((k) => k.pessoal.Benefícios), backgroundColor: corWarn }
        ]
      },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { legend: { display: true, position: 'bottom', labels: { color: corTexto() } }, tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${money.format(c.parsed.y)}` } } },
        scales: escalasPadrao({
          x: { stacked: true, ticks: { color: corTexto(), font: { size: 11 } }, grid: { display: false } },
          y: { stacked: true, ticks: { color: corTexto(), callback: (v) => moneyAbrev(v, money) }, grid: { color: corGrade() } }
        })
      })
    });

    criarGrafico('chart-conciliado', {
      type: 'line',
      data: { labels, datasets: [{ label: '% Conciliado', data: kpisPorMes.map((k) => k.percConciliado), borderColor: corOk, backgroundColor: `${corOk}22`, fill: true, tension: 0.3, pointRadius: 3 }] },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { tooltip: { callbacks: { label: (c) => `% Conciliado: ${c.parsed.y.toFixed(1).replace('.', ',')}%` } } },
        scales: escalasPadrao({ y: { min: 0, max: 100, ticks: { color: corTexto(), callback: (v) => `${v}%` }, grid: { color: corGrade() } } })
      })
    });

    criarGrafico('chart-variacao', {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          label: 'Variação R$',
          data: kpisPorMes.map((k, i) => i === 0 ? null : k.custoContabil - kpisPorMes[i - 1].custoContabil),
          backgroundColor: (c) => (c.raw >= 0 ? corOk : corCritical)
        }]
      },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { tooltip: { callbacks: { label: (c) => `Variação: ${moneyAbrev(c.parsed.y, money)}` } } },
        scales: escalasPadrao({ y: { ticks: { color: corTexto(), callback: (v) => moneyAbrev(v, money) }, grid: { color: corGrade() } } })
      })
    });

    criarGrafico('chart-dif-liquida', {
      type: 'bar',
      data: {
        labels,
        datasets: [{ label: 'Diferença Líquida', data: kpisPorMes.map((k) => k.diferencaLiquida), backgroundColor: (c) => (c.raw >= 0 ? corOk : corCritical) }]
      },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { tooltip: { callbacks: { label: (c) => `Diferença Líquida: ${moneyAbrev(c.parsed.y, money)}` } } },
        scales: escalasPadrao({ y: { ticks: { color: corTexto(), callback: (v) => moneyAbrev(v, money) }, grid: { color: corGrade() } } })
      })
    });

    criarGrafico('chart-exposicao', {
      type: 'line',
      data: { labels, datasets: [{ label: 'Exposição', data: kpisPorMes.map((k) => k.diferencaAbsoluta), borderColor: corCritical, backgroundColor: `${corCritical}22`, fill: true, tension: 0.3, pointRadius: 3 }] },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { tooltip: { callbacks: { label: (c) => `Exposição: ${moneyAbrev(c.parsed.y, money)}` } } },
        scales: escalasPadrao({ y: { ticks: { color: corTexto(), callback: (v) => moneyAbrev(v, money) }, grid: { color: corGrade() } } })
      })
    });

    criarGrafico('chart-pendencias', {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'Identificadas', data: kpisPorMes.map((k) => k.divergenciasNoUpload), borderColor: corNeutral, backgroundColor: `${corNeutral}22`, fill: false, tension: 0.25, pointRadius: 3 },
          { label: 'Pendentes', data: kpisPorMes.map((k) => k.divergenciasPendentes), borderColor: corCritical, backgroundColor: `${corCritical}22`, fill: false, tension: 0.25, pointRadius: 3 }
        ]
      },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { legend: { display: true, position: 'bottom', labels: { color: corTexto() } } },
        scales: escalasPadrao({ y: { ticks: { color: corTexto(), precision: 0 }, grid: { color: corGrade() } } })
      })
    });

    criarGrafico('chart-rescisoes', {
      type: 'bar',
      data: { labels, datasets: [{ label: 'Rescisões', data: kpisPorMes.map((k) => k.rescisoes ?? 0), backgroundColor: corFlag }] },
      options: Object.assign({}, OPCOES_BASE, {
        plugins: { tooltip: { callbacks: { label: (c) => `Rescisões: ${c.parsed.y}` } } },
        scales: escalasPadrao({ y: { ticks: { color: corTexto(), precision: 0 }, grid: { color: corGrade() } } })
      })
    });
  }

  function renderTabelaCustos(kpisPorMes, money) {
    const host = document.getElementById('tabela-custos');
    const linhas = [...kpisPorMes].reverse();
    const linhasHtml = linhas.map((k, i) => {
      const anterior = linhas[i + 1];
      const variacaoValor = anterior ? k.custoContabil - anterior.custoContabil : null;
      const variacaoPct = anterior && anterior.custoContabil ? (variacaoValor / anterior.custoContabil) * 100 : null;
      const custoPessoal = k.pessoal.Salário + k.pessoal.Encargos + k.pessoal.Benefícios;
      const claseVariacao = variacaoValor === null ? '' : (variacaoValor >= 0 ? 'delta-pos' : 'delta-neg');
      const claseDifLiquida = k.diferencaLiquida >= 0 ? 'delta-pos' : 'delta-neg';
      return `<tr>
        <td>${k.competencia}</td>
        <td class="num">${moneyAbrev(k.custoContabil, money)}</td>
        <td class="num ${claseVariacao}">${variacaoValor === null ? '—' : moneyAbrev(variacaoValor, money)}</td>
        <td class="num ${claseVariacao}">${variacaoPct === null ? '—' : pctFmt(variacaoPct)}</td>
        <td class="num">${k.percConciliado.toFixed(1).replace('.', ',')}%</td>
        <td class="num">${moneyAbrev(custoPessoal, money)}</td>
        <td class="num">${moneyAbrev(k.pessoal.Benefícios, money)}</td>
        <td class="num">${moneyAbrev(k.pessoal.Encargos, money)}</td>
        <td class="num">${k.rescisoes === null ? '—' : k.rescisoes}</td>
        <td class="num ${claseDifLiquida}">${moneyAbrev(k.diferencaLiquida, money)}</td>
        <td class="num">${moneyAbrev(k.diferencaAbsoluta, money)}</td>
        <td class="num">${k.divergenciasNoUpload}</td>
        <td class="num">${k.divergenciasPendentes}</td>
      </tr>`;
    }).join('');

    host.innerHTML = `<div class="table-wrap">
      <table>
        <thead><tr>
          <th>Competência</th><th class="num">Custo Contábil</th><th class="num">Variação R$</th><th class="num">Variação %</th>
          <th class="num">% Conciliado</th><th class="num">Custo de Pessoal</th><th class="num">Benefícios</th><th class="num">Encargos</th>
          <th class="num">Rescisões</th><th class="num">Dif. Líquida</th><th class="num">Dif. Absoluta</th>
          <th class="num">Divergências no Upload</th><th class="num">Pendentes</th>
        </tr></thead>
        <tbody>${linhasHtml}</tbody>
      </table>
    </div>`;
  }

  window.PainelCustos = { render };
})();
