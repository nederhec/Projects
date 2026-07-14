/* ==========================================================================
   report.js — Relatório Gerencial em PDF
   Consome os mesmos resultados já reconciliados pela tabela (state.resultados)
   — não recalcula nada por conta própria. Monta KPIs da competência mais
   recente e a evolução mês a mês, renderiza gráficos via Chart.js em canvas
   fora da tela e monta o PDF com jsPDF (ambos vendorizados em vendor/, sem
   CDN, mesmo princípio do ExcelJS).
   ========================================================================== */
(function () {
  'use strict';

  /** Classifica a descrição de uma conta de Pessoal em Salário/Encargos/
   *  Benefícios por palavra-chave. Não há fonte pronta com essa quebra no
   *  arquivo real (conferido: nem o BALANCETE, nem RESUMO FECHAMENTO FOPAG,
   *  nem a aba ENCARGOS trazem os três totais já separados) — por isso é uma
   *  heurística, não uma leitura direta. A legenda é impressa no próprio PDF
   *  pra quem for usar os números saber exatamente o critério. */
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
   *  competência (mês/ano) da coluna MÊS — devolve Map(rótulo da competência
   *  -> contagem). Sem essa aba no arquivo, devolve um Map vazio (KPI mostra
   *  "—" em vez de 0, pra não parecer "zero rescisões" quando na verdade é
   *  "sem como verificar"). */
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

  /** Agrega os KPIs de uma competência a partir das linhas já reconciliadas
   *  — mesma fonte que a tabela usa, nada recalculado aqui. */
  function computeKpisMes(state, competencia, tipoDivergencia, normalize, rescisoesPorCompetencia) {
    const linhas = state.resultados.filter((r) => r.competencia === competencia);
    let custoContabil = 0;
    let diferencaLiquida = 0;
    let diferencaAbsoluta = 0;
    let conciliadas = 0;
    let divergencias = 0;
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
      if (r.alertas.length > 0 || tipoDivergencia(r) !== 'Sem divergência') divergencias++;
    });

    return {
      competencia,
      custoContabil, diferencaLiquida, diferencaAbsoluta,
      percConciliado: linhas.length ? (conciliadas / linhas.length) * 100 : 0,
      divergencias,
      totalContas: linhas.length,
      pessoal,
      rescisoes: rescisoesPorCompetencia.has(competencia) ? rescisoesPorCompetencia.get(competencia) : null
    };
  }

  // -------------------------------------------------------------- gráficos

  function corTema(varName, fallback) {
    const v = getComputedStyle(document.documentElement).getPropertyValue(varName).trim();
    return v || fallback;
  }

  function criarCanvas(w, h) {
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    return canvas;
  }

  /** Fundo branco explícito via hook `beforeDraw` do próprio Chart.js — o
   *  Chart.js limpa o canvas (volta pra transparente) na própria
   *  inicialização, então preencher o canvas ANTES de criar o Chart não
   *  sobrevive; precisa entrar no ciclo de desenho da lib. Necessário porque
   *  JPEG (usado no embed do PDF por gerar arquivos bem menores que PNG pra
   *  esse tipo de imagem) não suporta transparência — sem isso, a área fora
   *  dos elementos do gráfico saía preta no PDF. */
  const PLUGIN_FUNDO_BRANCO = {
    id: 'fundoBranco',
    beforeDraw(chart) {
      const { ctx, width, height } = chart;
      ctx.save();
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, width, height);
      ctx.restore();
    }
  };

  async function aguardarPintura() {
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  }

  async function renderLinha(labels, dados, cor, titulo) {
    const canvas = criarCanvas(1000, 460);
    // eslint-disable-next-line no-new
    new window.Chart(canvas.getContext('2d'), {
      type: 'line',
      data: { labels, datasets: [{ label: titulo, data: dados, borderColor: cor, backgroundColor: `${cor}33`, fill: true, tension: 0.25, pointRadius: 4 }] },
      options: {
        responsive: false, animation: false, devicePixelRatio: 1,
        plugins: { legend: { display: false }, title: { display: true, text: titulo, font: { size: 16 } } },
        scales: { y: { beginAtZero: true } }
      },
      plugins: [PLUGIN_FUNDO_BRANCO]
    });
    await aguardarPintura();
    return canvas;
  }

  async function renderBarra(labels, dados, cor, titulo) {
    const canvas = criarCanvas(1000, 460);
    // eslint-disable-next-line no-new
    new window.Chart(canvas.getContext('2d'), {
      type: 'bar',
      data: { labels, datasets: [{ label: titulo, data: dados, backgroundColor: cor }] },
      options: {
        responsive: false, animation: false, devicePixelRatio: 1,
        plugins: { legend: { display: false }, title: { display: true, text: titulo, font: { size: 16 } } },
        scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
      },
      plugins: [PLUGIN_FUNDO_BRANCO]
    });
    await aguardarPintura();
    return canvas;
  }

  async function renderBarraEmpilhada(labels, kpisPorMes) {
    const canvas = criarCanvas(1000, 460);
    // eslint-disable-next-line no-new
    new window.Chart(canvas.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Salário', data: kpisPorMes.map((k) => k.pessoal.Salário), backgroundColor: corTema('--accent-strong', '#9C7418') },
          { label: 'Encargos', data: kpisPorMes.map((k) => k.pessoal.Encargos), backgroundColor: corTema('--flag', '#2E6660') },
          { label: 'Benefícios', data: kpisPorMes.map((k) => k.pessoal.Benefícios), backgroundColor: corTema('--warn', '#B15C1D') }
        ]
      },
      options: {
        responsive: false, animation: false, devicePixelRatio: 1,
        plugins: { title: { display: true, text: 'Custo de Pessoal por categoria', font: { size: 16 } } },
        scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }
      },
      plugins: [PLUGIN_FUNDO_BRANCO]
    });
    await aguardarPintura();
    return canvas;
  }

  // ------------------------------------------------------------------ PDF

  function desenhaKpiBox(doc, x, y, w, h, label, valor, sub) {
    doc.setDrawColor(203, 185, 141);
    doc.setFillColor(246, 236, 210);
    doc.roundedRect(x, y, w, h, 4, 4, 'FD');
    doc.setTextColor(110, 101, 80);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.text(label.toUpperCase(), x + 10, y + 15);
    doc.setTextColor(20, 33, 59);
    doc.setFontSize(14);
    doc.text(valor, x + 10, y + 33);
    if (sub) {
      doc.setTextColor(110, 101, 80);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7.5);
      doc.text(sub, x + 10, y + 45, { maxWidth: w - 20 });
    }
  }

  function montarPdf(competencias, kpisPorMes, graficos, ctx) {
    const { money } = ctx;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'pt', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 40;
    let y = margin;

    const ultima = kpisPorMes[kpisPorMes.length - 1];
    const penultima = kpisPorMes.length > 1 ? kpisPorMes[kpisPorMes.length - 2] : null;
    const variacaoValor = penultima ? ultima.custoContabil - penultima.custoContabil : null;
    const variacaoPct = penultima && penultima.custoContabil ? (variacaoValor / penultima.custoContabil) * 100 : null;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(17);
    doc.setTextColor(20, 33, 59);
    doc.text('Relatório Gerencial — Reconciliação Independente da Base', margin, y);
    y += 18;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor(110, 101, 80);
    doc.text(`Competência de referência: ${ultima.competencia} — gerado em ${new Date().toLocaleString('pt-BR')}`, margin, y);
    y += 22;

    const kpis = [
      ['Custo Contábil', money.format(ultima.custoContabil), 'Custo total reconhecido na competência'],
      ['Variação do Custo', variacaoValor === null ? '—' : `${money.format(variacaoValor)} (${variacaoPct >= 0 ? '+' : ''}${variacaoPct.toFixed(1)}%)`, 'Frente ao mês anterior'],
      ['% Conciliado', `${ultima.percConciliado.toFixed(1)}%`, `${ultima.totalContas} conta(s) na competência`],
      ['Custo de Pessoal', money.format(ultima.pessoal.Salário + ultima.pessoal.Encargos + ultima.pessoal.Benefícios), 'Salário + Encargos + Benefícios'],
      ['Benefícios', money.format(ultima.pessoal.Benefícios), 'Custo de benefícios aos colaboradores'],
      ['Encargos', money.format(ultima.pessoal.Encargos), 'Encargos patronais incidentes'],
      ['Rescisões', ultima.rescisoes === null ? '—' : String(ultima.rescisoes), ultima.rescisoes === null ? 'Aba de Rescisões não encontrada' : 'Pessoas desligadas na competência'],
      ['Diferença Líquida', money.format(ultima.diferencaLiquida), 'Saldo entre Folha e Contábil'],
      ['Diferença Absoluta', money.format(ultima.diferencaAbsoluta), 'Exposição total das divergências'],
      ['Divergências no Mês', String(ultima.divergencias), 'Contas com alerta, antes de qualquer ajuste']
    ];

    const cols = 2;
    const gap = 12;
    const boxW = (pageWidth - margin * 2 - gap * (cols - 1)) / cols;
    const boxH = 56;
    kpis.forEach(([label, valor, sub], i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = margin + col * (boxW + gap);
      const yy = y + row * (boxH + gap);
      desenhaKpiBox(doc, x, yy, boxW, boxH, label, valor, sub);
    });
    y += Math.ceil(kpis.length / cols) * (boxH + gap);

    y += 10;
    doc.setFont('helvetica', 'italic');
    doc.setFontSize(7.5);
    doc.setTextColor(110, 101, 80);
    doc.text(
      'Custo de Pessoal, Encargos e Benefícios são classificados por palavra-chave na descrição da conta contábil ' +
      '(Encargos: INSS/FGTS/PIS/GRRF; Benefícios: vale/assistência/seguro de vida/plano de saúde/auxílio; demais ' +
      'contas: Salário) — não há fonte pronta com essa quebra no arquivo. Confira a classificação antes de usar os números.',
      margin, y, { maxWidth: pageWidth - margin * 2 }
    );

    // ------------------------------------------------ páginas de gráficos
    const imgW = pageWidth - margin * 2;
    const imgH = imgW * (460 / 1000);
    graficos.forEach((canvas, i) => {
      if (i % 2 === 0) { doc.addPage(); y = margin; }
      doc.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', margin, y, imgW, imgH);
      y += imgH + 16;
    });

    const nomeArquivo = `relatorio-gerencial-${String(ultima.competencia).toLowerCase()}.pdf`;
    doc.save(nomeArquivo);
    return nomeArquivo;
  }

  async function gerar(ctx) {
    const { state, tipoDivergencia, normalize, colIdxToLetter } = ctx;
    if (!state.resultados.length) throw new Error('Nenhum resultado carregado.');

    const competencias = [...state.competencias].sort((a, b) => {
      const da = state.competenciaDatas.get(a);
      const db = state.competenciaDatas.get(b);
      return da - db;
    });

    const rescisoesPorCompetencia = contarRescisoesPorCompetencia(state, normalize, colIdxToLetter);
    const kpisPorMes = competencias.map((c) => computeKpisMes(state, c, tipoDivergencia, normalize, rescisoesPorCompetencia));

    const corAccent = corTema('--accent-strong', '#9C7418');
    const corCritical = corTema('--critical', '#A23B2E');
    const corFlag = corTema('--flag', '#2E6660');

    const graficos = [
      await renderLinha(competencias, kpisPorMes.map((k) => k.custoContabil), corAccent, 'Custo Contábil por competência'),
      await renderBarraEmpilhada(competencias, kpisPorMes),
      await renderBarra(competencias, kpisPorMes.map((k) => k.divergencias), corCritical, 'Divergências por competência'),
      await renderBarra(competencias, kpisPorMes.map((k) => k.rescisoes ?? 0), corFlag, 'Rescisões por competência')
    ];

    return montarPdf(competencias, kpisPorMes, graficos, ctx);
  }

  window.ReconReport = { gerar };
})();
