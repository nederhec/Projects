/* ==========================================================================
   Motor de Reconciliação Independente — engine.js
   Projeto novo, independente do painel de conciliação existente.

   Ideia central: a aba CHECK deixa de ser lida como fonte de verdade e
   passa a ser um dado a mais a reconciliar. Este módulo:
     1. classifica a proveniência de cada célula relevante da CHECK
        (fórmula viva / fórmula quebrada / valor digitado);
     2. recalcula Folha somando a FOPAG diretamente, por conta;
     3. recalcula Contábil somando o BALANCETE diretamente, por conta;
     4. cruza os três valores (declarado, recalc. Folha, recalc. Contábil)
        e separa divergência financeira real de divergência de preenchimento;
     5. produz um Score de Confiabilidade da Base e a lista de contas
        presentes nas fontes primárias mas ausentes da CHECK.

   Não depende de nenhuma biblioteca de planilha diretamente — recebe um
   WorkbookAdapter (ver JSDoc abaixo) para poder ser testado em Node com
   fixtures sintéticas e reutilizado no navegador sobre o ExcelJS.
   ========================================================================== */

/**
 * @typedef {Object} CellRaw
 * @property {*} value       valor calculado/cacheado da célula (número, texto, Date)
 * @property {string} [formula]  texto da fórmula, sem o "=", se houver
 */

/**
 * WorkbookAdapter — contrato mínimo que este módulo precisa da planilha.
 * @typedef {Object} WorkbookAdapter
 * @property {string[]} sheetNames
 * @property {(sheet: string, col: string, row: number) => (CellRaw|undefined)} cell
 *   Retorna undefined quando a célula nunca foi preenchida (distinção
 *   importante: 0 digitado != célula nunca preenchida).
 * @property {(sheet: string) => number} usedRowCount
 */

(function (root, factory) {
  const mod = factory();
  if (typeof module === 'object' && module.exports) module.exports = mod;
  else root.ReconEngine = mod;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const TOLERANCIA = 0.005; // mesma régua de zero-tolerância do painel original

  // ------------------------------------------------------------------ utils

  function normalize(s) {
    return String(s ?? '')
      .toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function isEmptyCell(cellRaw) {
    return cellRaw === undefined || cellRaw === null ||
      cellRaw.value === undefined || cellRaw.value === null || cellRaw.value === '';
  }

  function toNumber(value) {
    if (typeof value === 'number') return value;
    if (value instanceof Date) return value.getTime();
    if (typeof value === 'string') {
      const n = Number(value.replace(/\./g, '').replace(',', '.'));
      return Number.isFinite(n) ? n : NaN;
    }
    return NaN;
  }

  function sameCriteria(a, b) {
    if (a instanceof Date || b instanceof Date) {
      const da = a instanceof Date ? a : new Date(a);
      const db = b instanceof Date ? b : new Date(b);
      return da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth() && da.getDate() === db.getDate();
    }
    return normalize(a) === normalize(b);
  }

  /** Extrai o código de conta contábil de uma descrição, ex.:
   * "SALARIOS E ORDENADOS (3.1.01.001.001)" -> "3.1.01.001.001" */
  function parseAccountCode(descricao) {
    const match = String(descricao || '').match(/\(([\d]+(?:\.[\d]+)+)\)/);
    return match ? match[1] : '';
  }

  function parseRef(ref) {
    const m = String(ref || '').match(/\$?([A-Z]+)\$?(\d+)/);
    return m ? { col: m[1], row: Number(m[2]) } : null;
  }

  // -------------------------------------------------------- parsing de fórmula

  const SUMIFS_RE = /SUMIFS\(\s*'([^']+)'!\$?([A-Z]+)\$?\d+:\$?[A-Z]+\$?\d+\s*,\s*'\1'!\$?([A-Z]+)\$?\d+:\$?[A-Z]+\$?\d+\s*,\s*(\$?[A-Z]+\$?\d+)\s*\)/g;
  const DIRECT_REF_RE = /'([^']+)'!\$?([A-Z]+)\$?(\d+)/g;

  /**
   * Classifica sintaticamente uma fórmula (sem tocar a planilha ainda).
   * @param {string|undefined} formula
   */
  function parseFormula(formula) {
    if (!formula) return { kind: 'literal' };

    const sumifsTerms = [];
    let m;
    SUMIFS_RE.lastIndex = 0;
    while ((m = SUMIFS_RE.exec(formula))) {
      sumifsTerms.push({ sheet: m[1], valueCol: m[2], criteriaCol: m[3], criteriaRef: m[4].replace(/\$/g, '') });
    }
    if (sumifsTerms.length) return { kind: 'sumifs', terms: sumifsTerms };

    const refTerms = [];
    DIRECT_REF_RE.lastIndex = 0;
    while ((m = DIRECT_REF_RE.exec(formula))) {
      refTerms.push({ sheet: m[1], col: m[2], row: Number(m[3]) });
    }
    if (refTerms.length) return { kind: 'directRef', terms: refTerms };

    // fórmula intra-aba (ex.: =K5-L5, =SUM(C5:C8)) — nada externo a validar.
    return { kind: 'intra' };
  }

  // ------------------------------------------------------------ recálculo

  /** Soma uma fonte via SUMIFS reimplementado sobre o adapter. Devolve também
   *  o valor por termo (por rubrica/coluna), pra alimentar o drill-down. */
  function evaluateSumifsTerms(adapter, terms, criteriaSheet) {
    let total = 0;
    let resolvable = true;
    const termos = [];
    for (const term of terms) {
      if (!adapter.sheetNames.includes(term.sheet)) { resolvable = false; continue; }
      const criteriaRefParsed = parseRef(term.criteriaRef);
      const criteriaCell = criteriaRefParsed && adapter.cell(criteriaSheet, criteriaRefParsed.col, criteriaRefParsed.row);
      if (isEmptyCell(criteriaCell)) { resolvable = false; continue; }
      const rows = adapter.usedRowCount(term.sheet);
      let subtotal = 0;
      for (let r = 1; r <= rows; r++) {
        const crit = adapter.cell(term.sheet, term.criteriaCol, r);
        if (isEmptyCell(crit) || !sameCriteria(crit.value, criteriaCell.value)) continue;
        const val = adapter.cell(term.sheet, term.valueCol, r);
        if (!isEmptyCell(val)) subtotal += toNumber(val.value) || 0;
      }
      termos.push({ sheet: term.sheet, coluna: term.valueCol, rubrica: findColumnLabel(adapter, term.sheet, term.valueCol) || term.valueCol, value: subtotal });
      total += subtotal;
    }
    return { value: total, resolvable, termos };
  }

  /** Acha o rótulo de uma coluna pro drill-down sem assumir que o cabeçalho
   *  está na linha 1 — planilhas reais têm o cabeçalho em linhas diferentes
   *  (achado real: a FOPAG 2026 do arquivo do cliente tem cabeçalho na
   *  linha 3; sem isso, o drill-down mostrava letra de coluna em vez do
   *  nome da rubrica mesmo nas contas corretas). Cabeçalho é texto; linhas
   *  de dado acima dele (título, célula solta) tendem a ser vazias ou
   *  numéricas — por isso pega a primeira célula de texto encontrada.
   */
  function findColumnLabel(adapter, sheet, col, maxRow) {
    for (let r = 1; r <= (maxRow || 8); r++) {
      const cell = adapter.cell(sheet, col, r);
      if (cell && typeof cell.value === 'string' && cell.value.trim()) return cell.value.trim();
    }
    return null;
  }

  /** Confere se as referências diretas de uma fórmula apontam para células
   *  que realmente existem (distingue "0 real" de "célula nunca preenchida"). */
  function diagnoseDirectRefs(adapter, terms) {
    const broken = [];
    for (const term of terms) {
      if (!adapter.sheetNames.includes(term.sheet)) {
        broken.push({ ...term, motivo: 'aba inexistente' });
        continue;
      }
      const target = adapter.cell(term.sheet, term.col, term.row);
      if (isEmptyCell(target)) broken.push({ ...term, motivo: 'célula vazia / nunca preenchida' });
    }
    return { broken, ok: broken.length === 0 };
  }

  /**
   * Soma os termos de uma fórmula de referência direta lendo o valor de
   * verdade de cada célula (não o cache da célula que contém a fórmula).
   * Isso é o que permite reconhecer ajustes manuais do RAZÃO que a própria
   * fórmula da CHECK já embute (ex.: BALANCETE!J160 + RAZÃO!G700 + RAZÃO!G702)
   * em vez de descartá-los — a mesma lógica que já aplicamos ao FOPAG via
   * SUMIFS, agora aplicada às referências diretas do lado Contábil.
   */
  function evaluateDirectRefTerms(adapter, terms) {
    const diag = diagnoseDirectRefs(adapter, terms);
    if (!diag.ok) return { value: null, resolvable: false, termos: [], broken: diag.broken };
    const termos = terms.map((t) => {
      const cell = adapter.cell(t.sheet, t.col, t.row);
      return { ...t, value: toNumber(cell.value) || 0 };
    });
    const total = termos.reduce((sum, t) => sum + t.value, 0);
    return { value: total, resolvable: true, termos, broken: [] };
  }

  /**
   * Classifica a proveniência de uma célula da CHECK (FOPAG, CONTABIL ou
   * DIFERENÇA) em 'formula-viva' | 'formula-quebrada' | 'valor-digitado' | 'intra'.
   */
  function classifyProvenance(adapter, cellRaw) {
    const parsed = parseFormula(cellRaw && cellRaw.formula);
    if (parsed.kind === 'literal') return { status: 'valor-digitado', parsed };
    if (parsed.kind === 'intra') return { status: 'formula-viva', parsed };
    if (parsed.kind === 'sumifs') {
      // proveniência sintática apenas; "quebrada" para SUMIFS é auferida
      // indiretamente pelo recálculo (ver recalcFopag) — aqui já sinalizamos
      // aba inexistente, que é detectável sem varrer linhas.
      const missingSheet = parsed.terms.some((t) => !adapter.sheetNames.includes(t.sheet));
      return { status: missingSheet ? 'formula-quebrada' : 'formula-viva', parsed };
    }
    if (parsed.kind === 'directRef') {
      const diag = diagnoseDirectRefs(adapter, parsed.terms);
      return { status: diag.ok ? 'formula-viva' : 'formula-quebrada', parsed, diag };
    }
    return { status: 'formula-viva', parsed };
  }

  /** Recalcula o valor de Folha de uma célula CHECK a partir da FOPAG,
   *  ignorando o valor declarado — usa só a composição revelada pela própria
   *  fórmula (quais colunas da FOPAG somam aquela conta). */
  function recalcFopag(adapter, checkSheet, cellRaw) {
    const provenance = classifyProvenance(adapter, cellRaw);
    if (provenance.parsed.kind !== 'sumifs') {
      return { value: null, status: 'sem-verificacao', provenance: provenance.status, termos: [] };
    }
    const { value, resolvable, termos } = evaluateSumifsTerms(adapter, provenance.parsed.terms, checkSheet);
    return {
      value,
      status: resolvable ? 'verificado' : 'sem-verificacao',
      provenance: resolvable ? provenance.status : 'formula-quebrada',
      termos
    };
  }

  /** Busca uma conta diretamente no BALANCETE do mês (por código de conta) e
   *  devolve Débito − Crédito. Independente de para onde a fórmula da CHECK
   *  aponta — é a fonte de fallback quando não dá pra confiar na composição
   *  da própria fórmula (valor digitado, ou referência quebrada). */
  function lookupContaNoBalancete(adapter, balanceteSheet, contaCodigo, opts) {
    const cfg = Object.assign({ codigoCol: 'B', debitoCol: 'H', creditoCol: 'I' }, opts);
    if (!balanceteSheet || !adapter.sheetNames.includes(balanceteSheet)) {
      return { value: null, status: 'sem-fonte', row: null };
    }
    const rows = adapter.usedRowCount(balanceteSheet);
    for (let r = 1; r <= rows; r++) {
      const codeCell = adapter.cell(balanceteSheet, cfg.codigoCol, r);
      if (isEmptyCell(codeCell)) continue;
      if (String(codeCell.value).trim() !== contaCodigo) continue;
      const debito = adapter.cell(balanceteSheet, cfg.debitoCol, r);
      const credito = adapter.cell(balanceteSheet, cfg.creditoCol, r);
      const value = (isEmptyCell(debito) ? 0 : toNumber(debito.value) || 0) -
        (isEmptyCell(credito) ? 0 : toNumber(credito.value) || 0);
      return { value, status: 'verificado', row: r };
    }
    return { value: null, status: 'conta-nao-encontrada', row: null };
  }

  /**
   * Recalcula o valor Contábil de uma conta. Duas fontes, nesta ordem:
   *   1. Se a célula CONTABIL da CHECK for uma referência direta e todos os
   *      termos resolverem (aba e célula existem), soma os termos de
   *      verdade — isso reaproveita ajustes de reclassificação do RAZÃO que
   *      a própria fórmula já descreve (ex.: BALANCETE + 2 lançamentos do
   *      RAZÃO), sem depender do cache da célula.
   *   2. Senão (valor digitado, ou referência quebrada), cai no lookup
   *      direto da conta no BALANCETE por código — mesmo comportamento de
   *      antes, e o que continua pegando o caso Junho (referência quebrada).
   * Quando a fonte 1 tem mais de um termo, o excedente sobre o primeiro é
   * reportado como `ajuste` — não é escondido nem tratado como erro, é
   * exposto pra quem for revisar confirmar que a reclassificação é válida.
   */
  function recalcContabil(adapter, contabilCell, balanceteSheet, contaCodigo, opts) {
    const parsed = parseFormula(contabilCell && contabilCell.formula);
    if (parsed.kind === 'directRef') {
      const resolved = evaluateDirectRefTerms(adapter, parsed.terms);
      if (resolved.resolvable) {
        const ajuste = resolved.termos.length > 1
          ? { valor: resolved.value - resolved.termos[0].value, termos: resolved.termos.slice(1) }
          : null;
        return { value: resolved.value, status: 'verificado', origem: 'formula-composicao', ajuste, row: null };
      }
      // referência quebrada: cai no fallback por código de conta abaixo.
    }
    const base = lookupContaNoBalancete(adapter, balanceteSheet, contaCodigo, opts);
    return { ...base, origem: 'balancete-por-codigo', ajuste: null };
  }

  // --------------------------------------------------- reconciliação 3 vias

  /**
   * Reconcilia uma conta em uma competência: cruza o que a CHECK declara
   * com o que foi recalculado a partir da FOPAG e do BALANCETE, e classifica
   * em até dois alertas independentes: financeiro (Folha != Contábil de
   * verdade) e de preenchimento (CHECK != fonte, mesmo que por acaso bata).
   */
  function reconcileConta(params) {
    const { adapter, checkSheet, contaCodigo, descricao, competencia,
      fopagCell, contabilCell, diferencaCell, balanceteSheet } = params;

    const fopagDeclarado = fopagCell ? toNumber(fopagCell.value) || 0 : 0;
    const contabilDeclarado = contabilCell ? toNumber(contabilCell.value) || 0 : 0;
    const diferencaDeclarada = diferencaCell ? toNumber(diferencaCell.value) || 0 : (fopagDeclarado - contabilDeclarado);

    const fopagProv = classifyProvenance(adapter, fopagCell);
    const contabilProv = classifyProvenance(adapter, contabilCell);
    const diferencaProv = classifyProvenance(adapter, diferencaCell);

    const fopagRecalc = recalcFopag(adapter, checkSheet, fopagCell);
    const contabilRecalc = recalcContabil(adapter, contabilCell, balanceteSheet, contaCodigo);

    const temFonteCompleta = fopagRecalc.status === 'verificado' && contabilRecalc.status === 'verificado';
    const diferencaReal = temFonteCompleta ? (fopagRecalc.value - contabilRecalc.value) : null;

    const proveniencia = {
      fopag: fopagProv.status,
      contabil: contabilProv.status,
      diferenca: diferencaProv.status
    };

    const temProblemaPreenchimento =
      proveniencia.fopag === 'formula-quebrada' || proveniencia.fopag === 'valor-digitado' ||
      proveniencia.contabil === 'formula-quebrada' || proveniencia.contabil === 'valor-digitado' ||
      proveniencia.diferenca === 'valor-digitado';

    const alertas = [];
    if (temProblemaPreenchimento) alertas.push('preenchimento');
    if (temFonteCompleta && Math.abs(diferencaReal) >= TOLERANCIA) alertas.push('financeiro');
    // fallback: sem fonte pra verificar e a própria CHECK já declara diferença -> ainda financeiro, mas com selo "não verificado"
    if (!temFonteCompleta && !temProblemaPreenchimento && Math.abs(diferencaDeclarada) >= TOLERANCIA) alertas.push('financeiro-nao-verificado');

    return {
      codigo: contaCodigo,
      descricao,
      competencia,
      declarado: { fopag: fopagDeclarado, contabil: contabilDeclarado, diferenca: diferencaDeclarada },
      recalculado: {
        fopag: fopagRecalc.value,
        contabil: contabilRecalc.value,
        diferenca: diferencaReal,
        fonteContabil: balanceteSheet || null,
        statusFopag: fopagRecalc.status,
        statusContabil: contabilRecalc.status,
        origemContabil: contabilRecalc.origem || null,
        ajusteRazao: contabilRecalc.ajuste || null,
        fopagTermos: fopagRecalc.termos || []
      },
      proveniencia,
      alertas
    };
  }

  // ------------------------------------------------- score de confiabilidade

  /**
   * % (ponderado pelo valor financeiro da conta) de células cuja proveniência
   * é 'formula-viva' nas três colunas (FOPAG/CONTABIL/DIFERENÇA). Responde
   * "o resumo que eu li está certo?", separado do Score de Risco financeiro.
   */
  function computeConfiabilidade(contas) {
    if (!contas.length) return { score: null, celulasVivas: 0, celulasQuebradas: 0, celulasDigitadas: 0, contasAtivas: 0 };
    let pesoTotal = 0, pesoOk = 0, vivas = 0, quebradas = 0, digitadas = 0, ativas = 0;
    for (const conta of contas) {
      const peso = Math.max(Math.abs(conta.declarado.fopag), Math.abs(conta.declarado.contabil), 1);
      const estados = [conta.proveniencia.fopag, conta.proveniencia.contabil, conta.proveniencia.diferenca];
      estados.forEach((s) => {
        if (s === 'formula-viva' || s === 'intra') vivas++;
        else if (s === 'formula-quebrada') quebradas++;
        else if (s === 'valor-digitado') digitadas++;
      });
      const ok = estados.every((s) => s === 'formula-viva' || s === 'intra');
      if (conta.declarado.fopag !== 0 || conta.declarado.contabil !== 0) ativas++;
      pesoTotal += peso;
      pesoOk += ok ? peso : 0;
    }
    return {
      score: pesoTotal > 0 ? Math.round((pesoOk / pesoTotal) * 100) : 100,
      celulasVivas: vivas,
      celulasQuebradas: quebradas,
      celulasDigitadas: digitadas,
      contasAtivas: ativas
    };
  }

  // ------------------------------------------------------- cobertura (contas fantasmas)

  /**
   * Compara os códigos de conta presentes no BALANCETE (grupo de resultado,
   * prefixo configurável — "3." por padrão) contra os códigos já mapeados
   * na CHECK. Retorna as contas que existem na fonte primária e não têm
   * linha correspondente na CHECK.
   */
  /**
   * `profundidade` (opcional) restringe a contas com o mesmo número de
   * segmentos das já mapeadas na CHECK (ex.: 5, como em "3.1.01.001.001").
   * Sem isso, um código-pai/subtotal (ex.: "3.1.01.001", que soma as
   * próprias contas-filha que a CHECK já cobre individualmente) aparece
   * como se fosse uma conta "faltando" — não é: é só o agregado de contas
   * já conhecidas, listado de novo em outra granularidade.
   */
  function computeCobertura(adapter, balanceteSheet, codigosNaCheck, opts) {
    const cfg = Object.assign({ codigoCol: 'B', prefixo: '3.', profundidade: null }, opts);
    const fantasmas = [];
    if (!balanceteSheet || !adapter.sheetNames.includes(balanceteSheet)) return { contasFantasmas: fantasmas };
    const rows = adapter.usedRowCount(balanceteSheet);
    const vistos = new Set();
    for (let r = 1; r <= rows; r++) {
      const codeCell = adapter.cell(balanceteSheet, cfg.codigoCol, r);
      if (isEmptyCell(codeCell)) continue;
      const codigo = String(codeCell.value).trim();
      if (!codigo.startsWith(cfg.prefixo) || vistos.has(codigo)) continue;
      if (cfg.profundidade && codigo.split('.').length !== cfg.profundidade) continue;
      vistos.add(codigo);
      if (!codigosNaCheck.has(codigo)) fantasmas.push({ codigo, sheet: balanceteSheet, row: r });
    }
    return { contasFantasmas: fantasmas };
  }

  return {
    TOLERANCIA,
    parseAccountCode,
    parseRef,
    parseFormula,
    classifyProvenance,
    evaluateSumifsTerms,
    diagnoseDirectRefs,
    evaluateDirectRefTerms,
    recalcFopag,
    lookupContaNoBalancete,
    recalcContabil,
    reconcileConta,
    computeConfiabilidade,
    computeCobertura
  };
});
