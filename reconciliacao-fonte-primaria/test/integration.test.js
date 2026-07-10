'use strict';
/* Teste de integração: carrega a fixture sintética (test/fixtures/) via
 * ExcelJS de verdade e roda engine.js sobre ela — diferente dos testes de
 * unidade (que usam um adapter em memória escrito à mão), isso exercita o
 * parser de fórmula e o adapter contra um arquivo .xlsx real, construído de
 * forma independente do arquivo real do cliente (não versionado aqui).
 *
 * Não testa a detecção de abas/blocos do app.js (isso é DOM-coupled — sem
 * navegador, validado manualmente com Playwright, ver README) — testa que
 * engine.js interpreta corretamente fórmulas, referências e proveniência
 * lidas de um arquivo .xlsx de verdade, com uma estrutura diferente da
 * fixture sintética dos testes de unidade (família de conta, layout de
 * cabeçalho, hierarquia de contas).
 *
 * Requer a devDependency `exceljs` (npm install) — a mesma biblioteca que
 * o app.js usa no navegador (ver vendor/exceljs.min.js).
 */
const assert = require('node:assert/strict');
const path = require('node:path');
const ExcelJS = require('exceljs');
const engine = require('../engine.js');

let passed = 0;
let failed = 0;
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  ok — ${name}`);
  } catch (err) {
    failed++;
    console.error(`  FAIL — ${name}`);
    console.error(`         ${err.message}`);
  }
}

// Mesmo adapter que app.js monta sobre o ExcelJS no navegador.
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

(async () => {
  const fixturePath = path.join(__dirname, 'fixtures', 'fechamento-ficticio-teste.xlsx');
  const exceljsWorkbook = new ExcelJS.Workbook();
  await exceljsWorkbook.xlsx.readFile(fixturePath);
  const wb = normalizeWorkbook(exceljsWorkbook);
  const adapter = makeAdapter(wb);

  const CHECK = 'CHECK';
  // blocos conhecidos da fixture: Jan=C, Fev=G, Mar=K, Abr=O
  const blocks = {
    JANEIRO: { fopag: 'C', contabil: 'D', diferenca: 'E', balancete: null },
    FEVEREIRO: { fopag: 'G', contabil: 'H', diferenca: 'I', balancete: 'BALANCETE 02.2026' },
    MARÇO: { fopag: 'K', contabil: 'L', diferenca: 'M', balancete: 'Balancete 03-2026' },
    ABRIL: { fopag: 'O', contabil: 'P', diferenca: 'Q', balancete: 'BALANCETE 04.2026' }
  };
  const contas = [
    { codigo: '4.2.01.001.001', descricao: 'SALARIO BASE TESTE', row: 5 },
    { codigo: '4.2.01.001.002', descricao: 'HORAS EXTRAS TESTE', row: 6 },
    { codigo: '4.2.01.001.003', descricao: 'BENEFICIOS TESTE', row: 7 },
    { codigo: '4.2.01.004.001', descricao: 'ENCARGOS TESTE', row: 8 }
  ];

  function reconcile(mes, contaCodigo) {
    const block = blocks[mes];
    const conta = contas.find((c) => c.codigo === contaCodigo);
    return engine.reconcileConta({
      adapter, checkSheet: CHECK, contaCodigo: conta.codigo, descricao: conta.descricao, competencia: mes,
      fopagCell: adapter.cell(CHECK, block.fopag, conta.row),
      contabilCell: adapter.cell(CHECK, block.contabil, conta.row),
      diferencaCell: adapter.cell(CHECK, block.diferenca, conta.row),
      balanceteSheet: block.balancete
    });
  }

  console.log('sheets detectadas na fixture:', wb.SheetNames.join(', '));

  console.log('\nJaneiro — padrão "valor digitado" (família de conta diferente da fixture dos testes de unidade)');
  test('CONTABIL e DIFERENÇA digitados geram alerta de preenchimento', () => {
    const conta = reconcile('JANEIRO', '4.2.01.001.001');
    assert.equal(conta.proveniencia.contabil, 'valor-digitado');
    assert.ok(conta.alertas.includes('preenchimento'));
  });

  console.log('\nMarço — referência quebrada, com BALANCETE de nomenclatura alternativa ("Balancete 03-2026")');
  test('fórmula quebrada é detectada mesmo com separador de mês por hífen na aba', () => {
    const conta = reconcile('MARÇO', '4.2.01.001.001');
    assert.equal(conta.proveniencia.contabil, 'formula-quebrada');
    assert.ok(conta.alertas.includes('preenchimento'));
  });
  test('recálculo por código de conta encontra o BALANCETE de nomenclatura alternativa e não acha divergência real', () => {
    const conta = reconcile('MARÇO', '4.2.01.001.001');
    assert.equal(conta.recalculado.origemContabil, 'balancete-por-codigo');
    assert.ok(Math.abs(conta.recalculado.diferenca) < 0.01, 'a fonte real não tem divergência, só a fórmula da CHECK está quebrada');
  });

  console.log('\nAbril — ajuste do RAZÃO embutido na fórmula + rótulo de rubrica com cabeçalho fora da linha 1');
  test('ajuste do RAZÃO é reconhecido e detalhado', () => {
    const conta = reconcile('ABRIL', '4.2.01.004.001');
    assert.equal(conta.proveniencia.contabil, 'formula-viva');
    assert.ok(conta.recalculado.ajusteRazao, 'deveria reportar o ajuste');
    assert.equal(conta.recalculado.ajusteRazao.termos.length, 2);
  });
  test('rótulo da rubrica no drill-down usa o nome real, não a letra da coluna (cabeçalho na linha 2 da FOPAG TESTE)', () => {
    const conta = reconcile('ABRIL', '4.2.01.004.001');
    assert.equal(conta.recalculado.fopagTermos[0].rubrica, 'ENCARGOS');
  });

  console.log('\nFevereiro — cenário limpo com uma divergência financeira real de propósito');
  test('contas sem problema de preenchimento não disparam alerta', () => {
    const conta = reconcile('FEVEREIRO', '4.2.01.001.001');
    assert.deepEqual(conta.alertas, []);
  });
  test('a única divergência financeira real plantada é detectada, e só ela', () => {
    const encargos = reconcile('FEVEREIRO', '4.2.01.004.001');
    assert.ok(encargos.alertas.includes('financeiro'));
    assert.ok(Math.abs(encargos.recalculado.diferenca - (-250.75)) < 0.01);
  });

  console.log('\nCobertura — hierarquia pai/filho no BALANCETE (família de conta 4.2.01, diferente da real)');
  test('conta-pai não entra como fantasma, conta-filha ausente da CHECK entra', () => {
    const codigosNaCheck = new Set(contas.map((c) => c.codigo));
    const semProfundidade = engine.computeCobertura(adapter, 'BALANCETE 04.2026', codigosNaCheck, { prefixo: '4.2.01.' });
    const comProfundidade = engine.computeCobertura(adapter, 'BALANCETE 04.2026', codigosNaCheck, { prefixo: '4.2.01.', profundidade: 5 });
    assert.ok(semProfundidade.contasFantasmas.some((f) => f.codigo === '4.2.01.001'), 'sem filtro, o pai aparece (comportamento antigo)');
    assert.ok(!comProfundidade.contasFantasmas.some((f) => f.codigo === '4.2.01.001'), 'com filtro, o pai não deveria aparecer');
    assert.ok(comProfundidade.contasFantasmas.some((f) => f.codigo === '4.2.01.001.009'), 'a conta-folha de verdade deveria aparecer');
  });

  console.log(`\n${passed} passaram, ${failed} falharam`);
  process.exit(failed ? 1 : 0);
})();
