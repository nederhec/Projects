'use strict';
/* Testes de unidade do engine.js, com fixtures sintéticas — não depende de
 * nenhum arquivo .xlsx externo (o arquivo real do cliente usado para validar
 * o motor durante o desenvolvimento não é versionado neste repositório).
 * Cada fixture reproduz, de forma minimizada, um dos padrões encontrados no
 * arquivo real: valor digitado (Jan/Fev) e fórmula quebrada (Junho).
 */

const assert = require('node:assert/strict');
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

/** Adapter em memória para os testes — mesma interface que o app.js monta
 *  em cima do SheetJS no navegador. */
function makeAdapter(sheets) {
  return {
    sheetNames: Object.keys(sheets),
    cell(sheet, col, row) {
      const s = sheets[sheet];
      if (!s || !s[col] || !(row in s[col])) return undefined;
      return s[col][row];
    },
    usedRowCount(sheet) {
      const s = sheets[sheet];
      if (!s) return 0;
      let max = 0;
      for (const col of Object.values(s)) {
        for (const row of Object.keys(col)) max = Math.max(max, Number(row));
      }
      return max;
    }
  };
}

console.log('parseAccountCode');
test('extrai código entre parênteses', () => {
  assert.equal(engine.parseAccountCode('SALARIOS E ORDENADOS (3.1.01.001.001)'), '3.1.01.001.001');
});
test('retorna vazio quando não há código', () => {
  assert.equal(engine.parseAccountCode('TOTAL ENCARGOS'), '');
});

console.log('classifyProvenance — valor digitado (padrão Jan/Fev do arquivo real)');
test('célula sem fórmula é valor-digitado', () => {
  const adapter = makeAdapter({ CHECK: { E: { 5: { value: 0 } } } });
  const result = engine.classifyProvenance(adapter, adapter.cell('CHECK', 'E', 5));
  assert.equal(result.status, 'valor-digitado');
});

console.log('classifyProvenance / recalcContabil — fórmula quebrada (padrão Junho do arquivo real)');
test('referência para célula nunca preenchida é formula-quebrada', () => {
  const adapter = makeAdapter({
    CHECK: { X: { 5: { value: 0, formula: "'BALANCETE'!N136" } } },
    BALANCETE: {
      B: { 136: { value: '3.1.01.001.001' } },
      H: { 136: { value: 803770 } },
      I: { 136: { value: 0 } }
      // coluna N nunca foi preenchida -> célula ausente no adapter
    }
  });
  const prov = engine.classifyProvenance(adapter, adapter.cell('CHECK', 'X', 5));
  assert.equal(prov.status, 'formula-quebrada');
});
test('lookupContaNoBalancete acha a conta direto pelo código', () => {
  const adapter = makeAdapter({
    BALANCETE: {
      B: { 136: { value: '3.1.01.001.001' } },
      H: { 136: { value: 803770 } },
      I: { 136: { value: 0 } }
    }
  });
  const result = engine.lookupContaNoBalancete(adapter, 'BALANCETE', '3.1.01.001.001');
  assert.equal(result.status, 'verificado');
  assert.equal(result.value, 803770);
});

console.log('findBalanceteColunas — achado real: no arquivo do cliente, Jan/Fev têm "Classificação" na coluna A e Débito/Crédito em E/F, não a posição fixa B/H/I que o motor assumia antes');
test('detecta colunas pelo cabeçalho quando o layout não é o padrão', () => {
  const adapter = makeAdapter({
    BALANCETE: {
      A: { 3: { value: 'Classificação' }, 140: { value: '3.1.01.001.001' } },
      C: { 3: { value: 'Nome' } },
      E: { 3: { value: 'Débito' }, 140: { value: 100 } },
      F: { 3: { value: 'Crédito' }, 140: { value: 40 } }
    }
  });
  const colunas = engine.findBalanceteColunas(adapter, 'BALANCETE');
  assert.deepEqual(colunas, { codigoCol: 'A', debitoCol: 'E', creditoCol: 'F' });
  const result = engine.lookupContaNoBalancete(adapter, 'BALANCETE', '3.1.01.001.001');
  assert.equal(result.status, 'verificado');
  assert.equal(result.value, 60);
});
test('sem cabeçalho "Classificação" detectável, cai no padrão B/H/I (comportamento anterior preservado)', () => {
  const adapter = makeAdapter({
    BALANCETE: { B: { 136: { value: '3.1.01.001.001' } }, H: { 136: { value: 803770 } }, I: { 136: { value: 0 } } }
  });
  const colunas = engine.findBalanceteColunas(adapter, 'BALANCETE');
  assert.deepEqual(colunas, { codigoCol: 'B', debitoCol: 'H', creditoCol: 'I' });
});
test('recalcContabil ignora a referência quebrada da CHECK e busca a conta direto no BALANCETE', () => {
  const adapter = makeAdapter({
    CHECK: { D: { 5: { value: 0, formula: "'BALANCETE'!N136" } } },
    BALANCETE: {
      B: { 136: { value: '3.1.01.001.001' } },
      H: { 136: { value: 803770 } },
      I: { 136: { value: 0 } }
    }
  });
  const result = engine.recalcContabil(adapter, adapter.cell('CHECK', 'D', 5), 'BALANCETE', '3.1.01.001.001');
  assert.equal(result.status, 'verificado');
  assert.equal(result.value, 803770);
  assert.equal(result.origem, 'balancete-por-codigo');
});
test('recalcContabil sinaliza sem-fonte quando o mês não tem BALANCETE nem fórmula pra seguir', () => {
  const adapter = makeAdapter({ CHECK: { D: { 5: { value: 700 } } } });
  const result = engine.recalcContabil(adapter, adapter.cell('CHECK', 'D', 5), null, '3.1.01.001.001');
  assert.equal(result.status, 'sem-fonte');
  assert.equal(result.value, null);
});
test('recalcContabil segue a composição da fórmula quando ela resolve, incluindo ajuste do RAZÃO', () => {
  const adapter = makeAdapter({
    CHECK: { P: { 8: { value: 93738.08, formula: "'BALANCETE'!J160+'RAZAO'!G700+'RAZAO'!G702" } } },
    BALANCETE: { J: { 160: { value: 75815.70 } } },
    RAZAO: { G: { 700: { value: 12000 }, 702: { value: 5922.38 } } }
  });
  const result = engine.recalcContabil(adapter, adapter.cell('CHECK', 'P', 8), 'BALANCETE', '3.1.01.001.003');
  assert.equal(result.status, 'verificado');
  assert.equal(result.origem, 'formula-composicao');
  assert.ok(Math.abs(result.value - 93738.08) < 0.01, `esperava somar os 3 termos (93738.08), veio ${result.value}`);
  assert.ok(result.ajuste, 'deveria reportar o ajuste do razão');
  assert.equal(result.ajuste.termos.length, 2);
  assert.ok(Math.abs(result.ajuste.valor - 17922.38) < 0.01);
});

console.log('classifyProvenance — fórmula viva concordante (padrão Abril/Maio do arquivo real)');
test('referência que resolve para célula preenchida é formula-viva', () => {
  const adapter = makeAdapter({
    CHECK: { P: { 5: { value: 100, formula: "'BALANCETE'!J158" } } },
    BALANCETE: { J: { 158: { value: 100 } } }
  });
  const prov = engine.classifyProvenance(adapter, adapter.cell('CHECK', 'P', 5));
  assert.equal(prov.status, 'formula-viva');
});

console.log('recalcFopag — SUMIFS reimplementado sobre a FOPAG');
test('soma as colunas certas filtrando por mês', () => {
  const jan = new Date(2026, 0, 1);
  const adapter = makeAdapter({
    CHECK: { C: { 2: { value: jan }, 5: { value: 1000, formula: "SUMIFS('FOPAG'!$E$1:$E$10,'FOPAG'!$B$1:$B$10,C2)" } } },
    FOPAG: {
      B: { 1: { value: jan }, 2: { value: jan }, 3: { value: new Date(2026, 1, 1) } },
      E: { 1: { value: 600 }, 2: { value: 400 }, 3: { value: 999999 } } // linha de fevereiro não deve entrar
    }
  });
  const result = engine.recalcFopag(adapter, 'CHECK', adapter.cell('CHECK', 'C', 5));
  assert.equal(result.status, 'verificado');
  assert.equal(result.value, 1000);
});
test('rótulo de rubrica no drill-down acha o cabeçalho mesmo quando não está na linha 1', () => {
  // Achado real: a FOPAG 2026 do cliente tem cabeçalho na linha 3, não na
  // linha 1 — sem procurar em mais de uma linha, o drill-down mostrava a
  // letra da coluna ("G") em vez do nome da rubrica ("ENCARGOS").
  const jan = new Date(2026, 0, 1);
  const adapter = makeAdapter({
    CHECK: { C: { 2: { value: jan }, 5: { value: 1000, formula: "SUMIFS('FOPAG'!$E$1:$E$10,'FOPAG'!$B$1:$B$10,C2)" } } },
    FOPAG: {
      A: { 1: { value: 'planilha de teste' } }, // linha 1 tem só um título solto, não cabeçalho
      B: { 3: { value: 'MES' } },
      E: { 3: { value: 'ENCARGOS' }, 1: { value: jan } } // header real na linha 3
    }
  });
  const result = engine.recalcFopag(adapter, 'CHECK', adapter.cell('CHECK', 'C', 5));
  assert.equal(result.termos[0].rubrica, 'ENCARGOS');
});

console.log('reconcileConta — cenário completo');
test('Janeiro: valor digitado gera alerta de preenchimento mesmo com números plausíveis', () => {
  const jan = new Date(2026, 0, 1);
  const adapter = makeAdapter({
    CHECK: {
      A: { 5: { value: 'SALARIOS E ORDENADOS (3.1.01.001.001)' } },
      C: { 2: { value: jan }, 5: { value: 1000, formula: "SUMIFS('FOPAG'!$E$1:$E$10,'FOPAG'!$B$1:$B$10,C2)" } },
      D: { 5: { value: 1000 } }, // digitado, "por coincidência" bate com o recalculado
      E: { 5: { value: 0 } }    // digitado
    },
    FOPAG: { B: { 1: { value: jan } }, E: { 1: { value: 1000 } } }
  });
  const conta = engine.reconcileConta({
    adapter, checkSheet: 'CHECK', contaCodigo: '3.1.01.001.001', descricao: 'SALARIOS E ORDENADOS',
    competencia: 'Janeiro',
    fopagCell: adapter.cell('CHECK', 'C', 5),
    contabilCell: adapter.cell('CHECK', 'D', 5),
    diferencaCell: adapter.cell('CHECK', 'E', 5),
    balanceteSheet: null
  });
  assert.ok(conta.alertas.includes('preenchimento'), 'deveria sinalizar preenchimento mesmo com valores concordantes');
  assert.equal(conta.proveniencia.contabil, 'valor-digitado');
});

test('Junho: fórmula quebrada suprime a divergência financeira falsa e sinaliza preenchimento', () => {
  const jun = new Date(2026, 5, 1);
  const adapter = makeAdapter({
    CHECK: {
      A: { 5: { value: 'SALARIOS E ORDENADOS (3.1.01.001.001)' } },
      C: { 2: { value: jun }, 5: { value: 770174.61, formula: "SUMIFS('FOPAG'!$E$1:$E$10,'FOPAG'!$B$1:$B$10,C2)" } },
      D: { 5: { value: 0, formula: "'BALANCETE 05.2026'!N136" } }, // referência quebrada, replica o achado real
      E: { 5: { value: 770174.61, formula: '=C5-D5' } }
    },
    FOPAG: { B: { 1: { value: jun } }, E: { 1: { value: 770174.61 } } },
    'BALANCETE 05.2026': {
      B: { 136: { value: '3.1.01.001.001' } },
      H: { 136: { value: 803770 } },
      I: { 136: { value: 0 } }
    }
  });
  const conta = engine.reconcileConta({
    adapter, checkSheet: 'CHECK', contaCodigo: '3.1.01.001.001', descricao: 'SALARIOS E ORDENADOS',
    competencia: 'Junho',
    fopagCell: adapter.cell('CHECK', 'C', 5),
    contabilCell: adapter.cell('CHECK', 'D', 5),
    diferencaCell: adapter.cell('CHECK', 'E', 5),
    balanceteSheet: 'BALANCETE 05.2026'
  });
  assert.ok(conta.alertas.includes('preenchimento'), 'deveria sinalizar preenchimento pela referência quebrada');
  assert.equal(conta.recalculado.contabil, 803770, 'recálculo deve ignorar a referência quebrada e buscar a conta direto no BALANCETE');
  assert.equal(Math.round(conta.recalculado.diferenca), Math.round(770174.61 - 803770), 'diferença real recalculada, não os R$770 mil aparentes');
});

console.log('computeConfiabilidade');
test('conta 100% viva pontua 100', () => {
  const contas = [{
    declarado: { fopag: 100, contabil: 100 },
    proveniencia: { fopag: 'formula-viva', contabil: 'formula-viva', diferenca: 'intra' }
  }];
  assert.equal(engine.computeConfiabilidade(contas).score, 100);
});
test('mistura de proveniência derruba o score proporcionalmente ao peso financeiro', () => {
  const contas = [
    { declarado: { fopag: 900, contabil: 900 }, proveniencia: { fopag: 'valor-digitado', contabil: 'valor-digitado', diferenca: 'valor-digitado' } },
    { declarado: { fopag: 100, contabil: 100 }, proveniencia: { fopag: 'formula-viva', contabil: 'formula-viva', diferenca: 'intra' } }
  ];
  const result = engine.computeConfiabilidade(contas);
  assert.ok(result.score < 20, `esperava score baixo (conta de maior peso está toda digitada), veio ${result.score}`);
});

console.log('computeCobertura');
test('conta presente no BALANCETE e ausente da CHECK vira conta fantasma', () => {
  const adapter = makeAdapter({
    BALANCETE: {
      B: { 10: { value: '3.1.01.001.001' }, 11: { value: '3.1.01.099.001' } }
    }
  });
  const result = engine.computeCobertura(adapter, 'BALANCETE', new Set(['3.1.01.001.001']));
  assert.equal(result.contasFantasmas.length, 1);
  assert.equal(result.contasFantasmas[0].codigo, '3.1.01.099.001');
});
test('conta-pai/subtotal não vira fantasma quando profundidade filtra só contas-folha (padrão real observado no BALANCETE)', () => {
  // Replica o achado real: "3.1.01.003" é o subtotal de "3.1.01.003.003" +
  // "3.1.01.003.007" — sem o filtro de profundidade, o pai aparecia como
  // "conta faltando" mesmo já estando coberto pelas contas-filha.
  const adapter = makeAdapter({
    BALANCETE: {
      B: {
        10: { value: '3.1.01.003' },       // pai/subtotal (4 segmentos)
        11: { value: '3.1.01.003.003' },   // filha (5 segmentos) — na CHECK
        12: { value: '3.1.01.003.007' }    // filha (5 segmentos) — não está na CHECK
      }
    }
  });
  const semFiltro = engine.computeCobertura(adapter, 'BALANCETE', new Set(['3.1.01.003.003']));
  assert.equal(semFiltro.contasFantasmas.length, 2, 'sem profundidade, o pai também aparece (comportamento antigo)');
  const comFiltro = engine.computeCobertura(adapter, 'BALANCETE', new Set(['3.1.01.003.003']), { profundidade: 5 });
  assert.equal(comFiltro.contasFantasmas.length, 1);
  assert.equal(comFiltro.contasFantasmas[0].codigo, '3.1.01.003.007');
});

console.log(`\n${passed} passaram, ${failed} falharam`);
process.exit(failed ? 1 : 0);
