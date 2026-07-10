'use strict';
/* Gera test/fixtures/fechamento-ficticio-teste.xlsx — dado 100% fictício,
 * sem nenhum vínculo com cliente real. Reproduz a MESMA classe de estrutura
 * do arquivo real que motivou este projeto, mas com variações deliberadas
 * que aquele arquivo não cobria — serve pra provar que a detecção do
 * app.js generaliza, não está "decorada" pra um arquivo só:
 *
 *   - família de conta contábil diferente (4.2.01.* em vez de 3.1.01.*)
 *   - nome de aba FOPAG sem o ano no nome ("FOPAG TESTE")
 *   - cabeçalho da FOPAG na linha 2 (arquivo real tinha na linha 3)
 *   - um BALANCETE com separador de mês por hífen ("Balancete 03-2026")
 *     em vez de ponto (arquivo real só tinha "BALANCETE MM.AAAA")
 *   - "Centro de Custo" com espaçamento irregular no cabeçalho
 *   - hierarquia pai/filho no BALANCETE (testa o filtro de profundidade)
 *   - os 3 padrões de proveniência (digitado, quebrada, ajuste RAZÃO) +
 *     um mês limpo com uma divergência financeira real de propósito
 *
 * Rodar: npm install (uma vez, instala a devDependency xlsx) e depois
 * node test/fixtures/build-synthetic-fixture.js
 */
const path = require('path');
const XLSX = require('xlsx');

function cellNum(v, formula) {
  const c = { t: 'n', v };
  if (formula) c.f = formula;
  return c;
}
function cellStr(v) { return { t: 's', v }; }
function cellDate(date) { return { t: 'd', v: date }; }

function buildSheet(cells, dims) {
  const ws = {};
  Object.assign(ws, cells);
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: dims.rows - 1, c: dims.cols - 1 } });
  return ws;
}

const wb = { SheetNames: [], Sheets: {} };
function addSheet(name, ws) { wb.SheetNames.push(name); wb.Sheets[name] = ws; }

// ------------------------------------------------------------ FOPAG TESTE
const funcionarios = [
  { nome: 'ANA TESTE SILVA', cc: 'Comercial' },
  { nome: 'BRUNO TESTE COSTA', cc: 'Operações' },
  { nome: 'CARLA TESTE LIMA', cc: 'Comercial' }
];
const meses = [
  { label: 'JANEIRO', date: new Date(2026, 0, 1) },
  { label: 'FEVEREIRO', date: new Date(2026, 1, 1) },
  { label: 'MARÇO', date: new Date(2026, 2, 1) },
  { label: 'ABRIL', date: new Date(2026, 3, 1) }
];
const base = {
  'ANA TESTE SILVA': [8000, 400, 600, 1800],
  'BRUNO TESTE COSTA': [6500, 250, 500, 1450],
  'CARLA TESTE LIMA': [9200, 0, 700, 2100]
};

const fopagCells = {
  A1: cellStr('planilha de teste — dados 100% fictícios'),
  A2: cellStr('NOME'), B2: cellStr('MES'), C2: cellStr('Centro  de   Custo'),
  D2: cellStr('SALARIO'), E2: cellStr('HORAS EXTRAS'), F2: cellStr('BENEFICIOS'), G2: cellStr('ENCARGOS'),
  H2: cellStr('TOTAL PROVENTOS')
};
let row = 3;
const fopagRowsByMonth = {};
for (const mes of meses) {
  const key = `${mes.date.getFullYear()}-${String(mes.date.getMonth() + 1).padStart(2, '0')}`;
  fopagRowsByMonth[key] = [];
  for (const f of funcionarios) {
    const [sal, he, ben, enc] = base[f.nome];
    const mult = 1 + (mes.date.getMonth() * 0.01);
    const vals = [sal * mult, he * mult, ben, enc * mult].map((v) => Math.round(v * 100) / 100);
    fopagCells[`A${row}`] = cellStr(f.nome);
    fopagCells[`B${row}`] = cellDate(mes.date);
    fopagCells[`C${row}`] = cellStr(f.cc);
    fopagCells[`D${row}`] = cellNum(vals[0]);
    fopagCells[`E${row}`] = cellNum(vals[1]);
    fopagCells[`F${row}`] = cellNum(vals[2]);
    fopagCells[`G${row}`] = cellNum(vals[3]);
    fopagCells[`H${row}`] = cellNum(vals[0] + vals[1] + vals[2] + vals[3]);
    fopagRowsByMonth[key].push({ row, vals });
    row++;
  }
}
addSheet('FOPAG TESTE', buildSheet(fopagCells, { rows: row, cols: 8 }));

function sumFopagCol(monthKey, colIdx) {
  return fopagRowsByMonth[monthKey].reduce((s, r) => s + r.vals[colIdx], 0);
}

// ------------------------------------------------------------ BALANCETEs
function buildBalancete(entries) {
  const cells = { A1: cellStr('Código'), B1: cellStr('Classificação'), E1: cellStr('Nome'), H1: cellStr('Débito'), I1: cellStr('Crédito') };
  let r = 2;
  entries.forEach(([codigo, nome, debito, credito]) => {
    cells[`A${r}`] = cellNum(r);
    cells[`B${r}`] = cellStr(codigo);
    cells[`E${r}`] = cellStr(nome);
    cells[`H${r}`] = cellNum(debito);
    cells[`I${r}`] = cellNum(credito || 0);
    r++;
  });
  return buildSheet(cells, { rows: r, cols: 9 });
}

// Fevereiro: contábil bate exatamente com a fonte, exceto ENCARGOS TESTE que
// diverge de propósito (é a divergência financeira real do cenário limpo).
const fevContabil = {
  salario: sumFopagCol('2026-02', 0), horas: sumFopagCol('2026-02', 1),
  beneficios: sumFopagCol('2026-02', 2), encargos: sumFopagCol('2026-02', 3) + 250.75
};
addSheet('BALANCETE 02.2026', buildBalancete([
  ['4.2.01.001.001', 'SALARIO BASE TESTE', fevContabil.salario, 0],
  ['4.2.01.001.002', 'HORAS EXTRAS TESTE', fevContabil.horas, 0],
  ['4.2.01.001.003', 'BENEFICIOS TESTE', fevContabil.beneficios, 0],
  ['4.2.01.004.001', 'ENCARGOS TESTE', fevContabil.encargos, 0]
]));

// Março: BALANCETE existe (nomenclatura alternativa "Balancete MM-AAAA"),
// mas a fórmula da CHECK aponta pra fora da faixa preenchida — referência
// quebrada, igual ao achado real de Junho no arquivo do cliente.
addSheet('Balancete 03-2026', buildBalancete([
  ['4.2.01.001.001', 'SALARIO BASE TESTE', sumFopagCol('2026-03', 0), 0],
  ['4.2.01.001.002', 'HORAS EXTRAS TESTE', sumFopagCol('2026-03', 1), 0],
  ['4.2.01.001.003', 'BENEFICIOS TESTE', sumFopagCol('2026-03', 2), 0],
  ['4.2.01.004.001', 'ENCARGOS TESTE', sumFopagCol('2026-03', 3), 0]
]));

// Abril: BALANCETE com hierarquia pai/filho (testa o filtro de profundidade
// da cobertura) + uma conta-fantasma de verdade (folha, ausente da CHECK) +
// ENCARGOS ganha um ajuste do RAZÃO na própria fórmula da CHECK.
const abrSalario = sumFopagCol('2026-04', 0);
const abrHoras = sumFopagCol('2026-04', 1);
const abrBeneficios = sumFopagCol('2026-04', 2);
const abrEncargosBase = sumFopagCol('2026-04', 3);
const razaoAjuste1 = 180.4, razaoAjuste2 = 95.1;
addSheet('BALANCETE 04.2026', buildBalancete([
  ['4.2.01.001', 'GRUPO SALARIAL TESTE (pai)', abrSalario + 340.2, 0], // pai = soma das filhas abaixo
  ['4.2.01.001.001', 'SALARIO BASE TESTE', abrSalario, 0],
  ['4.2.01.001.002', 'HORAS EXTRAS TESTE', abrHoras, 0],
  ['4.2.01.001.003', 'BENEFICIOS TESTE', abrBeneficios, 0],
  ['4.2.01.001.009', 'AUXILIO CRECHE TESTE (fantasma)', 340.2, 0], // folha, ausente da CHECK
  ['4.2.01.004.001', 'ENCARGOS TESTE', abrEncargosBase, 0]
]));
addSheet('RAZÃO 04.2026', buildSheet({
  A1: cellStr('Data'), B1: cellStr('Histórico'), E1: cellStr('Contra'), F1: cellStr('Débito'), G1: cellStr('Crédito'),
  A10: cellDate(new Date(2026, 3, 12)), B10: cellStr('Reclassificação teste 1'), F10: cellNum(razaoAjuste1), G10: cellNum(0),
  A11: cellDate(new Date(2026, 3, 20)), B11: cellStr('Reclassificação teste 2'), F11: cellNum(razaoAjuste2), G11: cellNum(0)
}, { rows: 12, cols: 8 }));

// ------------------------------------------------------------------ CHECK
const blockStartCols = ['C', 'G', 'K', 'O']; // Jan, Fev, Mar, Abr
const contas = [
  { codigo: '4.2.01.001.001', desc: 'SALARIO BASE TESTE', fopagValueCol: 'D', idx: 0 },
  { codigo: '4.2.01.001.002', desc: 'HORAS EXTRAS TESTE', fopagValueCol: 'E', idx: 1 },
  { codigo: '4.2.01.001.003', desc: 'BENEFICIOS TESTE', fopagValueCol: 'F', idx: 2 },
  { codigo: '4.2.01.004.001', desc: 'ENCARGOS TESTE', fopagValueCol: 'G', idx: 3 }
];

const checkCells = { A4: cellStr('DESCRIÇÃO CONTA') };
contas.forEach((conta, i) => { checkCells[`A${5 + i}`] = cellStr(`${conta.desc} (${conta.codigo})`); });

meses.forEach((mes, blockIdx) => {
  const startCol = blockStartCols[blockIdx];
  const startIdx = XLSX.utils.decode_col(startCol) + 1;
  const fopagCol = XLSX.utils.encode_col(startIdx - 1);
  const contabilCol = XLSX.utils.encode_col(startIdx);
  const diferencaCol = XLSX.utils.encode_col(startIdx + 1);
  const monthKey = `${mes.date.getFullYear()}-${String(mes.date.getMonth() + 1).padStart(2, '0')}`;

  checkCells[`${fopagCol}1`] = cellStr(mes.label);
  checkCells[`${fopagCol}2`] = cellDate(mes.date);
  checkCells[`${fopagCol}3`] = cellStr('FOPAG');
  checkCells[`${contabilCol}3`] = cellStr('CONTABIL');
  checkCells[`${diferencaCol}3`] = cellStr('DIFERENÇA');

  contas.forEach((conta, i) => {
    const r = 5 + i;
    const fopagValue = sumFopagCol(monthKey, conta.idx);
    const fopagFormula = `SUMIFS('FOPAG TESTE'!$${conta.fopagValueCol}$3:$${conta.fopagValueCol}$100,'FOPAG TESTE'!$B$3:$B$100,${fopagCol}2)`;
    checkCells[`${fopagCol}${r}`] = cellNum(fopagValue, fopagFormula);

    if (mes.label === 'JANEIRO') {
      // padrão real Jan/Fev: contábil e diferença digitados à mão
      checkCells[`${contabilCol}${r}`] = cellNum(fopagValue);
      checkCells[`${diferencaCol}${r}`] = cellNum(0);
    } else if (mes.label === 'FEVEREIRO') {
      const contabilRef = `'BALANCETE 02.2026'!H${2 + i}`;
      const contabilValue = [fevContabil.salario, fevContabil.horas, fevContabil.beneficios, fevContabil.encargos][i];
      checkCells[`${contabilCol}${r}`] = cellNum(contabilValue, contabilRef);
      checkCells[`${diferencaCol}${r}`] = cellNum(fopagValue - contabilValue, `=${fopagCol}${r}-${contabilCol}${r}`);
    } else if (mes.label === 'MARÇO') {
      const contabilRef = `'Balancete 03-2026'!Z${2 + i}`; // coluna nunca preenchida
      checkCells[`${contabilCol}${r}`] = cellNum(0, contabilRef);
      checkCells[`${diferencaCol}${r}`] = cellNum(fopagValue, `=${fopagCol}${r}-${contabilCol}${r}`);
    } else if (mes.label === 'ABRIL') {
      if (conta.codigo === '4.2.01.004.001') {
        const contabilFormula = `'BALANCETE 04.2026'!H7+'RAZÃO 04.2026'!F10+'RAZÃO 04.2026'!F11`;
        const contabilValue = abrEncargosBase + razaoAjuste1 + razaoAjuste2;
        checkCells[`${contabilCol}${r}`] = cellNum(contabilValue, contabilFormula);
        checkCells[`${diferencaCol}${r}`] = cellNum(fopagValue - contabilValue, `=${fopagCol}${r}-${contabilCol}${r}`);
      } else {
        const balRow = { '4.2.01.001.001': 3, '4.2.01.001.002': 4, '4.2.01.001.003': 5 }[conta.codigo];
        const contabilRef = `'BALANCETE 04.2026'!H${balRow}`;
        const contabilValue = [abrSalario, abrHoras, abrBeneficios][i];
        checkCells[`${contabilCol}${r}`] = cellNum(contabilValue, contabilRef);
        checkCells[`${diferencaCol}${r}`] = cellNum(fopagValue - contabilValue, `=${fopagCol}${r}-${contabilCol}${r}`);
      }
    }
  });
});
addSheet('CHECK', buildSheet(checkCells, { rows: 9, cols: 20 }));

const outPath = path.join(__dirname, 'fechamento-ficticio-teste.xlsx');
XLSX.writeFile(wb, outPath, { bookType: 'xlsx' });
console.log('Fixture gerada em:', outPath);
