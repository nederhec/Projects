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
 * Rodar: npm install (uma vez, instala a devDependency exceljs) e depois
 * npm run fixtures:build
 */
const path = require('path');
const ExcelJS = require('exceljs');

const wb = new ExcelJS.Workbook();
function sheet(name) { return wb.addWorksheet(name); }
function set(ws, addr, value, formula) {
  ws.getCell(addr).value = formula ? { formula, result: value } : value;
}

// ------------------------------------------------------------ FOPAG TESTE
// Cabeçalho na LINHA 2 (não linha 3 como no arquivo real) — testa que a
// detecção de header não está fixa numa posição.
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

const fopag = sheet('FOPAG TESTE');
set(fopag, 'A1', 'planilha de teste — dados 100% fictícios');
set(fopag, 'A2', 'NOME'); set(fopag, 'B2', 'MES'); set(fopag, 'C2', 'Centro  de   Custo');
set(fopag, 'D2', 'SALARIO'); set(fopag, 'E2', 'HORAS EXTRAS'); set(fopag, 'F2', 'BENEFICIOS'); set(fopag, 'G2', 'ENCARGOS');
set(fopag, 'H2', 'TOTAL PROVENTOS');

let row = 3;
const fopagRowsByMonth = {};
for (const mes of meses) {
  const key = `${mes.date.getFullYear()}-${String(mes.date.getMonth() + 1).padStart(2, '0')}`;
  fopagRowsByMonth[key] = [];
  for (const f of funcionarios) {
    const [sal, he, ben, enc] = base[f.nome];
    const mult = 1 + (mes.date.getMonth() * 0.01);
    const vals = [sal * mult, he * mult, ben, enc * mult].map((v) => Math.round(v * 100) / 100);
    set(fopag, `A${row}`, f.nome);
    set(fopag, `B${row}`, mes.date);
    set(fopag, `C${row}`, f.cc);
    set(fopag, `D${row}`, vals[0]);
    set(fopag, `E${row}`, vals[1]);
    set(fopag, `F${row}`, vals[2]);
    set(fopag, `G${row}`, vals[3]);
    set(fopag, `H${row}`, vals[0] + vals[1] + vals[2] + vals[3]);
    fopagRowsByMonth[key].push({ row, vals });
    row++;
  }
}

function sumFopagCol(monthKey, colIdx) {
  return fopagRowsByMonth[monthKey].reduce((s, r) => s + r.vals[colIdx], 0);
}

// ------------------------------------------------------------ BALANCETEs
function balancete(name, entries) {
  const ws = sheet(name);
  set(ws, 'A1', 'Código'); set(ws, 'B1', 'Classificação'); set(ws, 'E1', 'Nome'); set(ws, 'H1', 'Débito'); set(ws, 'I1', 'Crédito');
  let r = 2;
  entries.forEach(([codigo, nome, debito, credito]) => {
    set(ws, `A${r}`, r);
    set(ws, `B${r}`, codigo);
    set(ws, `E${r}`, nome);
    set(ws, `H${r}`, debito);
    set(ws, `I${r}`, credito || 0);
    r++;
  });
  return ws;
}

// Fevereiro: contábil bate exatamente com a fonte, exceto ENCARGOS TESTE que
// diverge de propósito (é a divergência financeira real do cenário limpo).
const fevContabil = {
  salario: sumFopagCol('2026-02', 0), horas: sumFopagCol('2026-02', 1),
  beneficios: sumFopagCol('2026-02', 2), encargos: sumFopagCol('2026-02', 3) + 250.75
};
balancete('BALANCETE 02.2026', [
  ['4.2.01.001.001', 'SALARIO BASE TESTE', fevContabil.salario, 0],
  ['4.2.01.001.002', 'HORAS EXTRAS TESTE', fevContabil.horas, 0],
  ['4.2.01.001.003', 'BENEFICIOS TESTE', fevContabil.beneficios, 0],
  ['4.2.01.004.001', 'ENCARGOS TESTE', fevContabil.encargos, 0]
]);

// Março: BALANCETE existe (nomenclatura alternativa "Balancete MM-AAAA"),
// mas a fórmula da CHECK aponta pra fora da faixa preenchida — referência
// quebrada, igual ao achado real de Junho no arquivo do cliente.
balancete('Balancete 03-2026', [
  ['4.2.01.001.001', 'SALARIO BASE TESTE', sumFopagCol('2026-03', 0), 0],
  ['4.2.01.001.002', 'HORAS EXTRAS TESTE', sumFopagCol('2026-03', 1), 0],
  ['4.2.01.001.003', 'BENEFICIOS TESTE', sumFopagCol('2026-03', 2), 0],
  ['4.2.01.004.001', 'ENCARGOS TESTE', sumFopagCol('2026-03', 3), 0]
]);

// Abril: BALANCETE com hierarquia pai/filho (testa o filtro de profundidade
// da cobertura) + uma conta-fantasma de verdade (folha, ausente da CHECK) +
// ENCARGOS ganha um ajuste do RAZÃO na própria fórmula da CHECK.
const abrSalario = sumFopagCol('2026-04', 0);
const abrHoras = sumFopagCol('2026-04', 1);
const abrBeneficios = sumFopagCol('2026-04', 2);
const abrEncargosBase = sumFopagCol('2026-04', 3);
const razaoAjuste1 = 180.4, razaoAjuste2 = 95.1;
balancete('BALANCETE 04.2026', [
  ['4.2.01.001', 'GRUPO SALARIAL TESTE (pai)', abrSalario + 340.2, 0], // pai = soma das filhas abaixo
  ['4.2.01.001.001', 'SALARIO BASE TESTE', abrSalario, 0],
  ['4.2.01.001.002', 'HORAS EXTRAS TESTE', abrHoras, 0],
  ['4.2.01.001.003', 'BENEFICIOS TESTE', abrBeneficios, 0],
  ['4.2.01.001.009', 'AUXILIO CRECHE TESTE (fantasma)', 340.2, 0], // folha, ausente da CHECK
  ['4.2.01.004.001', 'ENCARGOS TESTE', abrEncargosBase, 0]
]);

const razao = sheet('RAZÃO 04.2026');
set(razao, 'A1', 'Data'); set(razao, 'B1', 'Histórico'); set(razao, 'E1', 'Contra'); set(razao, 'F1', 'Débito'); set(razao, 'G1', 'Crédito');
set(razao, 'A10', new Date(2026, 3, 12)); set(razao, 'B10', 'Reclassificação teste 1'); set(razao, 'F10', razaoAjuste1); set(razao, 'G10', 0);
set(razao, 'A11', new Date(2026, 3, 20)); set(razao, 'B11', 'Reclassificação teste 2'); set(razao, 'F11', razaoAjuste2); set(razao, 'G11', 0);

// ------------------------------------------------------------------ CHECK
const blockStartCols = ['C', 'G', 'K', 'O']; // Jan, Fev, Mar, Abr
const colIdx = { C: 3, G: 7, K: 11, O: 15 };
function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
const contas = [
  { codigo: '4.2.01.001.001', desc: 'SALARIO BASE TESTE', fopagValueCol: 'D', idx: 0 },
  { codigo: '4.2.01.001.002', desc: 'HORAS EXTRAS TESTE', fopagValueCol: 'E', idx: 1 },
  { codigo: '4.2.01.001.003', desc: 'BENEFICIOS TESTE', fopagValueCol: 'F', idx: 2 },
  { codigo: '4.2.01.004.001', desc: 'ENCARGOS TESTE', fopagValueCol: 'G', idx: 3 }
];

const check = sheet('CHECK');
set(check, 'A4', 'DESCRIÇÃO CONTA');
contas.forEach((conta, i) => { set(check, `A${5 + i}`, `${conta.desc} (${conta.codigo})`); });

meses.forEach((mes, blockIdx) => {
  const startCol = blockStartCols[blockIdx];
  const startIdx = colIdx[startCol];
  const fopagCol = colLetter(startIdx);
  const contabilCol = colLetter(startIdx + 1);
  const diferencaCol = colLetter(startIdx + 2);
  const monthKey = `${mes.date.getFullYear()}-${String(mes.date.getMonth() + 1).padStart(2, '0')}`;

  set(check, `${fopagCol}1`, mes.label);
  set(check, `${fopagCol}2`, mes.date);
  set(check, `${fopagCol}3`, 'FOPAG');
  set(check, `${contabilCol}3`, 'CONTABIL');
  set(check, `${diferencaCol}3`, 'DIFERENÇA');

  contas.forEach((conta, i) => {
    const r = 5 + i;
    const fopagValue = sumFopagCol(monthKey, conta.idx);
    const fopagFormula = `SUMIFS('FOPAG TESTE'!$${conta.fopagValueCol}$3:$${conta.fopagValueCol}$100,'FOPAG TESTE'!$B$3:$B$100,${fopagCol}2)`;
    set(check, `${fopagCol}${r}`, fopagValue, fopagFormula);

    if (mes.label === 'JANEIRO') {
      set(check, `${contabilCol}${r}`, fopagValue); // digitado
      set(check, `${diferencaCol}${r}`, 0); // digitado
    } else if (mes.label === 'FEVEREIRO') {
      const contabilRef = `'BALANCETE 02.2026'!H${2 + i}`;
      const contabilValue = [fevContabil.salario, fevContabil.horas, fevContabil.beneficios, fevContabil.encargos][i];
      set(check, `${contabilCol}${r}`, contabilValue, contabilRef);
      set(check, `${diferencaCol}${r}`, fopagValue - contabilValue, `${fopagCol}${r}-${contabilCol}${r}`);
    } else if (mes.label === 'MARÇO') {
      const contabilRef = `'Balancete 03-2026'!Z${2 + i}`; // coluna nunca preenchida
      set(check, `${contabilCol}${r}`, 0, contabilRef);
      set(check, `${diferencaCol}${r}`, fopagValue, `${fopagCol}${r}-${contabilCol}${r}`);
    } else if (mes.label === 'ABRIL') {
      if (conta.codigo === '4.2.01.004.001') {
        const contabilFormula = `'BALANCETE 04.2026'!H7+'RAZÃO 04.2026'!F10+'RAZÃO 04.2026'!F11`;
        const contabilValue = abrEncargosBase + razaoAjuste1 + razaoAjuste2;
        set(check, `${contabilCol}${r}`, contabilValue, contabilFormula);
        set(check, `${diferencaCol}${r}`, fopagValue - contabilValue, `${fopagCol}${r}-${contabilCol}${r}`);
      } else {
        const balRow = { '4.2.01.001.001': 3, '4.2.01.001.002': 4, '4.2.01.001.003': 5 }[conta.codigo];
        const contabilRef = `'BALANCETE 04.2026'!H${balRow}`;
        const contabilValue = [abrSalario, abrHoras, abrBeneficios][i];
        set(check, `${contabilCol}${r}`, contabilValue, contabilRef);
        set(check, `${diferencaCol}${r}`, fopagValue - contabilValue, `${fopagCol}${r}-${contabilCol}${r}`);
      }
    }
  });
});

const outPath = path.join(__dirname, 'fechamento-ficticio-teste.xlsx');
wb.xlsx.writeFile(outPath).then(() => console.log('Fixture gerada em:', outPath));
