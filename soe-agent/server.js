const http = require('http');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

function loadEnvFile() {
  const envPath = path.join(__dirname, '.env');
  if (!fs.existsSync(envPath)) return;
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const m = line.match(/^\s*([\w.-]+)\s*=\s*(.*)?\s*$/);
    if (!m || m[1].startsWith('#')) continue;
    const key = m[1];
    let val = (m[2] || '').trim();
    if (val.startsWith('"') && val.endsWith('"')) val = val.slice(1, -1);
    if (!(key in process.env)) process.env[key] = val;
  }
}
loadEnvFile();

const PORT = process.env.PORT || 4000;
const FRONTEND_ORIGIN = process.env.FRONTEND_ORIGIN || '*';
const BODY_LIMIT = 5 * 1024 * 1024; // 5MB — payload de dados consolidados + margem de segurança
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';
const ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-6';

const REQUIRED_ENV = [
  'AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET',
  'SHAREPOINT_SITE_ID', 'POWERBI_WORKSPACE_ID', 'POWERBI_DATASET_ID',
];
const OPTIONAL_ENV = ['ANTHROPIC_API_KEY', 'GRAPH_SENDER_UPN', 'ALERT_RECIPIENTS'];

function assertConfigured(keys) {
  const missing = keys.filter(k => !process.env[k]);
  if (missing.length) {
    throw new Error(`Variáveis de ambiente ausentes: ${missing.join(', ')}`);
  }
}

// --- Azure AD — client credentials flow ---
async function getAccessToken(scope) {
  assertConfigured(['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET']);
  const url = `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: process.env.AZURE_CLIENT_ID,
    client_secret: process.env.AZURE_CLIENT_SECRET,
    scope,
    grant_type: 'client_credentials',
  });
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(`Falha ao obter token (${res.status}): ${json.error_description || JSON.stringify(json)}`);
  }
  return json.access_token;
}

// --- Excel real (.xlsx) ---
const COLUMNS = [
  { header: 'SKU', key: 'SKU' },
  { header: 'Planta', key: 'Planta' },
  { header: 'Recurso', key: 'Recurso' },
  { header: 'Produto', key: 'Produto' },
  { header: 'Cliente', key: 'Cliente' },
  { header: 'Tipo_Cliente', key: 'Tipo_Cliente' },
  { header: 'Carteira_Total', key: 'Carteira_Total' },
  { header: 'Demanda_Media_Proxima_Janela', key: 'Demanda_Media_Proxima_Janela' },
  { header: 'Producao_Planejada', key: 'Producao_Planejada' },
  { header: 'Producao_Real', key: 'Producao_Real' },
  { header: 'Estoque_Final_Real', key: 'Estoque_Final_Real' },
  { header: 'Projecao_Estoque_SOE', key: 'Projecao_Estoque_SOE' },
  { header: 'Estoque_PVE', key: 'Estoque_PVE' },
  { header: 'Margem_EBITDA_Rt', key: 'Margem_EBITDA_Rt' },
  { header: 'Ciclo_Produtivo', key: 'Ciclo_Produtivo' },
  { header: 'Flag_Ruptura', key: 'Flag_Ruptura' },
  { header: 'Flag_Excesso', key: 'Flag_Excesso' },
  { header: 'Flag_Excecao_Lote', key: 'Flag_Excecao_Lote' },
  { header: 'Impacto_EBITDA', key: 'Impacto_EBITDA' },
  { header: 'Status', key: 'Status' },
  { header: 'Observacoes', key: 'Observacoes' },
];

async function buildWorkbook(rows) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'SOE_AGENT_V4';
  wb.created = new Date();
  const sheet = wb.addWorksheet('Base_Consolidada');
  sheet.columns = COLUMNS;
  sheet.getRow(1).font = { bold: true };
  sheet.addRows(rows);
  return wb.xlsx.writeBuffer();
}

// --- SharePoint via Microsoft Graph ---
async function uploadToSharePoint(buffer, filename) {
  assertConfigured(['SHAREPOINT_SITE_ID']);
  const token = await getAccessToken('https://graph.microsoft.com/.default');
  const siteId = process.env.SHAREPOINT_SITE_ID;
  const folderPath = process.env.SHAREPOINT_FOLDER_PATH || 'Comite_SOE/Base_Gerencial';
  // PUT direto no conteúdo só suporta até 4MB; arquivos maiores exigem upload session (createUploadSession).
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${folderPath}/${encodeURIComponent(filename)}:/content`;
  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    body: buffer,
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(`Falha no upload ao SharePoint (${res.status}): ${json.error?.message || JSON.stringify(json)}`);
  }
  return { webUrl: json.webUrl, id: json.id, lastModifiedDateTime: json.lastModifiedDateTime };
}

// --- Power BI REST API ---
async function refreshPowerBIDataset(datasetId) {
  assertConfigured(['POWERBI_WORKSPACE_ID']);
  const token = await getAccessToken('https://analysis.windows.net/powerbi/api/.default');
  const workspaceId = process.env.POWERBI_WORKSPACE_ID;
  const url = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/datasets/${datasetId}/refreshes`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ notifyOption: 'NoNotification' }),
  });
  if (res.status !== 202) {
    const json = await res.json().catch(() => ({}));
    throw new Error(`Falha ao iniciar refresh do dataset (${res.status}): ${json.error?.message || JSON.stringify(json)}`);
  }
  return { requestId: res.headers.get('RequestId') || null };
}

async function getLastRefreshStatus(datasetId) {
  const token = await getAccessToken('https://analysis.windows.net/powerbi/api/.default');
  const workspaceId = process.env.POWERBI_WORKSPACE_ID;
  const url = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/datasets/${datasetId}/refreshes?$top=1`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(`Falha ao consultar status do refresh (${res.status}): ${json.error?.message || JSON.stringify(json)}`);
  }
  return json.value?.[0] || null;
}

async function waitForRefreshCompletion(datasetId, { timeoutMs = 120_000, intervalMs = 3000 } = {}) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const status = await getLastRefreshStatus(datasetId);
    if (status?.status === 'Completed') return status;
    if (status?.status === 'Failed') {
      throw new Error(`Refresh do dataset falhou: ${status.serviceExceptionJson || 'sem detalhes'}`);
    }
    await new Promise(r => setTimeout(r, intervalMs));
  }
  throw new Error('Tempo limite excedido aguardando conclusão do refresh do Power BI');
}

// --- Análise Executiva via Claude API (server-side, chave nunca exposta ao browser) ---
function buildAnalysisSystemPrompt() {
  return `Você é o Agente S&OE V4 — Sales & Operations Execution. Análise técnica, executiva e defensável. Linguagem direta, sem prolixidade. Cada afirmação suportada pelos dados. Use markdown limpo. Nunca use bullet points decorativos — use dados.`;
}

function buildAnalysisUserMessage(data) {
  const criticals = data.filter(d => d.Status === 'CRITICO');
  const tacRisk = data.filter(d => d.Tipo_Cliente === 'TAC' && d.Flag_Ruptura === 1);
  const totalNeg = data.reduce((s, d) => s + (d.Impacto_EBITDA < 0 ? d.Impacto_EBITDA : 0), 0);

  return `BASE CONSOLIDADA PROCESSADA — JUN/2026 — Grupo Proteína Sul:

${data.map(d => `SKU: ${d.SKU} | Planta: ${d.Planta} | Cliente: ${d.Cliente} [${d.Tipo_Cliente}]
  PVE: ${d.Estoque_PVE}t | Proj.SOE: ${d.Projecao_Estoque_SOE}t | Est.Real: ${d.Estoque_Final_Real}t
  Prod.Plan: ${d.Producao_Planejada}t | Prod.Real: ${d.Producao_Real}t | Margem: R$${d.Margem_EBITDA_Rt}/t
  Status: ${d.Status} | Flag_Ruptura: ${d.Flag_Ruptura} | Flag_Lote: ${d.Flag_Excecao_Lote}
  Impacto_EBITDA: R$ ${d.Impacto_EBITDA.toLocaleString('pt-BR')}
  Raciocínio: ${d.Observacoes}`).join('\n\n')}

RESUMO EXECUTIVO:
- Total SKUs: ${data.length} | CRÍTICOS: ${criticals.length} | ATENÇÃO: ${data.filter(d => d.Status === 'ATENCAO').length}
- Clientes TAC com risco de ruptura: ${tacRisk.length}/${data.filter(d => d.Tipo_Cliente === 'TAC').length}
- Impacto EBITDA negativo total: R$ ${Math.abs(totalNeg).toLocaleString('pt-BR')}

Entregue EXATAMENTE nesta estrutura:

## Diagnóstico Executivo
(3-4 linhas objetivas citando os dados. Sem suavizar o que é crítico.)

## Top Ofensores PVE
(tabela markdown: SKU | PVE | Impacto EBITDA | Prioridade | Ação urgente)

## Clientes TAC
(analise cada cliente TAC separadamente. Status, risco e exposição.)

## Cenários de Decisão

### CONSERVADOR
Lógica: ... | Ação imediata: ... | Risco residual: ...

### EQUILIBRADO
Lógica: ... | Ação imediata: ... | Risco residual: ...

### AGRESSIVO
Lógica: ... | Ação imediata: ... | Risco residual: ...

## [DECISÃO RECOMENDADA]
Uma frase. Cenário + ação específica + justificativa técnica.`;
}

async function callAnthropicAnalysis(data) {
  if (!ANTHROPIC_API_KEY) throw new Error('ANTHROPIC_API_KEY não configurada');
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'anthropic-version': '2023-06-01',
      'x-api-key': ANTHROPIC_API_KEY,
    },
    body: JSON.stringify({
      model: ANTHROPIC_MODEL,
      max_tokens: 1000,
      system: buildAnalysisSystemPrompt(),
      messages: [{ role: 'user', content: buildAnalysisUserMessage(data) }],
    }),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(`Falha na API Anthropic (${res.status}): ${json.error?.message || JSON.stringify(json)}`);
  }
  return json.content?.[0]?.text || '';
}

// --- Alertas por e-mail via Microsoft Graph (sendMail) ---
function buildAlertEmail(data, { period, workspace } = {}) {
  const criticals = data.filter(d => d.Status === 'CRITICO');
  const atencao = data.filter(d => d.Status === 'ATENCAO');
  const rows = data
    .filter(d => d.Status !== 'OK')
    .sort((a, b) => a.Impacto_EBITDA - b.Impacto_EBITDA)
    .map(d => `<tr><td>${d.SKU}</td><td>${d.Status}</td><td>${d.Cliente}</td><td>${d.Estoque_PVE}t</td><td>R$ ${d.Impacto_EBITDA.toLocaleString('pt-BR')}</td><td>${d.Observacoes}</td></tr>`)
    .join('');
  const subject = `[SOE V4] ${criticals.length} CRÍTICOS | ${atencao.length} ATENÇÃO${period ? ` — ${period}` : ''}`;
  const html = `<h2>Alerta S&OE${workspace ? ` — ${workspace}` : ''}${period ? ` — ${period}` : ''}</h2>
    <p>${criticals.length} SKU(s) em status CRÍTICO, ${atencao.length} em ATENÇÃO.</p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:monospace;font-size:12px">
      <tr><th>SKU</th><th>Status</th><th>Cliente</th><th>PVE</th><th>Impacto EBITDA</th><th>Observações</th></tr>
      ${rows}
    </table>`;
  return { subject, html };
}

async function sendAlertEmail({ to, subject, html }) {
  assertConfigured(['GRAPH_SENDER_UPN']);
  const token = await getAccessToken('https://graph.microsoft.com/.default');
  const senderUpn = process.env.GRAPH_SENDER_UPN;
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderUpn)}/sendMail`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: to.split(',').map(addr => ({ emailAddress: { address: addr.trim() } })),
      },
      saveToSentItems: true,
    }),
  });
  if (res.status !== 202) {
    const json = await res.json().catch(() => ({}));
    throw new Error(`Falha ao enviar e-mail (${res.status}): ${json.error?.message || JSON.stringify(json)}`);
  }
}

// --- HTTP server ---
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': FRONTEND_ORIGIN,
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    let exceeded = false;
    req.on('data', chunk => {
      if (exceeded) return;
      body += chunk;
      if (Buffer.byteLength(body) > BODY_LIMIT) {
        exceeded = true;
        reject(Object.assign(new Error('Payload muito grande'), { statusCode: 413 }));
        req.destroy();
      }
    });
    req.on('end', () => {
      if (exceeded) return;
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(Object.assign(new Error('JSON inválido no corpo da requisição'), { statusCode: 400 }));
      }
    });
    req.on('error', reject);
  });
}

function sendJson(res, status, payload) {
  res.writeHead(status, { 'Content-Type': 'application/json', ...CORS_HEADERS });
  res.end(JSON.stringify(payload));
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') {
    res.writeHead(204, CORS_HEADERS);
    res.end();
    return;
  }

  if (req.method === 'POST' && req.url === '/api/soe/save-excel') {
    try {
      const { data } = await readJsonBody(req);
      if (!Array.isArray(data) || data.length === 0) {
        return sendJson(res, 400, { error: 'Campo "data" deve ser um array não vazio com a base consolidada' });
      }
      assertConfigured(REQUIRED_ENV.filter(k => k.startsWith('AZURE') || k.startsWith('SHAREPOINT')));
      const buffer = await buildWorkbook(data);
      const stamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 12);
      const filename = `Painel_SOE_Base_Atualizada_${stamp}.xlsx`;
      const uploadResult = await uploadToSharePoint(Buffer.from(buffer), filename);
      sendJson(res, 200, { ok: true, filename, ...uploadResult });
    } catch (err) {
      sendJson(res, err.statusCode || 500, { error: err.message });
    }
    return;
  }

  if (req.method === 'POST' && req.url === '/api/soe/refresh-powerbi') {
    try {
      const { datasetId, waitForCompletion } = await readJsonBody(req);
      const targetDatasetId = datasetId || process.env.POWERBI_DATASET_ID;
      if (!targetDatasetId) {
        return sendJson(res, 400, { error: 'datasetId não informado e POWERBI_DATASET_ID não configurado' });
      }
      assertConfigured(['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'POWERBI_WORKSPACE_ID']);
      await refreshPowerBIDataset(targetDatasetId);
      const finalStatus = waitForCompletion
        ? await waitForRefreshCompletion(targetDatasetId)
        : { status: 'Refresh iniciado — acompanhamento assíncrono' };
      sendJson(res, 200, { ok: true, datasetId: targetDatasetId, ...finalStatus });
    } catch (err) {
      sendJson(res, err.statusCode || 500, { error: err.message });
    }
    return;
  }

  if (req.method === 'POST' && req.url === '/api/soe/analise-executiva') {
    try {
      const { data } = await readJsonBody(req);
      if (!Array.isArray(data) || data.length === 0) {
        return sendJson(res, 400, { error: 'Campo "data" deve ser um array não vazio com a base consolidada' });
      }
      const text = await callAnthropicAnalysis(data);
      sendJson(res, 200, { ok: true, text });
    } catch (err) {
      sendJson(res, err.statusCode || 500, { error: err.message });
    }
    return;
  }

  if (req.method === 'POST' && req.url === '/api/soe/send-alert') {
    try {
      const { data, to, period, workspace } = await readJsonBody(req);
      if (!Array.isArray(data) || data.length === 0) {
        return sendJson(res, 400, { error: 'Campo "data" deve ser um array não vazio com a base consolidada' });
      }
      const recipients = to || process.env.ALERT_RECIPIENTS;
      if (!recipients) {
        return sendJson(res, 400, { error: 'Destinatário não informado ("to") e ALERT_RECIPIENTS não configurado' });
      }
      assertConfigured(['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'GRAPH_SENDER_UPN']);
      const { subject, html } = buildAlertEmail(data, { period, workspace });
      await sendAlertEmail({ to: recipients, subject, html });
      sendJson(res, 200, { ok: true, to: recipients, subject });
    } catch (err) {
      sendJson(res, err.statusCode || 500, { error: err.message });
    }
    return;
  }

  sendJson(res, 404, { error: 'Not found' });
});

server.listen(PORT, () => {
  console.log(`SOE Agent backend — http://localhost:${PORT}`);
  const missing = REQUIRED_ENV.filter(k => !process.env[k]);
  if (missing.length) {
    console.warn(`Aviso: variáveis de ambiente ausentes (configure em soe-agent/.env): ${missing.join(', ')}`);
  }
  const missingOptional = OPTIONAL_ENV.filter(k => !process.env[k]);
  if (missingOptional.length) {
    console.warn(`Aviso: Etapas 8 (análise) e 9 (alerta) ficarão indisponíveis sem: ${missingOptional.join(', ')}`);
  }
});
