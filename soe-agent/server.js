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

const REQUIRED_ENV = [
  'AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET',
  'SHAREPOINT_SITE_ID', 'POWERBI_WORKSPACE_ID', 'POWERBI_DATASET_ID',
];

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

  sendJson(res, 404, { error: 'Not found' });
});

server.listen(PORT, () => {
  console.log(`SOE Agent backend — http://localhost:${PORT}`);
  const missing = REQUIRED_ENV.filter(k => !process.env[k]);
  if (missing.length) {
    console.warn(`Aviso: variáveis de ambiente ausentes (configure em soe-agent/.env): ${missing.join(', ')}`);
  }
});
