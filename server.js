const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const ANTHROPIC_KEY    = process.env.ANTHROPIC_API_KEY   || '';
const OPENAI_KEY       = process.env.OPENAI_API_KEY      || '';
const OPENROUTER_KEY   = process.env.OPENROUTER_API_KEY  || '';
const OPENAI_MODEL     = process.env.OPENAI_MODEL        || 'gpt-4o-mini';
const OPENROUTER_MODEL = process.env.OPENROUTER_MODEL    || 'anthropic/claude-sonnet-4-6';
const BODY_LIMIT       = 16 * 1024;
const API_TIMEOUT_MS   = 90_000;

function getProvider() {
  if (ANTHROPIC_KEY)  return 'anthropic';
  if (OPENAI_KEY)     return 'openai';
  if (OPENROUTER_KEY) return 'openrouter';
  return null;
}

function buildSystemPrompt() {
  const year = new Date().getFullYear();
  return `Você é um especialista sênior em Arquitetura de Cargos e Remuneração, com expertise nas metodologias Hay, Mercer, Willis Towers Watson e Korn Ferry. Gera descrições de cargos no padrão das maiores consultorias de RH do Brasil.

Você deve retornar APENAS um JSON válido, sem nenhum texto antes ou depois, sem markdown, sem backticks. O JSON deve seguir exatamente esta estrutura:

{
  "cargo": "título exato do cargo",
  "codigo": "código interno no formato FAM-NNN (ex: TEC-042)",
  "versao": "1.0",
  "data_referencia": "${year}",
  "segmento": "segmento de mercado informado",
  "familia": "nome da família de cargos",
  "subfamilia": "subfamília específica",
  "nivel": "Júnior | Pleno | Sênior | Especialista | Gerência | Diretoria | C-Level",
  "grau": "Operacional | Tático | Estratégico | Estratégico Executivo",
  "proposito": "parágrafo de 2-3 frases sobre o propósito central do cargo",
  "contribuicao_negocio": "frase única e poderosa sobre o impacto no negócio",
  "responsabilidades": [
    { "area": "área/tema", "descricao": "responsabilidade específica com verbo no infinitivo" }
  ],
  "competencias_tecnicas": [
    { "nome": "nome da competência técnica", "nivel": "Básico | Intermediário | Avançado | Expert" }
  ],
  "competencias_comportamentais": [
    { "nome": "nome da competência comportamental", "nivel": "Básico | Intermediário | Avançado | Expert" }
  ],
  "requisitos": {
    "formacao_minima": "formação mínima exigida",
    "formacao_desejavel": "formação desejável / diferencial",
    "experiencia_minima": "tempo e tipo de experiência mínima",
    "experiencia_desejavel": "experiência diferencial desejável"
  },
  "complexidade": {
    "profundidade_analitica": 0,
    "densidade_tecnica": 0,
    "influencia_estrategica": 0,
    "transformacao_inovacao": 0,
    "autonomia_decisao": 0
  },
  "avaliacao_risco": {
    "VAR1": false,
    "VAR2": false,
    "VAR3": false,
    "VAR4": false,
    "VAR5": false,
    "VAR6": false
  }
}

Regras:
- responsabilidades: 7 a 10 itens variados e específicos
- competencias_tecnicas: 5 a 7 itens
- competencias_comportamentais: 4 a 6 itens
- complexidade: avalie cada dimensão de 0 a 100 conforme o cargo e segmento
  - profundidade_analitica: nível de raciocínio analítico e diagnóstico exigido
  - densidade_tecnica: volume e profundidade de conhecimento técnico/metodológico
  - influencia_estrategica: capacidade de influenciar decisões e alinhar stakeholders
  - transformacao_inovacao: capacidade de transformar processos e gerar inovação aplicada
  - autonomia_decisao: amplitude de autonomia e responsabilidade por decisões
- avaliacao_risco: responda true/false para cada variável:
  - VAR1: a atividade possui acesso a informações sensíveis de PLDFT
  - VAR2: possui alçada para deliberar reportes em relação a PLDFT
  - VAR3: possui alçada para autorizar operações financeiras
  - VAR4: possui contato direto com o cliente
  - VAR5: possui acesso para incluir/alterar dados de cliente/operações
  - VAR6: comercializa produtos/serviços ou oferta propostas de transações financeiras
- Todo o conteúdo em português do Brasil
- Use linguagem corporativa precisa, sem genéricos
- Calibre o nível ao cargo (pleno ≠ sênior ≠ especialista)
- Use o segmento para calibrar linguagem, requisitos e complexidade`;
}

const rateLimitMap = new Map();
const RATE_LIMIT_COUNT  = 10;
const RATE_LIMIT_WINDOW = 60_000;

setInterval(() => {
  const now = Date.now();
  for (const [ip, entry] of rateLimitMap) {
    if (now > entry.resetAt) rateLimitMap.delete(ip);
  }
}, 5 * 60_000).unref();

function isRateLimited(ip) {
  const now = Date.now();
  const entry = rateLimitMap.get(ip);
  if (!entry || now > entry.resetAt) {
    rateLimitMap.set(ip, { count: 1, resetAt: now + RATE_LIMIT_WINDOW });
    return false;
  }
  if (entry.count >= RATE_LIMIT_COUNT) return true;
  entry.count++;
  return false;
}

const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'X-XSS-Protection': '1; mode=block',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
};

function parseJsonResponse(raw) {
  const cleaned = raw.trim().replace(/^```[\w]*\n?/, '').replace(/\n?```$/, '').trim();
  return JSON.parse(cleaned);
}

function makeRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (res.statusCode !== 200) {
            reject(new Error(parsed.error?.message || `API error ${res.statusCode}`));
            return;
          }
          resolve(parsed);
        } catch {
          reject(new Error('Resposta da API em formato inesperado'));
        }
      });
    });
    req.setTimeout(API_TIMEOUT_MS, () => req.destroy(new Error('Tempo limite da API excedido')));
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function callAnthropic(cargo, segmento) {
  const userMsg = segmento
    ? `Gere a descrição completa para o cargo: ${cargo}\nSegmento de mercado: ${segmento}`
    : `Gere a descrição completa para o cargo: ${cargo}`;
  const body = JSON.stringify({
    model: 'claude-sonnet-4-6',
    max_tokens: 4096,
    system: buildSystemPrompt(),
    messages: [{ role: 'user', content: userMsg }]
  });
  return makeRequest({
    hostname: 'api.anthropic.com',
    path: '/v1/messages',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'anthropic-version': '2023-06-01',
      'x-api-key': ANTHROPIC_KEY,
      'Content-Length': Buffer.byteLength(body)
    }
  }, body).then(parsed => parseJsonResponse(
    (parsed.content || []).map(b => b.text || '').join('')
  ));
}

function callChatCompletions({ hostname, path: apiPath, apiKey, model, extraHeaders = {} }, cargo, segmento) {
  const userMsg = segmento
    ? `Gere a descrição completa para o cargo: ${cargo}\nSegmento de mercado: ${segmento}`
    : `Gere a descrição completa para o cargo: ${cargo}`;
  const body = JSON.stringify({
    model,
    max_tokens: 4096,
    messages: [
      { role: 'system', content: buildSystemPrompt() },
      { role: 'user', content: userMsg }
    ]
  });
  return makeRequest({
    hostname,
    path: apiPath,
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
      'Content-Length': Buffer.byteLength(body),
      ...extraHeaders
    }
  }, body).then(parsed => parseJsonResponse(
    parsed.choices?.[0]?.message?.content || ''
  ));
}

function callOpenAI(cargo, segmento) {
  return callChatCompletions({
    hostname: 'api.openai.com',
    path: '/v1/chat/completions',
    apiKey: OPENAI_KEY,
    model: OPENAI_MODEL,
  }, cargo, segmento);
}

function callOpenRouter(cargo, segmento) {
  return callChatCompletions({
    hostname: 'openrouter.ai',
    path: '/api/v1/chat/completions',
    apiKey: OPENROUTER_KEY,
    model: OPENROUTER_MODEL,
    extraHeaders: {
      'HTTP-Referer': 'https://arquitetura-de-cargos.local',
      'X-Title': 'Arquitetura de Cargos',
    }
  }, cargo, segmento);
}

function callLLM(cargo, segmento) {
  const provider = getProvider();
  if (provider === 'anthropic')  return callAnthropic(cargo, segmento);
  if (provider === 'openai')     return callOpenAI(cargo, segmento);
  if (provider === 'openrouter') return callOpenRouter(cargo, segmento);
  throw new Error('Nenhuma chave de API configurada.');
}

const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api/generate') {
    const ip = req.socket.remoteAddress || 'unknown';
    if (isRateLimited(ip)) {
      res.writeHead(429, { 'Content-Type': 'application/json', 'Retry-After': '60', ...SECURITY_HEADERS });
      res.end(JSON.stringify({ error: 'Muitas requisições. Aguarde um minuto antes de tentar novamente.' }));
      return;
    }

    let body = '';
    let exceeded = false;

    req.on('data', chunk => {
      if (exceeded) return;
      body += chunk;
      if (Buffer.byteLength(body) > BODY_LIMIT) {
        exceeded = true;
        res.writeHead(413, { 'Content-Type': 'application/json', ...SECURITY_HEADERS });
        res.end(JSON.stringify({ error: 'Payload muito grande' }));
        req.destroy();
      }
    });

    req.on('end', async () => {
      if (exceeded) return;
      try {
        let parsed;
        try {
          parsed = JSON.parse(body);
        } catch {
          res.writeHead(400, { 'Content-Type': 'application/json', ...SECURITY_HEADERS });
          res.end(JSON.stringify({ error: 'JSON inválido no corpo da requisição' }));
          return;
        }

        const { cargo, segmento } = parsed;
        if (!cargo || typeof cargo !== 'string' || cargo.trim().length === 0) {
          res.writeHead(400, { 'Content-Type': 'application/json', ...SECURITY_HEADERS });
          res.end(JSON.stringify({ error: 'Campo "cargo" é obrigatório' }));
          return;
        }
        if (!getProvider()) {
          res.writeHead(500, { 'Content-Type': 'application/json', ...SECURITY_HEADERS });
          res.end(JSON.stringify({ error: 'Nenhuma chave de API configurada. Defina ANTHROPIC_API_KEY, OPENAI_API_KEY ou OPENROUTER_API_KEY.' }));
          return;
        }

        const seg = typeof segmento === 'string' ? segmento.trim().slice(0, 100) : '';
        const doc = await callLLM(cargo.trim().slice(0, 200), seg);
        res.writeHead(200, { 'Content-Type': 'application/json', ...SECURITY_HEADERS });
        res.end(JSON.stringify(doc));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json', ...SECURITY_HEADERS });
        res.end(JSON.stringify({ error: err.message }));
      }
    });
    return;
  }

  if (req.method === 'GET' && (req.url === '/' || req.url === '/index.html')) {
    const filePath = path.join(__dirname, 'index.html');
    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404, SECURITY_HEADERS);
        res.end('Not found');
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', ...SECURITY_HEADERS });
      res.end(data);
    });
    return;
  }

  res.writeHead(404, SECURITY_HEADERS);
  res.end('Not found');
});

server.listen(PORT, () => {
  const provider = getProvider();
  const label = {
    anthropic:  'Anthropic (claude-sonnet-4-6)',
    openai:     `OpenAI (${OPENAI_MODEL})`,
    openrouter: `OpenRouter (${OPENROUTER_MODEL})`,
  }[provider] || null;
  console.log(`Arquitetura de Cargos v4 — http://localhost:${PORT}`);
  if (label) console.log(`Provedor: ${label}`);
  else console.warn('Aviso: nenhuma chave de API configurada. Defina ANTHROPIC_API_KEY, OPENAI_API_KEY ou OPENROUTER_API_KEY.');
});
