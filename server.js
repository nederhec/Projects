const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.ANTHROPIC_API_KEY || '';

const SYSTEM_PROMPT = `Você é um especialista sênior em Arquitetura de Cargos e Remuneração, com expertise nas metodologias Hay, Mercer, Willis Towers Watson e Korn Ferry. Gera descrições de cargos no padrão das maiores consultorias de RH do Brasil.

Você deve retornar APENAS um JSON válido, sem nenhum texto antes ou depois, sem markdown, sem backticks. O JSON deve seguir exatamente esta estrutura:

{
  "cargo": "título exato do cargo",
  "codigo": "código interno no formato FAM-NNN (ex: TEC-042)",
  "versao": "1.0",
  "data_referencia": "2025",
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
- responsabilidades: 6 a 8 itens variados e específicos
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

function callAnthropic(cargo, segmento, apiKey) {
  return new Promise((resolve, reject) => {
    const userMsg = segmento
      ? `Gere a descrição completa para o cargo: ${cargo}\nSegmento de mercado: ${segmento}`
      : `Gere a descrição completa para o cargo: ${cargo}`;
    const body = JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4096,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMsg }]
    });

    const req = https.request({
      hostname: 'api.anthropic.com',
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'anthropic-version': '2023-06-01',
        'x-api-key': apiKey,
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (res.statusCode !== 200) {
            reject(new Error(parsed.error?.message || `API error ${res.statusCode}`));
            return;
          }
          const raw = (parsed.content || []).map(b => b.text || '').join('').trim()
            .replace(/^```[\w]*\n?/, '').replace(/\n?```$/, '').trim();
          resolve(JSON.parse(raw));
        } catch (e) {
          reject(new Error('Falha ao processar resposta da API'));
        }
      });
    });

    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api/generate') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { cargo, segmento } = JSON.parse(body);
        if (!cargo || typeof cargo !== 'string' || cargo.trim().length === 0) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Campo "cargo" é obrigatório' }));
          return;
        }
        const key = API_KEY;
        if (!key) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'ANTHROPIC_API_KEY não configurada no servidor' }));
          return;
        }
        const seg = typeof segmento === 'string' ? segmento.trim().slice(0, 100) : '';
        const doc = await callAnthropic(cargo.trim().slice(0, 200), seg, key);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(doc));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      }
    });
    return;
  }

  if (req.method === 'GET' && (req.url === '/' || req.url === '/index.html')) {
    const filePath = path.join(__dirname, 'index.html');
    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        res.end('Not found');
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(data);
    });
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, () => {
  console.log(`Arquitetura de Cargos v4 — http://localhost:${PORT}`);
  if (!API_KEY) console.warn('Aviso: ANTHROPIC_API_KEY não definida. Defina antes de iniciar.');
});
