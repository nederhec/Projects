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
  "segmento": "segmento de mercado informado (ex: Varejo, Tecnologia, Saúde)",
  "familia": "nome da família de cargos (ex: Tecnologia, Finanças, Recursos Humanos)",
  "subfamilia": "subfamília específica (ex: Engenharia de Software, Remuneração & Benefícios)",
  "nivel": "Júnior | Pleno | Sênior | Especialista | Gerência | Diretoria | C-Level",
  "grau": "Operacional | Tático | Estratégico | Estratégico Executivo",
  "proposito": "parágrafo de 2-3 frases descrevendo o propósito central do cargo de forma precisa e objetiva",
  "contribuicao_negocio": "frase única e poderosa descrevendo o impacto no negócio",
  "dimensoes": [
    { "nome": "nome da dimensão", "pct": 30, "descricao": "descrição do tempo/energia dedicado a essa dimensão" }
  ],
  "responsabilidades": [
    { "area": "área/tema", "descricao": "responsabilidade específica e mensurável com verbo no infinitivo" }
  ],
  "competencias_tecnicas": [
    { "nome": "competência técnica", "descricao": "descrição breve de aplicação", "nivel": "Básico | Intermediário | Avançado | Expert" }
  ],
  "competencias_comportamentais": [
    { "nome": "competência comportamental", "descricao": "manifestação esperada no cargo", "nivel": "Básico | Intermediário | Avançado | Expert" }
  ],
  "requisitos": {
    "formacao_minima": "formação mínima exigida",
    "formacao_desejavel": "formação desejável / diferencial",
    "experiencia_minima": "tempo e tipo de experiência mínima",
    "experiencia_desejavel": "experiência diferencial desejável",
    "idiomas": ["idioma e nível (ex: Inglês avançado)"],
    "certificacoes": ["certificação 1", "certificação 2"]
  },
  "kpis": [
    { "nome": "nome do KPI", "descricao": "como é medido e qual a meta de referência" }
  ],
  "progressao": [
    { "cargo": "cargo anterior", "descricao": "perfil típico de quem vem deste cargo", "tipo": "origem" },
    { "cargo": "cargo atual", "descricao": "posição em foco neste documento", "tipo": "atual" },
    { "cargo": "cargo destino 1", "descricao": "progressão vertical típica", "tipo": "destino" },
    { "cargo": "cargo destino 2", "descricao": "progressão lateral possível", "tipo": "destino" }
  ]
}

Regras:
- dimensoes: 4 a 6 itens, soma dos pct = 100
- responsabilidades: 6 a 8 itens variados e específicos
- competencias_tecnicas: 5 a 7 itens relevantes para o cargo
- competencias_comportamentais: 4 a 6 itens
- kpis: 4 a 6 indicadores mensuráveis
- progressao: sempre 4 itens (1 origem, 1 atual, 2 destino)
- Todo o conteúdo em português do Brasil
- Use linguagem corporativa precisa, sem genéricos
- Calibre o nível ao cargo informado (pleno ≠ sênior ≠ especialista)
- Use o segmento informado para calibrar linguagem, KPIs, requisitos e exemplos com a realidade daquele mercado`;

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
