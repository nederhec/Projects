/* Cloudflare Pages Functions middleware — exige Basic Auth antes de servir
 * qualquer arquivo estático deste projeto (index.html, app.js, engine.js,
 * vendor/, etc). Só entra em vigor quando este diretório é implantado como
 * o Pages project (ver README, seção "Hospedagem"); rodando localmente
 * (npx serve, python -m http.server) ou via file:// isso não existe.
 *
 * As credenciais NUNCA ficam no repositório — vêm de variáveis de ambiente
 * configuradas no painel do Cloudflare Pages (Settings > Environment
 * variables), como "Secret": BASIC_AUTH_USER e BASIC_AUTH_PASS.
 */

function decodeBasicAuth(base64) {
  const binary = atob(base64);
  const bytes = Uint8Array.from(binary, (c) => c.charCodeAt(0));
  const decoded = new TextDecoder().decode(bytes);
  const sep = decoded.indexOf(':');
  if (sep === -1) return null;
  return { user: decoded.slice(0, sep), pass: decoded.slice(sep + 1) };
}

const REALM = 'Reconciliação Independente da Base';

export async function onRequest(context) {
  const { request, next, env } = context;

  if (!env.BASIC_AUTH_USER || !env.BASIC_AUTH_PASS) {
    return new Response(
      'Basic Auth não configurado: defina BASIC_AUTH_USER e BASIC_AUTH_PASS ' +
      'nas variáveis de ambiente do projeto no Cloudflare Pages (ver README).',
      { status: 500 }
    );
  }

  const header = request.headers.get('Authorization') || '';
  const [scheme, encoded] = header.split(' ');
  if (scheme === 'Basic' && encoded) {
    const credenciais = decodeBasicAuth(encoded);
    if (credenciais && credenciais.user === env.BASIC_AUTH_USER && credenciais.pass === env.BASIC_AUTH_PASS) {
      return next();
    }
  }

  return new Response('Autenticação necessária.', {
    status: 401,
    headers: { 'WWW-Authenticate': `Basic realm="${REALM}", charset="UTF-8"` }
  });
}
