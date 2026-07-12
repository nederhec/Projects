/* Cloudflare Worker — exige Basic Auth antes de servir os arquivos
 * estáticos deste projeto (index.html, app.js, engine.js, vendor/) via
 * o binding `env.ASSETS`. Só entra em vigor quando implantado como
 * Worker no Cloudflare (ver README, seção "Hospedagem"); rodando
 * localmente (npx serve, python -m http.server) ou via file:// isso
 * não existe.
 *
 * `run_worker_first: true` no wrangler.jsonc é obrigatório — o padrão
 * dos Workers é servir os assets estáticos ANTES de rodar o script
 * (o oposto do antigo Pages Functions, que rodava o middleware antes
 * dos assets por padrão). Sem essa opção, este arquivo nunca roda pra
 * requisições que batem em um arquivo estático, e o Basic Auth vira
 * decoração — nada realmente ficaria protegido.
 *
 * As credenciais NUNCA ficam no repositório — vêm de variáveis de
 * ambiente configuradas em Settings > Variables and Secrets do Worker
 * no painel do Cloudflare (runtime), não nas "Build variables" (só
 * existem durante o build, não em produção), como BASIC_AUTH_USER e
 * BASIC_AUTH_PASS.
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

export default {
  async fetch(request, env) {
    if (!env.BASIC_AUTH_USER || !env.BASIC_AUTH_PASS) {
      return new Response(
        'Basic Auth não configurado: defina BASIC_AUTH_USER e BASIC_AUTH_PASS ' +
        'em Settings > Variables and Secrets do Worker (ver README).',
        { status: 500 }
      );
    }

    const header = request.headers.get('Authorization') || '';
    const [scheme, encoded] = header.split(' ');
    if (scheme === 'Basic' && encoded) {
      const credenciais = decodeBasicAuth(encoded);
      if (credenciais && credenciais.user === env.BASIC_AUTH_USER && credenciais.pass === env.BASIC_AUTH_PASS) {
        return env.ASSETS.fetch(request);
      }
    }

    return new Response('Autenticação necessária.', {
      status: 401,
      headers: { 'WWW-Authenticate': `Basic realm="${REALM}", charset="UTF-8"` }
    });
  }
};
