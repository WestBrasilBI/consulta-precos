'use strict';
/* Portão de login Microsoft (MSAL, SPA com PKCE, sem segredo no app).
 * Só autentica e confere o domínio — todos têm o mesmo acesso; é só um filtro de segurança.
 * Usa POPUP (loginPopup): o fluxo acontece numa janelinha e devolve a conta aqui mesmo,
 * sem depender de fragmento/redirect na URL. */

const Auth = (() => {
  const CFG = window.CONFIG || {};
  const ativo = !!(CFG.LOGIN_ATIVO && CFG.CLIENT_ID && CFG.TENANT_ID && window.msal);
  const redirectUri = location.origin + location.pathname.replace(/index\.html$/, '');
  let app = null, conta = null;

  function contaPermitida(c) {
    const email = String(c.username || '').toLowerCase();
    const tid = String(c.tenantId || '').toLowerCase();
    return email.endsWith('@' + String(CFG.DOMINIO_PERMITIDO).toLowerCase()) && tid === String(CFG.TENANT_ID).toLowerCase();
  }

  // Guarda o resultado da última tentativa (diagnóstico leve, se precisar).
  function diag(o) {
    try { localStorage.setItem('consulta_login_diag', JSON.stringify(Object.assign({ quando: new Date().toLocaleString('pt-BR') }, o))); } catch (e) {}
  }

  // Na carga: só checa se já existe uma conta válida logada (SSO silencioso). Não redireciona.
  async function iniciar() {
    if (!ativo) return null;
    app = new msal.PublicClientApplication({
      auth: { clientId: CFG.CLIENT_ID, authority: `https://login.microsoftonline.com/${CFG.TENANT_ID}`, redirectUri },
      cache: { cacheLocation: 'localStorage' },
    });
    await app.initialize();
    await app.handleRedirectPromise().catch(() => {});   // compat, caso algum retorno de redirect
    conta = app.getAllAccounts().find(contaPermitida) || null;
    if (conta) app.setActiveAccount(conta);
    return conta;
  }

  // Clique em "Entrar": abre o popup, confere o domínio e devolve a conta.
  async function entrar() {
    const res = await app.loginPopup({
      scopes: CFG.ESCOPO_LOGIN || ['User.Read'],
      prompt: 'select_account',
      extraQueryParameters: { domain_hint: CFG.DOMINIO_PERMITIDO },
    });
    const c = (res && res.account) || null;
    diag({ via: 'popup', ok: !!(c && contaPermitida(c)), conta: c && c.username, tenant: c && c.tenantId });
    if (c && contaPermitida(c)) { conta = c; app.setActiveAccount(c); return c; }
    if (c) {
      await app.clearCache({ account: c }).catch(() => {});
      throw new Error(`a conta ${c.username} não é da West Brasil. Entre com o seu e-mail @${CFG.DOMINIO_PERMITIDO}.`);
    }
    throw new Error('login não concluído');
  }

  const sair = () => app.logoutPopup({ account: conta }).catch(() => {});

  return {
    get ativo() { return ativo; },
    get conta() { return conta; },
    iniciar, entrar, sair,
    ultimoDiag() { try { return JSON.parse(localStorage.getItem('consulta_login_diag')); } catch (e) { return null; } },
  };
})();
window.Auth = Auth;   // expõe pro script do index (const não vira propriedade de window)
