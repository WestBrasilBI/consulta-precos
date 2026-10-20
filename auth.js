'use strict';
/* Portão de login Microsoft (MSAL, SPA com PKCE, sem segredo no app).
 * Só autentica e confere o domínio — o Consulta de Preços não acessa banco nem arquivos.
 * Mesmo padrão do app Pesquisa de Preços. */

const Auth = (() => {
  const CFG = window.CONFIG || {};
  const ativo = !!(CFG.LOGIN_ATIVO && CFG.CLIENT_ID && CFG.TENANT_ID && window.msal);
  // redirect = endereço do app sem o index.html (tem que bater com o cadastrado na TI)
  const redirectUri = location.origin + location.pathname.replace(/index\.html$/, '');
  let app = null, conta = null;

  function contaPermitida(c) {
    const email = String(c.username || '').toLowerCase();
    return email.endsWith('@' + String(CFG.DOMINIO_PERMITIDO).toLowerCase()) && c.tenantId === CFG.TENANT_ID;
  }

  async function iniciar() {
    if (!ativo) return null;                 // login desligado = libera o app
    app = new msal.PublicClientApplication({
      auth: { clientId: CFG.CLIENT_ID, authority: `https://login.microsoftonline.com/${CFG.TENANT_ID}`, redirectUri },
      cache: { cacheLocation: 'localStorage' },
    });
    await app.initialize();
    const r = await app.handleRedirectPromise();
    // só contas @westbrasil.com.br do tenant da West; convidados/externos são recusados
    const contas = [(r && r.account), ...app.getAllAccounts()].filter(Boolean);
    const recusada = contas.find(c => !contaPermitida(c));
    conta = contas.find(contaPermitida) || null;
    if (recusada && !conta) {
      await app.clearCache({ account: recusada }).catch(() => {});
      throw new Error(`a conta ${recusada.username} não é da West Brasil. Entre com o seu e-mail @${CFG.DOMINIO_PERMITIDO}.`);
    }
    if (conta) app.setActiveAccount(conta);
    return conta;
  }

  const entrar = () => app.loginRedirect({
    scopes: CFG.ESCOPO_LOGIN || ['User.Read'],
    prompt: 'select_account',
    extraQueryParameters: { domain_hint: CFG.DOMINIO_PERMITIDO },
  });
  const sair = () => app.logoutRedirect({ account: conta, postLogoutRedirectUri: redirectUri });

  return {
    get ativo() { return ativo; },
    get conta() { return conta; },
    iniciar, entrar, sair,
  };
})();
window.Auth = Auth;   // expõe pro script do index (const não vira propriedade de window)
