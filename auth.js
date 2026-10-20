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
    const tid = String(c.tenantId || '').toLowerCase();
    return email.endsWith('@' + String(CFG.DOMINIO_PERMITIDO).toLowerCase()) && tid === String(CFG.TENANT_ID).toLowerCase();
  }

  // Grava o que aconteceu na última tentativa (ajuda a diagnosticar sem precisar do console).
  function diag(o) {
    try { localStorage.setItem('consulta_login_diag', JSON.stringify(Object.assign({ quando: new Date().toLocaleString('pt-BR'), href: location.href }, o))); } catch (e) {}
  }

  // Log interno do MSAL guardado na tela (anel de ~40 linhas) p/ diagnóstico.
  function logMsal(msg) {
    try {
      const arr = JSON.parse(localStorage.getItem('consulta_msal_log') || '[]');
      arr.push(String(msg).replace(/^\[[^\]]*\]\s*/, '').slice(0, 180));
      while (arr.length > 40) arr.shift();
      localStorage.setItem('consulta_msal_log', JSON.stringify(arr));
    } catch (e) {}
  }

  function contaKeys(store) {
    try { return Object.keys(store).filter(k => /msal/i.test(k)).length; } catch (e) { return -1; }
  }

  async function iniciar() {
    if (!ativo) return null;                 // login desligado = libera o app
    const hashIni = location.hash || '';
    const hashTemCode = /[#&?](code|error)=/.test(hashIni + location.search);
    app = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${CFG.TENANT_ID}`,
        redirectUri,
        navigateToLoginRequestUrl: false,   // processa a resposta aqui mesmo, sem navegar de novo
      },
      // storeAuthStateInCookie: o estado do login (state/nonce/PKCE) vai tbm pra cookie,
      // sobrevivendo ao vai-e-volta mesmo quando o sessionStorage se perde.
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true },
      system: {
        loggerOptions: {
          loggerCallback: (level, message, containsPii) => { if (!containsPii) logMsal(message); },
          piiLoggingEnabled: false,
          logLevel: (msal.LogLevel ? msal.LogLevel.Verbose : 3),
        },
      },
    });
    try { localStorage.removeItem('consulta_msal_log'); } catch (e) {}  // só o log desta tentativa
    await app.initialize();
    let r = null, erro = null;
    try {
      r = await app.handleRedirectPromise();
    } catch (e) {
      erro = String((e && e.message) || e);
    }
    // só contas @westbrasil.com.br do tenant da West; convidados/externos são recusados
    const todas = app.getAllAccounts();
    const contas = [(r && r.account), ...todas].filter(Boolean);
    const recusada = contas.find(c => !contaPermitida(c));
    conta = contas.find(contaPermitida) || null;
    diag({ hashTemCode, hashIni: hashIni.slice(0, 18), teveResposta: !!r, nContas: todas.length,
           contas: contas.map(c => ({ u: c.username, t: c.tenantId })),
           tenantCfg: CFG.TENANT_ID, aceita: !!conta,
           msalLS: contaKeys(localStorage), msalSS: contaKeys(sessionStorage), erro });
    if (erro && !conta) throw new Error('não consegui concluir o login: ' + erro);
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
    ultimoDiag() { try { return JSON.parse(localStorage.getItem('consulta_login_diag')); } catch (e) { return null; } },
    msalLog() { try { return JSON.parse(localStorage.getItem('consulta_msal_log') || '[]'); } catch (e) { return []; } },
  };
})();
window.Auth = Auth;   // expõe pro script do index (const não vira propriedade de window)
