/* Login Microsoft (Entra ID) do app Consulta de Preços.
 *
 * >>> AÇÃO DA TI (obrigatória para o login funcionar) <<<
 * Autorizar o endereço do app como "redirect URI" (plataforma SPA) num registro de app do Entra ID:
 *     https://westbrasilbi.github.io/consulta-precos/
 * Caminho mais simples: reaproveitar o MESMO registro do app Pesquisa de Preços
 * ("West Brasil - Pesquisa de Preços", CLIENT_ID abaixo) e só ADICIONAR esse redirect.
 * O Consulta NÃO chama o Fabric (os preços ficam embutidos em data/data.js); ele só
 * autentica o usuário e confere o domínio. Basta o login (User.Read já está consentido).
 * Se a TI preferir um registro NOVO só para o Consulta, troque o CLIENT_ID abaixo pelo novo.
 */
window.CONFIG = {
  // Liga/desliga a exigência de login. Deixe false para publicar SEM login enquanto a TI
  // não autoriza o redirect; vire true quando o redirect estiver autorizado.
  LOGIN_ATIVO: true,

  TENANT_ID: '1516fbda-e5c5-429f-8c31-bf8dc4dedc55',   // WEST
  CLIENT_ID: '7b109c34-9e32-47c8-970d-4f03ed0c0d08',   // West Brasil - Pesquisa de Preços (reaproveitado)

  // Só entra conta deste domínio (e do tenant acima). Convidados externos são recusados.
  DOMINIO_PERMITIDO: 'westbrasil.com.br',

  // Escopo só para entrar e identificar o usuário (sem acesso a dados).
  ESCOPO_LOGIN: ['User.Read'],
};
