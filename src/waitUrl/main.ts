import 'dotenv/config';


export const aguardarURLCorreta = async (pagina, urlEsperada) => {
  console.log(`Aguardando a navegação manual para a URL: ${urlEsperada}`);
  await pagina.waitForFunction(
    (url) => window.location.href === url,
    { timeout: 100000 },
    urlEsperada
  );
  console.log("Navegação para a URL esperada detectada!");
};