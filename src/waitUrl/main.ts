import 'dotenv/config';
import { Page } from 'puppeteer';


export const aguardarURLCorreta = async (pagina: Page, urlEsperada: string) => {
  console.log(`Aguardando a navegação manual para a URL: ${urlEsperada}`);
  await pagina.waitForFunction(
    (url) => window.location.href === url,
    { timeout: 100000 },
    urlEsperada
  );
  console.log("Navegação para a URL esperada detectada!");
};