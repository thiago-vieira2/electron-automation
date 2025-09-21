import 'dotenv/config';
const puppeteer = require("puppeteer");



export const iniciarNavegador = async () => {

      const navegador = await puppeteer.launch({
        headless: false,  
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized'],

        defaultViewport: null
      });

      navegador.on('targetcreated', async (target) => {
        const url = target.url();
        console.log(`Nova aba criada: ${url}`); // Log aqui para ver a URL
        if (url.includes("PoliticaPrivacidade")) {
          const page = await target.page();
          if (page) {
            await page.close();
          }
          console.log(`Nova aba com URL ${url} foi fechada.`);
        }
      });
      
      const pagina = await navegador.newPage();
      
      const urlInicial = "https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2fEntidadesFilantropicas%2fCadastroNotaEntidade.aspx";
      console.log(`Abrindo URL inicial: ${urlInicial}`);
      await pagina.goto(urlInicial, { waitUntil: "domcontentloaded" });

      return { navegador, pagina };
};