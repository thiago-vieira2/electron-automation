const puppeteer = require("puppeteer");
const xlsx = require("xlsx");

const lerPrimeiraColuna = (caminho) => {
  const planilha = xlsx.readFile(caminho); 
  const aba = planilha.SheetNames[0];  
  const dados = xlsx.utils.sheet_to_json(planilha.Sheets[aba], { header: 1 }); 

  const primeiraColuna = dados.map(linha => linha[0]);  

  return primeiraColuna;
};

const iniciarNavegador = async () => {
  const navegador = await puppeteer.launch({
    headless: false,  
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const pagina = await navegador.newPage();
  return { navegador, pagina };
};

const executarAutomacao = async (codigoNota, pagina) => {
  try {
    if (!codigoNota || typeof codigoNota !== 'string') {
      throw new Error("O código da nota não é válido.");
    }

    await pagina.goto("https://www.youtube.com/", { waitUntil: "domcontentloaded" });

    await pagina.waitForSelector('[name="search_query"]', { visible: true, timeout: 5000 });
    await pagina.type('[name="search_query"]', codigoNota, { delay: 50 });

    await pagina.waitForSelector("#search-icon-legacy", { visible: true, timeout: 5000 });
    await pagina.click("#search-icon-legacy");

    await pagina.waitForNavigation({ waitUntil: "networkidle2" });

    console.log(`Pesquisa realizada para: ${codigoNota}`);
    await pagina.waitForTimeout(500);
  } catch (erro) {
    console.error(`Erro no processo: ${erro}`);
  }
};

const handler = async ({ path: filePath }) => {
  if (!filePath) {
    throw new Error("Nenhum arquivo fornecido para processamento.");
  }

  try {
    const primeiraColuna = lerPrimeiraColuna(filePath);

    const { navegador, pagina } = await iniciarNavegador();

    for (const codigoNota of primeiraColuna) {
      if (!codigoNota || typeof codigoNota !== "string") {
        console.log(`Valor inválido para Código da Nota: ${codigoNota}`);
        continue;
      }
      await executarAutomacao(codigoNota, pagina);
    }

    await navegador.close();
    console.log("Automação concluída com sucesso.");
    return "Automação concluída com sucesso!";
  } catch (erro: any) {
    console.error("Erro no processo:", erro);
    throw new Error("Erro ao processar o arquivo: " + erro.message);
  }
};

module.exports = {
  handler,
  lerPrimeiraColuna,
};