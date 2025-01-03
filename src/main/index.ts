import 'dotenv/config';
import { app, shell, BrowserWindow, ipcMain } from 'electron'
import { join } from 'path'
import { electronApp, optimizer, is } from '@electron-toolkit/utils'


const puppeteer = require("puppeteer");
const xlsx = require("xlsx");


const handlePrimeiraColuna = (planilha) => {
  const aba = planilha.SheetNames.length > 0 ? planilha.SheetNames[0] : null;
  if (!aba) {
    console.error("A planilha não possui abas.");
    return [];
  }

  console.log(`Aba selecionada: ${aba}`);
  
  const dados = xlsx.utils.sheet_to_json(planilha.Sheets[aba], { header: 1 });
  if (!dados || dados.length === 0) {
    console.error(`A aba ${aba} não contém dados válidos.`);
    return [];
  }
  console.log("Dados lidos:", dados);

  const primeiraColuna = dados
    .map((linha) => Array.isArray(linha) ? linha[0] : null)
    .filter(value => value !== null);  // Filtra valores nulos
  
  console.log("Primeira coluna:", primeiraColuna);

  return primeiraColuna;
}


const iniciarNavegador = async () => {

      const navegador = await puppeteer.launch({
        headless: false,  
        args: ['--no-sandbox', '--disable-setuid-sandbox']
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



const aguardarURLCorreta = async (pagina, urlEsperada) => {
  console.log(`Aguardando a navegação manual para a URL: ${urlEsperada}`);
  await pagina.waitForFunction(
    (url) => window.location.href === url,
    { timeout: 100000 },
    urlEsperada
  );
  console.log("Navegação para a URL esperada detectada!");
};

const executarAutomacao = async (codigoNota, pagina) => { 
  try {
    if (!codigoNota || typeof codigoNota !== 'string') {
      throw new Error('O código da nota não é válido.');
    }


   
    await pagina.waitForSelector('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]', { visible: true, timeout: 5000 });  

    await new Promise(resolve => setTimeout(resolve, 7000));
    
    await pagina.evaluate((codigo) => {
      navigator.clipboard.writeText(codigo);
    }, codigoNota);
    
  
    await pagina.focus('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
    await pagina.keyboard.down('Control'); 
    await pagina.keyboard.press('V');
    await pagina.keyboard.up('Control'); 

    await pagina.waitForSelector('[value="Salvar Nota"]', { visible: true, timeout: 5000 });
    
    await pagina.click('[value="Salvar Nota"]', { visible: true, timeout: 5000 });

    console.log(`nota salva para: ${codigoNota}`);

    if(pagina.click){
      console.log('nota cadastrada')
    }


  } catch (erro) {
    console.error(`Erro no processo: ${erro}`);
  }
};

// Função principal de manipulação da planilha
const handler = async (planilha) => {
  if (!planilha) {
    throw new Error("Nenhum arquivo fornecido para processamento.");
  }

  try {
    // Aguarda a execução de handlePrimeiraColuna para garantir que a primeira coluna seja extraída corretamente
    const primeiraColuna = await handlePrimeiraColuna(planilha);

    const { pagina } = await iniciarNavegador();

    const urlInicial = "https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2fEntidadesFilantropicas%2fCadastroNotaEntidade.aspx";
    await pagina.goto(urlInicial, { waitUntil: "domcontentloaded" });


    const urlEsperada = "https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx";
    await aguardarURLCorreta(pagina, urlEsperada);


    for (const codigoNota of primeiraColuna) {
      if (!codigoNota || typeof codigoNota !== "string" || codigoNota.trim() === "") {
        console.log(`Valor inválido para Código da Nota: ${codigoNota}`);
        continue;
      }

      try {
        await executarAutomacao(codigoNota, pagina); // Chama a automação para cada código
      } catch (erro) {
        console.error(`Erro ao processar o código ${codigoNota}:`, erro);
      }
    }

    console.log("Automação concluída com sucesso.");
    return "Automação concluída com sucesso!";
  } catch (erro) {
    console.error("Erro no processo:", erro);
    throw new Error("Erro ao processar o arquivo: " + erro);
  }
};


ipcMain.handle('iniciar-navegador', async (_, buffer) => {
  try {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    await handler(workbook);
    console.log("Automação concluída com sucesso.");
  } catch (e: unknown) {
    if (e instanceof Error) {
      console.error('Algo deu errado ao processar a planilha: ', e.message);
      throw new Error(`Erro ao processar a planilha: ${e.message}`);
    }
    console.error('Erro desconhecido:', e);
    throw new Error("Erro desconhecido ao processar a planilha.");
  }
});

// Função para criar a janela principal do Electron
function createWindow() {

  const mainWindow = new BrowserWindow({
    width: 1900,
    height: 2300,
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      sandbox: false,
    },
  });

  mainWindow.on('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url);
    return { action: 'deny' };
  });

  if (is.dev && process.env['ELECTRON_RENDERER_URL']) {
    mainWindow.loadURL(process.env['ELECTRON_RENDERER_URL']);
  } else {
    mainWindow.loadFile(join(__dirname, '../renderer/index.html'));
  }
}

// Quando o app estiver pronto, executa
app.whenReady().then(() => {
  electronApp.setAppUserModelId('com.electron');

  app.on('browser-window-created', (_, window) => {
    optimizer.watchWindowShortcuts(window);
  });

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

// Quando todas as janelas forem fechadas
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
