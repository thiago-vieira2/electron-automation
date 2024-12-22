import { app, shell, BrowserWindow, ipcMain } from 'electron'
import { join } from 'path'
import { electronApp, optimizer, is } from '@electron-toolkit/utils'

import 'dotenv/config';

const puppeteer = require("puppeteer");
const xlsx = require("xlsx");


const handlePrimeiraColuna = (planilha) => {
  const aba = planilha.SheetNames[0];  //leituradosdados provalmente nao e ela o arrombado que ta corropendo 
  console.log(`os dados da variavel aba deu essa bosta aqui ${aba}`);
  
  const dados = xlsx.utils.sheet_to_json(planilha.Sheets[aba], { header: 1 });  //provalmente e esse merdinha 90% de certeza 
  console.log("Dados lidos da variavel dados foi:", dados);


  const primeiraColuna = dados.map((linha) => Array.isArray(linha) ? linha[0] : undefined);
  console.log("Primeira coluna da variavel primeira coluna foi:", primeiraColuna);

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

const aguardarURLCorreta = async (pagina, urlEsperada) => {
  console.log(`Aguardando a navegação manual para a URL: ${urlEsperada}`);
  await pagina.waitForFunction(
    (url) => window.location.href === url,
    { timeout: 0 },
    urlEsperada
  );
  console.log("Navegação para a URL esperada detectada!");
};

const executarAutomacao = async (codigoNota, pagina) => {
  try {
    if (!codigoNota || typeof codigoNota !== 'string') {
      throw new Error('O código da nota não é válido.');
    }

    await pagina.goto("https://www.youtube.com/", { waitUntil: 'domcontentloaded' }); 

    await pagina.waitForSelector('[name="search_query"]', { visible: true, timeout: 5000 });  

    await pagina.type('[name="search_query"]', codigoNota, { delay: 50 });  

    await pagina.waitForSelector('.ytSearchboxComponentSearchButton', { visible: true, timeout: 5000 });
    await pagina.click('.ytSearchboxComponentSearchButton');

    await pagina.waitForNavigation({ waitUntil: 'networkidle2', timeout: 1000 });

    console.log(`Pesquisa realizada para: ${codigoNota}`);

    await pagina.waitForTimeout(500); 

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
    const primeiraColuna = handlePrimeiraColuna(planilha);

    const { navegador, pagina } = await iniciarNavegador();

    for (const codigoNota of primeiraColuna) {
      if (!codigoNota || typeof codigoNota !== "string") {
        console.log(`Valor inválido para Código da Nota: ${codigoNota}`);
        continue;
      }
      await executarAutomacao(codigoNota, pagina);
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
    // Lê o buffer da planilha
    const workbook = xlsx.read(buffer, { type: "buffer" });
    
    // Aguarda a execução da função handler de forma assíncrona
    await handler(workbook);

    console.log("Automação concluída com sucesso.");
  } catch (e: any) {
    console.error('Algo deu errado ao processar a planilha: ', e.message);
    // Pode lançar o erro novamente ou retornar uma mensagem para o frontend, se necessário
    throw new Error(`Erro ao processar a planilha: ${e.message}`);
  }
});
// Função para criar a janela principal do Electron
function createWindow() {

  const mainWindow = new BrowserWindow({
    width: 1900,
    height: 1300,
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
