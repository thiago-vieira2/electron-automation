import 'dotenv/config';
import xlsx from 'xlsx';
import { app, shell, BrowserWindow, ipcMain } from 'electron'
import { join } from 'path'
import { electronApp, optimizer, is } from '@electron-toolkit/utils';
import { handlePrimeiraColuna } from '../ReadFirstColumn/main';
import { iniciarNavegador } from '../initBrowser/main';
import { aguardarURLCorreta } from '../waitUrl/main';
import { executarAutomacao } from '../executeAutomation/main';



// Função principal de manipulação da planilha
const handler = async (planilha) => {
  if (!planilha) {
    throw new Error("Nenhum arquivo fornecido para processamento.");
  }

  try {
    const primeiraColuna = await handlePrimeiraColuna(planilha);
    const notasFiltradas = primeiraColuna.filter((nota, index, self) => self.indexOf(nota) === index);
    console.log(`Notas únicas a serem processadas: ${notasFiltradas.length}`)
    const { navegador, pagina } = await iniciarNavegador();

    const urlInicial = "https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2fEntidadesFilantropicas%2fCadastroNotaEntidade.aspx";
    await pagina.goto(urlInicial, { waitUntil: "domcontentloaded" });

    const urlEsperada = "https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx";
    await aguardarURLCorreta(pagina, urlEsperada);

    for (const [index, codigoNota] of notasFiltradas.entries()) {
      console.log(`Processando nota ${index + 1}/${notasFiltradas.length}: ${codigoNota}`);

      try {
        await executarAutomacao(codigoNota, pagina);
        await new Promise(resolve => setTimeout(resolve, 2000)); // Pausa de 2 segundos entre notas
      } catch (erro) {
        console.error(`Erro ao processar a nota ${codigoNota}:`, erro);
      }
    }

    console.log("Automação concluída com sucesso.");
    await navegador.close();
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
    width: 1300,
    height: 700,
   
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