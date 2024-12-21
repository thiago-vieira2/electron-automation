import { app, shell, BrowserWindow, ipcMain } from 'electron'
import { join } from 'path'
import { electronApp, optimizer, is } from '@electron-toolkit/utils'
import icon from '../../resources/icon.png?asset'
import 'dotenv/config';

const puppeteer = require("puppeteer");
const xlsx = require("xlsx");

const handlePrimeiraColuna = async (planilha) => {
  const aba = planilha.SheetNames[0];  
  const dados = xlsx.utils.sheet_to_json(planilha.Sheets[aba], { header: 1 }); 

  const primeiraColuna = dados.map(linha => linha[0]);  
  console.log(primeiraColuna);
  return primeiraColuna;
}

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
    if (!codigoNota || typeof codigoNota !== "string") {
      throw new Error("O código da nota não é válido.");
    }

    await pagina.waitForSelector('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]', {
      visible: true,
      timeout: 10000,
    });

    await pagina.evaluate((codigo: string) => {
      try {
        const input = document.querySelector<HTMLInputElement>('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
        if (input) {
          const nativeValueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
          if (nativeValueSetter) {
            nativeValueSetter.call(input, codigo);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            console.log('Valor definido com sucesso no campo de entrada.');
          } else {
            console.error('Não foi possível definir o valor: setter nativo não encontrado.');
          }
        } else {
          console.error('Elemento de entrada não encontrado.');
        }
      } catch (error) {
        console.error('Erro ao tentar definir o valor no campo de entrada:', error);
      }
    }, codigoNota);

    await pagina.waitForSelector('[value="Salvar Nota"]', { visible: true, timeout: 1000 });
    await pagina.click('[value="Salvar Nota"]');

    await pagina.evaluate(() => {
      const input = document.querySelector<HTMLInputElement>('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
      if (input) {
        const nativeValueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
        if (nativeValueSetter) {
          nativeValueSetter.call(input, "");
          input.dispatchEvent(new Event('input', { bubbles: true }));
        } else {
          console.error('Não foi possível acessar o setter nativo para o valor do input.');
        }
      } else {
        console.error('Elemento de entrada não encontrado.');
      }
    });

    await pagina.click('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
    await new Promise(resolve => setTimeout(resolve, 2500)); 

  } catch (erro) {
    console.error(`Erro no processo: ${erro}`);
  }
};

const handler = async (planilha) => {
  if (!planilha) {
    throw new Error("Nenhum arquivo fornecido para processamento.");
  }

  try {
    // Aguarda a execução de handlePrimeiraColuna para garantir que a primeira coluna seja extraída corretamente
    const primeiraColuna = await handlePrimeiraColuna(planilha);

    const { navegador, pagina } = await iniciarNavegador();

    const urlInicial = "https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2fEntidadesFilantropicas%2fCadastroNotaEntidade.aspx";
    await pagina.goto(urlInicial, { waitUntil: "domcontentloaded" });

    const urlEsperada = "https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx";
    await aguardarURLCorreta(pagina, urlEsperada);

    // Loop para processar cada código de nota
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
  } catch (erro: any) {
    console.error("Erro no processo:", erro);
    throw new Error("Erro ao processar o arquivo: " + erro.message);
  }
};




//----------------------------------------------------------
function createWindow(): void {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 900,
    height: 670,
    show: false,
    autoHideMenuBar: true,
    ...(process.platform === 'linux' ? { icon } : {}),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      sandbox: false
    }
  })

  mainWindow.on('ready-to-show', () => {
    mainWindow.show()
  })

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url)
    return { action: 'deny' }
  })

  // HMR for renderer base on electron-vite cli.
  // Load the remote URL for development or the local html file for production.
  if (is.dev && process.env['ELECTRON_RENDERER_URL']) {
    mainWindow.loadURL(process.env['ELECTRON_RENDERER_URL'])
  } else {
    mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

ipcMain.handle('iniciar-navegador', async (_, buffer) => {
  try {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    handler(workbook);
  }catch(e: any) {
    console.log('Something went wrong! ', e.message);
  }
});

app.whenReady().then(() => {
  // Set app user model id for windows
  electronApp.setAppUserModelId('com.electron')

  // Default open or close DevTools by F12 in development
  // and ignore CommandOrControl + R in production.
  // see https://github.com/alex8088/electron-toolkit/tree/master/packages/utils
  app.on('browser-window-created', (_, window) => {
    optimizer.watchWindowShortcuts(window)
  })

  createWindow()

  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})