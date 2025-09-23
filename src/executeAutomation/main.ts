import { Page } from "puppeteer";




export const executarAutomacao = async (codigoNota: string, pagina: Page) => { 

 
  try {

    let contador = 0

    if (!codigoNota || typeof codigoNota !== 'string') {
      throw new Error('O código da nota não é válido.');
    }

    // aguarda o seletor do input
    await pagina.waitForSelector('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]', { visible: true, timeout: 5000 });

    await pagina.focus('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');

  
    await pagina.keyboard.down('Control');
    await pagina.keyboard.press('A'); // seleciona todo o texto no campo
    await pagina.keyboard.up('Control');
    await pagina.keyboard.press('Backspace'); // apaga o texto selecionado
    

    await new Promise(resolve => setTimeout(resolve, 2000));
    // insere a nova nota
    await pagina.evaluate((codigo) => {
      navigator.clipboard.writeText(codigo);
    },codigoNota);

    await pagina.click('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
    await pagina.keyboard.down('Control'); 
    await pagina.keyboard.press('V'); // cola a nova nota
    await pagina.keyboard.up('Control'); 
    

    await new Promise(resolve => setTimeout(resolve, 3000));
    await pagina.evaluate(() => {
      window.scrollBy(0, 500); // cola 500 pixels para baixo
    });

    // aguarda o botão de salvar e clica
    await pagina.waitForSelector('[value="Salvar Nota"]', { visible: true, timeout: 4000 });

    await new Promise(resolve => setTimeout(resolve, 4000));

    await pagina.click('[value="Salvar Nota"]', { visible: true, timeout: 4000 });

    // verifica o texto do span para erros
    const spanText = await pagina.evaluate(() => {
      const span = document.querySelector('#lblErro'); 
      return span && span.textContent ? span.textContent.trim() : null;
    });

    if (spanText && spanText.includes('Este pedido já existe no sistema. Favor inserir uma nova nota.')) { // Ajuste a mensagem específica do erro
      contador++
      console.log(`Nota ${codigoNota} já foi cadastrada. Pulando para a próxima. ${contador}`);
      return; // sai da função e passa para a próxima nota
    }

    

    console.log(`Nota cadastrada com sucesso: ${codigoNota}`);

  } catch (erro) {
    console.error(`Erro no processo para a nota ${codigoNota}:`, erro);
  }
};