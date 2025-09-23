import 'dotenv/config';

import xlsx from 'xlsx'; 


export const handlePrimeiraColuna = (planilha) => {
  const aba = planilha.SheetNames.length > 0 ? planilha.SheetNames[0] : null;
  if (!aba) {
    console.error("A planilha não possui abas.");
    return [];
  }

  console.log(`Aba selecionada: ${aba}`);
  
  // Lê os dados da aba como array de arrays
  const dados = xlsx.utils.sheet_to_json(planilha.Sheets[aba], { header: 1});
  if (!dados || dados.length === 0) {
    console.error(`A aba ${aba} não contém dados válidos.`);
    return [];
  }

  console.log("Dados lidos da planilha:", dados);

  // Mapeia a primeira coluna, incluindo valores nulos ou vazios
  const primeiraColuna = dados.map((linha, index) => {
    const valor = Array.isArray(linha) ? linha[0] : null;
    console.log(`Linha ${index + 1}, Coluna 1:`, valor); // Log para depuração
    return valor;
  });

  console.log("Primeira coluna completa (incluindo nulos):", primeiraColuna);

  // Filtra valores apenas se necessário
  const primeiraColunaFiltrada = primeiraColuna.filter(value => value !== null && value !== '');
  console.log("Primeira coluna filtrada (sem nulos ou vazios):", primeiraColunaFiltrada);

  return primeiraColunaFiltrada;
};