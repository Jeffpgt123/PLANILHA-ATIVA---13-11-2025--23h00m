// function compararEInserirMarcacao() {
//   const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARTEIRA");

//   // CONFIGURAÇÕES AJUSTÁVEIS
//   const linhaInicial = 4;              // Primeira linha com dados
//   const colunaBase = 'A';              // Onde estão os nomes "referência"
//   const colunaComparada = 'H';         // Onde estão os nomes a serem comparados
//   const colunaMarcacao = 'F';          // Onde será escrita a marcação
//   const textoMarcacao = 'ENCARTEIRADO';
//   const letrasIguais = 20;             // Quantas letras iniciais devem ser iguais

//   const ultimaLinha = aba.getLastRow();
//   const dadosBase = aba.getRange(colunaBase + linhaInicial + ':' + colunaBase + ultimaLinha).getValues().flat();
//   const dadosComparados = aba.getRange(colunaComparada + linhaInicial + ':' + colunaComparada + ultimaLinha).getValues().flat();

//   for (let i = 0; i < dadosComparados.length; i++) {
//     const nomeH = normalizarTexto(dadosComparados[i]);

//     if (nomeH) {
//       for (let nomeA of dadosBase) {
//         const nomeBase = normalizarTexto(nomeA);
//         if (nomeBase.substring(0, letrasIguais) === nomeH.substring(0, letrasIguais)) {
//           aba.getRange(linhaInicial + i, letraParaNumero(colunaMarcacao)).setValue(textoMarcacao);
//           break; // evita múltiplas marcações
//         }
//       }
//     }
//   }
// }

// // Converte letra de coluna (ex: 'F') para número (ex: 6)
// function letraParaNumero(letra) {
//   let numero = 0;
//   for (let i = 0; i < letra.length; i++) {
//     numero = numero * 26 + (letra.charCodeAt(i) - 64);
//   }
//   return numero;
// }

// // Remove acentos, espaços extras e converte para maiúsculas
// function normalizarTexto(texto) {
//   return texto
//     ? texto.toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim()
//     : '';
// }
