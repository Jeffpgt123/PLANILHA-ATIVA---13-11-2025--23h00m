// // === Módulo: Complementos ===
// // Arquivo sugerido: Complementos.gs ou Utilitarios.gs
// // Finalidade: incluir ações complementares como geração de documentos e marcação de campos

// const Complementos = {
//   /**
//    * Marca a célula de checkbox como "true" caso ainda não esteja marcada.
//    */
//   marcarCheckboxSeguro: function(sheet, linha, valor) {
//     try {
//       if (valor !== true) return;

//       const celulaStatus = sheet.getRange(linha, 15); // Coluna O
//       const statusAtual = celulaStatus.getValue();

//       if (statusAtual === true) {
//         Logger?.registrarLog?.('INFO', 'marcarCheckboxSeguro', `Já estava marcada. Linha ${linha}`);
//       } else {
//         const rangeA1 = `'${sheet.getName()}'!O${linha}`;
//         BatchOperations.add('setValue', rangeA1, true);
//         Logger?.registrarLog?.('INFO', 'marcarCheckboxSeguro', `Checkbox marcada na linha ${linha}`);
//       }
//     } catch (e) {
//       Logger?.registrarLog?.('ERRO', 'marcarCheckboxSeguro', e.message);
//     }
//   },

//   /**
//    * Cria um documento com base nos dados da linha editada
//    */
//   gerarDocumentoPendencias: function(sheet, range, linha) {
//     try {
//       const dados = sheet.getRange(linha, 1, 1, sheet.getLastColumn()).getValues()[0];
//       const nomeCliente = dados[1]; // Exemplo: Coluna B
//       const tipoPendencia = dados[3]; // Exemplo: Coluna D

//       const doc = DocumentApp.create(`Pendência - ${nomeCliente}`);
//       doc.getBody().appendParagraph(`Relatório de Pendência`);
//       doc.getBody().appendParagraph(`Cliente: ${nomeCliente}`);
//       doc.getBody().appendParagraph(`Tipo: ${tipoPendencia}`);
//       doc.getBody().appendParagraph(`Gerado em: ${new Date().toLocaleString()}`);
//       doc.saveAndClose();

//       Logger?.registrarLog?.('INFO', 'gerarDocumentoPendencias', `Criado documento: ${doc.getUrl()}`);
//     } catch (e) {
//       Logger?.registrarLog?.('ERRO', 'gerarDocumentoPendencias', e.message);
//     }
//   }
// };
