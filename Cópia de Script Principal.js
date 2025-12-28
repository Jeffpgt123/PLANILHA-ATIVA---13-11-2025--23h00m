// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO 1: CONFIGURAÇÃO GLOBAL OTIMIZADA
// // ████████████████████████████████████████████████████████████████████████

// const CONFIG = {  
//   // ►►► [CONCEITO] SHEET_NAMES: Lista das abas da planilha  
//   SHEET_NAMES: ['PROSPECCAO', 'INTERNALIZADO', 'EM ANALISE', 'CONTRATADO/LIBERACAO', 'INDEFERIDO', 'SEGUROS', 'LIMITES', 'CADASTROS', 'SEG - HISTORICO', 'LC - HISTORICO'],  
   
//   // ►►► [CONCEITO] FOLDER_ID: ID da pasta do Google Drive para salvar documentos  
//   FOLDER_ID: '1tgyuubBbUb8_bY5OxqCGXHYTv4baKTov',  

//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [CONCEITO] MODELO_DOCUMENTO: Configura o layout dos documentos gerados  
//   // ████████████████████████████████████████████████████████████████████████  
//   MODELO_DOCUMENTO: {  
//     CABECALHO: "PENDENCIAS",  
//     ESTILOS: {  
//       FONTE: "Arial",  
//       COR_TITULO: "#2c3e50"  
//     }  
//   },  

//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [CONCEITO] SEGURANCA: Configurações de segurança e acesso  
//   // ████████████████████████████████████████████████████████████████████████  
//   SEGURANCA: {  
//     VALIDACAO_NOMES: {  
//       TAMANHO_MAXIMO: 100,  
//       CARACTERES_INVALIDOS: /[\\/:*?"<>|]/g  
//     },  
//     NIVEL_ACESSO: "ANYONE",  
//     NOTIFICACOES: true  
//   },  

//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [MANTIDO] BATCH_OPS: Configuração original preservada
//   // ████████████████████████████████████████████████████████████████████████
//   BATCH_OPS: {
//     ENABLED: true,
//     DASHBOARD_SHEET: 'Dashboard_Batch',
//     MAX_QUEUE_SIZE: 500,          // ✅ MANTIDO: Valor original
//     LOG_PERFORMANCE: true
//   },

//   PRODUTOS_SEGUROS: {
//     enabled: true,
//     origem: 'EM ANALISE',
//     destino: 'SEGUROS',
//     colunaOrigem: 'J',
//     colunaProdutoDestinoNum: 3, // C
//     startRow: 4,
//     // mapa “EM ANALISE (letra)” → “SEGUROS (letra)”
//     mapeamento: { 'A':'A', 'B':'B', 'F':'D', 'G':'G'},
//     governanca: false
//   },
//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [CORREÇÃO] ACOMPANHAMENTO_PARCELAS: Adicionadas configurações necessárias para ParcelasManager
//   // ████████████████████████████████████████████████████████████████████████
//   ACOMPANHAMENTO_PARCELAS: {
//     ENABLED: false, // ← nova flag: defina false só durante o teste
//     ABA: 'SEGUROS',
//     LINHA_INICIO: 4,
//     COLUNA_DATA_CONTRATO: 13,
//     COLUNA_DIA_VENCTO: 14,
//     DIAS_ALERTA: 7,
//     // ✅ NECESSÁRIO: Usado por ParcelasManager.desativarAcompanhamento()
//     TAB_CHECKBOX_FORMATO: {
//       BACKGROUND: "#FFFFFF",
//       FONT_FAMILY: "Arial",
//       FONT_SIZE: 10
//     },
//     // ✅ NECESSÁRIO: Usado por ParcelasManager.atualizarCores()
//     COLUNA_DEF_ACOMPANHAMENTO: "I",  // Coluna I que define se tem acompanhamento (status do seguro)
//     VALOR_SEM_ACOMPANHAMENTO: ["A CONTRATAR"],
//     VALOR_COM_ACOMPANHAMENTO: ["EM CONTRATAÇÃO", "EM CONTRATAÇÃO-DILIGÊNCIA", "PAGO/CONTRATADO"],
//     COLUNAS_PARCELAS: [
//       { col: 16, mes: 1 },         // Janeiro
//       { col: 17, mes: 2 },         // Fevereiro
//       { col: 18, mes: 3 },         // Março
//       { col: 19, mes: 4 },         // Abril
//       { col: 20, mes: 5 },         // Maio
//       { col: 21, mes: 6 }          // Junho
//     ],
//     CORES: {
//       PAGO: '#2da544',      // Verde
//       ALERTA: '#ffff00',    // Amarelo
//       ATRASADO: '#ff0000',  // Vermelho
//       NEUTRO: '#FFFFFF'     // Branco
//     }
//   },

//   SHEETS: {
//     // ████████████████████████████████████████████████████████████████████████
//     // ►►► [CONCEITO] Configuração da aba "PROSPECCAO" (Clientes Potenciais)
//     // ████████████████████████████████████████████████████████████████████████
//       PROSPECCAO: {
//         START_ROW: 4,

//         MOVER_LINHA_CONFIG: {
//           13: {
//             true: {
//               destino: 'CADASTROS',
//               APAGAR_LINHA_ORIGEM: false,
//               linhaDestino: (aba) => {
//                 const valores = aba.getRange("A4:A").getValues();
//                 const index = valores.findIndex(row => row[0] === '');
//                 return index === -1 ? aba.getLastRow() + 1 : index + 4;
//               },
//               colunas: {
//                 origem:  [1, 2],
//                 destino: [1, 2]
//               }
//             }
//           }
//         },

//         COLUNAS_MONITORADAS: ["G", "I", "J", "K", "L", "M", "N"],
//         COLUNA_TIMESTAMP: "R",
//         FORMATO: "dd/MM/yyyy",
//         TRAVAR_HORA_EM_ZERO: true,

//         VALIDACAO: {
//           ativo: true,
//           colInicio: 1,
//           colFim: 18,
//           regrasPorColuna: {
//             1:  ['cpfOuCnpjValido'],
//             11: ['formatarData'],
//             15: ['formatarData']
//           }
//         },

//         // ▶ VISOES — UI-only (COMPLETO="TODAS"; OPERACIONAL aguarda lista)
//         VISOES: {
//           COMPLETO: "TODAS",
//           OPERACIONAL: ["B", "G", "H", "I", "J", "N", "O" ] // defina letras/intervalos quando quiser (ex.: ["A","C:F","H"])
//         }
//       }, 

//     // ████████████████████████████████████████████████████████████████████████
//     // ►►► [CONCEITO] Configuração da aba "INTERNALIZADO" (Processos Internos)
//     // ████████████████████████████████████████████████████████████████████████
//          INTERNALIZADO: {
//         START_ROW: 4,

//         MOVER_LINHA_CONFIG: {
//           11: {
//             "EM ANALISE": {
//               destino: 'EM ANALISE',
//               APAGAR_LINHA_ORIGEM: true,
//               linhaDestino: (aba) => {
//                 const valores = aba.getRange("A4:A").getValues();
//                 const index = valores.findIndex(row => row[0] === '');
//                 return index === -1 ? aba.getLastRow() + 1 : index + 4;
//               },
//               colunas: { origem: [1,2,3,4,5,6,7,8,9], destino: [1,2,3,4,5,6,7,8,9] }
//             }
//           },
//           12: {
//             "ATUALIZAR CADASTRO": {
//               destino: 'CADASTROS',
//               APAGAR_LINHA_ORIGEM: false,
//               linhaDestino: (aba) => {
//                 const valores = aba.getRange("A4:A").getValues();
//                 const index = valores.findIndex(row => row[0] === '');
//                 return index === -1 ? aba.getLastRow() + 1 : index + 4;
//               },
//               colunas: { origem: [1,2], destino: [1,2] }
//             }
//           },          
//         },

//         GOOGLE_DOC_CONFIG: { 
//           FILE_NAME: 6,
//           LINK: 14,
//           FILE_ID: 20,
//           TRIGGER_COL: 13
//         },

//         COLUNAS_MONITORADAS: ["J","K","L","M","N"],
//         COLUNA_TIMESTAMP: "O",
//         FORMATO: "dd/MM/yyyy",
//         TRAVAR_HORA_EM_ZERO: true,

//         VALIDACAO: {
//           ativo: true,
//           colInicio: 1,
//           colFim: 14,
//           regrasPorColuna: {
//             1:  ['cpfOuCnpjValido'],
//             6:  ['MultiFormatoProcesso'],
//             7:  ['FormatoDocsflow'],
//             9:  ['formatarData'],
//             14: ['formatarData']
//           }
//         },

//         // ▶ VISOES — UI-only (COMPLETO="TODAS"; OPERACIONAL aguarda lista)
//         VISOES: {
//           COMPLETO: "TODAS",
//           OPERACIONAL: ["B","E","H","J","K","L","M","N"]
//         },

//         // ▶ LISTAS_SUSPENSAS — coloracao: match exato em I; J herda cor quando herdarCorSubcategoria !== false; se false, J usa coresSubcategoria.
//         LISTAS_SUSPENSAS: [
//           // {
//           //   // D → E (fonte DADOS J → K)
//           //   colunaCategoria: "D",
//           //   colunaSubcategoria: "E",
//           //   fonte: { nomeAba: "DADOS", colunaCategoria: "J", colunaSubcategoria: "K", linhaInicial: 4 },
//           //   cores: { /* sem cores mapeadas para D (mantem comportamento original) */ },
//           //   herdarCorSubcategoria: true,   // J herda cor de D (nao ha cor mapeada, entao sem efeito visual)
//           //   coresSubcategoria: {           // usado apenas se herdarCorSubcategoria=false
//           //     // "EXEMPLO DETALHE": { fundo: "#eeeeee", texto: "#000000" }
//           //   }
//           // },
//           {
//             // K → L (fonte DADOS M → N)
//             colunaCategoria: "K",
//             colunaSubcategoria: "L",
//             fonte: { nomeAba: "DADOS", colunaCategoria: "M", colunaSubcategoria: "N", linhaInicial: 4 },
//             cores: { // match exato de texto em K
//               "A INTERNALIZAR": { fundo: "#ffe598", texto: "#000000" },
//               "INTERNALIZADO": { fundo: "#fdbf12", texto: "#000000" },
//               "CONFEC. PROPOSTA": { fundo: "#3d9aff", texto: "#FFFFFF" },
//               "CONFEC. PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#000000" },
//               "ASSINATURA PROPOSTA": { fundo: "#007aff", texto: "#FFFFFF" },
//               "ASSINATURA PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#FFFFFF" }
//             },
//             herdarCorSubcategoria: true,   // J (L) herda a cor de K (padrao)
//             coresSubcategoria: {           // usado se quiser desligar heranca (herdarCorSubcategoria=false)
//               // "PENDENTE BANCARIO": { fundo: "#ffe9a8", texto: "#000000" },
//               // "AGUARDANDO ASSINATURA": { fundo: "#a8d0ff", texto: "#000000" }
//             }
//           }
//         ]
//       },

//     // ████████████████████████████████████████████████████████████████████████
//     // ►►► [CONCEITO] Configuração da aba "EM ANALISE" 
//     // ████████████████████████████████████████████████████████████████████████
//       "EM ANALISE": {
//         START_ROW: 4,

//         MOVER_LINHA_CONFIG: {
//         11: {
//           "CONTRATADO": {
//             destino: 'CONTRATADO/LIBERACAO',
//             APAGAR_LINHA_ORIGEM: true,
//             linhaDestino: (aba) => {
//               const valores = aba.getRange('A4:A').getValues();
//               const index = valores.findIndex(row => row[0] === '');
//               return index === -1 ? aba.getLastRow() + 1 : index + 4;
//             },
//             colunas: { origem: [1,2,3,4,5,6,7,8,9,10], destino: [1,2,3,4,5,6,7,8,9,10] }
//           },
//           "INDEFERIDO": {
//             destino: 'INDEFERIDO',
//             APAGAR_LINHA_ORIGEM: true,
//             linhaDestino: (aba) => {
//               const valores = aba.getRange('A4:A').getValues();
//               const index = valores.findIndex(row => row[0] === '');
//               return index === -1 ? aba.getLastRow() + 1 : index + 4;
//             },
//             colunas: { origem: [1,2,3,4,5,6,7,8,9,10], destino: [1,2,3,4,5,6,7,8,9,10] }
//           }
//         }
//       },

//         GOOGLE_DOC_CONFIG: { 
//           FILE_NAME: 6,
//           LINK: 17,
//           FILE_ID: 20,
//           TRIGGER_COL: 16
//         },

//         COLUNAS_MONITORADAS: ["J","K","L","M","N"],
//         COLUNA_TIMESTAMP: "O",
//         FORMATO: "dd/MM/yyyy",
//         TRAVAR_HORA_EM_ZERO: true,

//         VALIDACAO: {
//           ativo: true,
//           colInicio: 1,
//           colFim: 14,
//           regrasPorColuna: {
//             1:  ['cpfOuCnpjValido'],
//             6:  ['MultiFormatoProcesso'],
//             7:  ['FormatoDocsflow'],
//             9:  ['formatarData'],
//             14: ['formatarData']
//           }
//         },

//         // ▶ VISOES — UI-only (COMPLETO="TODAS"; OPERACIONAL aguarda lista)
//         VISOES: {
//           COMPLETO: "TODAS",
//           OPERACIONAL: ["B","E","F","G","H","J","K","L","M","N"]
//         },

//         // ▶ LISTAS_SUSPENSAS — Coloracao: match exato em categoria; J herda cor quando herdarCorSubcategoria !== false; se false, J usa coresSubcategoria.
//         LISTAS_SUSPENSAS: [
//           // {
//           //   // D → E (fonte DADOS J → K)
//           //   colunaCategoria: "D",
//           //   colunaSubcategoria: "E",
//           //   fonte: { nomeAba: "DADOS", colunaCategoria: "J", colunaSubcategoria: "K", linhaInicial: 4 },
//           //   cores: { /* sem cores mapeadas para D (mantem comportamento original) */ },
//           //   herdarCorSubcategoria: true,   // J herda a cor de D (nao ha cor mapeada, entao sem efeito visual)
//           //   coresSubcategoria: {           // usado somente se herdarCorSubcategoria=false
//           //     // "EXEMPLO DETALHE": { fundo: "#eeeeee", texto: "#000000" }
//           //   }
//           // },
//           {
//             // N → O (fonte DADOS P → Q)
//             colunaCategoria: "K",
//             colunaSubcategoria: "L",
//             fonte: { nomeAba: "DADOS", colunaCategoria: "P", colunaSubcategoria: "Q", linhaInicial: 4 },
//               cores: {
//                     // Bloco INTERNALIZADO
//                     "__INTERNALIZADO__":                 { fundo: "#FDBF12", texto: "#000000" },
//                     "EM FILA PARA CONFECCAO":              { fundo: "#FDBF12", texto: "#000000" },
//                     "CONFEC. PROPOSTA":                     { fundo: "#3d9aff", texto: "#FFFFFF" },
//                     "CONFEC. PROPOSTA - PENDENCIAS":        { fundo: "#FD7E12", texto: "#000000" },
//                     "ASSINATURA PROPOSTA":                  { fundo: "#3d9aff", texto: "#FFFFFF" },
//                     "ASSINATURA PROPOSTA - PENDENCIAS":     { fundo: "#FD7E12", texto: "#000000" },

//                     // Bloco ANALISE
//                     "__ANALISE__":                        { fundo: "#3d9aff", texto: "#FFFFFF" },
//                     "EM ANALISE - DILIGENCIA":              { fundo: "#FD7E12", texto: "#000000" },
//                     "ESTRUTURACAO":                         { fundo: "#3d9aff", texto: "#FFFFFF" },
//                     "ESTRUTURACAO - DILIGENCIA":            { fundo: "#FDBF12", texto: "#000000" },
//                     "DESPACHO DO COMITE COMPETENTE":        { fundo: "#3d9aff", texto: "#FFFFFF" },
//                     "PRE CONTRATACAO":                      { fundo: "#3d9aff", texto: "#FFFFFF" },
//                     "PRE CONTRATACAO - DILIGENCIA":         { fundo: "#FD7E12", texto: "#000000" },
//                      // variação sem espaço

//                     // Finais
//                     "CONTRATADO":                           { fundo: "#2DA544", texto: "#FFFFFF" },
//                     "INDEFERIDO":                           { fundo: "#DC2626", texto: "#FFFFFF" },
//                     "RECONSIDERACAO":                       { fundo: "#FD7E12", texto: "#000000" }
//                   },
//             herdarCorSubcategoria: true,   // O herda a cor de N (padrao)
//             coresSubcategoria: {           // usado se herdarCorSubcategoria=false
//               // "PENDENTE BANCARIO": { fundo: "#ffe9a8", texto: "#000000" },
//               // "AGUARDANDO ASSINATURA": { fundo: "#a8d0ff", texto: "#000000" }
//             }
//           }
//         ]
//       },

//     // ████████████████████████████████████████████████████████████████████████
//     // ►►► [CONCEITO] Configuração da aba "CONTRATADO/LIBERAÇÃO"
//     // ████████████████████████████████████████████████████████████████████████
//       "CONTRATADO/LIBERACAO": { 
//         START_ROW: 4,

//         MOVER_LINHA_CONFIG: {
//           11: {
//             "PROPOSTA CONTRATADA": {
//               destino: 'SEGUROS',
//               APAGAR_LINHA_ORIGEM: true,
//               colunas: {
//                 origem:  [1, 2, 3],
//                 destino: [3, 1, 2]
//               }
//             }
//           }
//         },

//         GOOGLE_DOC_CONFIG: { 
//           FILE_NAME: 6,
//           LINK: 14,        
//           FILE_ID: 20,     
//           TRIGGER_COL: 13  
//         },

//         COLUNAS_MONITORADAS: ["K", "L", "M"],
//         COLUNA_TIMESTAMP: "N",
//         FORMATO: "dd/MM/yyyy",
//         TRAVAR_HORA_EM_ZERO: true,

//         VALIDACAO: {
//           ativo: true,
//           colInicio: 1,
//           colFim: 14,
//           regrasPorColuna: {
//             11: ['FormatoDocsflow'],          
//             12: ['formatarData']
//           }
//         },

//         // ▶ VISOES — UI-only (COMPLETO="TODAS"; OPERACIONAL aguarda lista)
//         VISOES: {
//           COMPLETO: "TODAS",
//           OPERACIONAL: ["B", "E", "F", "G", "H", "J","K", "L", "M", "N"]
//         },

//         // =============================================================================
//         // LISTAS_SUSPENSAS (DESATIVADO NESTA ABA)
//         // -----------------------------------------------------------------------------
//         // Observacao: este trecho esta totalmente comentado porque a aba nao utiliza
//         // listas/coloração no momento. Quando decidir ativar, remova os "//" e ajuste:
//         // - colunaCategoria / colunaSubcategoria (ex.: "K" → "L")
//         // - fonte.nomeAba / fonte.colunaCategoria / fonte.colunaSubcategoria
//         // - mapeamentos de cores (match EXATO de texto)
//         // - herdarCorSubcategoria: true (J herda cor de I) ou false (J tem cores proprias)
//         // -----------------------------------------------------------------------------
//         // LISTAS_SUSPENSAS: [
//         //   {
//         //     colunaCategoria: "K",          // TODO: defina a coluna da categoria
//         //     colunaSubcategoria: "L",       // TODO: defina a coluna do dependente
//         //     fonte: {
//         //       nomeAba: "DADOS",            // TODO: defina a aba origem
//         //       colunaCategoria: "?",        // TODO: ex.: "M"
//         //       colunaSubcategoria: "?",     // TODO: ex.: "N"
//         //       linhaInicial: 4
//         //     },
//         //     // Coloracao: aplicada pelo DropdownManager com MATCH EXATO do valor em K.
//         //     // Se herdarCorSubcategoria !== false, L herda a mesma cor de K.
//         //     cores: {
//         //       // "STATUS EXEMPLO": { fundo: "#007aff", texto: "#FFFFFF" },
//         //       // "STATUS EXEMPLO - PENDENCIAS": { fundo: "#fdbf12", texto: "#000000" }
//         //     },
//         //     herdarCorSubcategoria: true,   // true: L herda cor de K; false: L usa coresSubcategoria
//         //     coresSubcategoria: {           // usado APENAS quando herdarCorSubcategoria=false (match EXATO em L)
//         //       // "DETALHE EXEMPLO": { fundo: "#ffe9a8", texto: "#000000" }
//         //     }
//         //   }
//         // ]
//         // =============================================================================
//       },

//     // ████████████████████████████████████████████████████████████████████████
//     // ►►► [CONCEITO] Configuração da aba "SEGUROS"
//     // ████████████████████████████████████████████████████████████████████████
//     SEGUROS: {  // aba SEGUROS (modelo por-aba, com listas e cores locais)
//       START_ROW: 4,  // primeira linha util de dados

//       GOOGLE_DOC_CONFIG: { 
//         FILE_NAME: [2, 3, 7],  // colunas que compoem o nome do arquivo
//         LINK: 14,              // coluna do link do doc
//         FILE_ID: 20,           // coluna do fileId
//         TRIGGER_COL: 13        // coluna que aciona a geracao do doc
//       },

//       COLUNAS_MONITORADAS: ["D", "E", "F", "G", "H", "I", "J", "K", "L"], // colunas que disparam validacao/timestamp
//       COLUNA_TIMESTAMP: "V",  // coluna de timestamp
//       FORMATO: "dd/MM/yyyy",  // formato de data padrao
//       TRAVAR_HORA_EM_ZERO: true, // timestamp so com data

//       VALIDACAO: {
//         ativo: true,    // liga validacao linha-a-linha
//         colInicio: 1,   // coluna inicial do bloco validado
//         colFim: 14,     // coluna final do bloco validado
//         regrasPorColuna: {
//           1: ['cpfOuCnpjValido'], // valida doc
//           4:  ['MultiFormatoProcesso'],
//           8: ['formatarData'],    // data
//           11: ['formatarData'],
//           12: ['formatarData']    // data
//         }
//       },

//       // ▶ VISOES — UI-only (COMPLETO="TODAS"; OPERACIONAL aguarda lista)
//       VISOES: {
//         COMPLETO: "TODAS",                            // mostra todas as colunas
//         OPERACIONAL: ["B","C","E","I","J","L","N","O"] // colunas foco em operacao
//       },

//       // ▶ LISTAS_SUSPENSAS — definicao local (I → J, fonte DADOS U → V)
//       LISTAS_SUSPENSAS: [
//         {
//           colunaCategoria: "I",    // coluna da categoria (situacao)
//           colunaSubcategoria: "J", // coluna do detalhe (dependente)
//           fonte: {
//             nomeAba: "DADOS",          // aba de origem das listas
//             colunaCategoria: "U",      // categoria na origem
//             colunaSubcategoria: "V",   // subcategoria na origem
//             linhaInicial: 4            // primeira linha de dados na origem
//           },
//           // REGRAS DE COLORACAO (primaria I e opcional heranca para J)
//           cores: { // mapeamento de cores aplicado ao valor EXATO de I (match exato de texto)
//             "A CONTRATAR": { fundo: "#FDBF12", texto: "#000000" },            // cinza
//             "EM CONTRATACAO": { fundo: "#3d9aff", texto: "#FFFFFF" },         // azul
//             "EM CONTRATACAO-DILIGENCIA": { fundo: "#FD7E12", texto: "#000000" }, // amarelo
//             "PAGO/CONTRATADO": { fundo: "#2da544", texto: "#FFFFFF" },         // verde
//             "PROP CANCELADA": { fundo: "#DD3544", texto: "#FFFFFF" }         // verde
//           },
//           herdarCorSubcategoria: true, // true=J herda a cor de I (padrao); false=J usa suas proprias cores
//           coresSubcategoria: {         // usado SOMENTE quando herdarCorSubcategoria=false (match exato de texto em J)
//             //"PENDENTE BANCARIO": { fundo: "#ffe9a8", texto: "#000000" },      // amarelo claro
//            // "AGUARDANDO ASSINATURA": { fundo: "#a8d0ff", texto: "#000000" },  // azul claro
//            // "COMPLEMENTAR DOCUMENTACAO": { fundo: "#ffd6a8", texto: "#000000" }, // laranja claro
//             // "CANCELADO": { fundo: "#8b0000", texto: "#FFFFFF" }               // vermelho escuro
//           }
//         }
//       ]
//     },
//     // ████████████████████████████████████████████████████████████████████████
//     // ►►► [CONCEITO] Configuração da aba "LIMITES"
//     // ████████████████████████████████████████████████████████████████████████
//       LIMITES: {
//         START_ROW: 4,
//         GOOGLE_DOC_CONFIG: { 
//           FILE_NAME: 2,
//           LINK: 11,
//           FILE_ID: 20,
//           TRIGGER_COL: 10
//         },
//         COLUNAS_MONITORADAS: ["H", "I", "J", "K", "L", "M", "N", "O"],
//         COLUNA_TIMESTAMP: "P",
//         FORMATO: "dd/MM/yyyy",
//         TRAVAR_HORA_EM_ZERO: true,
//         VALIDACAO: {
//           ativo: true,
//           colInicio: 1,
//           colFim: 14,
//           regrasPorColuna: {
//             1: ['cpfOuCnpjValido'],  
//             6: ['FormatoDocsflow'],
//             7: ['formatarData'],
//             13: ['formatarData'],
//             14: ['formatarData'],
//           }
//         },

//         // ▶ LISTAS_SUSPENSAS — H→I (fonte DADOS Z→AA). Coloracao: match exato em H; I herda cor quando herdarCorSubcategoria !== false; se false, I usa coresSubcategoria.
//         LISTAS_SUSPENSAS: [
//           {
//             colunaCategoria: "H",
//             colunaSubcategoria: "I",
//             fonte: {
//               nomeAba: "DADOS",
//               colunaCategoria: "Z",
//               colunaSubcategoria: "AA",
//               linhaInicial: 4
//             },
//             cores: { // match EXATO de texto em H
//               "INTERNALIZADO": { fundo: "#6c747d", texto: "#FFFFFF" },
//               "EM CONFECÇÃO": { fundo: "#fd7e12", texto: "#FFFFFF" },
//               "EM CONFECÇÃO - PENDÊNCIAS": { fundo: "#fdbf12", texto: "#000000" },
//               "EM ANALISE": { fundo: "#3d9aff", texto: "#FFFFFF" },
//               "EM ANALISE - DILIGENCIA": { fundo: "#fdbf12", texto: "#000000" },
//               "DEFERIDO": { fundo: "#2da544", texto: "#FFFFFF" }
//             },
//             herdarCorSubcategoria: true, // true: I herda cor de H (padrao); false: I usa coresSubcategoria
//             coresSubcategoria: {         // usado SOMENTE quando herdarCorSubcategoria=false (match EXATO em I)
//               // "DETALHE EXEMPLO": { fundo: "#a8d0ff", texto: "#000000" }
//             }
//           }
//         ]
//       },

//     // ████████████████████████████████████████████████████████████████████████  
//     // ►►► [CONCEITO] Configuração da aba "CADASTROS" (Dados Cadastrais)  
//     // ████████████████████████████████████████████████████████████████████████  
//     CADASTROS: {
//       START_ROW: 4,

//       GOOGLE_DOC_CONFIG: { 
//         FILE_NAME: 2,
//         LINK: 10,
//         FILE_ID: 20,
//         TRIGGER_COL: 9
//       },

//       COLUNAS_MONITORADAS: ["E", "F", "G", "H", "I"],
//       COLUNA_TIMESTAMP: "J",
//       FORMATO: "dd/MM/yyyy",  
//       TRAVAR_HORA_EM_ZERO: true,

//       VALIDACAO: {
//         ativo: true,
//         colInicio: 1,
//         colFim: 14,
//         regrasPorColuna: {
//           1:  ['cpfOuCnpjValido'],
//           4:  ['formatarData'],
//           5:  ['FormatoDocsflow'],
//           9: ['formatarData']
//         }
//       },

//       // ▶ VISOES — UI-only (COMPLETO="TODAS"; OPERACIONAL aguarda lista)
//       VISOES: {
//         COMPLETO: "TODAS",
//         OPERACIONAL: ["B", "C", "E", "F", "G", "H", "I", "J"]
//       },

//       // ▶ LISTAS_SUSPENSAS — G→H (fonte DADOS AD→AE). Coloracao: match EXATO em G; H herda cor quando herdarCorSubcategoria !== false; se false, H usa coresSubcategoria.
//       LISTAS_SUSPENSAS: [
//         {
//           colunaCategoria: "F",
//           colunaSubcategoria: "G",
//           fonte: {
//             nomeAba: "DADOS",
//             colunaCategoria: "AD",
//             colunaSubcategoria: "AE",
//             linhaInicial: 4
//           },
//           cores: { // match EXATO de texto em G
//             "A ATUALIZAR": { fundo: "#FDBF12", texto: "#000000" },
//             "EM ATUALIZAÇÃO": { fundo: "#87CEEB", texto: "#FFFFFF" },
//             "EM ATUALIZAÇÃO - DILIGÊNCIA": { fundo: "#fdbf12", texto: "#000000" },
//             "ATUALIZADO/CONCLUIDO": { fundo: "#2da544", texto: "#FFFFFF" }
//           },
//           herdarCorSubcategoria: true, // true: H herda a cor de G (padrao); false: H usa coresSubcategoria
//           coresSubcategoria: {         // usado SOMENTE quando herdarCorSubcategoria=false (match EXATO em H)
//             // "DETALHE EXEMPLO": { fundo: "#a8d0ff", texto: "#000000" }
//            }
//          }
//        ]
//       }
//     }, 
//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [CONCEITO] MOVED_COLUMNS - Colunas movidas entre abas por status  
//   // ████████████████████████████████████████████████████████████████████████  
//   MOVED_COLUMNS: {
//     'EM ANALISE': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
//     'CONTRATADO/LIBERACAO': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
//     INDEFERIDO: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
//     CADASTROS: [1, 2]
//   },

//   EMAIL: [
//     // ████████████████████████████████████████████████████████████████████████  
//     // ►►► [CONCEITO] Configuração de e-mails automáticos para Limite de Crédito (LC)  
//     // ████████████████████████████████████████████████████████████████████████  
//     {
//       TYPE: "LC",
//       DESTINATARIOS: ["aaa@aaa.com", "bbb@bbb.com"],
//       PLANILHA_HISTORICO: "LC - HISTORICO",
//       START_ROW: 4,
//       COLUMNS: {
//         NOME_CLIENTE: 2,
//         TIPO_LC: 5,
//         VALOR_LIMITE: 12,
//         DOCSFLOW: "F",
//         DATA_VENCIMENTO: 14
//       },  
//       DIAS_ANTES: 40
//     },  

//     // ████████████████████████████████████████████████████████████████████████  
//     // ►►► [CONCEITO] Configuração de e-mails automáticos para Seguros (SEG)  
//     // ████████████████████████████████████████████████████████████████████████  
//     {
//       TYPE: "SEG",
//       DESTINATARIOS: ["ccc@ccc.com", "ddd@ddd.com"],
//       PLANILHA_HISTORICO: "SEG - HISTORICO",
//       START_ROW: 4,  
//       COLUMNS: {  
//         NOME_CLIENTE: 2,
//         TIPO_SEG: "C",
//         VALOR_SEG: 5,
//         DATA_VENCIMENTO: 12
//       },  
//       DIAS_ANTES: 40
//     }  
//   ],  

//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [CORREÇÃO] LISTAS_SUSPENSAS: Configuração otimizada para melhor performance
//   // ████████████████████████████████████████████████████████████████████████  


//   // ████████████████████████████████████████████████████████████████████████  
//   // ►►► [CONCEITO] Regra de expiração de linhas  
//   // ████████████████████████████████████████████████████████████████████████  
//   EXPIRED: { 
//     THRESHOLD_DAYS: 2,
//     COLOR: "#FFCCCC"
//    } 
//   };



// // ████████████████████████
// // ▶ CONFIG.UI — parâmetros de validação (toast + visuais)
// CONFIG.UI = CONFIG.UI || {};
// CONFIG.UI.VALIDACAO = CONFIG.UI.VALIDACAO || {};
// CONFIG.UI.VALIDACAO.TOAST = {
//   ativar: true,          // manter toast em DEV/PROD
//   limitePorExecucao: 3,  // no máx. 3 toasts por onEdit
//   cooldownMs: 1500       // suprimir toasts repetidos da mesma célula por ~1,5s
// };
// // (nenhum '};' extra depois do bloco)
// // ████████████████████████


// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO 2: ACOMPANHAMENTO DE PARCELAS 
// //    
// //    
// // ████████████████████████████████████████████████████████████████████████

// const ParcelasManager = {
//   /**
//    * Remove checkboxes, limpa validações e redefine formatação das colunas de parcelas
//    * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba "SEGUROS"
//    * @param {number} linha - Linha a ser processada
//    */
//   desativarAcompanhamento: (sheet, linha) => {
//   const config = CONFIG.ACOMPANHAMENTO_PARCELAS;

//   config.COLUNAS_PARCELAS.forEach(({ col }) => {
//     try {
//       const celula = sheet.getRange(linha, col);

//       try { celula.removeCheckbox(); } catch(e) {}
//       celula.clearDataValidations();
//       SpreadsheetApp.flush();

//       celula.clearContent();
//       celula.clearFormat();

//       // Remove cor personalizada, retornando à cor padrão da planilha
//       celula.setBackground(null); 
//       celula.setFontFamily(config.TAB_CHECKBOX_FORMATO.FONT_FAMILY);
//       celula.setFontSize(config.TAB_CHECKBOX_FORMATO.FONT_SIZE);
//       celula.setNote('');

//     } catch (ex) {  // ✅ CORRIGIDO: } adicionado antes do catch
//       const errorMsg = `(${linha},${col}) ${ex?.message || ''} | ${ex?.stack || ''}`;
//       try {
//         Logger.registrarLogBatch([['ERRO','ParcelasManager.desativarAcompanhamento', errorMsg]]);
//       } catch (_) {
//         console.error('ERRO ParcelasManager.desativarAcompanhamento:', errorMsg);
//       }
//     }
//   }); // ✅ CORRIGIDO: }); adicionado para fechar forEach

//   SpreadsheetApp.flush(); // ✅ CORRIGIDO: Movido para fora do forEach
// },

//   /**
//    * Insere checkboxes e aplica formatação padrão nas colunas de parcelas
//    * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba "SEGUROS"
//    * @param {number} linha - Linha a ser processada
//    */
//   reativarAcompanhamento: (sheet, linha) => {
//     const config = CONFIG.ACOMPANHAMENTO_PARCELAS;
//     const colunas = config.COLUNAS_PARCELAS.map(p => p.col);
//     const range = sheet.getRange(linha, Math.min(...colunas), 1, colunas.length);
    
//     range.clearDataValidations().clearFormat();
//     const validation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
//     range.setDataValidation(validation)
//       .setBackground(config.TAB_CHECKBOX_FORMATO.BACKGROUND)
//       .setFontFamily(config.TAB_CHECKBOX_FORMATO.FONT_FAMILY)
//       .setFontSize(config.TAB_CHECKBOX_FORMATO.FONT_SIZE);
//     SpreadsheetApp.flush();
//   },

//   /**
//    * Calcula a data de vencimento de uma parcela
//    * @param {Date} dataContrato - Data base do contrato
//    * @param {number} mesParcela - Número da parcela (1-6)
//    * @param {number} diaVencto - Dia de vencimento configurado
//    */
//   calcularVencimento: (dataContrato, mesParcela, diaVencto) => {
//     if (!dataContrato || isNaN(dataContrato.getTime())) return null;
//     const vencimento = new Date(dataContrato);

//     // Lógica de cálculo (mantida igual à original)
//     if (mesParcela === 1) vencimento.setDate(vencimento.getDate() + 3);
//     else if (mesParcela === 2) {
//       const candidate = new Date(dataContrato);
//       candidate.setMonth(candidate.getMonth() + 1);
//       candidate.setDate(diaVencto);
//       const diffDays = Math.ceil((candidate - dataContrato) / (1000 * 60 * 60 * 24));
//       vencimento.setTime((diffDays >= 29 && diffDays <= 31) ? candidate.getTime() : candidate.setMonth(candidate.getMonth() + 1));
//     } else {
//       vencimento.setMonth(vencimento.getMonth() + mesParcela);
//       vencimento.setDate(diaVencto);
//       while (vencimento.getDate() !== diaVencto) vencimento.setDate(vencimento.getDate() - 1);
//     }
//     return vencimento;
//   },

//   /**
//    * Define a cor da célula com base no status do pagamento
//    * @param {Date} vencimento - Data de vencimento
//    * @param {boolean} pago - Indica se a parcela está paga
//    */
//   definirCorParcela: (vencimento, pago) => {
//     if (!vencimento) return CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.NEUTRO;
//     const diffDias = Math.ceil((vencimento - new Date()) / (1000 * 60 * 60 * 24));
    
//     return pago ? CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.PAGO :
//            diffDias < 0 ? CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.ATRASADO :
//            diffDias <= CONFIG.ACOMPANHAMENTO_PARCELAS.DIAS_ALERTA ? CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.ALERTA :
//            CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.NEUTRO;
//   },

//   /**
//    * Atualiza cores e notas de todas as células de parcelas em uma linha
//    * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba "SEGUROS"
//    * @param {number} linha - Linha processada
//    */
//   atualizarCores: (sheet, linha) => {
//     const config = CONFIG.ACOMPANHAMENTO_PARCELAS;
//     const categoria = sheet.getRange(linha, Utils.colunaParaIndice(config.COLUNA_DEF_ACOMPANHAMENTO)).getDisplayValue().trim();
    
//     if (!config.VALOR_COM_ACOMPANHAMENTO.includes(categoria)) return;
    
//     const dataContrato = sheet.getRange(linha, config.COLUNA_DATA_CONTRATO).getValue();
//     const diaVencto = sheet.getRange(linha, config.COLUNA_DIA_VENCTO).getValue();
    
//     if (!dataContrato || !diaVencto) {
//       config.COLUNAS_PARCELAS.forEach(({ col }) => {
//         sheet.getRange(linha, col)
//           .setBackground(config.TAB_CHECKBOX_FORMATO.BACKGROUND)
//           .setNote('Preencha a data de vencimento na coluna N!');
//       });
//       return;
//     }

//     config.COLUNAS_PARCELAS.forEach(({ col, mes }) => {
//       const vencimento = ParcelasManager.calcularVencimento(dataContrato, mes, diaVencto);
//       const pago = sheet.getRange(linha, col).isChecked();
//       const cor = ParcelasManager.definirCorParcela(vencimento, pago);
      
//       sheet.getRange(linha, col)
//         .setBackground(cor)
//         .setNote(vencimento ? `Vencimento: ${Utilities.formatDate(vencimento, Utils.getTimezone(), 'dd/MM/yyyy')}` : '');
//     });
//   }
// };



// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO 3: UTILITÁRIOS CORRIGIDOS E OTIMIZADOS
// // ████████████████████████████████████████████████████████████████████████
// const Utils = {
//   _cache: {
//     // Unificado tudo em um só lugar
//     spreadsheet: null,
//     sheets: {},
//     configs: {},
//     statusValues: {},
//     lastProcessedRow: 0,
//     // ✅ NOVO: Controle de cache otimizado
//     timestamps: {},
//     TTL: 5 * 60 * 1000  // 5 minutos TTL
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ FUNÇÕES DE CONVERSÃO OTIMIZADAS
//   // ████████████████████████████████████████████████████████████████████████

//   /**
//    * ✅ FUNÇÃO PRINCIPAL: Converte letras de coluna para número (Ex: "A"→1, "AB"→28)
//    * Mantém compatibilidade com código existente
//    */
//   colunaParaIndice: (coluna) => {
//     try {
//       if (typeof coluna !== 'string') throw new Error('Tipo inválido para conversão');
//       coluna = coluna.toString().toUpperCase().replace(/[^A-Z]/g, '');
//       let indice = 0;
//       for (let i = 0; i < coluna.length; i++) {
//         indice = indice * 26 + (coluna.charCodeAt(i) - 64);
//       }
//       return indice;
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.colunaParaIndice', `Falha na conversão: ${erro.message}`]]);
//       return 0;
//     }
//   },

//   /**
//    * ✅ FUNÇÃO SECUNDÁRIA: Alias para compatibilidade (mantém código existente funcionando)
//    */
//   columnLetterToNumber: (column) => {
//     return Utils.colunaParaIndice(column);
//   },

//   /**
//    * ✅ CORRIGIDO: Converte número para letra (Ex: 1→"A", 28→"AB")
//    */
//   colunaParaLetra: (coluna) => {
//     try {
//       if (typeof coluna !== 'number' || coluna < 1) throw new Error('Número inválido para conversão');
//       let letra = '';
//       while (coluna > 0) {
//         const temp = (coluna - 1) % 26;
//         letra = String.fromCharCode(temp + 65) + letra;
//         coluna = Math.floor((coluna - temp - 1) / 26);
//       }
//       return letra || 'A';
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.colunaParaLetra', `Falha na conversão: ${erro.message}`]]);
//       return 'A';
//     }
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ SISTEMA DE CACHE OTIMIZADO
//   // ████████████████████████████████████████████████████████████████████████

//   /** 
//    * ✅ OTIMIZADO: Retorna a planilha ativa (com cache e TTL)
//    */
//   getCachedSpreadsheet: () => {
//     const now = Date.now();
//     const key = 'spreadsheet';
    
//     // Verificar TTL
//     if (Utils._cache.timestamps[key] && (now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
//       Utils._cache.spreadsheet = null;
//       delete Utils._cache.timestamps[key];
//     }
    
//     if (!Utils._cache.spreadsheet) {
//       try {
//         Utils._cache.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//         Utils._cache.timestamps[key] = now;
//       } catch (erro) {
//         Logger.registrarLogBatch([['ERRO', 'Utils.getCachedSpreadsheet', `Falha ao acessar planilha: ${erro.message}`]]);
//         return null;
//       }
//     }
//     return Utils._cache.spreadsheet;
//   },
//     getTimezone: () => {
//       try {
//         const ssTz = SpreadsheetApp.getActiveSpreadsheet()?.getSpreadsheetTimeZone();
//         return ssTz || Session.getScriptTimeZone() || 'GMT-3';
//       } catch (_) {
//         return 'GMT-3';
//       }
//     },
//   /** 
//    * ✅ OTIMIZADO: Retorna uma aba pelo nome (com cache e TTL)
//    */
//   getCachedSheet: (sheetName) => {
//     if (!sheetName) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.getCachedSheet', 'Nome da aba não fornecido']]);
//       return null;
//     }

//     const now = Date.now();
//     const key = `sheet_${sheetName}`;
    
//     // Verificar TTL
//     if (Utils._cache.timestamps[key] && (now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
//       delete Utils._cache.sheets[sheetName];
//       delete Utils._cache.timestamps[key];
//     }
    
//     if (!Utils._cache.sheets[sheetName]) {
//       try {
//         const ss = Utils.getCachedSpreadsheet();
//         if (!ss) return null;
        
//         Utils._cache.sheets[sheetName] = ss.getSheetByName(sheetName);
//         Utils._cache.timestamps[key] = now;
        
//         if (!Utils._cache.sheets[sheetName]) {
//           Logger.registrarLogBatch([['AVISO', 'Utils.getCachedSheet', `Aba "${sheetName}" não encontrada`]]);
//           return null;
//         }
//       } catch (erro) {
//         Logger.registrarLogBatch([['ERRO', 'Utils.getCachedSheet', `Falha ao acessar aba "${sheetName}": ${erro.message}`]]);
//         return null;
//       }
//     }
//     return Utils._cache.sheets[sheetName];
//   },

//   /** 
//    * ✅ OTIMIZADO: Retorna valores da coluna de status (com cache e TTL)
//    */
//   getCachedStatusValues: (sheetName, column) => {
//     if (!sheetName || !column) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.getCachedStatusValues', 'Parâmetros inválidos']]);
//       return [];
//     }

//     const now = Date.now();
//     const key = `${sheetName}_${column}`;
    
//     // Verificar TTL
//     if (Utils._cache.timestamps[key] && (now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
//       delete Utils._cache.statusValues[key];
//       delete Utils._cache.timestamps[key];
//     }
    
//     if (!Utils._cache.statusValues[key]) {
//       try {
//         const sheet = Utils.getCachedSheet(sheetName);
//         if (!sheet) return [];
        
//         const colIndex = typeof column === 'string' ? Utils.colunaParaIndice(column) : column;
//         Utils._cache.statusValues[key] = sheet.getRange(1, colIndex, sheet.getLastRow()).getValues();
//         Utils._cache.timestamps[key] = now;
//       } catch (erro) {
//         Logger.registrarLogBatch([['ERRO', 'Utils.getCachedStatusValues', `Falha ao ler valores: ${erro.message}`]]);
//         return [];
//       }
//     }
//     return Utils._cache.statusValues[key];
//   },

//   /**
//    * ✅ NOVO: Limpeza automática de cache expirado
//    */
//   limparCacheExpirado: () => {
//     try {
//       const now = Date.now();
//       let itensLimpos = 0;
      
//       Object.keys(Utils._cache.timestamps).forEach(key => {
//         if ((now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
//           // Limpar timestamp
//           delete Utils._cache.timestamps[key];
          
//           // Limpar dados relacionados
//           if (key === 'spreadsheet') {
//             Utils._cache.spreadsheet = null;
//           } else if (key.startsWith('sheet_')) {
//             const sheetName = key.replace('sheet_', '');
//             delete Utils._cache.sheets[sheetName];
//           } else {
//             delete Utils._cache.statusValues[key];
//           }
//           itensLimpos++;
//         }
//       });
      
//       if (itensLimpos > 0) {
//         Logger.registrarLogBatch([['INFO', 'Utils.limparCacheExpirado', `${itensLimpos} itens de cache expirados limpos`]]);
//       }
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.limparCacheExpirado', `Falha na limpeza: ${erro.message}`]]);
//     }
//   },

//   /**
//    * ✅ MELHORADO: Invalidação de cache com tipos específicos
//    */
//   invalidateCache: (key) => {
//     try {
//       if (key === 'all') {
//         Utils._cache.sheets = {};
//         Utils._cache.configs = {};
//         Utils._cache.statusValues = {};
//         Utils._cache.timestamps = {};
//         Utils._cache.spreadsheet = null;
//         Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', 'Todo o cache foi limpo']]);
//       } else if (key === 'sheets') {
//         Utils._cache.sheets = {};
//         // Limpar timestamps de sheets
//         Object.keys(Utils._cache.timestamps).forEach(k => {
//           if (k.startsWith('sheet_')) delete Utils._cache.timestamps[k];
//         });
//         Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', 'Cache de abas limpo']]);
//       } else if (key === 'configs') {
//         Utils._cache.configs = {};
//         Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', 'Cache de configurações limpo']]);
//       } else if (key) {
//         delete Utils._cache.sheets[key];
//         delete Utils._cache.configs[key];
//         delete Utils._cache.statusValues[key];
//         delete Utils._cache.timestamps[`sheet_${key}`];
//         Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', `Cache de "${key}" limpo`]]);
//       }
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.invalidateCache', `Falha na invalidação: ${erro.message}`]]);
//     }
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ FUNÇÕES DE CONFIGURAÇÃO E VALIDAÇÃO
//   // ████████████████████████████████████████████████████████████████████████

//   /**
//    * ✅ OTIMIZADO: Obter config da aba com cache
//    */
//   obterConfigAba: (nomeAba) => {
//     if (!nomeAba) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.obterConfigAba', 'Nome da aba não fornecido']]);
//       return {};
//     }

//     try {
//       if (!Utils._cache.configs[nomeAba]) {
//         Utils._cache.configs[nomeAba] = CONFIG?.SHEETS?.[nomeAba] || { 
//           START_ROW: 4,
//           COLUNAS_MONITORADAS: [],
//           COLUNA_TIMESTAMP: "O",
//           FORMATO: "dd/MM/yyyy",
//           TRAVAR_HORA_EM_ZERO: true
//         };
//       }
//       return Utils._cache.configs[nomeAba];
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.obterConfigAba', `Falha ao obter configuração para "${nomeAba}": ${erro.message}`]]);
//       return {};
//     }
//   },

//   /**
//    * ✅ CORRIGIDO: Verificar se linha está vazia (correção crítica)
//    */
//   linhaVazia: (sheet, linha) => {
//     try {
//       if (!sheet || !linha) {
//         Logger.registrarLogBatch([['ERRO', 'Utils.linhaVazia', 'Parâmetros inválidos']]);
//         return false;
//       }
//       const nome = sheet.getName();
//       const cfg = Utils.obterConfigAba(nome) || {};
//       const cachedSheet = Utils.getCachedSheet(nome);
//       if (!cachedSheet) return false;
//       const colFimValid = (cfg.VALIDACAO && cfg.VALIDACAO.colFim) || 15;
//       const maxMonit = Array.isArray(cfg.COLUNAS_MONITORADAS)
//         ? cfg.COLUNAS_MONITORADAS.map(Utils.colunaParaIndice).reduce((a,b)=>Math.max(a,b), 0)
//         : 0;
//       const numColunas = Math.max(colFimValid, maxMonit, 15);
//       const intervalo = cachedSheet.getRange(linha, 1, 1, numColunas);
//       const valores = intervalo.getValues()[0];
//       return valores.every(cel => !cel || String(cel).trim() === '');
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.linhaVazia', `Falha na verificação linha ${linha}: ${erro.message}`]]);
//       return false;
//     }
//   },
//   /**
//    * ✅ MANTIDO: Obtém valores de colunas específicas de uma linha
//    */
//   getColumnValues: (sheet, linha, colunas) => {
//     try {
//       if (!sheet || !linha || !Array.isArray(colunas)) {
//         Logger.registrarLogBatch([['ERRO', 'Utils.getColumnValues', 'Parâmetros inválidos']]);
//         return [];
//       }

//       return colunas.map(col => {
//         try {
//           const valor = sheet.getRange(linha, col).getDisplayValue();
//           return valor === '' ? null : valor;
//         } catch (erro) {
//           Logger.registrarLogBatch([['ERRO', 'Utils.getColumnValues', `Falha ao ler coluna ${col}: ${erro.message}`]]);
//           return null;
//         }
//       });
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.getColumnValues', `Falha geral: ${erro.message}`]]);
//       return [];
//     }
//   },

//   /**
//    * ✅ MANTIDO: Validação de nome de arquivo
//    */
//   validarNomeArquivo: (nome) => {
//     try {
//       const cfg = CONFIG?.SEGURANCA?.VALIDACAO_NOMES;
//       if (!cfg) {
//         throw new Error("Configuração de validação não encontrada");
//       }
      
//       if (!nome || nome.trim().length === 0) {
//         throw new Error("Nome do arquivo não pode ser vazio");
//       }
      
//       if (nome.length > cfg.TAMANHO_MAXIMO) {
//         throw new Error(`Nome excede ${cfg.TAMANHO_MAXIMO} caracteres`);
//       }
      
//       if (cfg.CARACTERES_INVALIDOS.test(nome)) {
//         throw new Error("Nome contém caracteres inválidos: \\ / : * ? \" < > |");
//       }
      
//       return nome.trim();
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.validarNomeArquivo', `Falha na validação: ${erro.message}`]]);
//       throw erro;
//     }
//   },

//   /**
//    * ✅ MANTIDO: Tratamento centralizado de erros do Google Drive
//    */
//   handleDriveError: (erro, contexto) => {
//     try {
//       const errosConhecidos = {
//         "LIMIT_EXCEEDED": "Limite de operações excedido no Drive",
//         "NO_ACCESS": "Acesso negado à pasta destino",
//         "INVALID_INPUT": "Parâmetros inválidos para criação do arquivo",
//         "FILE_NOT_FOUND": "Arquivo não encontrado",
//         "PERMISSION_DENIED": "Permissão negada"
//       };

//       const mensagem = errosConhecidos[erro.message] || erro.message;
//      const logEntry = [
//         Utilities.formatDate(new Date(), Utils.getTimezone(), 'dd/MM/yyyy'),
//         'ERRO_DRIVE',
//         `${contexto || 'Operação Drive'}: ${mensagem}`
//       ];

//       Logger.registrarLogBatch([logEntry]);
      
//       // Tentar mostrar alerta (pode falhar em execução automática)
//       try {
//         SpreadsheetApp.getUi().alert(`Erro no Drive: ${mensagem}`);
//       } catch (e) {
//         // Ignorar se não conseguir mostrar alerta (execução automática)
//       }
//     } catch (erro2) {
//       console.error('Erro no tratamento de erro do Drive:', erro2);
//     }
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ FUNÇÕES DE TESTE E DIAGNÓSTICO
//   // ████████████████████████████████████████████████████████████████████████

//   /**
//    * ✅ NOVO: Função de teste para validar Utils
//    */
//   testarSistema: () => {
//     try {
//       Logger.registrarLogBatch([['INFO', 'Utils.testarSistema', '🧪 Iniciando testes do Utils']]);
      
//       const testes = [];
      
//       // Teste 1: Conversões de coluna
//       const testeA = Utils.colunaParaIndice('A');
//       const testeZ = Utils.colunaParaIndice('Z');
//       const testeAA = Utils.colunaParaIndice('AA');
//       testes.push(`Conversões: A=${testeA}, Z=${testeZ}, AA=${testeAA}`);
      
//       // Teste 2: Conversões inversas
//       const letraA = Utils.colunaParaLetra(1);
//       const letraZ = Utils.colunaParaLetra(26);
//       const letraAA = Utils.colunaParaLetra(27);
//       testes.push(`Conversões inversas: 1=${letraA}, 26=${letraZ}, 27=${letraAA}`);
      
//       // Teste 3: Cache de planilha
//       const ss = Utils.getCachedSpreadsheet();
//       testes.push(`Planilha: ${ss ? '✅ OK' : '❌ ERRO'}`);
      
//       // Teste 4: Cache de aba
//       const sheet = Utils.getCachedSheet('DADOS');
//       testes.push(`Aba DADOS: ${sheet ? '✅ OK' : '❌ ERRO'}`);
      
//       // Teste 5: Configuração
//       const config = Utils.obterConfigAba('INTERNALIZADO');
//       testes.push(`Config INTERNALIZADO: ${config.START_ROW ? '✅ OK' : '❌ ERRO'}`);
      
//       // Teste 6: Limpeza de cache
//       Utils.limparCacheExpirado();
//       testes.push(`Limpeza de cache: ✅ Executada`);
      
//       Logger.registrarLogBatch([
//         ['INFO', 'Utils.testarSistema', '📊 Resultados dos testes:'],
//         ...testes.map(teste => ['INFO', 'Utils.testarSistema', `  ${teste}`]),
//         ['INFO', 'Utils.testarSistema', '🏁 Testes concluídos']
//       ]);
      
//       return {
//         sucesso: true,
//         testes: testes.length,
//         detalhes: testes
//       };
      
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'Utils.testarSistema', `Falha nos testes: ${erro.message}`]]);
//       return {
//         sucesso: false,
//         erro: erro.message
//       };
//     }
//   }
// };

// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO LOGGER CORRIGIDO E APRIMORADO
// // ████████████████████████████████████████████████████████████████████████

// const Logger = {
//     _cache: {
//         logSheet: null,
//         logBuffer: [],
//         lastFlush: 0,
//         MAX_BUFFER_SIZE: 5,
//         _logAtivo: true,
//         _modo: 'planilha' 
//     },

//     criarAbaLog: () => {
//         try {
//             if (!Logger._cache.logSheet) {
//                 const ss = SpreadsheetApp.getActive();
//                 if (!ss) {
//                     console.error("❌ Erro crítico: Não foi possível acessar a planilha ativa");
//                     return null;
//                 }
                
//                 Logger._cache.logSheet = ss.getSheetByName('Log do Sistema') || ss.insertSheet('Log do Sistema');
//                 const cabecalho = ['Timestamp', 'Tipo', 'Origem', 'Mensagem'];
//                 if (Logger._cache.logSheet.getLastRow() === 0) {
//                     Logger._cache.logSheet.getRange(1, 1, 1, cabecalho.length)
//                         .setValues([cabecalho])
//                         .setFontWeight('bold');
//                 }
//             }
//             return Logger._cache.logSheet;
//         } catch (erro) {
//             console.error(`❌ Erro crítico no Logger.criarAbaLog: ${erro.message}`);
//             return null;
//         }
//     },

//     registrarLogBatch: (entradas, forcePlanilha = false) => {
//         try {
//             if (!Logger._cache._logAtivo || !entradas?.length) return;

//             const tz = Utils.getTimezone();
//             const timestamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
//             const modo = forcePlanilha ? 'planilha' : Logger._cache._modo;

//             if (modo === 'console') {
//                 entradas.forEach((linha) => {
//                     const tipo = linha[0] || 'INFO';
//                     const origem = linha[1] || '';
//                     const mensagem = linha[2] || JSON.stringify(linha);
//                     console.log(`[${timestamp}] [${tipo}] [${origem}] ${mensagem}`);
//                 });
//                 return;
//             }

//             // modo planilha: bufferiza e grava em lote
//             entradas.forEach((linha) => {
//                 const tipo = linha[0] || 'INFO';
//                 const origem = linha[1] || '';
//                 const mensagem = linha[2] || JSON.stringify(linha);
//                 Logger._cache.logBuffer.push([timestamp, tipo, origem, mensagem]);
//             });

//             if (Logger._cache.logBuffer.length >= Logger._cache.MAX_BUFFER_SIZE) {
//                 Logger._flushLogBuffer();
//             }
//         } catch (erro) {
//             console.error(`❌ [Logger] Erro ao registrar logs: ${erro.message}`);
//         }
//     },

//     _flushLogBuffer: () => {
//         try {
//             if (!Logger._cache.logBuffer.length) return;
//             const sheet = Logger.criarAbaLog();
            
//             // ✅ VERIFICAÇÃO ROBUSTA: Verifica se a sheet existe antes de tentar gravar
//             if (!sheet) {
//                 console.error("❌ Não foi possível criar/acessar a aba de log. Buffer será limpo para evitar overflow.");
//                 Logger._cache.logBuffer = []; // Limpa buffer para evitar crescimento infinito
//                 return;
//             }
            
//             const lastRow = sheet.getLastRow();
//             const range = sheet.getRange(lastRow + 1, 1, Logger._cache.logBuffer.length, 4);
//             range.setValues(Logger._cache.logBuffer);
            
//             Logger._cache.logBuffer = [];
//             Logger._cache.lastFlush = Date.now();
            
//         } catch (e) {
//             console.error(`❌ Erro ao gravar logs na planilha: ${e.message}`);
//             // Em caso de erro, limpa buffer para evitar crescimento infinito
//             Logger._cache.logBuffer = [];
//         }
//     },

//     getModo: () => {
//         return Logger._cache._modo || 'planilha';
//     },

//     /**
//      * ✅ FUNÇÃO CORRIGIDA: alternarModo com tratamento de erro robusto
//      */
//     alternarModo: () => {
//         try {
//             // 1. Flush no buffer da planilha antes de mudar o modo
//             if (Logger._cache._modo === 'planilha') {
//                 Logger._flushLogBuffer();
//             }

//             // 2. Alterna o estado
//             Logger._cache._modo = Logger._cache._modo === 'planilha' ? 'console' : 'planilha';

//             const modoAtual = Logger._cache._modo;
//             const statusMsg = modoAtual === 'planilha'
//                 ? '📄 Planilha (Log do Sistema)'
//                 : '🖥️ Console (Execução)';

//             // 3. ✅ CORREÇÃO: Notificação com tratamento de erro
//             try {
//                 const ui = SpreadsheetApp.getUi();
//                 if (ui) {
//                     ui.alert(`✅ Modo de log alterado para: ${statusMsg}`);
//                 } else {
//                     console.log(`✅ Modo de log alterado para: ${statusMsg}`);
//                 }
//             } catch (uiError) {
//                 console.log(`✅ Modo de log alterado para: ${statusMsg} (UI não disponível)`);
//             }

//             // 4. ✅ CORREÇÃO: Recria o menu com verificação de existência
//             try {
//                 if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
//                     Triggers.recriarMenu();
//                 } else {
//                     console.log("⚠️ Triggers.recriarMenu não disponível. Menu não foi atualizado.");
//                 }
//             } catch (menuError) {
//                 console.error(`⚠️ Erro ao recriar menu: ${menuError.message}`);
//             }

//             // 5. ✅ Log da operação
//             Logger.registrarLogBatch([["INFO", "Logger.alternarModo", `Modo alterado para: ${modoAtual}`]]);

//         } catch (error) {
//             console.error(`❌ Erro crítico ao alternar modo do Logger: ${error.message}`);
//             console.error("Stack trace:", error.stack);
            
//             // ✅ CORREÇÃO: Fallback para notificação de erro
//             try {
//                 const ui = SpreadsheetApp.getUi();
//                 if (ui) {
//                     ui.alert(`❌ Erro ao alternar modo do log.\n\nDetalhes: ${error.message}\n\nVerifique o console para mais informações.`);
//                 }
//             } catch (fallbackError) {
//                 console.error("❌ Não foi possível nem exibir alerta de erro:", fallbackError.message);
//             }
//         }
//     },

//     getStatusLog: () => Logger._cache._logAtivo,
    
//     ativarLog: () => {
//         try {
//             Logger._cache._logAtivo = true;
//             Logger.registrarLogBatch([["INFO", "Logger", "🟢 Log ativado"]]);
            
//             try {
//                 SpreadsheetApp.getUi().alert("🟢 Registro de LOG ATIVADO.");
//             } catch (uiError) {
//                 console.log("🟢 Registro de LOG ATIVADO. (UI não disponível)");
//             }
            
//             if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
//                 Triggers.recriarMenu();
//             }
//         } catch (error) {
//             console.error(`❌ Erro ao ativar log: ${error.message}`);
//         }
//     },

//     desativarLog: () => {
//         try {
//             // Garante que o buffer pendente seja gravado antes de desligar
//             Logger._flushLogBuffer();
//             Logger._cache._logAtivo = false;
            
//             try {
//                 SpreadsheetApp.getUi().alert("🔴 Registro de LOG DESATIVADO.");
//             } catch (uiError) {
//                 console.log("🔴 Registro de LOG DESATIVADO. (UI não disponível)");
//             }
            
//             if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
//                 Triggers.recriarMenu();
//             }
//         } catch (error) {
//             console.error(`❌ Erro ao desativar log: ${error.message}`);
//         }
//     },

//     alternarLogAtivo: () => {
//         try {
//             Logger._cache._logAtivo = !Logger._cache._logAtivo;
//             const status = Logger._cache._logAtivo ? '🟢 ATIVADO' : '🔴 DESATIVADO';
            
//             Logger.registrarLogBatch([["INFO", "Logger", `Status alterado para: ${status}`]]);
            
//             try {
//                 SpreadsheetApp.getUi().alert(`📝 Registro de LOG está agora: ${status}`);
//             } catch (uiError) {
//                 console.log(`📝 Registro de LOG está agora: ${status} (UI não disponível)`);
//             }
            
//             if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
//                 Triggers.recriarMenu();
//             }
//         } catch (error) {
//             console.error(`❌ Erro ao alternar status do log: ${error.message}`);
//         }
//     },

//     /**
//      * ✅ NOVA FUNÇÃO: Força flush manual do buffer
//      */
//     forceFlush: () => {
//         try {
//             if (Logger._cache._modo === 'planilha') {
//                 Logger._flushLogBuffer();
//                 console.log("✅ Buffer de logs gravado manualmente");
//             } else {
//                 console.log("ℹ️ Modo console ativo - não há buffer para gravar");
//             }
//         } catch (error) {
//             console.error(`❌ Erro ao forçar flush: ${error.message}`);
//         }
//     }
// };

// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO 7: GERENCIAMENTO DE TIMESTAMP
// //    
// //    
// // ████████████████████████████████████████████████████████████████████████
// const TimestampManager = {
//   registrarTimestamp: (sheet, range) => {
//     try {
//       const nomeAba = sheet.getName();
//       const configAba = CONFIG.SHEETS[nomeAba];
//       if (!configAba) {
//         Logger.registrarLogBatch([['AVISO', `Aba ${nomeAba} não possui configuração de timestamp`]]);
//         return;
//       }

//       const linha = range.getRow();
//       const coluna = range.getColumn();
//       const colunaLetra = Utils.colunaParaLetra(coluna);

//       if (!configAba.COLUNAS_MONITORADAS.includes(colunaLetra)) {
//         return; // Não é uma coluna monitorada
//       }

//       const timestampCol = Utils.columnLetterToNumber(configAba.COLUNA_TIMESTAMP);
//       const data = new Date();
      
//       // Respeita a configuração TRAVAR_HORA_EM_ZERO
//       if (configAba.TRAVAR_HORA_EM_ZERO) {
//         data.setHours(0, 0, 0, 0);
//       }

//       // Formatação da data 
//      sheet.getRange(linha, timestampCol)
//       .setValue(Utilities.formatDate(data, Utils.getTimezone(), configAba.FORMATO));


//     } catch (erro) {
//       Logger.registrarLogBatch([
//         ['ERRO', 'TimestampManager', `Falha ao registrar timestamp: ${erro.message}`]
//       ]);
//     }
//   }
// };

// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO DE VALIDAÇÃO DE DADOS – VERSÃO FINAL COM CONTROLE DE RECURSIVIDADE
// //    Integrado à estrutura CONFIG.SHEETS.<ABA>.VALIDACAO
// //    Inclui suporte a validações múltiplas, formatação e proteção contra loops de edição
// // ████████████████████████████████████████████████████████████████████████

// const DataValidator = {
//   // ── Constantes/estado interno ────────────────────────────────
//   COR_ERRO: '#FFE6E6',                 // cor aplicada em dispararErro()
//   PREFIXO_NOTA_ERRO: /^erro:/i,        // identifica notas criadas pelo validador
//   DEFAULT_VALIDACAO_CACHE_TTL: 10,     // segundos (default otimizado; pode ser sobrescrito por CONFIG.VALIDACAO_CACHE_TTL)
//   _toastCount: 0,                      // contador de toasts por execução
//   _summaryToastShown: false,           // mostra 1 toast-resumo quando exceder limite
//   // ▶ Função principal que é chamada pelo onEdit
//   // Faz checagens iniciais e aplica regras definidas por coluna na aba correspondente
//     validarEdicao: function(e) {
//       try {
//         if (!e || !e.range) return true;
//       const range = e.range;
//       // === BLOCO 6 (NOVO): Colagem 2D — validar célula a célula ===
//       if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
//         const nR = range.getNumRows(), nC = range.getNumColumns();
//         for (let i = 0; i < nR; i++) {
//           for (let j = 0; j < nC; j++) {
//             const sub = range.offset(i, j, 1, 1);
//             DataValidator.validarEdicao({ range: sub, oldValue: null });
//           }
//         }
//         return true;
//       }
//       // === FIM BLOCO 6 ===
//       const valorDepois = range.getValue();

//       const valorAntes = e.oldValue ?? null;
//       const aba = range.getSheet().getName();
//       const linha = range.getRow();
//       const coluna = range.getColumn();

//       // Acesso às configurações da aba a partir do objeto global CONFIG
//       const configAba = CONFIG?.SHEETS?.[aba];
//       if (!configAba?.VALIDACAO?.ativo || linha < configAba.START_ROW) return true;
//       const configValidacao = configAba.VALIDACAO;
//       if (coluna < configValidacao.colInicio || coluna > configValidacao.colFim) return true;

//       const regras = configValidacao.regrasPorColuna?.[coluna];
//       if (!regras || regras.length === 0) return true;

//       // ▶ Controle de recursividade usando CacheService para evitar loops ao formatar
//       const chaveCache = `validacao_${range.getSheet().getSheetId()}_${linha}_${coluna}`;
//       const cache = CacheService.getScriptCache();
//       if (cache.get(chaveCache)) {
//         cache.remove(chaveCache);
//         return true; // Ignora validações se for reentrada gerada pelo próprio script
//       }

//       for (const regra of regras) {
//         switch (regra) {

//           // ▶ Validação de CPF ou CNPJ (com formatação automática)
//           case 'cpfOuCnpjValido': {
//             if (!valorDepois) {
//               this.limparFormatacaoErro(range, true);
//               return true;
//             }

//             const doc = String(valorDepois).replace(/[^\d]/g, '');
//             const isCpf = doc.length === 11 && this.validarCPF(doc);
//             const isCnpj = doc.length === 14 && this.validarCNPJ(doc);

//             if (isCpf || isCnpj) {
//               const valorFormatado = isCpf ? this.formatarCPF(doc) : this.formatarCNPJ(doc);
//               const TTL = (CONFIG && CONFIG.VALIDACAO_CACHE_TTL) ? CONFIG.VALIDACAO_CACHE_TTL : DataValidator.DEFAULT_VALIDACAO_CACHE_TTL;
//               cache.put(chaveCache, '1', TTL);
//               SpreadsheetApp.flush();
//               range.setValue(valorFormatado);
//               this.limparFormatacaoErro(range);
//               return true;
//             }

//             this.dispararErro(range, "CPF/CNPJ inválido.");
//             return false;
//           }

//           // ▶ Validação de datas inseridas como texto (dd/mm/aaaa ou dd/mm/aa)
//           case 'formatarData': {
//             if (!valorDepois) {
//               this.limparFormatacaoErro(range, true);
//               return true;
//             }
//             const dataValida = this.parsearData(valorDepois);
//             if (dataValida) {
//               const TTL = (CONFIG && CONFIG.VALIDACAO_CACHE_TTL) ? CONFIG.VALIDACAO_CACHE_TTL : DataValidator.DEFAULT_VALIDACAO_CACHE_TTL;
//               cache.put(chaveCache, '1', TTL);
//               SpreadsheetApp.flush();
//               range.setValue(dataValida);
//               try { range.setNumberFormat('dd/mm/yyyy'); } catch (_) {}
//               this.limparFormatacaoErro(range);
//               return true;
//             }
//             this.dispararErro(range, "Data inválida.");
//             return false;
//           }

//           // ▶ Validação de número Docsflow com 14 dígitos (formato: 0000/0000000000)
//           case 'FormatoDocsflow': {
//             if (!valorDepois) {
//               this.limparFormatacaoErro(range, true);
//               return true;
//             }
//             const doc = String(valorDepois).replace(/[^\d]/g, '');
//             if (doc.length === 14) {
//               const formatado = this.formatarDocsflow(doc);
//               const TTL = (CONFIG && CONFIG.VALIDACAO_CACHE_TTL) ? CONFIG.VALIDACAO_CACHE_TTL : DataValidator.DEFAULT_VALIDACAO_CACHE_TTL;
//               cache.put(chaveCache, '1', TTL);
//               SpreadsheetApp.flush();
//               range.setValue(formatado);
//               this.limparFormatacaoErro(range);
//               return true;
//             }
//             this.dispararErro(range, "Número Docsflow inválido (esperado 14 dígitos).");
//             return false;
//           }

//           // ▶  TRATAMENTO MULTIFORMATO: ACEITA 3 PADRÕES DE PROCESSOS
//           // 1. CONTROPER (10 DÍGITOS): Formato XXX-XX-XXXX-X (Ex: '1234567890' → '123-45-6789-0')
//           // 2. AMZCRED7 (7 DÍGITOS): Formato XXXXXX/X (Ex: '1234567' → '123456/7')
//           // 3. NÚMEROS SIMPLES (4 DÍGITOS): Mantém formato original (Ex: '1234')
//           // ⚠️ OUTROS FORMATOS SÃO BLOQUEADOS COM ERRO VISÍVEL
//           case 'MultiFormatoProcesso': {
//             if (!valorDepois) {
//               this.limparFormatacaoErro(range, true);
//               return true;
//             }
//             const doc = String(valorDepois).replace(/[^\d]/g, '');
//             let valorFinal = null;

//             if (doc.length === 10) {
//               valorFinal = this.formatarControper(doc);
//             } else if (doc.length === 7) {
//               valorFinal = this.formatarAmzcred7(doc);
//             } else if (doc.length === 4) {
//               valorFinal = doc;
//             }

//             if (valorFinal !== null) {
//               const TTL = (CONFIG && CONFIG.VALIDACAO_CACHE_TTL) ? CONFIG.VALIDACAO_CACHE_TTL : DataValidator.DEFAULT_VALIDACAO_CACHE_TTL;
//               cache.put(chaveCache, '1', TTL);
//               SpreadsheetApp.flush();
//               range.setValue(valorFinal);
//               this.limparFormatacaoErro(range);
//               return true;
//             }

//             this.dispararErro(range, "Formato de número de operação inválido.");
//             return false;
//           }

//         } // ← Fim do switch
//       } // ← Fim do loop de regras
//       return true;

//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO_CRITICO', 'DataValidator', `Falha inesperada: ${erro.message}`]], true);
//       return true;
//     }
//   },

//     // ▶ Função auxiliar para destacar o erro visualmente e deixar mensagem
//     dispararErro: function(range, tipoErro) {
//       const msg = `${tipoErro}`;

//       // log existente (mantido)
//       Logger.registrarLogBatch(
//         [['ERRO', 'DataValidator', `Dado inválido na célula ${range.getA1Notation()}. Tipo: ${tipoErro}`]],
//         true
//       );

//       // marcação visual (agora usando constante)
//       range.setBackground(this.COR_ERRO);
//       range.setNote(`ERRO: ${msg}`);

//       // ─────────── freios via CONFIG.UI.VALIDACAO.TOAST ───────────
//       try {
//         const cfg = (CONFIG && CONFIG.UI && CONFIG.UI.VALIDACAO && CONFIG.UI.VALIDACAO.TOAST) || {};
//         if (!cfg.ativar) return; // só cor/nota

//         // contador por execução (reinicia a cada onEdit)
//         this._toastCount = (typeof this._toastCount === 'number') ? this._toastCount : 0;
//         this._summaryToastShown = (typeof this._summaryToastShown === 'boolean') ? this._summaryToastShown : false;

//         const limite = (typeof cfg.limitePorExecucao === 'number') ? cfg.limitePorExecucao : 3;

//         // Se já atingiu o limite, mostra UM resumo e encerra
//         if (this._toastCount >= limite) {
//           if (!this._summaryToastShown) {
//             try {
//               const ui = (SpreadsheetApp.getActive() || range.getSheet().getParent());
//               ui.toast('Múltiplos erros encontrados. Verifique as células destacadas em vermelho.', 'Validação', 8);
//             } catch (_) {}
//             this._summaryToastShown = true; // garantir 1 só vez por execução
//           }
//           return;
//         }

//         // cooldown por célula (CacheService; TTL em segundos)
//         const sheet = range.getSheet();
//         const key = [
//           'toastCooldown',
//           sheet.getParent().getId(),
//           sheet.getSheetId(),
//           range.getRow(),
//           range.getColumn()
//         ].join(':');

//         const cache = CacheService.getScriptCache();
//         if (cache.get(key)) return; // em cooldown
//         const ttlSec = Math.max(1, Math.ceil(((cfg.cooldownMs || 1500) / 1000)));
//         cache.put(key, '1', ttlSec);

//         // exibir toast (classe correta + fallback)
//         try { (SpreadsheetApp.getActive() || sheet.getParent()).toast(msg, 'Erro de Validação', 5); } catch (_) {}
//         this._toastCount++;
//       } catch (_) {
//         // execução sem interface: ignore o toast
//       }
//     },

//   // ▶ Limpa marcações visuais de erro (cor e nota “ERRO:”)
//     limparFormatacaoErro: function(range, limparNota = false) {
//       try {
//         // 1) Remover a cor de erro, se for exatamente a usada pelo validador
//         const bg = (range.getBackground() || '').toLowerCase();
//         if (bg === this.COR_ERRO.toLowerCase()) {
//           range.setBackground(null);
//         }

//         // 2) Nota atual (se existir)
//         const notaAtual = (typeof range.getNote === 'function' ? (range.getNote() || '') : '').trim();

//         // 3) Remoção de nota:
//         //    a) Quando explicitamente solicitado (ex.: célula vazia): limparNota === true
//         //    b) Sempre que a nota for do validador (prefixo "ERRO:")
//         if (limparNota || this.PREFIXO_NOTA_ERRO.test(notaAtual)) {
//           range.clearNote(); // usar clearNote() para remover completamente
//         }

//         // 4) (Opcional) Se também houver cor de FONTE em erro, limpar aqui:
//         // const corFonteErro = '#9C0006';
//         // if ((range.getFontColor() || '').toLowerCase() === corFonteErro.toLowerCase()) range.setFontColor(null);

//       } catch (_) {
//         // Não bloquear o fluxo em caso de falha de limpeza
//       }
//     },
//   // ▶ Formatação de campos específicos
//   formatarDocsflow: function(doc) {
//     return doc.replace(/(\d{4})(\d{10})/, '$1/$2');
//   },
//   formatarControper: function(doc) {
//     return doc.replace(/(\d{3})(\d{2})(\d{4})(\d{1})/, '$1-$2-$3-$4');
//   },
//   formatarAmzcred7: function(doc) {
//     return doc.replace(/(\d{6})(\d{1})/, '$1/$2');
//   },
//   formatarCPF: function(cpf) {
//     return cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
//   },
//   formatarCNPJ: function(cnpj) {
//     return cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
//   },

//   // ▶ Conversão de textos para datas válidas, nos formatos ddmmaaaa ou ddmmaa
//   parsearData: function(textoData) {
//     // 1) Se já é Date válido, apenas retorne
//     if (textoData instanceof Date && !isNaN(textoData.getTime())) {
//       return textoData;
//     }

//     // 2) Se é número (data serial), tente converter
//     if (typeof textoData === 'number') {
//       const dataNum = new Date(textoData);
//       if (!isNaN(dataNum.getTime())) return dataNum;
//     }

//     // 3) Parse por dígitos (casos: "ddmmaaaa" ou "ddmmaa")
//     const digitos = String(textoData).replace(/[^\d]/g, '');
//     if (digitos.length === 8) {
//       const dia = digitos.substring(0, 2), mes = digitos.substring(2, 4), ano = digitos.substring(4, 8);
//       const dataObj = new Date(ano, mes - 1, dia);
//       if (dataObj.getFullYear() == ano && dataObj.getMonth() == (mes - 1) && dataObj.getDate() == dia) return dataObj;
//     } else if (digitos.length === 6) {
//       const dia = digitos.substring(0, 2), mes = digitos.substring(2, 4), ano = '20' + digitos.substring(4, 6);
//       const dataObj = new Date(ano, mes - 1, dia);
//       if (dataObj.getFullYear() == ano && dataObj.getMonth() == (mes - 1) && dataObj.getDate() == dia) return dataObj;
//     }
//     return null;
//   },

//   // ▶ Validação de CPF
//   validarCPF: function(cpf) {
//     if (typeof cpf !== 'string' || !cpf || cpf.length !== 11 || /^(\d)\1+$/.test(cpf)) return false;
//     let soma = 0, resto;
//     for (let i = 1; i <= 9; i++) soma += parseInt(cpf.substring(i-1, i)) * (11 - i);
//     resto = (soma * 10) % 11;
//     if (resto >= 10) resto = 0;
//     if (resto !== parseInt(cpf.substring(9, 10))) return false;
//     soma = 0;
//     for (let i = 1; i <= 10; i++) soma += parseInt(cpf.substring(i-1, i)) * (12 - i);
//     resto = (soma * 10) % 11;
//     if (resto >= 10) resto = 0;
//     return resto === parseInt(cpf.substring(10, 11));
//   },

//   // ▶ Validação de CNPJ
//   validarCNPJ: function(cnpj) {
//     if (typeof cnpj !== 'string' || !cnpj || cnpj.length !== 14 || /^(\d)\1+$/.test(cnpj)) return false;
//     let tamanho = cnpj.length - 2;
//     let numeros = cnpj.substring(0, tamanho);
//     let digitos = cnpj.substring(tamanho);
//     let soma = 0;
//     let pos = tamanho - 7;

//     for (let i = tamanho; i >= 1; i--) {
//       soma += parseInt(numeros.charAt(tamanho - i)) * pos--;
//       if (pos < 2) pos = 9;
//     }

//     let resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
//     if (resultado != parseInt(digitos.charAt(0))) return false;

//     tamanho += 1;
//     numeros = cnpj.substring(0, tamanho);
//     soma = 0;
//     pos = tamanho - 7;

//     for (let i = tamanho; i >= 1; i--) {
//       soma += parseInt(numeros.charAt(tamanho - i)) * pos--;
//       if (pos < 2) pos = 9;
//     }

//     resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
//     return resultado == parseInt(digitos.charAt(1));
//   }

// };

// // 🔗 Este módulo deve ser invocado a partir do gatilho onEdit no bloco superior
// //     Exemplo:
// //     if (typeof DataValidator !== 'undefined' && DataValidator.validarEdicao) {
// //        DataValidator.validarEdicao(e);
// //     }


// // ████████████████████████████████████████████████████████████████████████
// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO 9: GERENCIAMENTO DE E-MAILS (com formatação pt-BR e bloco completo)
// // ████████████████████████████████████████████████████████████████████████
// const EmailManager = {
//   /**
//    * Verifica vencimentos e envia e-mails para todas as configurações
//    */
//   verificarEViarAlertas: () => {
//     try {
//       if (CONFIG.EMAIL && Array.isArray(CONFIG.EMAIL)) {
//         CONFIG.EMAIL.forEach(config => {
//           const sheet = Utils.getCachedSheet(config.PLANILHA_HISTORICO);
//           if (!sheet) {
//             Logger.registrarLogBatch([['ERRO', 'EmailManager', `Planilha ${config.PLANILHA_HISTORICO} não encontrada`]]);
//             return;
//           }

//           const dados = sheet.getDataRange().getValues();
//           for (let i = config.START_ROW - 1; i < dados.length; i++) {
//             const linha = dados[i];
//             if (EmailManager.deveEnviarAlerta(linha, config)) {
//               EmailManager.enviarEmail(linha, config);
//             }
//           }
//         });
//       }
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'EmailManager', `Falha geral no envio de e-mails: ${erro.message}`]]);
//     }
//   },

//   /**
//    * Verifica se a linha precisa de alerta
//    */
//   deveEnviarAlerta: (linha, config) => {
//     try {
//       const colunas = config.COLUMNS;
//       const idxVenc = EmailManager.obterIndice(colunas.DATA_VENCIMENTO, config);
//       const dataVencimento = new Date(linha[idxVenc]);

//       if (!dataVencimento || isNaN(dataVencimento.getTime())) return false;

//       const hoje = new Date();
//       const diffDias = Math.floor((dataVencimento - hoje) / (1000 * 60 * 60 * 24));

//       return diffDias <= config.DIAS_ANTES && diffDias >= 0;
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'EmailManager', `Erro na verificação: ${erro.message}`]]);
//       return false;
//     }
//   },

//   /**
//    * Converte colunas (letras ou números) para índices (base 0)
//    */
//   obterIndice: (coluna, _config) => {
//     return typeof coluna === 'string'
//       ? Utils.colunaParaIndice(coluna) - 1
//       : coluna - 1;
//   },

//   /**
//    * Constrói dados e envia o e-mail (valor formatado em pt-BR)
//    */
//     enviarEmail: (linha, config) => {
//     try {
//       const col = config.COLUMNS;
//        // ❌ Removido: deduplicação via CacheService (24h por chave)
//        // Mantido: parsing numérico pt-BR e timezone centralizado
//        const idxNome = EmailManager.obterIndice(col.NOME_CLIENTE, config);
//        const idxTipo = EmailManager.obterIndice(config.TYPE === 'LC' ? col.TIPO_LC : col.TIPO_SEG, config);
//        const idxVenc = EmailManager.obterIndice(col.DATA_VENCIMENTO, config);
//        const nomeCliente = String(linha[idxNome] || '').trim();
//        const tipoOper   = String(linha[idxTipo] || '').trim();
//        const dtVenc     = new Date(linha[idxVenc]);      

//       // Valor bruto vindo da linha (pode ser número ou string com separadores)
//       const brutoValor = linha[EmailManager.obterIndice(
//         config.TYPE === 'LC' ? col.VALOR_LIMITE : col.VALOR_SEG,
//         config
//       )];

//       // 1) Tenta converter para número (aceita "1.234,56" → 1234.56)
//       let valorNum = Number(brutoValor);
//       if (isNaN(valorNum) && typeof brutoValor === 'string') {
//         const normalizado = brutoValor
//           .toString()
//           .replace(/[^\d.,-]/g, '')   // mantém dígitos, ponto, vírgula e sinal
//           .replace(/\./g, '')         // remove pontos (milhar)
//           .replace(',', '.');         // vírgula → ponto (decimal)
//         valorNum = Number(normalizado);
//       }

//       // 2) Formata em pt-BR se for numérico; senão, usa o texto original
//       const valorFmt = isFinite(valorNum)
//         ? valorNum.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
//         : String(brutoValor || '');

//       const dados = {
//         nome: nomeCliente,
//         tipo: tipoOper,
//         valor: valorFmt, // <- já vem formatado pt-BR
//         vencimento: Utilities.formatDate(
//           dtVenc,
//           Utils.getTimezone(), // <- timezone centralizado
//           'dd/MM/yyyy'
//         ),
//         docsflow: config.TYPE === 'LC' ? linha[EmailManager.obterIndice(col.DOCSFLOW, config)] : null
//       };

//       const assunto = `⏰ Alerta: Vencimento ${config.TYPE} - ${dados.nome}`;
//       const corpo = EmailManager.construirCorpoEmail(dados, config.TYPE);

//       MailApp.sendEmail({
//         to: config.DESTINATARIOS.join(','),
//         subject: assunto,
//         htmlBody: corpo
//       });

//       // ❌ Removido: marcação em cache (deduplicação 24h)

//       Logger.registrarLogBatch([['INFO', 'EmailManager', `E-mail enviado para ${dados.nome}`]]);
//     } catch (erro) {
//       Logger.registrarLogBatch([['ERRO', 'EmailManager', `Falha no envio: ${erro.message}`]]);
//     }
//   },

//   /**
//    * Monta o HTML do e-mail (usa o valor já formatado)
//    */
//   construirCorpoEmail: (dados, tipo) => {
//     return `
//       <html>
//         <body style="font-family: Arial, sans-serif; padding: 20px;">
//           <h2 style="color: #2c3e50;">Alerta de Vencimento</h2>

//           <div style="background: #f8f9fa; padding: 15px; border-radius: 8px;">
//             <p><strong>Nome do Cliente:</strong> ${dados.nome}</p>
//             <p><strong>Tipo ${tipo}:</strong> ${dados.tipo}</p>
//             <p><strong>Valor:</strong> R$ ${dados.valor}</p>
//             ${tipo === 'LC' ? `<p><strong>Docsflow:</strong> ${dados.docsflow || 'N/A'}</p>` : ''}
//             <p><strong>Data de Vencimento:</strong> ${dados.vencimento}</p>
//           </div>

//           <p style="margin-top: 20px; color: #7f8c8d;">
//             Este é um alerta automático. Favor não responder este e-mail.
//           </p>
//         </body>
//       </html>
//     `;
//   }
// };
// // ████████████████████████████████████████████████████████████████████████
// // ▶ VISOES — Wrapper mínimo (submenu + aplicação UI-only)
// // ████████████████████████████████████████████████████████████████████████

// /** Normaliza nome de aba para chave canônica (SEM acento, upper) */
// function VISOES__normalizeAbaName(nome) {
//   try {
//     return nome
//       .toString()
//       .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
//       .replace(/ç/gi, 'c')
//       .toUpperCase()
//       .trim();
//   } catch (e) {
//     return (nome || '').toString().toUpperCase().trim();
//   }
// }

// /** Converte tokens de coluna (ex.: 'A', 'C:D', 3, '3:8') para índices (1-based) */
// function VISOES__expandTokensToIndices(tokens, maxCols) {
//   const out = new Set();
//   if (!Array.isArray(tokens)) return out;
//   tokens.forEach(raw => {
//     if (raw == null) return;
//     const t = String(raw).trim().toUpperCase();
//     const addIndex = (i) => { if (i >= 1 && i <= maxCols) out.add(i); };
//     if (/^\d+$/.test(t)) { addIndex(Number(t)); return; }
//     if (/^\d+:\d+$/.test(t)) {
//       let [a,b] = t.split(':').map(n => Number(n));
//       if (a>b) [a,b] = [b,a];
//       for (let i=a;i<=b;i++) addIndex(i);
//       return;
//     }
//     if (/^[A-Z]+$/.test(t)) { addIndex(Utils.colunaParaIndice(t)); return; }
//     if (/^[A-Z]+:[A-Z]+$/.test(t)) {
//       let [a,b] = t.split(':');
//       let ai = Utils.colunaParaIndice(a);
//       let bi = Utils.colunaParaIndice(b);
//       if (ai>bi) [ai,bi] = [bi,ai];
//       for (let i=ai;i<=bi;i++) addIndex(i);
//       return;
//     }
//   });
//   return out;
// }

// /** Aplica visão na aba ativa (UI-only). visao='OPERACIONAL'|'COMPLETO' */
// function VISOES__aplicarVisaoMenu(visao) {
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const sheet = ss.getActiveSheet();
//     const abaReal = sheet.getName();
//     const abaKey = VISOES__normalizeAbaName(abaReal);
//     // 1ª fonte: bloco da própria aba
//     let cfgAba = (CONFIG && CONFIG.SHEETS && CONFIG.SHEETS[abaKey] && CONFIG.SHEETS[abaKey].VISOES)
//       ? CONFIG.SHEETS[abaKey].VISOES
//       : null;
//     // 2ª fonte: bloco central (se existir)
//     if (!cfgAba) {
//       const cent = (CONFIG && CONFIG.VISOES && CONFIG.VISOES.ABAS) ? CONFIG.VISOES.ABAS : null;
//       if (cent && cent[abaKey]) cfgAba = cent[abaKey];
//     }
//     if (!cfgAba) {
//       if (CONFIG?.VISOES?.DEBUG && typeof Logger?.registrarLogBatch === 'function') {
//         Logger.registrarLogBatch([["DEBUG","VISOES",`Sem configuração de visões para: ${abaReal} (${abaKey})`]]);
//       }
//       return; // silencioso
//     }
//     const maxCols = sheet.getMaxColumns();

//     // Desejado: matriz booleana (true=visível)
//     let desejadoVisivel = new Array(maxCols).fill(true);
//     if (visao === 'OPERACIONAL') {
//       const set = VISOES__expandTokensToIndices(cfgAba.OPERACIONAL || [], maxCols);
//       desejadoVisivel = desejadoVisivel.map((_,i)=> set.has(i+1)); // visível apenas o que consta
//     } else {
//       // COMPLETO: mantém tudo visível (ou se cfgAba.COMPLETO==="TODAS")
//       desejadoVisivel = desejadoVisivel.map(()=> true);
//     }

//     // Estado atual
//     const atualOculto = new Array(maxCols);
//     for (let c=1;c<=maxCols;c++) atualOculto[c-1] = sheet.isColumnHiddenByUser(c);

//     // Calcular blocos de toggle (contiguidade por alvo)
//     const hideBlocks = [];
//     const showBlocks = [];
//     let i = 1;
//     while (i <= maxCols) {
//       const desiredHidden = !desejadoVisivel[i-1];
//       const isHidden = atualOculto[i-1];
//       if (desiredHidden !== isHidden) {
//         const target = desiredHidden; // queremos chegar neste estado
//         let j = i;
//         while (j <= maxCols) {
//           const jDesiredHidden = !desejadoVisivel[j-1];
//           const jIsHidden = atualOculto[j-1];
//           if (jDesiredHidden !== jIsHidden && jDesiredHidden === target) {
//             j++;
//           } else {
//             break;
//           }
//         }
//         const count = j - i;
//         (target ? hideBlocks : showBlocks).push({ start: i, count });
//         i = j;
//       } else {
//         i++;
//       }
//     }

//     // Aplicar
//     hideBlocks.forEach(b => sheet.hideColumns(b.start, b.count));
//     showBlocks.forEach(b => sheet.showColumns(b.start, b.count));

//     if (CONFIG?.VISOES?.DEBUG && typeof Logger?.registrarLogBatch === 'function') {
//       const fmt = (arr)=>arr.map(b=>`${b.start}-${b.start+b.count-1}`).join(',');
//       Logger.registrarLogBatch([["DEBUG","VISOES",`Aba=${abaReal} (${abaKey}) | Visao=${visao} | hide=[${fmt(hideBlocks)}] | show=[${fmt(showBlocks)}] | ops=${hideBlocks.length+showBlocks.length}`]]);
//     }
//   } catch (e) {
//     try {
//       Logger.registrarLogBatch([["ERRO","VISOES",`Falha ao aplicar visão: ${e.message}`]]);
//     } catch (_e) {
//       console.error(e);
//     }
//   }
// }

// // Wrappers públicos para o menu (sem parâmetros no Apps Script)
// function VISOES__aplicarVisaoMenu_Operacional() { VISOES__aplicarVisaoMenu('OPERACIONAL'); }
// function VISOES__aplicarVisaoMenu_Completo() { VISOES__aplicarVisaoMenu('COMPLETO'); }

// // ████████████████████████████████████████████████████████████████████████
// // ▶ MÓDULO: TRIGGERS CORRIGIDO CONSERVADOR - MELHORIAS MÍNIMAS
// // ▶ BASEADO NO CÓDIGO ORIGINAL COM APENAS AJUSTES NECESSÁRIOS
// // ████████████████████████████████████████████████████████████████████████

// const Triggers = {
//   _cache: {
//     linhasProcessadas: new Set(),
//     menuConfigurado: false,
//     lastClean: 0,
//     // ✅ ADICIONADO: Controle básico de reentrada
//     ultimaExecucao: 0
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ ONEDIT: MANTIDO ORIGINAL + MELHORIAS MÍNIMAS
//   // ████████████████████████████████████████████████████████████████████████
  

// onEdit: (e) => {
//   try {
//     // ███ INÍCIO DA MODIFICAÇÃO (Flush) ███
//     // Força a planilha a salvar todas as edições pendentes ANTES de qualquer script de leitura.
//     // Isso corrige o problema de "Leitura Vencida" (Stale Read).
//     try {
//       SpreadsheetApp.flush();
//     } catch (f) {}
//     // ███ FIM DA MODIFICAÇÃO ███
//     /* <-- INÍCIO DO COMENTÁRIO
//     // ✅ MELHORADO: Controle básico de reentrada (evitar execuções muito rápidas)
//     const agora = Date.now();
//     if (agora - Triggers._cache.ultimaExecucao < 50) { // 50ms mínimo
//       return;
//     }
//     Triggers._cache.ultimaExecucao = agora;
//     FIM DO COMENTÁRIO --> */
//     // ⇩⇩⇩ NOVO: zera limite de toasts por execução (evita carregar contagem entre edições)
//     if (typeof DataValidator !== 'undefined') {
//       DataValidator._toastCount = 0;
//       DataValidator._summaryToastShown = false; // garante resumo novamente por edição
//     }

//     Logger.registrarLogBatch([
//       ["INFO", "Triggers.onEdit", "🟢 Iniciou Triggers.onEdit"]
//     ]);
//     if (!e || !e.range) return;

//     // ... (fim do bloco na linha 1559)

//     // ✅ MANTIDO: DataValidator primeiro (ordem original correta)
//     if (typeof DataValidator !== 'undefined' && DataValidator.validarEdicao) {
//       const resultadoVal = DataValidator.validarEdicao(e);
//       if (resultadoVal === false) {
//         Logger.registrarLogBatch([
//           ["DEBUG", "Triggers.onEdit", "⚠️ DataValidator retornou false - parando execução"]
//         ]);
//         return;
//       }
//       Logger.registrarLogBatch([
//         ["DEBUG", "Triggers.onEdit", "✅ DataValidator executado com sucesso"]
//       ]);
//     }

//     /* <-- INÍCIO DO COMENTÁRIO (MODIFICAÇÃO 2 - REMOÇÃO DA CHAMADA ANTIGA)

//     // ▶ Produtos → Seguros (MVP simplificado)
//     try {
//       const cfg = (typeof CONFIG !== 'undefined' && CONFIG.PRODUTOS_SEGUROS) || {};
//       if (cfg.enabled && typeof ProdutosSegurosSimples !== 'undefined' && ProdutosSegurosSimples.processar) {
//         ProdutosSegurosSimples.processar(e);
//         Logger.registrarLogBatch([["DEBUG","Triggers.onEdit","✅ ProdutosSegurosSimples executado"]]);
//       }
//     } catch (erroPSS) {
//       Logger.registrarLogBatch([["ERRO","Triggers.onEdit",`❌ Erro no ProdutosSegurosSimples: ${erroPSS.message}`]]);
//     }
//           FIM DO COMENTÁRIO --> */

//     // ✅ MANTIDO: DropdownManager segundo (integração original correta)
//     if (typeof DropdownManager !== 'undefined' && DropdownManager.processarEdicao) {
//       try {
//         DropdownManager.processarEdicao(e);
//         Logger.registrarLogBatch([
//           ["DEBUG", "Triggers.onEdit", "✅ DropdownManager executado"]
//         ]);
//       } catch (erroDropdown) {
//         Logger.registrarLogBatch([
//           ["ERRO", "Triggers.onEdit", `❌ Erro no DropdownManager: ${erroDropdown.message}`]
//         ]);
//       }
//     }

//     // ✅ MANTIDO: onEditHandler legado (compatibilidade)
//     if (typeof onEditHandler !== 'undefined' && onEditHandler.gerenciarEdicoes) {
//       try {
//         onEditHandler.gerenciarEdicoes(e);
//         Logger.registrarLogBatch([
//           ["DEBUG", "Triggers.onEdit", "✅ onEditHandler legado executado"]
//         ]);
//       } catch (erroLegado) {
//         Logger.registrarLogBatch([
//           ["ERRO", "Triggers.onEdit", `❌ Erro no onEditHandler: ${erroLegado.message}`]
//         ]);
//       }
//     }

//     // ✅ MANTIDO: TimestampManager (lógica original correta)
//     const sheet = e.range.getSheet();
//     const coluna = e.range.getColumn();
//     const configAba = CONFIG?.SHEETS?.[sheet.getName()];
//     if (configAba?.COLUNAS_MONITORADAS?.includes(Utils.colunaParaLetra(coluna))) {
//       try {
//         TimestampManager.registrarTimestamp(sheet, e.range);
//         Logger.registrarLogBatch([
//           ["DEBUG", "Triggers.onEdit", "✅ Timestamp registrado"]
//         ]);
//       } catch (erroTimestamp) {
//         Logger.registrarLogBatch([
//           ["ERRO", "Triggers.onEdit", `❌ Erro no TimestampManager: ${erroTimestamp.message}`]
//         ]);
//       }
//     }

//     // ✅ MANTIDO: BatchOperations por último (lógica original correta)
//     if (CONFIG.BATCH_OPS?.ENABLED && typeof BatchOperations?.execute === 'function') {
//       try {
//         BatchOperations.execute('onEdit');
//         Logger.registrarLogBatch([
//           ["DEBUG", "Triggers.onEdit", "✅ BatchOperations executado"]
//         ]);
//       } catch (erroBatch) {
//         Logger.registrarLogBatch([
//           ["ERRO", "Triggers.onEdit", `❌ Erro no BatchOperations: ${erroBatch.message}`]
//         ]);
//       }
//     } else {
//       Logger.registrarLogBatch([
//         ["DEBUG", "Triggers.onEdit", "ℹ️ BatchOperations não habilitado ou não disponível"]
//       ]);
//     }

//     Logger.registrarLogBatch([
//       ["INFO", "Triggers.onEdit", "🏁 Cadeia onEdit concluída"]
//     ]);

//   } catch (erro) {
//     Logger.registrarLogBatch([
//       ["ERRO_CRITICO", "Triggers.onEdit", `❌ Erro crítico: ${erro.message} | Stack: ${erro.stack}`]
//     ]);
//   }
// },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ ONOPEN: MANTIDO ORIGINAL + MENU MELHORADO
//   // ████████████████████████████████████████████████████████████████████████
  
//   onOpen: () => {
//     try {
//       if (Triggers._cache.menuConfigurado) return;
//       const ui = SpreadsheetApp.getUi();

//       const status = typeof Logger.getStatusLog === 'function'
//         ? (Logger.getStatusLog() ? '✅ Ativo' : '❌ Inativo')
//         : '✅ Ativo';

//       const modo = typeof Logger.getModo === 'function'
//         ? (Logger.getModo() === 'console' ? '🖥️ Console' : '📒 Planilha')
//         : '📒 Planilha';

//       // ✅ MELHORADO: Menu mais organizado com itens de teste
//       const menu = ui.createMenu('⚙️ Sistema GAS')
//         .addItem(`🔁 Alternar Modo de Log (${modo})`, 'Logger.alternarModo')
//         .addSeparator()
//         .addItem('🔄 Recriar Listas Suspensas', 'DropdownManager.criarListasSuspensas')
//         .addItem('🧹 Limpar Cache Dropdown', 'DropdownManager.limparAbaTemporariaDropdown')
//         .addSeparator()
//         .addItem('🧪 Testar DropdownManager', 'testarDropdownManager')
//         .addItem('🔍 Debug Sistema Completo', 'debugSistemaCompleto')
//         .addSeparator()
//         .addItem('📤 Enviar Alertas de Vencimento', 'EmailManager.verificarEViarAlertas');
//         try {
//           // considera visões por aba (e mantém compatibilidade com bloco central, se existir)
//           const sheetsCfg = (CONFIG && CONFIG.SHEETS) ? CONFIG.SHEETS : {};
//           const hasVisoesAba = Object.keys(sheetsCfg).some(k => {
//             const v = sheetsCfg[k]?.VISOES;
//             return v && (v.COMPLETO || (Array.isArray(v.OPERACIONAL) && v.OPERACIONAL.length));
//           });
//           const hasVisoesCentral = !!(CONFIG?.VISOES?.ABAS) && Object.keys(CONFIG.VISOES.ABAS).some(k => {
//             const v = CONFIG.VISOES.ABAS[k];
//             return v && (v.COMPLETO || (Array.isArray(v.OPERACIONAL) && v.OPERACIONAL.length));
//           });
//             if (hasVisoesAba || hasVisoesCentral) {
//               ui.createMenu('👁️ Visões')
//                 .addItem('Operacional', 'VISOES__aplicarVisaoMenu_Operacional')
//                 .addItem('Completo', 'VISOES__aplicarVisaoMenu_Completo')
//                 .addToUi();
//             }
          
//         } catch (_e) { /* silencioso */ }
//       // ✅ MANTIDO: Integração BatchOperations original
//       if (CONFIG.BATCH_OPS?.ENABLED) {
//         try {
//           BatchOperations.init(CONFIG);
//           Logger.registrarLogBatch([["INFO", "Triggers.onOpen", "BatchOperations inicializado"]]);
//           menu.addItem('⚡ Executar Operações em Lote', 'BatchOperations.execute');
//         } catch (erroBatch) {
//           Logger.registrarLogBatch([["ERRO", "Triggers.onOpen", `Erro BatchOperations: ${erroBatch.message}`]]);
//           menu.addItem('❌ Erro no Batch', 'Logger.alternarModo');
//         }
//       }
//       try {
//               ui.createMenu('🧮 Calculadora')
//                 .addItem('1. Preparar/Limpar Aba', 'prepararCalculadora')
//                 .addItem('2. Calcular Prêmio', 'calcularPremio')
//                 .addToUi();
//             } catch (calcError) {
//               Logger.registrarLogBatch([["ERRO", "Triggers.onOpen", `Erro ao criar menu Calculadora: ${calcError.message}`]]);
//             }
//       menu.addToUi();
//       Triggers._cache.menuConfigurado = true;
//       Logger.registrarLogBatch([["INFO", "Triggers.onOpen", "Menu configurado com sucesso"]]);
//     } catch (erro) {
//       try {
//         Logger.registrarLogBatch([["ERRO", "Triggers.onOpen", `Erro: ${erro.message}`]]);
//       } catch (erroLogger) {
//         console.error("Erro crítico no onOpen:", erro);
//       }
//     }
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ OUTRAS FUNÇÕES: MANTIDAS ORIGINAIS
//   // ████████████████████████████████████████████████████████████████████████

//   /**
//    * ✅ MANTIDO: onChangeRemoveRow original
//    */
//   onChangeRemoveRow: (e) => {
//     try {
//       if (e.changeType !== 'REMOVE_ROW') return;
//       const docsOrfaos = [];
//       const abaLog = Utils.getCachedSheet('Log do Sistema');
//       if (!abaLog) return;
      
//       const dadosLog = abaLog.getRange(2, 1, abaLog.getLastRow() - 1, 3).getValues();

//       dadosLog.forEach((linha) => {
//         const mensagem = linha[2];
//         if (mensagem && mensagem.includes('Documento criado')) {
//           const match = mensagem.match(/d\/(.+?)\//);
//           if (match) docsOrfaos.push(match[1]);
//         }
//       });

//       if (docsOrfaos.length > 0) {
//         docsOrfaos.forEach(docId => {
//           try {
//             const file = DriveApp.getFileById(docId);
//             file.setTrashed(true);
//           } catch (erro2) {
//             Logger.registrarLogBatch([["AVISO", "Triggers.onChangeRemoveRow", `Documento não encontrado: ${docId}`]]);
//           }
//         });
//         Logger.registrarLogBatch([["INFO", "Triggers.onChangeRemoveRow", `${docsOrfaos.length} documentos órfãos removidos`]]);
//       }
//     } catch (erro) {
//       Logger.registrarLogBatch([["ERRO", "Triggers.onChangeRemoveRow", erro.message]]);
//     }
//   },

//   /**
//    * ✅ MANTIDO: atualizacaoDiaria original
//    */
//   atualizacaoDiaria: () => {
//     try {
//       const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.ACOMPANHAMENTO_PARCELAS.ABA);
//       if (!sheet) return;
//       const ultimaLinha = sheet.getLastRow();
//       for (let linha = CONFIG.ACOMPANHAMENTO_PARCELAS.LINHA_INICIO; linha <= ultimaLinha; linha++) {
//         ParcelasManager.atualizarCores(sheet, linha);
//       }
//       Logger.registrarLogBatch([["INFO", "Triggers.atualizacaoDiaria", `Atualizadas ${ultimaLinha} linhas`]]);
//     } catch (erro) {
//       Logger.registrarLogBatch([["ERRO", "Triggers.atualizacaoDiaria", erro.message]]);
//     }
//   },

//   /**
//    * ✅ MANTIDO: adicionarTriggersEmail original
//    */
//   adicionarTriggersEmail: () => {
//     try {
//       ScriptApp.newTrigger('EmailManager.verificarEViarAlertas')
//         .timeBased()
//         .everyDays(1)
//         .atHour(9)
//         .create();
//       Logger.registrarLogBatch([["INFO", "Triggers.adicionarTriggersEmail", "Trigger de email configurado"]]);
//     } catch (erro) {
//       Logger.registrarLogBatch([["ERRO", "Triggers.adicionarTriggersEmail", erro.message]]);
//     }
//   },

//   // ████████████████████████████████████████████████████████████████████████
//   // ▶ FUNÇÕES NOVAS: APENAS UTILIDADES BÁSICAS
//   // ████████████████████████████████████████████████████████████████████████

//   /**
//    * ✅ NOVO: Recriar menu manualmente
//    */
//   recriarMenu: () => {
//     try {
//       Triggers._cache.menuConfigurado = false;
//       Triggers.onOpen();
//       Logger.registrarLogBatch([["INFO", "Triggers.recriarMenu", "Menu recriado"]]);
//     } catch (erro) {
//       Logger.registrarLogBatch([["ERRO", "Triggers.recriarMenu", erro.message]]);
//     }
//   }
// };

// // ████████████████████████████████████████████████████████████████████████
// // ▶ FUNÇÕES GLOBAIS: MANTIDAS ORIGINAIS
// // ████████████████████████████████████████████████████████████████████████

// function onEditManual(e) { 
//   // único ponto de verdade: sempre aciona o orquestrador
//   return Triggers.onEdit(e);
// }

// // onEdit global opcional (não usar trigger simples). Se existir, apenas encaminha.
// function _onEdit_DESATIVADO(e) { // <-- FUNÇÃO RENOMEADA
//   // Esta função não será mais acionada automaticamente pelo Google.
//   return Triggers.onEdit(e);
// }
// function onChange(e) { 
//   return (Triggers.onChangeRemoveRow && Triggers.onChangeRemoveRow(e));
// }

// function onOpen(e) { 


//   return (Triggers.onOpen && Triggers.onOpen());
// }