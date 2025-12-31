// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO 1: CONFIGURAÇÃO GLOBAL OTIMIZADA
// ████████████████████████████████████████████████████████████████████████

const CONFIG = {  
  // ►►► [CONCEITO] SHEET_NAMES: Lista das abas da planilha  
  SHEET_NAMES: [
    'PROSPECCAO🔍', // ✅ SUBSTITUÍDO
    'INTERNALIZADO🎯', // ✅ SUBSTITUÍDO
    'EM ANALISE📊', // ✅ SUBSTITUÍDO
    'CONTRATADO/LIBERACAO✅', // ✅ SUBSTITUÍDO
    'INDEFERIDO', 
    'SEGUROS🛡️', // ✅ SUBSTITUÍDO
    'LIMITES', 
    'CADASTROS🧑‍💻', // ✅ SUBSTITUÍDO
    'SEG - HISTORICO', 
    'LC - HISTORICO',
    'ENCARTEIRAMENTO🧑‍💻',
    'METAS🎯',
    'DEMANDAS DIVERSAS🔧'
    ],  
   
  // ►►► [CONCEITO] FOLDER_ID: ID da pasta do Google Drive para salvar documentos  
  FOLDER_ID: '1W6GIVcrR3KUK7ambfRJCpxA_I9LBaS9x',

  // ████████████████████████████████████████████████████████████████████████  
  // ►►► [CONCEITO] MODELO_DOCUMENTO: Configura o layout dos documentos gerados  
  // ████████████████████████████████████████████████████████████████████████  
  MODELO_DOCUMENTO: {  
    CABECALHO: "PENDENCIAS",  
    ESTILOS: {  
      FONTE: "Arial",  
      COR_TITULO: "#2c3e50"  
    }  
  },  

  // ████████████████████████████████████████████████████████████████████████  
  // ►►► [CONCEITO] SEGURANCA: Configurações de segurança e acesso  
  // ████████████████████████████████████████████████████████████████████████  
  SEGURANCA: {  
    VALIDACAO_NOMES: {  
      TAMANHO_MAXIMO: 100,  
      CARACTERES_INVALIDOS: /[\\/:*?"<>|]/g  
    },  
    NIVEL_ACESSO: "ANYONE",  
    NOTIFICACOES: true  
  },  

  // ████████████████████████████████████████████████████████████████████████  
  // ►►► [MANTIDO] BATCH_OPS: Configuração original preservada
  // ████████████████████████████████████████████████████████████████████████
  BATCH_OPS: {
    ENABLED: true,
    DASHBOARD_SHEET: 'Dashboard_Batch',
    MAX_QUEUE_SIZE: 500,          // ✅ MANTIDO: Valor original
    LOG_PERFORMANCE: true
  },

  PRODUTOS_SEGUROS: {
    enabled: true,
    origem: 'EM ANALISE📊', // ✅ SUBSTITUÍDO
    destino: 'SEGUROS🛡️', // ✅ SUBSTITUÍDO
    colunaOrigem: 'J',
    colunaProdutoDestinoNum: 3, // C
    startRow: 4,
    // mapa “EM ANALISE (letra)” → “SEGUROS (letra)”
    mapeamento: { 'A':'A', 'B':'B', 'F':'D', 'G':'G'},
    governanca: false
  },
  // ████████████████████████████████████████████████████████████████████████  
  // ►►► [CORREÇÃO] ACOMPANHAMENTO_PARCELAS: Adicionadas configurações necessárias para ParcelasManager
  // ████████████████████████████████████████████████████████████████████████
  ACOMPANHAMENTO_PARCELAS: {
    ENABLED: false, // ← nova flag: defina false só durante o teste
    ABA: 'SEGUROS🛡️', // ✅ SUBSTITUÍDO
    LINHA_INICIO: 4,
    COLUNA_DATA_CONTRATO: 13,
    COLUNA_DIA_VENCTO: 14,
    DIAS_ALERTA: 7,
    // ✅ NECESSÁRIO: Usado por ParcelasManager.desativarAcompanhamento()
    TAB_CHECKBOX_FORMATO: {
      BACKGROUND: "#FFFFFF",
      FONT_FAMILY: "Arial",
      FONT_SIZE: 10
    },
    // ✅ NECESSÁRIO: Usado por ParcelasManager.atualizarCores()
    COLUNA_DEF_ACOMPANHAMENTO: "I",  // Coluna I que define se tem acompanhamento (status do seguro)
    VALOR_SEM_ACOMPANHAMENTO: ["A CONTRATAR"],
    VALOR_COM_ACOMPANHAMENTO: ["EM CONTRATAÇÃO", "EM CONTRATAÇÃO-DILIGÊNCIA", "PAGO/CONTRATADO"],
    COLUNAS_PARCELAS: [
      { col: 16, mes: 1 },         // Janeiro
      { col: 17, mes: 2 },         // Fevereiro
      { col: 18, mes: 3 },         // Março
      { col: 19, mes: 4 },         // Abril
      { col: 20, mes: 5 },         // Maio
      { col: 21, mes: 6 }          // Junho
    ],
    CORES: {
      PAGO: '#2da544',      // Verde
      ALERTA: '#ffff00',    // Amarelo
      ATRASADO: '#ff0000',  // Vermelho
      NEUTRO: '#FFFFFF'     // Branco
    }
  },

METAS_OPERACOES_FORA_BORA: {
  ENABLED: true,

  DESTINO: {
    ABA: 'METAS🎯',
    LINHA_CABECALHO: 50,  // cabeçalho da tabela contínua OPERACOES_FORA_BORA
    PRIMEIRA_LINHA_DADOS: 51,
    COLUNA_INICIO: 'A',
    COLUNA_FIM: 'J',
  },

  ORIGENS: {
    // ==========================
    // ORIGEM: CONTRATADO/LIBERACAO✅
    // ==========================
    CONTRATADO_LIBERACAO: {
      ABA: 'CONTRATADO/LIBERACAO✅',

      FILTRO: {
        COLUNA_ELEGIVEL: 'L',     // L = FALSE → linha entra no fluxo
        VALOR_ELEGIVEL: false,
      },

      COLUNA_GRUPO: 'E',         // agrupa/classifica pelos valores da coluna E (produto)

      GRUPOS: {
        FNO: [
          'FNO - GIRO MPE',
          'FNO - GIRO ISOLADO',
          'FNO - INV - FIXO/MISTO (PROJETO)',
          'FNO - INV - MAQ E EQUIP.',
          'FNO - INV - VEICULO',
          'FNO - INV - ENERGIA VERDE',
        ],
        RPL: [
          'RPL - GIRO ESSENCIAL',
          'RPL - GIRO AMAZÔNIA',
          'RPL - CHEQUE ESPECIAL',
          'RPL - PRONAMPE',
        ],
        OUTRAS_FONTES: [
          'FUNGETUR - GIRO ISOLADO',
          'BNDES - GIRO ISOLADO',
          'BNDES - INV - MAQ E EQUIP.',
        ],
      },

      // Mapeamento origem (CONTRATADO/LIBERACAO✅) → destino (METAS🎯)
      MAPEAMENTO: {
        indicador:       { origem: null, destino: 'A' }, // nome do grupo: FNO, RPL, OUTRAS_FONTES
        cnpj:            { origem: 'A', destino: 'B' },
        cliente:         { origem: 'B', destino: 'C' },
        produto:         { origem: 'E', destino: 'D' },
        numeroOperacao:  { origem: 'F', destino: 'E' },
        docsflow:        { origem: 'G', destino: 'F' },
        valor:           { origem: 'H', destino: 'G' },
        dataInternacao:  { origem: 'I', destino: 'H' },
        dataContratacao: { origem: 'J', destino: 'I' },
        comentario:      { origem: 'K', destino: 'J' },
      },

      // Lista de métricas de "OPERAÇÕES/VALORES NÃO COMPUTADOS NO BORA"
      // Cada entrada define: indicador, como calcular (VALOR ou LINHAS),
      // de qual coluna somar e em qual célula gravar o resultado.
      SOMA_VALORES_NAO_COMP_BORA: {
        fno_contratado_valor: {
          indicador:    'FNO',   // linhas com A = "FNO"
          modo:         'VALOR', // somar valores
          colunaValor:  'G',     // somar valores da coluna G
          celulaDestino:'E4',    // gravar soma em METAS!E4
        },
        outras_fontes_valor: {
          indicador:    'OUTRAS_FONTES',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino:'E5',    // METAS!E5
        },
        capital_de_giro_comercial_valor: {
          indicador:    'RPL',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino:'E10',   // METAS!E10
        },
      },
    },

    // ==========================
    // ORIGEM: SEGUROS🛡️
    // ==========================
    SEGUROS: {
      ABA: 'SEGUROS🛡️',

      FILTRO: {
        COLUNA_ELEGIVEL: 'N',            // L = FALSE
        VALOR_ELEGIVEL: false,
        COLUNA_STATUS: 'I',              // I = 'PAGO/CONTRATADO'
        STATUS_CONTRATADO: 'PAGO/CONTRATADO',
      },

      COLUNA_GRUPO: 'C',                 // agrupa pelos valores da coluna C (produto)

      GRUPOS: {
        CAPITALIZACAO: [
          'CAP. PREMIUM',
          'CAP. VERDE CAPITALIZAÇAO',
        ],
        RISCOS_EMPRESARIAIS: [
          'SEG. RD (EQUIPAMENTOS/VEICULO)',
          'SEG. VEICULO',
          'SEG. EMP. (PRÉDIO/BENS)',
        ],
        SEGURO_VIDA: [
          'SEG. VIDA EMPRESARIAL',
        ],
        PRESTAMISTA: [
          'SEG. PRESTAMISTA',
        ],
        CONSORCIO: [
          'CONSORCIO',
        ],
        MAQUINA_ADQUIRENCIA: [
          'MAQUINA DE CARTÃO',
        ],
      },

      // Mapeamento origem (SEGUROS🛡️) → destino (METAS🎯)
      // Atualizado conforme especificação: A:B, B:C, C:D, F:E, G:F, E:G, H:H, K:I, M:J
      MAPEAMENTO: {
        indicador:       { origem: null, destino: 'A' }, // CAPITALIZACAO, RISCOS_EMPRESARIAIS, etc.
        cnpj:            { origem: 'A', destino: 'B' }, // A:B
        cliente:         { origem: 'B', destino: 'C' }, // B:C
        produto:         { origem: 'C', destino: 'D' }, // C:D
        numeroOperacao:  { origem: 'D', destino: 'E' }, // F:E
        docsflow:        { origem: 'G', destino: 'F' }, // G:F
        valor:           { origem: 'E', destino: 'G' }, // E:G (uso duplo de E conforme roteiro)
        dataInternacao:  { origem: 'H', destino: 'H' }, // H:H
        dataContratacao: { origem: 'K', destino: 'I' }, // K:I
        comentario:      { origem: 'M', destino: 'J' }, // M:J
      },

      // Lista de métricas de "OPERAÇÕES/VALORES NÃO COMPUTADOS NO BORA" para SEGUROS
      SOMA_VALORES_NAO_COMP_BORA: {
        capitalizacao_valor: {
          indicador:    'CAPITALIZACAO',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino: null,   // célula ainda a definir
        },
        riscos_empresariais_valor: {
          indicador:    'RISCOS_EMPRESARIAIS',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino:'E11',   // METAS!E11
        },
        seguro_de_vida_valor: {
          indicador:    'SEGURO_VIDA',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino:'E12',   // METAS!E12
        },
        prestamista_valor: {
          indicador:    'PRESTAMISTA',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino:'E15',   // METAS!E15
        },
        consorcio_valor: {
          indicador:    'CONSORCIO',
          modo:         'VALOR',
          colunaValor:  'G',
          celulaDestino:'E17',   // METAS!E16 (soma dos valores do CONSORCIO)
        },
        consorcio_qtd: {
          indicador:    'CONSORCIO',
          modo:         'LINHAS', // contar linhas, não somar valores
          colunaValor:  null,
          celulaDestino:'E16',   // METAS!E17 (quantidade de linhas CONSORCIO)
        },
        maquina_adquirencia_qtd: {
          indicador:    'MAQUINA_ADQUIRENCIA',
          modo:         'LINHAS',
          colunaValor:  null,
          celulaDestino:'E18',   // METAS!E18 (quantidade de linhas MAQUINA_ADQUIRENCIA)
        },
      },
    },

        ENCARTEIRAMENTO: {
          ABA: 'ENCARTEIRAMENTO🧑‍💻',

          FILTRO: {
            // 1) Coluna G (7) = FALSE
            COLUNA_ELEGIVEL: 'G',
            VALOR_ELEGIVEL: false,

            // 2) Coluna C (3) = 'SOL. ENCARTEIRAMENTO' ou 'ENCARTEIRADO'
            COLUNA_STATUS: 'C',
            STATUS_VALORES: ['SOL. ENCARTEIRAMENTO', 'ENCARTEIRADO'],
          },

          // Mapeamento origem (ENCARTEIRAMENTO🧑‍💻) → destino (METAS🎯)
          // A:B, B:C, E:H, F:I, D:J
          MAPEAMENTO: {
            indicador:       { origem: null, destino: 'A' }, // preenchido no código com "ENCARTEIRAMENTO"

            cnpj:            { origem: 'A', destino: 'B' },  // A → B
            cliente:         { origem: 'B', destino: 'C' },  // B → C

            // D, E, F, G da METAS não são usados neste fluxo de ENCARTEIRAMENTO
            dataInternacao:  { origem: 'D', destino: 'H' },  // E → H
            dataContratacao: { origem: 'F', destino: 'I' },  // F → I
            comentario:      { origem: 'E', destino: 'J' },  // D → J
          },

          // Métrica de "OPERAÇÕES NÃO COMPUTADAS NO BORA" para ENCARTEIRAMENTO
          SOMA_VALORES_NAO_COMP_BORA: {
            encarteiramento_qtd: {
              indicador:    'ENCARTEIRAMENTO', // usa a coluna A da METAS
              modo:         'LINHAS',          // conta número de linhas
              colunaValor:  null,              // não usado em 'LINHAS'
              celulaDestino:'E13',             // METAS🎯!E19 (ajuste se já estiver ocupada)
            },
          },
        },
      
  },
},
SHEETS: {
    // 1. PROSPECCAO
    'PROSPECCAO🔍': {
      START_ROW: 4,
      
      MOVER_LINHA_CONFIG: {
        13: {
          true: {
            destino: 'CADASTROS🧑‍💻',
            APAGAR_LINHA_ORIGEM: false,
            linhaDestino: (aba) => {
              const valores = aba.getRange("A4:A").getValues();
              const index = valores.findIndex(row => row[0] === '');
              return index === -1 ? aba.getLastRow() + 1 : index + 4;
            },
            colunas: {
              origem: [1, 2],
              destino: [1, 2]
            }
          }
        }
      },

      COLUNAS_MONITORADAS: ["G", "I", "J", "K", "L", "M", "N"],
      COLUNA_TIMESTAMP: "R",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 18,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          11: ['formatarData'],
          15: ['formatarData']
        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["B", "G", "H", "I", "J", "N", "O"]
      }
    },

    // 2. INTERNALIZADO
    'INTERNALIZADO🎯': {
      START_ROW: 4,

      MOVER_LINHA_CONFIG: {
        11: {
            "EM ANALISE": {
              destino: 'EM ANALISE📊',
              APAGAR_LINHA_ORIGEM: false,
              linhaDestino: (aba) => {
                const valores = aba.getRange("A4:A").getValues();
                const index = valores.findIndex(row => row[0] === '');
                return index === -1 ? aba.getLastRow() + 1 : index + 4;
              },
              colunas: { origem: [1,2,3,4,5,6,7,8,9], destino: [1,2,3,4,5,6,7,8,9] }
            },
            "ATUALIZAR CADASTRO": {
              destino: 'CADASTROS🧑‍💻',
              APAGAR_LINHA_ORIGEM: false,
              linhaDestino: (aba) => {
                const valores = aba.getRange("A4:A").getValues();
                const index = valores.findIndex(row => row[0] === '');
                return index === -1 ? aba.getLastRow() + 1 : index + 4;
              },
              colunas: { origem: [1, 2], destino: [1, 2] }
            },
            "SOL. DEMANDA PARALELA": {
              destino: 'DEMANDAS DIVERSAS🔧',
              APAGAR_LINHA_ORIGEM: false,
              linhaDestino: (aba) => {
                const valores = aba.getRange("A4:A").getValues();
                const index = valores.findIndex(row => row[0] === '');
                return index === -1 ? aba.getLastRow() + 1 : index + 4;
              },
              colunas: { origem: [1, 2], destino: [1, 2] }
            }
          },

        12: {
          "ATUALIZAR CADASTRO": {
            destino: 'CADASTROS🧑‍💻',
            APAGAR_LINHA_ORIGEM: false,
            linhaDestino: (aba) => {
              const valores = aba.getRange("A4:A").getValues();
              const index = valores.findIndex(row => row[0] === '');
              return index === -1 ? aba.getLastRow() + 1 : index + 4;
            },
            colunas: { origem: [1, 2], destino: [1, 2] }
          },
        }
      },

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: 6,
        LINK: 13,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 13, // Coluna M (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["J", "K", "L", "M", "N"],
      COLUNA_TIMESTAMP: "O",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 14,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          6: ['MultiFormatoProcesso'],
          7: ['FormatoDocsflow'],
          9: ['formatarData'],
          14: ['formatarData']
        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["A", "B", "E", "H", "J", "K", "L", "M", "O"]
      },

      LISTAS_SUSPENSAS: [
        {
          colunaCategoria: "K",
          colunaSubcategoria: "L",
          fonte: { nomeAba: "DADOS", colunaCategoria: "M", colunaSubcategoria: "N", linhaInicial: 6 },
          cores: {
            "A INTERNALIZAR": { fundo: "#ffe598", texto: "#000000" },
            "INTERNALIZADO": { fundo: "#fdbf12", texto: "#000000" },
            "CONFEC. PROPOSTA": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "CONFEC. PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#000000" },
            "ASSINATURA PROPOSTA": { fundo: "#007aff", texto: "#FFFFFF" },
            "ASSINATURA PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#FFFFFF" },
            "EM ANALISE": { fundo: "#2DA544", texto: "#FFFFFF" }
          },
          herdarCorSubcategoria: true,
          coresSubcategoria: {}
        }
      ]
    },

    // 3. EM ANALISE
    'EM ANALISE📊': {
      START_ROW: 4,

      MOVER_LINHA_CONFIG: {
  11: {
    "CONTRATADO": {
      destino: 'CONTRATADO/LIBERACAO✅',
      APAGAR_LINHA_ORIGEM: true,
      linhaDestino: (aba) => {
        const valores = aba.getRange('A4:A').getValues();
        const index = valores.findIndex(row => row[0] === '');
        return index === -1 ? aba.getLastRow() + 1 : index + 4;
      },
      colunas: {
        origem: [
          1, 2, 3, 4, 5, 6, 7, 8, 9,
          (sheet, linha) => {
            const valTexto = sheet.getRange(linha, 14).getDisplayValue();
            const valLimpo = valTexto ? valTexto.trim() : "";
            if (/^\\d{1,2}[\\/.-]\\d{1,2}[\\/.-]\\d{4}$/.test(valLimpo)) {
              return valLimpo;
            }
            return new Date();
          }
        ],
        destino: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
      },
      _deveMover: (sheet, linha) => {
        const valorN = sheet.getRange(linha, 14).getValue();
        return valorN && (valorN.toString().trim() !== '');
      }
    },

    "INDEFERIDO": {
      destino: 'INDEFERIDO',
      APAGAR_LINHA_ORIGEM: true,
      linhaDestino: (aba) => {
        const valores = aba.getRange('A4:A').getValues();
        const index = valores.findIndex(row => row[0] === '');
        return index === -1 ? aba.getLastRow() + 1 : index + 4;
      },
      colunas: {
        origem: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        destino: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
      }
    },

    "ATUALIZAR CADASTRO": {
      destino: 'CADASTROS🧑‍💻',
      APAGAR_LINHA_ORIGEM: false,
      linhaDestino: (aba) => {
        const valores = aba.getRange("A4:A").getValues();
        const index = valores.findIndex(row => row[0] === '');
        return index === -1 ? aba.getLastRow() + 1 : index + 4;
      },
      colunas: { origem: [1, 2], destino: [1, 2] }
    },

    "SOL. DEMANDA PARALELA": {
      destino: 'DEMANDAS DIVERSAS🔧',
      APAGAR_LINHA_ORIGEM: false,
      linhaDestino: (aba) => {
        const valores = aba.getRange("A4:A").getValues();
        const index = valores.findIndex(row => row[0] === '');
        return index === -1 ? aba.getLastRow() + 1 : index + 4;
      },
      colunas: { origem: [1, 2], destino: [1, 2] }
    }
  }
},

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: 6,
        LINK: 13,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 13, // Coluna P (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["J", "K", "L", "M", "N"],
      COLUNA_TIMESTAMP: "O",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 14,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          6: ['MultiFormatoProcesso'],
          7: ['FormatoDocsflow'],
          9: ['formatarData'],
          14: ['formatarData']
        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["B", "E", "F", "G", "H", "J", "K", "L", "M", "O"]
      },

      LISTAS_SUSPENSAS: [
        {
          colunaCategoria: "K",
          colunaSubcategoria: "L",
          fonte: { nomeAba: "DADOS", colunaCategoria: "P", colunaSubcategoria: "Q", linhaInicial: 4 },
          cores: {
            "INTERNALIZADO": { fundo: "#FDBF12", texto: "#000000" },
            "EM FILA PARA CONFECCAO": { fundo: "#FDBF12", texto: "#000000" },
            "CONFEC. PROPOSTA": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "CONFEC. PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#000000" },
            "ASSINATURA PROPOSTA": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "ASSINATURA PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#000000" },
            "_EM ANALISE_": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "EM ANALISE - DILIGENCIA": { fundo: "#FD7E12", texto: "#000000" },
            "ESTRUTURACAO": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "ESTRUTURACAO - DILIGENCIA": { fundo: "#FDBF12", texto: "#000000" },
            "DESPACHO DO COMITE COMPETENTE": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "PRE CONTRATACAO": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "PRE CONTRATACAO - DILIGENCIA": { fundo: "#FD7E12", texto: "#000000" },
            "CONTRATADO": { fundo: "#2DA544", texto: "#FFFFFF" },
            "INDEFERIDO": { fundo: "#DC2626", texto: "#FFFFFF" },
            "RECONSIDERACAO": { fundo: "#FD7E12", texto: "#000000" }
          },
          herdarCorSubcategoria: true,
          coresSubcategoria: {}
        }
      ]
    },

    // 4. CONTRATADO/LIBERACAO
    'CONTRATADO/LIBERACAO✅': {
      START_ROW: 4,

      MOVER_LINHA_CONFIG: {
        11: {
          "PROPOSTA CONTRATADA": {
            destino: 'SEGUROS🛡️',
            APAGAR_LINHA_ORIGEM: true,
            colunas: {
              origem: [1, 2, 3],
              destino: [3, 1, 2]
            }
          }
        }
      },

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: 6,
        LINK: 11,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 11, // Coluna M (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["K", "L", "M"],
      COLUNA_TIMESTAMP: "O",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 14,
        regrasPorColuna: {
          6: ['MultiFormatoProcesso'],
          9: ['formatarData'],
          10: ['formatarData'],
          11: ['FormatoDocsflow'],
          12: ['formatarData']
        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["B", "E", "F", "G", "H", "J", "K", "L", "M", "N"]
      }
    },

    // 5. SEGUROS
    'SEGUROS🛡️': {
      START_ROW: 4,

      MOVER_LINHA_CONFIG: {
        9: {
          "ATUALIZAR CADASTRO": {
            destino: 'CADASTROS🧑‍💻',
            APAGAR_LINHA_ORIGEM: false,
            linhaDestino: (aba) => {
              const valores = aba.getRange("A4:A").getValues();
              const index = valores.findIndex(row => row[0] === '');
              return index === -1 ? aba.getLastRow() + 1 : index + 4;
            },
            colunas: { origem: [1, 2], destino: [1, 2] }
          },
          "SOL. DEMANDA PARALELA": {
            destino: 'DEMANDAS DIVERSAS🔧',
            APAGAR_LINHA_ORIGEM: false,
            linhaDestino: (aba) => {
              const valores = aba.getRange("A4:A").getValues();
              const index = valores.findIndex(row => row[0] === '');
              return index === -1 ? aba.getLastRow() + 1 : index + 4;
            },
            colunas: { origem: [1, 2], destino: [1, 2] }
          }
        }
      },

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: [2, 3, 7],
        LINK: 13,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 13, // Coluna M (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"],
      COLUNA_TIMESTAMP: "O",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 14,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          4: ['MultiFormatoProcesso'],
          8: ['formatarData'],
          11: ['formatarData'],
          12: ['formatarData']
        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["B", "C", "D", "E", "F", "I", "J", "k", "L", "M", "N", "O"]
      },

      LISTAS_SUSPENSAS: [
        {
          colunaCategoria: "I",
          colunaSubcategoria: "J",
          fonte: {
            nomeAba: "DADOS",
            colunaCategoria: "U",
            colunaSubcategoria: "V",
            linhaInicial: 6
          },
          cores: {
            "A CONTRATAR": { fundo: "#FDBF12", texto: "#000000" },
            "EM CONTRATACAO": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "EM CONTRATACAO-DILIGENCIA": { fundo: "#FD7E12", texto: "#000000" },
            "PAGO/CONTRATADO": { fundo: "#2da544", texto: "#FFFFFF" },
            "PROP CANCELADA": { fundo: "#DD3544", texto: "#FFFFFF" }
          },
          herdarCorSubcategoria: true,
          coresSubcategoria: {}
        }
      ]
    },

    // 6. LIMITES
    LIMITES: {
      START_ROW: 4,

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: 2,
        LINK: 10,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 10, // Coluna J (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["H", "I", "J", "K", "L", "M", "N", "O"],
      COLUNA_TIMESTAMP: "P",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 14,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          6: ['FormatoDocsflow'],
          7: ['formatarData'],
          13: ['formatarData'],
          14: ['formatarData'],
        }
      },

      LISTAS_SUSPENSAS: [
        {
          colunaCategoria: "H",
          colunaSubcategoria: "I",
          fonte: {
            nomeAba: "DADOS",
            colunaCategoria: "Z",
            colunaSubcategoria: "AA",
            linhaInicial: 6
          },
          cores: {
            "INTERNALIZADO": { fundo: "#6c747d", texto: "#FFFFFF" },
            "EM CONFECÇÃO": { fundo: "#fd7e12", texto: "#FFFFFF" },
            "EM CONFECÇÃO - PENDÊNCIAS": { fundo: "#fdbf12", texto: "#000000" },
            "EM ANALISE": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "EM ANALISE - DILIGENCIA": { fundo: "#fdbf12", texto: "#000000" },
            "DEFERIDO": { fundo: "#2da544", texto: "#FFFFFF" }
          },
          herdarCorSubcategoria: true,
          coresSubcategoria: {}
        }
      ]
    },

    // 7. CADASTROS
    'CADASTROS🧑‍💻': {
      START_ROW: 4,

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: 2,
        LINK: 8,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 8, // Coluna I (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["B", "C", "E", "F", "G", "H", "I"],
      COLUNA_TIMESTAMP: "J",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 14,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          4: ['formatarData'],
          5: ['FormatoDocsflow'],
          9: ['formatarData']
        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["B", "C", "E", "F", "G", "H", "I", "J"]
      },

      LISTAS_SUSPENSAS: [
        {
          colunaCategoria: "F",
          colunaSubcategoria: "G",
          fonte: {
            nomeAba: "DADOS",
            colunaCategoria: "AD",
            colunaSubcategoria: "AE",
            linhaInicial: 6
          },
          cores: {
            "A ATUALIZAR": { fundo: "#FDBF12", texto: "#000000" },
            "EM ATUALIZAÇÃO": { fundo: "#87CEEB", texto: "#FFFFFF" },
            "EM ATUALIZAÇÃO - DILIGÊNCIA": { fundo: "#fd7e12", texto: "#000000" },
            "ATUALIZADO/CONCLUIDO": { fundo: "#2da544", texto: "#FFFFFF" }
          },
          herdarCorSubcategoria: true,
          coresSubcategoria: {}
        }
      ]
    },
    // 2. INTERNALIZADO
    'DEMANDAS DIVERSAS🔧': {
      START_ROW: 4,

      // MOVER_LINHA_CONFIG: {
      //   11: {
      //     "EM ANALISE": {
      //       destino: 'EM ANALISE📊',
      //       APAGAR_LINHA_ORIGEM: false,
      //       linhaDestino: (aba) => {
      //         const valores = aba.getRange("A4:A").getValues();
      //         const index = valores.findIndex(row => row[0] === '');
      //         return index === -1 ? aba.getLastRow() + 1 : index + 4;
      //       },
      //       colunas: { origem: [1, 2, 3, 4, 5, 6, 7, 8, 9], destino: [1, 2, 3, 4, 5, 6, 7, 8, 9] }
      //     }
      //   },
      //   12: {
      //     "ATUALIZAR CADASTRO": {
      //       destino: 'CADASTROS🧑‍💻',
      //       APAGAR_LINHA_ORIGEM: false,
      //       linhaDestino: (aba) => {
      //         const valores = aba.getRange("A4:A").getValues();
      //         const index = valores.findIndex(row => row[0] === '');
      //         return index === -1 ? aba.getLastRow() + 1 : index + 4;
      //       },
      //       colunas: { origem: [1, 2], destino: [1, 2] }
      //     }
      //   }
      // },

      GOOGLE_DOC_CONFIG: {
        FILE_NAME: 6,
        LINK: 8,        // ✅ AJUSTADO: Mesmo índice da TRIGGER_COL
        TRIGGER_COL: 8, // Coluna M (Comentários)
        FILE_ID: 20
      },

      COLUNAS_MONITORADAS: ["A", "B", "C", "D", "E", "J","K", "L", "M", "N"],
      COLUNA_TIMESTAMP: "K",
      FORMATO: "dd/MM/yyyy",
      TRAVAR_HORA_EM_ZERO: true,

      VALIDACAO: {
        ativo: true,
        colInicio: 1,
        colFim: 11,
        regrasPorColuna: {
          1: ['cpfOuCnpjValido'],
          6: ['FormatoDocsflow'],
          5 : ['formatarData'],
          10: ['formatarData'],
          4: ['MultiFormatoProcesso'],

        }
      },

      VISOES: {
        COMPLETO: "TODAS",
        OPERACIONAL: ["B", "C", "E", "F", "G", "H", "J", "K"]
      },

      LISTAS_SUSPENSAS: [
        {
          colunaCategoria: "G",
          colunaSubcategoria: "H",
          fonte: { nomeAba: "DADOS", colunaCategoria: "F", colunaSubcategoria: "G", linhaInicial: 6 },
          cores: {
            "A INTERNALIZAR": { fundo: "#ffe598", texto: "#000000" },
            "INTERNALIZADO": { fundo: "#fdbf12", texto: "#000000" },
            "CONFEC. PROPOSTA": { fundo: "#3d9aff", texto: "#FFFFFF" },
            "CONFEC. PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#000000" },
            "ASSINATURA PROPOSTA": { fundo: "#007aff", texto: "#FFFFFF" },
            "ASSINATURA PROPOSTA - PENDENCIAS": { fundo: "#FD7E12", texto: "#FFFFFF" },
            "EM ANALISE": { fundo: "#2DA544", texto: "#FFFFFF" }
          },
          herdarCorSubcategoria: true,
          coresSubcategoria: {}
        }
      ]
    },
  },
  // ████████████████████████████████████████████████████████████████████████  
  // ►►► [CONCEITO] MOVED_COLUMNS - Colunas movidas entre abas por status  
  // ████████████████████████████████████████████████████████████████████████  
MOVED_COLUMNS: {
    'EM ANALISE📊': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], // ✅ SUBSTITUÍDO
    'CONTRATADO/LIBERACAO✅': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], // ✅ SUBSTITUÍDO
    INDEFERIDO: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], // (MANTIDO - Não faz parte do mapeamento)
    'CADASTROS🧑‍💻': [1, 2] // ✅ SUBSTITUÍDO
  },

  // ████████████████████████████████████████████████████████████████████████  
  // ►►► [CONCEITO] Regra de expiração de linhas  
  // ████████████████████████████████████████████████████████████████████████  
  EXPIRED: { 
    THRESHOLD_DAYS: 2,
    COLOR: "#FFCCCC"
   } 
  };





// ████████████████████████
// ▶ CONFIG.UI — parâmetros de validação (toast + visuais)
CONFIG.UI = CONFIG.UI || {};
CONFIG.UI.VALIDACAO = CONFIG.UI.VALIDACAO || {};
CONFIG.UI.VALIDACAO.TOAST = {
  ativar: true,          // manter toast em DEV/PROD
  limitePorExecucao: 3,  // no máx. 3 toasts por onEdit
  cooldownMs: 1500       // suprimir toasts repetidos da mesma célula por ~1,5s
};
// (nenhum '};' extra depois do bloco)
// ████████████████████████


// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO 2: ACOMPANHAMENTO DE PARCELAS 
//    
//    
// ████████████████████████████████████████████████████████████████████████

const ParcelasManager = {
  /**
   * Remove checkboxes, limpa validações e redefine formatação das colunas de parcelas
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba "SEGUROS"
   * @param {number} linha - Linha a ser processada
   */
  desativarAcompanhamento: (sheet, linha) => {
  const config = CONFIG.ACOMPANHAMENTO_PARCELAS;

  config.COLUNAS_PARCELAS.forEach(({ col }) => {
    try {
      const celula = sheet.getRange(linha, col);

      try { celula.removeCheckbox(); } catch(e) {}
      celula.clearDataValidations();
      SpreadsheetApp.flush();

      celula.clearContent();
      celula.clearFormat();

      // Remove cor personalizada, retornando à cor padrão da planilha
      celula.setBackground(null); 
      celula.setFontFamily(config.TAB_CHECKBOX_FORMATO.FONT_FAMILY);
      celula.setFontSize(config.TAB_CHECKBOX_FORMATO.FONT_SIZE);
      celula.setNote('');

    } catch (ex) {  // ✅ CORRIGIDO: } adicionado antes do catch
      const errorMsg = `(${linha},${col}) ${ex?.message || ''} | ${ex?.stack || ''}`;
      try {
        Logger.registrarLogBatch([['ERRO','ParcelasManager.desativarAcompanhamento', errorMsg]]);
      } catch (_) {
        console.error('ERRO ParcelasManager.desativarAcompanhamento:', errorMsg);
      }
    }
  }); // ✅ CORRIGIDO: }); adicionado para fechar forEach

  SpreadsheetApp.flush(); // ✅ CORRIGIDO: Movido para fora do forEach
},

  /**
   * Insere checkboxes e aplica formatação padrão nas colunas de parcelas
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba "SEGUROS"
   * @param {number} linha - Linha a ser processada
   */
  reativarAcompanhamento: (sheet, linha) => {
    const config = CONFIG.ACOMPANHAMENTO_PARCELAS;
    const colunas = config.COLUNAS_PARCELAS.map(p => p.col);
    const range = sheet.getRange(linha, Math.min(...colunas), 1, colunas.length);
    
    range.clearDataValidations().clearFormat();
    const validation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    range.setDataValidation(validation)
      .setBackground(config.TAB_CHECKBOX_FORMATO.BACKGROUND)
      .setFontFamily(config.TAB_CHECKBOX_FORMATO.FONT_FAMILY)
      .setFontSize(config.TAB_CHECKBOX_FORMATO.FONT_SIZE);
    SpreadsheetApp.flush();
  },

  /**
   * Calcula a data de vencimento de uma parcela
   * @param {Date} dataContrato - Data base do contrato
   * @param {number} mesParcela - Número da parcela (1-6)
   * @param {number} diaVencto - Dia de vencimento configurado
   */
  calcularVencimento: (dataContrato, mesParcela, diaVencto) => {
    if (!dataContrato || isNaN(dataContrato.getTime())) return null;
    const vencimento = new Date(dataContrato);

    // Lógica de cálculo (mantida igual à original)
    if (mesParcela === 1) vencimento.setDate(vencimento.getDate() + 3);
    else if (mesParcela === 2) {
      const candidate = new Date(dataContrato);
      candidate.setMonth(candidate.getMonth() + 1);
      candidate.setDate(diaVencto);
      const diffDays = Math.ceil((candidate - dataContrato) / (1000 * 60 * 60 * 24));
      vencimento.setTime((diffDays >= 29 && diffDays <= 31) ? candidate.getTime() : candidate.setMonth(candidate.getMonth() + 1));
    } else {
      vencimento.setMonth(vencimento.getMonth() + mesParcela);
      vencimento.setDate(diaVencto);
      while (vencimento.getDate() !== diaVencto) vencimento.setDate(vencimento.getDate() - 1);
    }
    return vencimento;
  },

  /**
   * Define a cor da célula com base no status do pagamento
   * @param {Date} vencimento - Data de vencimento
   * @param {boolean} pago - Indica se a parcela está paga
   */
  definirCorParcela: (vencimento, pago) => {
    if (!vencimento) return CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.NEUTRO;
    const diffDias = Math.ceil((vencimento - new Date()) / (1000 * 60 * 60 * 24));
    
    return pago ? CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.PAGO :
           diffDias < 0 ? CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.ATRASADO :
           diffDias <= CONFIG.ACOMPANHAMENTO_PARCELAS.DIAS_ALERTA ? CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.ALERTA :
           CONFIG.ACOMPANHAMENTO_PARCELAS.CORES.NEUTRO;
  },

  /**
   * Atualiza cores e notas de todas as células de parcelas em uma linha
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba "SEGUROS"
   * @param {number} linha - Linha processada
   */
  atualizarCores: (sheet, linha) => {
    const config = CONFIG.ACOMPANHAMENTO_PARCELAS;
    const categoria = sheet.getRange(linha, Utils.colunaParaIndice(config.COLUNA_DEF_ACOMPANHAMENTO)).getDisplayValue().trim();
    
    if (!config.VALOR_COM_ACOMPANHAMENTO.includes(categoria)) return;
    
    const dataContrato = sheet.getRange(linha, config.COLUNA_DATA_CONTRATO).getValue();
    const diaVencto = sheet.getRange(linha, config.COLUNA_DIA_VENCTO).getValue();
    
    if (!dataContrato || !diaVencto) {
      config.COLUNAS_PARCELAS.forEach(({ col }) => {
        sheet.getRange(linha, col)
          .setBackground(config.TAB_CHECKBOX_FORMATO.BACKGROUND)
          .setNote('Preencha a data de vencimento na coluna N!');
      });
      return;
    }

    config.COLUNAS_PARCELAS.forEach(({ col, mes }) => {
      const vencimento = ParcelasManager.calcularVencimento(dataContrato, mes, diaVencto);
      const pago = sheet.getRange(linha, col).isChecked();
      const cor = ParcelasManager.definirCorParcela(vencimento, pago);
      
      sheet.getRange(linha, col)
        .setBackground(cor)
        .setNote(vencimento ? `Vencimento: ${Utilities.formatDate(vencimento, Utils.getTimezone(), 'dd/MM/yyyy')}` : '');
    });
  }
};



// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO 3: UTILITÁRIOS CORRIGIDOS E OTIMIZADOS
// ████████████████████████████████████████████████████████████████████████
const Utils = {
  _cache: {
    // Unificado tudo em um só lugar
    spreadsheet: null,
    sheets: {},
    configs: {},
    statusValues: {},
    lastProcessedRow: 0,
    // ✅ NOVO: Controle de cache otimizado
    timestamps: {},
    TTL: 5 * 60 * 1000  // 5 minutos TTL
  },

  // ████████████████████████████████████████████████████████████████████████
  // ▶ FUNÇÕES DE CONVERSÃO OTIMIZADAS
  // ████████████████████████████████████████████████████████████████████████

  /**
   * ✅ FUNÇÃO PRINCIPAL: Converte letras de coluna para número (Ex: "A"→1, "AB"→28)
   * Mantém compatibilidade com código existente
   */
  colunaParaIndice: (coluna) => {
    try {
      if (typeof coluna !== 'string') throw new Error('Tipo inválido para conversão');
      coluna = coluna.toString().toUpperCase().replace(/[^A-Z]/g, '');
      let indice = 0;
      for (let i = 0; i < coluna.length; i++) {
        indice = indice * 26 + (coluna.charCodeAt(i) - 64);
      }
      return indice;
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.colunaParaIndice', `Falha na conversão: ${erro.message}`]]);
      return 0;
    }
  },

  /**
   * ✅ FUNÇÃO SECUNDÁRIA: Alias para compatibilidade (mantém código existente funcionando)
   */
  columnLetterToNumber: (column) => {
    return Utils.colunaParaIndice(column);
  },

  /**
   * ✅ CORRIGIDO: Converte número para letra (Ex: 1→"A", 28→"AB")
   */
  colunaParaLetra: (coluna) => {
    try {
      if (typeof coluna !== 'number' || coluna < 1) throw new Error('Número inválido para conversão');
      let letra = '';
      while (coluna > 0) {
        const temp = (coluna - 1) % 26;
        letra = String.fromCharCode(temp + 65) + letra;
        coluna = Math.floor((coluna - temp - 1) / 26);
      }
      return letra || 'A';
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.colunaParaLetra', `Falha na conversão: ${erro.message}`]]);
      return 'A';
    }
  },

  // ████████████████████████████████████████████████████████████████████████
  // ▶ SISTEMA DE CACHE OTIMIZADO
  // ████████████████████████████████████████████████████████████████████████

  /** 
   * ✅ OTIMIZADO: Retorna a planilha ativa (com cache e TTL)
   */
  getCachedSpreadsheet: () => {
    const now = Date.now();
    const key = 'spreadsheet';
    
    // Verificar TTL
    if (Utils._cache.timestamps[key] && (now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
      Utils._cache.spreadsheet = null;
      delete Utils._cache.timestamps[key];
    }
    
    if (!Utils._cache.spreadsheet) {
      try {
        Utils._cache.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        Utils._cache.timestamps[key] = now;
      } catch (erro) {
        Logger.registrarLogBatch([['ERRO', 'Utils.getCachedSpreadsheet', `Falha ao acessar planilha: ${erro.message}`]]);
        return null;
      }
    }
    return Utils._cache.spreadsheet;
  },
    getTimezone: () => {
      try {
        const ssTz = SpreadsheetApp.getActiveSpreadsheet()?.getSpreadsheetTimeZone();
        return ssTz || Session.getScriptTimeZone() || 'GMT-3';
      } catch (_) {
        return 'GMT-3';
      }
    },
  /** 
   * ✅ OTIMIZADO: Retorna uma aba pelo nome (com cache e TTL)
   */
  getCachedSheet: (sheetName) => {
    if (!sheetName) {
      Logger.registrarLogBatch([['ERRO', 'Utils.getCachedSheet', 'Nome da aba não fornecido']]);
      return null;
    }

    const now = Date.now();
    const key = `sheet_${sheetName}`;
    
    // Verificar TTL
    if (Utils._cache.timestamps[key] && (now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
      delete Utils._cache.sheets[sheetName];
      delete Utils._cache.timestamps[key];
    }
    
    if (!Utils._cache.sheets[sheetName]) {
      try {
        const ss = Utils.getCachedSpreadsheet();
        if (!ss) return null;
        
        Utils._cache.sheets[sheetName] = ss.getSheetByName(sheetName);
        Utils._cache.timestamps[key] = now;
        
        if (!Utils._cache.sheets[sheetName]) {
          Logger.registrarLogBatch([['AVISO', 'Utils.getCachedSheet', `Aba "${sheetName}" não encontrada`]]);
          return null;
        }
      } catch (erro) {
        Logger.registrarLogBatch([['ERRO', 'Utils.getCachedSheet', `Falha ao acessar aba "${sheetName}": ${erro.message}`]]);
        return null;
      }
    }
    return Utils._cache.sheets[sheetName];
  },

  /** 
   * ✅ OTIMIZADO: Retorna valores da coluna de status (com cache e TTL)
   */
  getCachedStatusValues: (sheetName, column) => {
    if (!sheetName || !column) {
      Logger.registrarLogBatch([['ERRO', 'Utils.getCachedStatusValues', 'Parâmetros inválidos']]);
      return [];
    }

    const now = Date.now();
    const key = `${sheetName}_${column}`;
    
    // Verificar TTL
    if (Utils._cache.timestamps[key] && (now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
      delete Utils._cache.statusValues[key];
      delete Utils._cache.timestamps[key];
    }
    
    if (!Utils._cache.statusValues[key]) {
      try {
        const sheet = Utils.getCachedSheet(sheetName);
        if (!sheet) return [];
        
        const colIndex = typeof column === 'string' ? Utils.colunaParaIndice(column) : column;
        Utils._cache.statusValues[key] = sheet.getRange(1, colIndex, sheet.getLastRow()).getValues();
        Utils._cache.timestamps[key] = now;
      } catch (erro) {
        Logger.registrarLogBatch([['ERRO', 'Utils.getCachedStatusValues', `Falha ao ler valores: ${erro.message}`]]);
        return [];
      }
    }
    return Utils._cache.statusValues[key];
  },

  /**
   * ✅ NOVO: Limpeza automática de cache expirado
   */
  limparCacheExpirado: () => {
    try {
      const now = Date.now();
      let itensLimpos = 0;
      
      Object.keys(Utils._cache.timestamps).forEach(key => {
        if ((now - Utils._cache.timestamps[key]) > Utils._cache.TTL) {
          // Limpar timestamp
          delete Utils._cache.timestamps[key];
          
          // Limpar dados relacionados
          if (key === 'spreadsheet') {
            Utils._cache.spreadsheet = null;
          } else if (key.startsWith('sheet_')) {
            const sheetName = key.replace('sheet_', '');
            delete Utils._cache.sheets[sheetName];
          } else {
            delete Utils._cache.statusValues[key];
          }
          itensLimpos++;
        }
      });
      
      if (itensLimpos > 0) {
        Logger.registrarLogBatch([['INFO', 'Utils.limparCacheExpirado', `${itensLimpos} itens de cache expirados limpos`]]);
      }
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.limparCacheExpirado', `Falha na limpeza: ${erro.message}`]]);
    }
  },

  /**
   * ✅ MELHORADO: Invalidação de cache com tipos específicos
   */
  invalidateCache: (key) => {
    try {
      if (key === 'all') {
        Utils._cache.sheets = {};
        Utils._cache.configs = {};
        Utils._cache.statusValues = {};
        Utils._cache.timestamps = {};
        Utils._cache.spreadsheet = null;
        Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', 'Todo o cache foi limpo']]);
      } else if (key === 'sheets') {
        Utils._cache.sheets = {};
        // Limpar timestamps de sheets
        Object.keys(Utils._cache.timestamps).forEach(k => {
          if (k.startsWith('sheet_')) delete Utils._cache.timestamps[k];
        });
        Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', 'Cache de abas limpo']]);
      } else if (key === 'configs') {
        Utils._cache.configs = {};
        Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', 'Cache de configurações limpo']]);
      } else if (key) {
        delete Utils._cache.sheets[key];
        delete Utils._cache.configs[key];
        delete Utils._cache.statusValues[key];
        delete Utils._cache.timestamps[`sheet_${key}`];
        Logger.registrarLogBatch([['INFO', 'Utils.invalidateCache', `Cache de "${key}" limpo`]]);
      }
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.invalidateCache', `Falha na invalidação: ${erro.message}`]]);
    }
  },

  // ████████████████████████████████████████████████████████████████████████
  // ▶ FUNÇÕES DE CONFIGURAÇÃO E VALIDAÇÃO
  // ████████████████████████████████████████████████████████████████████████

  /**
   * ✅ OTIMIZADO: Obter config da aba com cache
   */
  obterConfigAba: (nomeAba) => {
    if (!nomeAba) {
      Logger.registrarLogBatch([['ERRO', 'Utils.obterConfigAba', 'Nome da aba não fornecido']]);
      return {};
    }

    try {
      if (!Utils._cache.configs[nomeAba]) {
        Utils._cache.configs[nomeAba] = CONFIG?.SHEETS?.[nomeAba] || { 
          START_ROW: 4,
          COLUNAS_MONITORADAS: [],
          COLUNA_TIMESTAMP: "O",
          FORMATO: "dd/MM/yyyy",
          TRAVAR_HORA_EM_ZERO: true
        };
      }
      return Utils._cache.configs[nomeAba];
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.obterConfigAba', `Falha ao obter configuração para "${nomeAba}": ${erro.message}`]]);
      return {};
    }
  },

  /**
   * ✅ CORRIGIDO: Verificar se linha está vazia (correção crítica)
   */
  linhaVazia: (sheet, linha) => {
    try {
      if (!sheet || !linha) {
        Logger.registrarLogBatch([['ERRO', 'Utils.linhaVazia', 'Parâmetros inválidos']]);
        return false;
      }
      const nome = sheet.getName();
      const cfg = Utils.obterConfigAba(nome) || {};
      const cachedSheet = Utils.getCachedSheet(nome);
      if (!cachedSheet) return false;
      const colFimValid = (cfg.VALIDACAO && cfg.VALIDACAO.colFim) || 15;
      const maxMonit = Array.isArray(cfg.COLUNAS_MONITORADAS)
        ? cfg.COLUNAS_MONITORADAS.map(Utils.colunaParaIndice).reduce((a,b)=>Math.max(a,b), 0)
        : 0;
      const numColunas = Math.max(colFimValid, maxMonit, 15);
      const intervalo = cachedSheet.getRange(linha, 1, 1, numColunas);
      const valores = intervalo.getValues()[0];
      return valores.every(cel => !cel || String(cel).trim() === '');
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.linhaVazia', `Falha na verificação linha ${linha}: ${erro.message}`]]);
      return false;
    }
  },
  /**
   * ✅ MANTIDO: Obtém valores de colunas específicas de uma linha
   */
  getColumnValues: (sheet, linha, colunas) => {
    try {
      if (!sheet || !linha || !Array.isArray(colunas)) {
        Logger.registrarLogBatch([['ERRO', 'Utils.getColumnValues', 'Parâmetros inválidos']]);
        return [];
      }

      return colunas.map(col => {
        try {
          const valor = sheet.getRange(linha, col).getDisplayValue();
          return valor === '' ? null : valor;
        } catch (erro) {
          Logger.registrarLogBatch([['ERRO', 'Utils.getColumnValues', `Falha ao ler coluna ${col}: ${erro.message}`]]);
          return null;
        }
      });
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.getColumnValues', `Falha geral: ${erro.message}`]]);
      return [];
    }
  },

  /**
   * ✅ MANTIDO: Validação de nome de arquivo
   */
  validarNomeArquivo: (nome) => {
    try {
      const cfg = CONFIG?.SEGURANCA?.VALIDACAO_NOMES;
      if (!cfg) {
        throw new Error("Configuração de validação não encontrada");
      }
      
      if (!nome || nome.trim().length === 0) {
        throw new Error("Nome do arquivo não pode ser vazio");
      }
      
      if (nome.length > cfg.TAMANHO_MAXIMO) {
        throw new Error(`Nome excede ${cfg.TAMANHO_MAXIMO} caracteres`);
      }
      
      if (cfg.CARACTERES_INVALIDOS.test(nome)) {
        throw new Error("Nome contém caracteres inválidos: \\ / : * ? \" < > |");
      }
      
      return nome.trim();
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.validarNomeArquivo', `Falha na validação: ${erro.message}`]]);
      throw erro;
    }
  },

  /**
   * ✅ MANTIDO: Tratamento centralizado de erros do Google Drive
   */
  handleDriveError: (erro, contexto) => {
    try {
      const errosConhecidos = {
        "LIMIT_EXCEEDED": "Limite de operações excedido no Drive",
        "NO_ACCESS": "Acesso negado à pasta destino",
        "INVALID_INPUT": "Parâmetros inválidos para criação do arquivo",
        "FILE_NOT_FOUND": "Arquivo não encontrado",
        "PERMISSION_DENIED": "Permissão negada"
      };

      const mensagem = errosConhecidos[erro.message] || erro.message;
     const logEntry = [
        Utilities.formatDate(new Date(), Utils.getTimezone(), 'dd/MM/yyyy'),
        'ERRO_DRIVE',
        `${contexto || 'Operação Drive'}: ${mensagem}`
      ];

      Logger.registrarLogBatch([logEntry]);
      
      // Tentar mostrar alerta (pode falhar em execução automática)
      try {
        SpreadsheetApp.getUi().alert(`Erro no Drive: ${mensagem}`);
      } catch (e) {
        // Ignorar se não conseguir mostrar alerta (execução automática)
      }
    } catch (erro2) {
      console.error('Erro no tratamento de erro do Drive:', erro2);
    }
  },

  // ████████████████████████████████████████████████████████████████████████
  // ▶ FUNÇÕES DE TESTE E DIAGNÓSTICO
  // ████████████████████████████████████████████████████████████████████████

  /**
   * ✅ NOVO: Função de teste para validar Utils
   */
  testarSistema: () => {
    try {
      Logger.registrarLogBatch([['INFO', 'Utils.testarSistema', '🧪 Iniciando testes do Utils']]);
      
      const testes = [];
      
      // Teste 1: Conversões de coluna
      const testeA = Utils.colunaParaIndice('A');
      const testeZ = Utils.colunaParaIndice('Z');
      const testeAA = Utils.colunaParaIndice('AA');
      testes.push(`Conversões: A=${testeA}, Z=${testeZ}, AA=${testeAA}`);
      
      // Teste 2: Conversões inversas
      const letraA = Utils.colunaParaLetra(1);
      const letraZ = Utils.colunaParaLetra(26);
      const letraAA = Utils.colunaParaLetra(27);
      testes.push(`Conversões inversas: 1=${letraA}, 26=${letraZ}, 27=${letraAA}`);
      
      // Teste 3: Cache de planilha
      const ss = Utils.getCachedSpreadsheet();
      testes.push(`Planilha: ${ss ? '✅ OK' : '❌ ERRO'}`);
      
      // Teste 4: Cache de aba
      const sheet = Utils.getCachedSheet('DADOS');
      testes.push(`Aba DADOS: ${sheet ? '✅ OK' : '❌ ERRO'}`);
      
      // Teste 5: Configuração
      const config = Utils.obterConfigAba('INTERNALIZADO');
      testes.push(`Config INTERNALIZADO: ${config.START_ROW ? '✅ OK' : '❌ ERRO'}`);
      
      // Teste 6: Limpeza de cache
      Utils.limparCacheExpirado();
      testes.push(`Limpeza de cache: ✅ Executada`);
      
      Logger.registrarLogBatch([
        ['INFO', 'Utils.testarSistema', '📊 Resultados dos testes:'],
        ...testes.map(teste => ['INFO', 'Utils.testarSistema', `  ${teste}`]),
        ['INFO', 'Utils.testarSistema', '🏁 Testes concluídos']
      ]);
      
      return {
        sucesso: true,
        testes: testes.length,
        detalhes: testes
      };
      
    } catch (erro) {
      Logger.registrarLogBatch([['ERRO', 'Utils.testarSistema', `Falha nos testes: ${erro.message}`]]);
      return {
        sucesso: false,
        erro: erro.message
      };
    }
  }
};

// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO LOGGER CORRIGIDO E APRIMORADO
// ████████████████████████████████████████████████████████████████████████

const Logger = {
    _cache: {
        logSheet: null,
        logBuffer: [],
        lastFlush: 0,
        MAX_BUFFER_SIZE: 5,
        _logAtivo: true,
        _modo: 'planilha' 
    },

    criarAbaLog: () => {
        try {
            if (!Logger._cache.logSheet) {
                const ss = SpreadsheetApp.getActive();
                if (!ss) {
                    console.error("❌ Erro crítico: Não foi possível acessar a planilha ativa");
                    return null;
                }
                
                Logger._cache.logSheet = ss.getSheetByName('Log do Sistema') || ss.insertSheet('Log do Sistema');
                const cabecalho = ['Timestamp', 'Tipo', 'Origem', 'Mensagem'];
                if (Logger._cache.logSheet.getLastRow() === 0) {
                    Logger._cache.logSheet.getRange(1, 1, 1, cabecalho.length)
                        .setValues([cabecalho])
                        .setFontWeight('bold');
                }
            }
            return Logger._cache.logSheet;
        } catch (erro) {
            console.error(`❌ Erro crítico no Logger.criarAbaLog: ${erro.message}`);
            return null;
        }
    },

    registrarLogBatch: (entradas, forcePlanilha = false) => {
        try {
            if (!Logger._cache._logAtivo || !entradas?.length) return;

            const tz = Utils.getTimezone();
            const timestamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
            const modo = forcePlanilha ? 'planilha' : Logger._cache._modo;

            if (modo === 'console') {
                entradas.forEach((linha) => {
                    const tipo = linha[0] || 'INFO';
                    const origem = linha[1] || '';
                    const mensagem = linha[2] || JSON.stringify(linha);
                    console.log(`[${timestamp}] [${tipo}] [${origem}] ${mensagem}`);
                });
                return;
            }

            // modo planilha: bufferiza e grava em lote
            entradas.forEach((linha) => {
                const tipo = linha[0] || 'INFO';
                const origem = linha[1] || '';
                const mensagem = linha[2] || JSON.stringify(linha);
                Logger._cache.logBuffer.push([timestamp, tipo, origem, mensagem]);
            });

            if (Logger._cache.logBuffer.length >= Logger._cache.MAX_BUFFER_SIZE) {
                Logger._flushLogBuffer();
            }
        } catch (erro) {
            console.error(`❌ [Logger] Erro ao registrar logs: ${erro.message}`);
        }
    },

    _flushLogBuffer: () => {
        try {
            if (!Logger._cache.logBuffer.length) return;
            const sheet = Logger.criarAbaLog();
            
            // ✅ VERIFICAÇÃO ROBUSTA: Verifica se a sheet existe antes de tentar gravar
            if (!sheet) {
                console.error("❌ Não foi possível criar/acessar a aba de log. Buffer será limpo para evitar overflow.");
                Logger._cache.logBuffer = []; // Limpa buffer para evitar crescimento infinito
                return;
            }
            
            const lastRow = sheet.getLastRow();
            const range = sheet.getRange(lastRow + 1, 1, Logger._cache.logBuffer.length, 4);
            range.setValues(Logger._cache.logBuffer);
            
            Logger._cache.logBuffer = [];
            Logger._cache.lastFlush = Date.now();
            
        } catch (e) {
            console.error(`❌ Erro ao gravar logs na planilha: ${e.message}`);
            // Em caso de erro, limpa buffer para evitar crescimento infinito
            Logger._cache.logBuffer = [];
        }
    },

    getModo: () => {
        return Logger._cache._modo || 'planilha';
    },

    /**
     * ✅ FUNÇÃO CORRIGIDA: alternarModo com tratamento de erro robusto
     */
    alternarModo: () => {
        try {
            // 1. Flush no buffer da planilha antes de mudar o modo
            if (Logger._cache._modo === 'planilha') {
                Logger._flushLogBuffer();
            }

            // 2. Alterna o estado
            Logger._cache._modo = Logger._cache._modo === 'planilha' ? 'console' : 'planilha';

            const modoAtual = Logger._cache._modo;
            const statusMsg = modoAtual === 'planilha'
                ? '📄 Planilha (Log do Sistema)'
                : '🖥️ Console (Execução)';

            // 3. ✅ CORREÇÃO: Notificação com tratamento de erro
            try {
                const ui = SpreadsheetApp.getUi();
                if (ui) {
                    ui.alert(`✅ Modo de log alterado para: ${statusMsg}`);
                } else {
                    console.log(`✅ Modo de log alterado para: ${statusMsg}`);
                }
            } catch (uiError) {
                console.log(`✅ Modo de log alterado para: ${statusMsg} (UI não disponível)`);
            }

            // 4. ✅ CORREÇÃO: Recria o menu com verificação de existência
            try {
                if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
                    Triggers.recriarMenu();
                } else {
                    console.log("⚠️ Triggers.recriarMenu não disponível. Menu não foi atualizado.");
                }
            } catch (menuError) {
                console.error(`⚠️ Erro ao recriar menu: ${menuError.message}`);
            }

            // 5. ✅ Log da operação
            Logger.registrarLogBatch([["INFO", "Logger.alternarModo", `Modo alterado para: ${modoAtual}`]]);

        } catch (error) {
            console.error(`❌ Erro crítico ao alternar modo do Logger: ${error.message}`);
            console.error("Stack trace:", error.stack);
            
            // ✅ CORREÇÃO: Fallback para notificação de erro
            try {
                const ui = SpreadsheetApp.getUi();
                if (ui) {
                    ui.alert(`❌ Erro ao alternar modo do log.\n\nDetalhes: ${error.message}\n\nVerifique o console para mais informações.`);
                }
            } catch (fallbackError) {
                console.error("❌ Não foi possível nem exibir alerta de erro:", fallbackError.message);
            }
        }
    },

    getStatusLog: () => Logger._cache._logAtivo,
    
    ativarLog: () => {
        try {
            Logger._cache._logAtivo = true;
            Logger.registrarLogBatch([["INFO", "Logger", "🟢 Log ativado"]]);
            
            try {
                SpreadsheetApp.getUi().alert("🟢 Registro de LOG ATIVADO.");
            } catch (uiError) {
                console.log("🟢 Registro de LOG ATIVADO. (UI não disponível)");
            }
            
            if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
                Triggers.recriarMenu();
            }
        } catch (error) {
            console.error(`❌ Erro ao ativar log: ${error.message}`);
        }
    },

    desativarLog: () => {
        try {
            // Garante que o buffer pendente seja gravado antes de desligar
            Logger._flushLogBuffer();
            Logger._cache._logAtivo = false;
            
            try {
                SpreadsheetApp.getUi().alert("🔴 Registro de LOG DESATIVADO.");
            } catch (uiError) {
                console.log("🔴 Registro de LOG DESATIVADO. (UI não disponível)");
            }
            
            if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
                Triggers.recriarMenu();
            }
        } catch (error) {
            console.error(`❌ Erro ao desativar log: ${error.message}`);
        }
    },

    alternarLogAtivo: () => {
        try {
            Logger._cache._logAtivo = !Logger._cache._logAtivo;
            const status = Logger._cache._logAtivo ? '🟢 ATIVADO' : '🔴 DESATIVADO';
            
            Logger.registrarLogBatch([["INFO", "Logger", `Status alterado para: ${status}`]]);
            
            try {
                SpreadsheetApp.getUi().alert(`📝 Registro de LOG está agora: ${status}`);
            } catch (uiError) {
                console.log(`📝 Registro de LOG está agora: ${status} (UI não disponível)`);
            }
            
            if (typeof Triggers !== 'undefined' && typeof Triggers.recriarMenu === 'function') {
                Triggers.recriarMenu();
            }
        } catch (error) {
            console.error(`❌ Erro ao alternar status do log: ${error.message}`);
        }
    },

    /**
     * ✅ NOVA FUNÇÃO: Força flush manual do buffer
     */
    forceFlush: () => {
        try {
            if (Logger._cache._modo === 'planilha') {
                Logger._flushLogBuffer();
                console.log("✅ Buffer de logs gravado manualmente");
            } else {
                console.log("ℹ️ Modo console ativo - não há buffer para gravar");
            }
        } catch (error) {
            console.error(`❌ Erro ao forçar flush: ${error.message}`);
        }
    }
};

// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO 7: GERENCIAMENTO DE TIMESTAMP
//    
//    
// ████████████████████████████████████████████████████████████████████████
const TimestampManager = {
  registrarTimestamp: (sheet, range) => {
    try {
      const nomeAba = sheet.getName();
      const configAba = CONFIG.SHEETS[nomeAba];
      if (!configAba) {
        Logger.registrarLogBatch([['AVISO', `Aba ${nomeAba} não possui configuração de timestamp`]]);
        return;
      }

      const linha = range.getRow();
      const coluna = range.getColumn();
      const colunaLetra = Utils.colunaParaLetra(coluna);

      if (!configAba.COLUNAS_MONITORADAS.includes(colunaLetra)) {
        return; // Não é uma coluna monitorada
      }

      const timestampCol = Utils.columnLetterToNumber(configAba.COLUNA_TIMESTAMP);
      const data = new Date();
      
      // Respeita a configuração TRAVAR_HORA_EM_ZERO
      if (configAba.TRAVAR_HORA_EM_ZERO) {
        data.setHours(0, 0, 0, 0);
      }

      // Formatação da data 
     sheet.getRange(linha, timestampCol)
      .setValue(Utilities.formatDate(data, Utils.getTimezone(), configAba.FORMATO));


    } catch (erro) {
      Logger.registrarLogBatch([
        ['ERRO', 'TimestampManager', `Falha ao registrar timestamp: ${erro.message}`]
      ]);
    }
  }
};

// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO DE VALIDAÇÃO DE DADOS – VERSÃO FINAL COM CONTROLE DE RECURSIVIDADE
//    Integrado à estrutura CONFIG.SHEETS.<ABA>.VALIDACAO
//    Inclui suporte a validações múltiplas, formatação e proteção contra loops de edição
// ████████████████████████████████████████████████████████████████████████

const DataValidator = {
  // ── Constantes/estado interno ────────────────────────────────
  COR_ERRO: '#FFE6E6',                 
  PREFIXO_NOTA_ERRO: /^erro:/i,        
  DEFAULT_VALIDACAO_CACHE_TTL: 10,     
  _toastCount: 0,                      
  _summaryToastShown: false,           

  // ▶ Função principal chamada pelo onEdit
  validarEdicao: function(e) {
    try {
      if (!e || !e.range) return true;
      const range = e.range;

      // === Validação de Colagem em Massa (Célula a Célula) ===
      if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
        const nR = range.getNumRows(), nC = range.getNumColumns();
        for (let i = 0; i < nR; i++) {
          for (let j = 0; j < nC; j++) {
            const sub = range.offset(i, j, 1, 1);
            // Recursão controlada para células individuais
            const resultado = DataValidator.validarEdicao({ range: sub, oldValue: null });
            if (resultado === false) return false; // Falha em bloco se uma falhar
          }
        }
        return true;
      }

      const valorDepois = range.getValue();
      const aba = range.getSheet().getName();
      const linha = range.getRow();
      const coluna = range.getColumn();

      const configAba = CONFIG?.SHEETS?.[aba];
      if (!configAba?.VALIDACAO?.ativo || linha < configAba.START_ROW) return true;
      const configValidacao = configAba.VALIDACAO;
      if (coluna < configValidacao.colInicio || coluna > configValidacao.colFim) return true;

      const regras = configValidacao.regrasPorColuna?.[coluna];
      if (!regras || regras.length === 0) return true;

      // ▶ Controle de recursividade (CacheService)
      const chaveCache = `validacao_${range.getSheet().getSheetId()}_${linha}_${coluna}`;
      const cache = CacheService.getScriptCache();
      if (cache.get(chaveCache)) {
        cache.remove(chaveCache); // Libera para próxima edição real
        return true; 
      }

      for (const regra of regras) {
        switch (regra) {

          // ▶ CPF / CNPJ
          case 'cpfOuCnpjValido': {
            if (!valorDepois) {
              this.limparFormatacaoErro(range, true);
              continue; 
            }
            const doc = String(valorDepois).replace(/[^\d]/g, '');
            const isCpf = doc.length === 11 && this.validarCPF(doc);
            const isCnpj = doc.length === 14 && this.validarCNPJ(doc);

            if (isCpf || isCnpj) {
              const valorFormatado = isCpf ? this.formatarCPF(doc) : this.formatarCNPJ(doc);
              this._aplicarValorFormatado(range, valorFormatado, chaveCache);
              continue; 
            }
            this.dispararErro(range, "CPF/CNPJ inválido.");
            return false;
          }

          // ▶ DATA (Com correção para 7 dígitos)
          case 'formatarData': {
            if (!valorDepois) {
              this.limparFormatacaoErro(range, true);
              return true;
            }
            const dataValida = this.parsearData(valorDepois);
            if (dataValida) {
              // 1. Converte para String formatada
              const textoFormatado = Utilities.formatDate(dataValida, Session.getScriptTimeZone(), 'dd/MM/yyyy');
              
              const TTL = (CONFIG && CONFIG.VALIDACAO_CACHE_TTL) ? CONFIG.VALIDACAO_CACHE_TTL : DataValidator.DEFAULT_VALIDACAO_CACHE_TTL;
              cache.put(chaveCache, '1', TTL);
              
              // 2. FORÇA FORMATO DE TEXTO (@) ANTES DE ESCREVER
              // Isso impede que o Sheets converta para Número/Data e cause o bug de 1969
              range.setNumberFormat('@'); 
              range.setValue(textoFormatado);
              
              this.limparFormatacaoErro(range);
              return true;
            }
            this.dispararErro(range, "Data inválida.");
            return false;
          }

          // ▶ DOCSFLOW
          case 'FormatoDocsflow': {
            if (!valorDepois) {
              this.limparFormatacaoErro(range, true);
              continue;
            }
            const doc = String(valorDepois).replace(/[^\d]/g, '');
            if (doc.length === 14) {
              const formatado = this.formatarDocsflow(doc);
              this._aplicarValorFormatado(range, formatado, chaveCache);
              continue;
            }
            this.dispararErro(range, "Docsflow inválido (requer 14 dígitos).");
            return false;
          }

          // ▶ MULTIFORMATO (Com correção de fluxo de erro)
          case 'MultiFormatoProcesso': {
            if (!valorDepois) {
              this.limparFormatacaoErro(range, true);
              continue;
            }
            const doc = String(valorDepois).replace(/[^\d]/g, '');
            let valorFinal = null;

            if (doc.length === 10) valorFinal = this.formatarControper(doc);
            else if (doc.length === 7) valorFinal = this.formatarAmzcred7(doc);
            else if (doc.length === 4) valorFinal = doc;

            if (valorFinal !== null) {
              this._aplicarValorFormatado(range, valorFinal, chaveCache);
              continue; // Sucesso: segue para próxima regra
            }

            // Falha: bloqueia execução
            this.dispararErro(range, "Formato de processo inválido (esperado 4, 7 ou 10 dígitos).");
            return false; 
          }
        } 
      } 
      return true;

    } catch (erro) {
      Logger.registrarLogBatch([['ERRO_CRITICO', 'DataValidator', erro.message]], true);
      return true; // Em erro crítico de script, permite continuar para não travar o usuário
    }
  },

  // ── Helpers Privados ──────────────────────────────────────

  _aplicarValorFormatado: function(range, valor, chaveCache, formatoNum = null) {
    const TTL = (CONFIG && CONFIG.VALIDACAO_CACHE_TTL) || this.DEFAULT_VALIDACAO_CACHE_TTL;
    CacheService.getScriptCache().put(chaveCache, '1', TTL);
    SpreadsheetApp.flush();
    range.setValue(valor);
    if (formatoNum) { try { range.setNumberFormat(formatoNum); } catch (_) {} }
    this.limparFormatacaoErro(range);
  },

  dispararErro: function(range, tipoErro) {
    // 1. Log
    Logger.registrarLogBatch([['ERRO', 'DataValidator', `${range.getA1Notation()}: ${tipoErro}`]], true);
    // 2. Visual
    range.setBackground(this.COR_ERRO);
    range.setNote(`ERRO: ${tipoErro}`);
    // 3. Toast (com limitador)
    this._exibirToastLimitado(range, tipoErro);
  },

  _exibirToastLimitado: function(range, msg) {
    const cfg = CONFIG?.UI?.VALIDACAO?.TOAST || {};
    if (!cfg.ativar) return;
    this._toastCount = this._toastCount || 0;
    const limite = cfg.limitePorExecucao || 3;

    if (this._toastCount >= limite) {
      if (!this._summaryToastShown) {
        try { SpreadsheetApp.getActive().toast('Vários erros detectados. Verifique as células vermelhas.', 'Validação'); } catch(_){}
        this._summaryToastShown = true;
      }
      return;
    }
    // Cooldown por célula
    const key = `toast_${range.getA1Notation()}`;
    if (CacheService.getScriptCache().get(key)) return;
    
    try { SpreadsheetApp.getActive().toast(msg, 'Erro Validação'); } catch(_){}
    CacheService.getScriptCache().put(key, '1', 2); // 2 seg cooldown
    this._toastCount++;
  },

  limparFormatacaoErro: function(range, limparNota = false) {
    try {
      if ((range.getBackground() || '').toLowerCase() === this.COR_ERRO.toLowerCase()) {
        range.setBackground(null);
      }
      const nota = (range.getNote() || '').trim();
      if (limparNota || this.PREFIXO_NOTA_ERRO.test(nota)) {
        range.clearNote();
      }
    } catch (_) {}
  },

  // ── Formatters & Parsers ──────────────────────────────────

  parsearData: function(textoData) {
    if (textoData instanceof Date && !isNaN(textoData.getTime())) return textoData;
    if (typeof textoData === 'number') {
      const d = new Date(textoData);
      if (!isNaN(d.getTime())) return d;
    }

    const digitos = String(textoData).replace(/[^\d]/g, '');
    let dia, mes, ano;

    if (digitos.length === 8) { // 01012025
      dia = digitos.substring(0, 2); mes = digitos.substring(2, 4); ano = digitos.substring(4, 8);
    } else if (digitos.length === 7) { // 1012025 -> 01/01/2025
      dia = '0' + digitos.substring(0, 1); mes = digitos.substring(1, 3); ano = digitos.substring(3, 7);
    } else if (digitos.length === 6) { // 010125
      dia = digitos.substring(0, 2); mes = digitos.substring(2, 4); ano = '20' + digitos.substring(4, 6);
    } else {
      return null;
    }

    const dataObj = new Date(ano, mes - 1, dia);
    if (dataObj.getFullYear() == ano && dataObj.getMonth() == (mes - 1) && dataObj.getDate() == parseInt(dia, 10)) {
      return dataObj;
    }
    return null;
  },

  formatarCPF: (v) => v.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4'),
  formatarCNPJ: (v) => v.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5'),
  formatarDocsflow: (v) => v.replace(/(\d{4})(\d{10})/, '$1/$2'),
  formatarControper: (v) => v.replace(/(\d{3})(\d{2})(\d{4})(\d{1})/, '$1-$2-$3-$4'),
  formatarAmzcred7: (v) => v.replace(/(\d{6})(\d{1})/, '$1/$2'),

  validarCPF: function(cpf) {
    if (typeof cpf !== 'string' || !cpf || cpf.length !== 11 || /^(\d)\1+$/.test(cpf)) return false;
    let soma = 0, resto;
    for (let i = 1; i <= 9; i++) soma += parseInt(cpf.substring(i-1, i)) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto >= 10) resto = 0;
    if (resto !== parseInt(cpf.substring(9, 10))) return false;
    soma = 0;
    for (let i = 1; i <= 10; i++) soma += parseInt(cpf.substring(i-1, i)) * (12 - i);
    resto = (soma * 10) % 11;
    if (resto >= 10) resto = 0;
    return resto === parseInt(cpf.substring(10, 11));
  },

  validarCNPJ: function(cnpj) {
    if (typeof cnpj !== 'string' || !cnpj || cnpj.length !== 14 || /^(\d)\1+$/.test(cnpj)) return false;
    let tamanho = cnpj.length - 2;
    let numeros = cnpj.substring(0, tamanho);
    let digitos = cnpj.substring(tamanho);
    let soma = 0;
    let pos = tamanho - 7;
    for (let i = tamanho; i >= 1; i--) {
      soma += parseInt(numeros.charAt(tamanho - i)) * pos--;
      if (pos < 2) pos = 9;
    }
    let resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
    if (resultado != parseInt(digitos.charAt(0))) return false;
    tamanho += 1;
    numeros = cnpj.substring(0, tamanho);
    soma = 0;
    pos = tamanho - 7;
    for (let i = tamanho; i >= 1; i--) {
      soma += parseInt(numeros.charAt(tamanho - i)) * pos--;
      if (pos < 2) pos = 9;
    }
    resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
    return resultado == parseInt(digitos.charAt(1));
  }
};

// 🔗 Este módulo deve ser invocado a partir do gatilho onEdit no bloco superior
//     Exemplo:
//     if (typeof DataValidator !== 'undefined' && DataValidator.validarEdicao) {
//        DataValidator.validarEdicao(e);
//     }


// ████████████████████████████████████████████████████████████████████████
// ▶ VISOES — Wrapper mínimo (submenu + aplicação UI-only)
// ████████████████████████████████████████████████████████████████████████

/** Normaliza nome de aba para chave canônica (SEM acento, upper) */
function VISOES__normalizeAbaName(nome) {
  try {
    return nome
      .toString()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/ç/gi, 'c')
      .toUpperCase()
      .trim();
  } catch (e) {
    return (nome || '').toString().toUpperCase().trim();
  }
}

/** Converte tokens de coluna (ex.: 'A', 'C:D', 3, '3:8') para índices (1-based) */
function VISOES__expandTokensToIndices(tokens, maxCols) {
  const out = new Set();
  if (!Array.isArray(tokens)) return out;
  tokens.forEach(raw => {
    if (raw == null) return;
    const t = String(raw).trim().toUpperCase();
    const addIndex = (i) => { if (i >= 1 && i <= maxCols) out.add(i); };
    if (/^\d+$/.test(t)) { addIndex(Number(t)); return; }
    if (/^\d+:\d+$/.test(t)) {
      let [a,b] = t.split(':').map(n => Number(n));
      if (a>b) [a,b] = [b,a];
      for (let i=a;i<=b;i++) addIndex(i);
      return;
    }
    if (/^[A-Z]+$/.test(t)) { addIndex(Utils.colunaParaIndice(t)); return; }
    if (/^[A-Z]+:[A-Z]+$/.test(t)) {
      let [a,b] = t.split(':');
      let ai = Utils.colunaParaIndice(a);
      let bi = Utils.colunaParaIndice(b);
      if (ai>bi) [ai,bi] = [bi,ai];
      for (let i=ai;i<=bi;i++) addIndex(i);
      return;
    }
  });
  return out;
}

/** Aplica visão na aba ativa (UI-only). visao='OPERACIONAL'|'COMPLETO' */
function VISOES__aplicarVisaoMenu(visao) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    const abaReal = sheet.getName();
    const abaKey = VISOES__normalizeAbaName(abaReal);
    // 1ª fonte: bloco da própria aba
    let cfgAba = (CONFIG && CONFIG.SHEETS && CONFIG.SHEETS[abaKey] && CONFIG.SHEETS[abaKey].VISOES)
      ? CONFIG.SHEETS[abaKey].VISOES
      : null;
    // 2ª fonte: bloco central (se existir)
    if (!cfgAba) {
      const cent = (CONFIG && CONFIG.VISOES && CONFIG.VISOES.ABAS) ? CONFIG.VISOES.ABAS : null;
      if (cent && cent[abaKey]) cfgAba = cent[abaKey];
    }
    if (!cfgAba) {
      if (CONFIG?.VISOES?.DEBUG && typeof Logger?.registrarLogBatch === 'function') {
        Logger.registrarLogBatch([["DEBUG","VISOES",`Sem configuração de visões para: ${abaReal} (${abaKey})`]]);
      }
      return; // silencioso
    }
    const maxCols = sheet.getMaxColumns();

    // Desejado: matriz booleana (true=visível)
    let desejadoVisivel = new Array(maxCols).fill(true);
    if (visao === 'OPERACIONAL') {
      const set = VISOES__expandTokensToIndices(cfgAba.OPERACIONAL || [], maxCols);
      desejadoVisivel = desejadoVisivel.map((_,i)=> set.has(i+1)); // visível apenas o que consta
    } else {
      // COMPLETO: mantém tudo visível (ou se cfgAba.COMPLETO==="TODAS")
      desejadoVisivel = desejadoVisivel.map(()=> true);
    }

    // Estado atual
    const atualOculto = new Array(maxCols);
    for (let c=1;c<=maxCols;c++) atualOculto[c-1] = sheet.isColumnHiddenByUser(c);

    // Calcular blocos de toggle (contiguidade por alvo)
    const hideBlocks = [];
    const showBlocks = [];
    let i = 1;
    while (i <= maxCols) {
      const desiredHidden = !desejadoVisivel[i-1];
      const isHidden = atualOculto[i-1];
      if (desiredHidden !== isHidden) {
        const target = desiredHidden; // queremos chegar neste estado
        let j = i;
        while (j <= maxCols) {
          const jDesiredHidden = !desejadoVisivel[j-1];
          const jIsHidden = atualOculto[j-1];
          if (jDesiredHidden !== jIsHidden && jDesiredHidden === target) {
            j++;
          } else {
            break;
          }
        }
        const count = j - i;
        (target ? hideBlocks : showBlocks).push({ start: i, count });
        i = j;
      } else {
        i++;
      }
    }

    // Aplicar
    hideBlocks.forEach(b => sheet.hideColumns(b.start, b.count));
    showBlocks.forEach(b => sheet.showColumns(b.start, b.count));

    if (CONFIG?.VISOES?.DEBUG && typeof Logger?.registrarLogBatch === 'function') {
      const fmt = (arr)=>arr.map(b=>`${b.start}-${b.start+b.count-1}`).join(',');
      Logger.registrarLogBatch([["DEBUG","VISOES",`Aba=${abaReal} (${abaKey}) | Visao=${visao} | hide=[${fmt(hideBlocks)}] | show=[${fmt(showBlocks)}] | ops=${hideBlocks.length+showBlocks.length}`]]);
    }
  } catch (e) {
    try {
      Logger.registrarLogBatch([["ERRO","VISOES",`Falha ao aplicar visão: ${e.message}`]]);
    } catch (_e) {
      console.error(e);
    }
  }
}

// Wrappers públicos para o menu (sem parâmetros no Apps Script)
function VISOES__aplicarVisaoMenu_Operacional() { VISOES__aplicarVisaoMenu('OPERACIONAL'); }
function VISOES__aplicarVisaoMenu_Completo() { VISOES__aplicarVisaoMenu('COMPLETO'); }

// ████████████████████████████████████████████████████████████████████████
// ▶ MÓDULO: TRIGGERS CORRIGIDO - COM SUPORTE A UI E ONEDITMANUAL
// ████████████████████████████████████████████████████████████████████████

const Triggers = {
  _cache: {
    linhasProcessadas: new Set(),
    menuConfigurado: false,
    lastClean: 0,
    ultimaExecucao: 0
  },

  // ████████████████████████████████████████████████████████████████████████
  // ▶ ONEDIT: ORQUESTRADOR CENTRAL (Chamado pelo onEditManual)
  // ████████████████████████████████████████████████████████████████████████
  onEdit: (e) => {
    try {
      // 1. Flush preventivo (Correção de Leitura Vencida)
      try { SpreadsheetApp.flush(); } catch (f) {}

      // 2. Reset de contadores de validação
      if (typeof DataValidator !== 'undefined') {
        DataValidator._toastCount = 0;
        DataValidator._summaryToastShown = false;
      }

      if (typeof Logger !== 'undefined' && Logger.registrarLogBatch) {
        Logger.registrarLogBatch([["INFO", "Triggers.onEdit", "🟢 Iniciou Triggers.onEdit"]]);
      }
      
      if (!e || !e.range) return;

      // 3. Validação de Dados (DataValidator) - Prioridade Alta
      if (typeof DataValidator !== 'undefined' && DataValidator.validarEdicao) {
        const resultadoVal = DataValidator.validarEdicao(e);
        if (resultadoVal === false) {
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "⚠️ DataValidator retornou false"]]);
          return; // Interrompe se inválido
        }
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "✅ DataValidator executado"]]);
      }

      // 4. DropdownManager (Atualizado para Cores UI)
      if (typeof DropdownManager !== 'undefined' && DropdownManager.processarEdicao) {
        try {
          DropdownManager.processarEdicao(e);
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "✅ DropdownManager executado"]]);
        } catch (erroDropdown) {
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.onEdit", `❌ Erro no DropdownManager: ${erroDropdown.message}`]]);
        }
      }

      // 5. Handler Legado (Movimentação, Docs, etc.)
      if (typeof onEditHandler !== 'undefined' && onEditHandler.gerenciarEdicoes) {
        try {
          onEditHandler.gerenciarEdicoes(e);
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "✅ onEditHandler legado executado"]]);
        } catch (erroLegado) {
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.onEdit", `❌ Erro no onEditHandler: ${erroLegado.message}`]]);
        }
      }

      // 6. Timestamp (Se configurado na aba)
      const sheet = e.range.getSheet();
      const coluna = e.range.getColumn();
      const configAba = CONFIG?.SHEETS?.[sheet.getName()];
      if (configAba?.COLUNAS_MONITORADAS?.includes(Utils.colunaParaLetra(coluna))) {
        try {
          TimestampManager.registrarTimestamp(sheet, e.range);
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "✅ Timestamp registrado"]]);
        } catch (erroTimestamp) {
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.onEdit", `❌ Erro no TimestampManager: ${erroTimestamp.message}`]]);
        }
      }

      // 7. Batch Operations (Execução final em lote)
      if (CONFIG.BATCH_OPS?.ENABLED && typeof BatchOperations?.execute === 'function') {
        try {
          BatchOperations.execute('onEdit');
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "✅ BatchOperations executado"]]);
        } catch (erroBatch) {
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.onEdit", `❌ Erro no BatchOperations: ${erroBatch.message}`]]);
        }
      } else {
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["DEBUG", "Triggers.onEdit", "ℹ️ BatchOperations inativo"]]);
      }

      if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["INFO", "Triggers.onEdit", "🏁 Cadeia onEdit concluída"]]);

    } catch (erro) {
      if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO_CRITICO", "Triggers.onEdit", `❌ Erro crítico: ${erro.message}`]]);
    }
  },

  // ████████████████████████████████████████████████████████████████████████
  // ▶ ONOPEN: CONSTRUÇÃO DO MENU (COM NOVO ITEM UI)
  // ████████████████████████████████████████████████████████████████████████

    onOpen: () => {
      try {
        if (Triggers._cache.menuConfigurado) return;
        
        const ui = SpreadsheetApp.getUi();
        const menu = ui.createMenu('⚙️ Sistema GAS');
        
        // Menu: Logs
        try { menu.addItem('🔁 Alternar Modo de Log (Logger)', 'Logger.alternarModo'); } catch (e) {}
        menu.addSeparator();
        
        // Menu: Dropdown & Cores
        try { menu.addItem('🔄 Recriar Listas Suspensas', 'executarDropdownCompleto'); } catch (e) {}
        try { menu.addItem('🧹 Resetar Cache Dropdown', 'resetDropdownCompleto'); } catch (e) {}
        
        // ▶▶▶ CONFIGURAÇÃO VISUAL DE CORES ◀◀◀
        try { 
          menu.addItem('🎨 Configurar Cores (Visual)', 'abrirConfiguradorCores'); 
        } catch (e) { 
          console.error("Falha ao adicionar menu de cores:", e); 
        }
        
        menu.addSeparator();
        
        // Menu: Ferramentas
        try { menu.addItem('🧪 Teste Rápido Dropdown', 'testeRapidoDropdown'); } catch (e) {}
        try { menu.addItem('🔍 Diagnóstico Dropdown', 'diagnosticoDropdownCompleto'); } catch (e) {}
        
        menu.addSeparator();
          
        // Menu: Visões
        try {
          const hasVisoes = Object.keys(CONFIG?.SHEETS || {}).some(k => CONFIG.SHEETS[k]?.VISOES);
          if (hasVisoes) {
            ui.createMenu('👁️ Visões')
              .addItem('Operacional', 'VISOES__aplicarVisaoMenu_Operacional')
              .addItem('Completo', 'VISOES__aplicarVisaoMenu_Completo')
              .addToUi();
          }
        } catch (e) {}
        
        // Menu: Batch Ops
        if (CONFIG.BATCH_OPS?.ENABLED) {
          try {
            BatchOperations.init(CONFIG);
            menu.addItem('⚡ Executar Operações em Lote', 'BatchOperations.execute');
          } catch (e) {}
        }
        
        menu.addToUi();
        Triggers._cache.menuConfigurado = true;
      
      } catch (erro) {
        console.error("Erro crítico no onOpen:", erro);
      }
  },
    // ████████████████████████████████████████████████████████████████████████
    // ▶ OUTRAS FUNÇÕES AUXILIARES
    // ████████████████████████████████████████████████████████████████████████

    onChangeRemoveRow: (e) => {
      try {
        if (e.changeType !== 'REMOVE_ROW') return;
        const docsOrfaos = [];
        const abaLog = Utils.getCachedSheet('Log do Sistema');
        if (!abaLog) return;
        
        const dadosLog = abaLog.getRange(2, 1, abaLog.getLastRow() - 1, 3).getValues();

        dadosLog.forEach((linha) => {
          const mensagem = linha[2];
          if (mensagem && mensagem.includes('Documento criado')) {
            const match = mensagem.match(/d\/(.+?)\//);
            if (match) docsOrfaos.push(match[1]);
          }
        });

        if (docsOrfaos.length > 0) {
          docsOrfaos.forEach(docId => {
            try {
              const file = DriveApp.getFileById(docId);
              file.setTrashed(true);
            } catch (erro2) {
              if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["AVISO", "Triggers.onChangeRemoveRow", `Documento não encontrado: ${docId}`]]);
            }
          });
          if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["INFO", "Triggers.onChangeRemoveRow", `${docsOrfaos.length} documentos órfãos removidos`]]);
        }
      } catch (erro) {
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.onChangeRemoveRow", erro.message]]);
      }
    },

    atualizacaoDiaria: () => {
      try {
        if (!CONFIG.ACOMPANHAMENTO_PARCELAS) return;
        const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.ACOMPANHAMENTO_PARCELAS.ABA);
        if (!sheet) return;
        const ultimaLinha = sheet.getLastRow();
        for (let linha = CONFIG.ACOMPANHAMENTO_PARCELAS.LINHA_INICIO; linha <= ultimaLinha; linha++) {
          ParcelasManager.atualizarCores(sheet, linha);
        }
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["INFO", "Triggers.atualizacaoDiaria", `Atualizadas ${ultimaLinha} linhas`]]);
      } catch (erro) {
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.atualizacaoDiaria", erro.message]]);
      }
    },

    recriarMenu: () => {
      try {
        Triggers._cache.menuConfigurado = false;
        Triggers.onOpen();
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["INFO", "Triggers.recriarMenu", "Menu recriado"]]);
      } catch (erro) {
        if (typeof Logger !== 'undefined') Logger.registrarLogBatch([["ERRO", "Triggers.recriarMenu", erro.message]]);
      }
    }
  };

// ████████████████████████████████████████████████████████████████████████
// ▶ FUNÇÕES GLOBAIS (PONTOS DE ENTRADA DO GOOGLE)
// ████████████████████████████████████████████████████████████████████████

// 1. Gatilho Simples ao Abrir
function onOpen(e) {
  try {
    if (Triggers && Triggers.onOpen) Triggers.onOpen();
  } catch (err) {
    console.error("Falha no onOpen global:", err);
  }
}

// 2. Gatilho Instalável Manual (ESSENCIAL PARA SEU SETUP)
function onEditManual(e) { 
  if (Triggers && Triggers.onEdit) {
    return Triggers.onEdit(e); 
  }
}

// 3. Gatilho de Mudança (Exclusão de linhas)
function onChange(e) { 
  if (Triggers && Triggers.onChangeRemoveRow) {
    return Triggers.onChangeRemoveRow(e);
  }
}

// 4. Legado (Desativado)
function _onEdit_DESATIVADO(e) { 
  return Triggers.onEdit(e);
}

function abrirConfiguradorCores() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SidebarCores')
        .setTitle('🎨 Configurador de Cores')
        .setWidth(600);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (erro) {
    SpreadsheetApp.getUi().alert('Erro ao abrir configurador: ' + erro.message);
    console.error('Erro em abrirConfiguradorCores:', erro);
  }
}