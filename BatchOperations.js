const BatchOperations = {
  queue: [],
  CONFIG: {
    GLOBAL_CONFIG: null,
    BATCH_METHODS: {
      'setValue':       { batch: 'setValues' },
      'setValues':      { batch: 'setValues' },
      'setBackground':  { batch: 'setBackgrounds' },
      'setFontColor':   { batch: 'setFontColors' },
      'setFormula':     { batch: 'setFormulas' },
      'clearContent':   { batch: 'clear' },
      // NOVO: Suporte a validações de dados
      'setDataValidation': { batch: 'setDataValidations' } 
    },
    SINGLE_EXECUTION_OPS: ['insertRowAfter', 'deleteRow', 'setValues']
  },

  init: function(globalConfig) {
    this.CONFIG.GLOBAL_CONFIG = globalConfig;
  },

  add: function(tipo, destino, valor) {
    if (!tipo || !destino) return;
    this.queue.push({ tipo, destino, valor });
  },

  execute: function(origem = 'desconhecida') {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const agrupadas = {};
    const tempoInicio = Date.now();
    let totalOperacoes = 0;
    let detalhes = '';
    let erros = '';

    try {
      // Processa todas as operações na fila
      this.queue.forEach(op => {
        const match = op.destino.match(/^'?(.*?)'?!([A-Z]+)([0-9]+)$/);
        if (!match) return;

        const aba = match[1], col = match[2], row = parseInt(match[3]);

        if (!agrupadas[aba]) agrupadas[aba] = {};
        if (!agrupadas[aba][op.tipo]) agrupadas[aba][op.tipo] = [];

        agrupadas[aba][op.tipo].push({ 
          row, 
          col, 
          valor: op.valor 
        });
      });

      // Executa operações agrupadas por aba e tipo
      for (let abaNome in agrupadas) {
        const aba = planilha.getSheetByName(abaNome);
        if (!aba) continue;

        const operacoesPorTipo = agrupadas[abaNome];
        for (let tipo in operacoesPorTipo) {
          const operacoes = operacoesPorTipo[tipo];
          totalOperacoes += operacoes.length;
          detalhes += ` [${abaNome}:${tipo}:${operacoes.length}] `;

          // Execução direta para tipos não agrupáveis por coluna
          if (this.CONFIG.SINGLE_EXECUTION_OPS.includes(tipo)) {
            if (tipo === 'deleteRow') {
              // Apagar de baixo para cima
              operacoes.sort((a,b) => b.row - a.row).forEach(op => {
                aba.deleteRow(op.row);
              });
              continue;
            }
            if (tipo === 'setValues') {
              operacoes.forEach(op => {
                const startCol = this.colunaNumero(op.col);
                const values = Array.isArray(op.valor) ? op.valor : [[op.valor]];
                const numRows = values.length;
                const numCols = Array.isArray(values[0]) ? values[0].length : 1;
                aba.getRange(op.row, startCol, numRows, numCols).setValues(values);
              });
              continue;
            }

            if (tipo === 'insertRowAfter') {
              // Inserir de cima para baixo
              operacoes.sort((a,b) => a.row - b.row).forEach(op => {
                aba.insertRowAfter(op.row);
              });
              continue;
            }

          }
          
          const colIndexMap = {};
          operacoes.forEach(op => {
            const colIndex = this.colunaNumero(op.col);
            if (!colIndexMap[colIndex]) colIndexMap[colIndex] = [];
            colIndexMap[colIndex].push(op);
          });

          for (let colIndex in colIndexMap) {
            const ops = colIndexMap[colIndex];
            const startRow = Math.min(...ops.map(o => o.row));
            const endRow = Math.max(...ops.map(o => o.row));
            const numRows = endRow - startRow + 1;

            const valores = Array(numRows).fill().map(() => []);
            ops.forEach(op => {
              const idx = op.row - startRow;
              // TRATAMENTO ESPECIAL PARA VALIDAÇÕES
              valores[idx][0] = op.valor;
            });

            const range = aba.getRange(startRow, parseInt(colIndex), numRows, 1);
            
            // EXECUÇÃO DAS OPERAÇÕES EM LOTE
            switch (tipo) {
              case 'setValue':       range.setValues(valores); break;
              case 'setBackground':  range.setBackgrounds(valores); break;
              case 'setFontColor':   range.setFontColors(valores); break;
              case 'setFormula':     range.setFormulas(valores); break;
              case 'clearContent':   range.clearContent(); break;
              case 'setDataValidation': 
                // Conversão especial para validações
                range.setDataValidations(valores.map(row => [row[0]])); 
                break;
            }
          }
        }
      }

    } catch (erro) {
      erros = erro.message;
      Logger.registrarLogBatch([['ERRO', 'BatchOperations', erro.message]]);
    }

    this.queue = [];
    
    // Log de desempenho
    if (CONFIG?.BATCH_OPS?.LOG_PERFORMANCE) {
      try {
        const tempoFinal = Date.now();
        const duracao = tempoFinal - tempoInicio;
        const abaDashboard = planilha.getSheetByName(CONFIG.BATCH_OPS.DASHBOARD_SHEET);
        if (abaDashboard) {
          abaDashboard.appendRow([
            new Date(),
            origem,
            duracao,
            totalOperacoes,
            detalhes.trim(),
            erros
          ]);
        }
      } catch (erroDashboard) {
        Logger.registrarLogBatch([['ERRO', 'BatchOperations', erroDashboard.message]]);
      }
    }
  },

    colunaNumero: function(letra) {
        let numero = 0;
        for (let i = 0; i < letra.length; i++) {
          numero = numero * 26 + letra.charCodeAt(i) - 64;
        }
        return numero;
      },

      

  // ... (métodos restantes permanecem iguais) ...
};