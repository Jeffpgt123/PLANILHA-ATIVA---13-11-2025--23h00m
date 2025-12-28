/**
 * MetasOperacoesManager
 * Gerencia a importação e o cálculo de resumos na aba METAS🎯.
 * * ESTRUTURA REFATORADA (V2):
 * 1. importarDados() -> Botão 1: Traz os dados das origens.
 * 2. calcularResumos() -> Botão 2: Calcula os totais na tabela existente.
 */
const MetasOperacoesManager = {
  
  /**
   * BOTÃO 1: Importar Dados
   * Limpa a tabela, busca dados das origens e preenche a lista.
   * NÃO calcula os somatórios finais (apenas limpa os antigos para evitar confusão).
   */
  importarDados: function () {
    const cfg = CONFIG?.METAS_OPERACOES_FORA_BORA;
    if (!cfg || cfg.ENABLED === false) return;

    const ss = SpreadsheetApp.getActive();
    const shDestino = ss.getSheetByName(cfg.DESTINO.ABA);
    if (!shDestino) return;

    // 0) Limpar células de resumo (para evitar dados desatualizados visíveis)
    this._limparResumoValoresNaoComputadosBora_(shDestino, cfg);

    // 1) Limpar tabela destino (mantém apenas cabeçalho)
    this._limparTabelaDestino(shDestino, cfg.DESTINO);

    let proximaLinhaDestino = cfg.DESTINO.PRIMEIRA_LINHA_DADOS;
    const numColsDestino = this._numCols(cfg.DESTINO.COLUNA_INICIO, cfg.DESTINO.COLUNA_FIM);

    // 2) Processar origem CONTRATADO/LIBERACAO✅
    const cfgContratado = cfg.ORIGENS.CONTRATADO_LIBERACAO;
    if (cfgContratado) {
      const shContratado = ss.getSheetByName(cfgContratado.ABA);
      if (shContratado) {
        proximaLinhaDestino = this._processarOrigemContratado_(
          shContratado,
          shDestino,
          cfgContratado,
          cfg.DESTINO,
          proximaLinhaDestino,
          numColsDestino
        );
      }
    }

    // 3) Processar origem SEGUROS🛡️
    const cfgSeguros = cfg.ORIGENS.SEGUROS;
    if (cfgSeguros) {
      const shSeguros = ss.getSheetByName(cfgSeguros.ABA);
      if (shSeguros) {
        proximaLinhaDestino = this._processarOrigemSeguros_(
          shSeguros,
          shDestino,
          cfgSeguros,
          cfg.DESTINO,
          proximaLinhaDestino,
          numColsDestino
        );
      }
    }

    // 3b) Processar origem ENCARTEIRAMENTO🧑‍💻
    const cfgEncarteiramento = cfg.ORIGENS.ENCARTEIRAMENTO;
    if (cfgEncarteiramento) {
      const shEncarteiramento = ss.getSheetByName(cfgEncarteiramento.ABA);
      if (shEncarteiramento) {
        proximaLinhaDestino = this._processarOrigemEncarteiramento_(
          shEncarteiramento,
          shDestino,
          cfgEncarteiramento,
          cfg.DESTINO,
          proximaLinhaDestino,
          numColsDestino
        );
      }
    }

    // 4) Aplicar formato de moeda BRL na coluna G (Valor)
    // Apenas formata visualmente a lista importada.
    const primeiraLinhaDados = cfg.DESTINO.PRIMEIRA_LINHA_DADOS;
    const ultimaLinhaPreenchida = proximaLinhaDestino - 1;

    if (ultimaLinhaPreenchida >= primeiraLinhaDados) {
      this._aplicarFormatoMoedaBRL_(
        shDestino,
        primeiraLinhaDados,
        ultimaLinhaPreenchida
      );
    }
    
    SpreadsheetApp.getActive().toast("Dados importados! Clique no botão de Soma para atualizar os totais.", "Sucesso");
  },

  /**
   * BOTÃO 2: Calcular Resumos
   * Lê a tabela existente em METAS🎯 e atualiza os campos de soma/contagem.
   * Não altera as linhas da tabela, apenas lê e soma.
   */
  calcularResumos: function () {
    const cfg = CONFIG?.METAS_OPERACOES_FORA_BORA;
    if (!cfg || cfg.ENABLED === false) return;

    const ss = SpreadsheetApp.getActive();
    const shDestino = ss.getSheetByName(cfg.DESTINO.ABA);
    if (!shDestino) return;

    // Detectar intervalo de dados atual dinamicamente
    const primeiraLinhaDados = cfg.DESTINO.PRIMEIRA_LINHA_DADOS;
    const ultimaLinhaPreenchida = shDestino.getLastRow();

    // Se não houver dados (tabela vazia), limpa os resumos e avisa
    if (ultimaLinhaPreenchida < primeiraLinhaDados) {
      this._limparResumoValoresNaoComputadosBora_(shDestino, cfg);
      SpreadsheetApp.getActive().toast("Tabela vazia. Resumos limpos.", "Aviso");
      return;
    }

    // Executa a lógica de soma existente
    this._aplicarResumosValoresNaoComputadosBora_(
      shDestino,
      cfg,
      primeiraLinhaDados,
      ultimaLinhaPreenchida
    );

    SpreadsheetApp.getActive().toast("Cálculo de resumos atualizado com sucesso!", "Concluído");
  },

  /**
   * Limpa apenas a tabela OPERACOES_FORA_BORA e os resumos.
   */
  limparApenasTabelaEResumos: function () {
    const cfg = CONFIG?.METAS_OPERACOES_FORA_BORA;
    if (!cfg || cfg.ENABLED === false) return;

    const ss = SpreadsheetApp.getActive();
    const shDestino = ss.getSheetByName(cfg.DESTINO.ABA);
    if (!shDestino) return;

    this._limparResumoValoresNaoComputadosBora_(shDestino, cfg);
    this._limparTabelaDestino(shDestino, cfg.DESTINO);
    SpreadsheetApp.getActive().toast("Tabela e resumos limpos.", "Limpeza");
  },

  // ========================================================================
  // MÉTODOS PRIVADOS AUXILIARES (Lógica Original Mantida Integralmente)
  // ========================================================================

  /**
   * Limpa as células de destino configuradas em SOMA_VALORES_NAO_COMP_BORA
   */
  _limparResumoValoresNaoComputadosBora_: function (sheetDestino, cfg) {
    if (!cfg || !cfg.ORIGENS) return;
    var origens = cfg.ORIGENS;
    var celulas = [];
    Object.keys(origens).forEach(function (origemKey) {
      var origemCfg = origens[origemKey];
      if (!origemCfg || !origemCfg.SOMA_VALORES_NAO_COMP_BORA) return;
      var resumoCfg = origemCfg.SOMA_VALORES_NAO_COMP_BORA;
      Object.keys(resumoCfg).forEach(function (nomeMetrica) {
        var def = resumoCfg[nomeMetrica];
        if (!def || !def.celulaDestino) return;
        var cel = String(def.celulaDestino).trim();
        if (cel && celulas.indexOf(cel) === -1) celulas.push(cel);
      });
    });
    if (!celulas.length) return;
    celulas.forEach(function (ref) {
      sheetDestino.getRange(ref).clearContent();
    });
  },

  /**
   * Apaga todas as linhas de dados abaixo do cabeçalho.
   */
  _limparTabelaDestino: function (sheetDestino, cfgDestino) {
    const headerRow = cfgDestino.LINHA_CABECALHO;
    const lastRow = sheetDestino.getLastRow();
    if (lastRow > headerRow) {
      const numRows = lastRow - headerRow;
      sheetDestino
        .getRange(headerRow + 1, 1, numRows, sheetDestino.getMaxColumns())
        .clearContent();
    }
  },

  /**
   * Aplica formato de moeda BRL na coluna G (Valor)
   */
  _aplicarFormatoMoedaBRL_: function (sheetDestino, linhaInicial, linhaFinal) {
    const numRows = linhaFinal - linhaInicial + 1;
    if (numRows <= 0) return;
    const colValorIndex = this._colLetterToIndex('G'); 
    const rangeValor = sheetDestino.getRange(linhaInicial, colValorIndex, numRows, 1);
    rangeValor.setNumberFormat('"R$" #,##0.00');
  },

  /**
   * Calcula os resumos (Soma ou Contagem)
   */
  _aplicarResumosValoresNaoComputadosBora_: function (sheetDestino, cfg, linhaInicial, linhaFinal) {
    var origens = cfg.ORIGENS || {};
    var metricas = [];

    // 1) Montar a lista de métricas
    Object.keys(origens).forEach(function (origemKey) {
      var origemCfg = origens[origemKey];
      if (!origemCfg || !origemCfg.SOMA_VALORES_NAO_COMP_BORA) return;
      var resumoCfg = origemCfg.SOMA_VALORES_NAO_COMP_BORA;
      Object.keys(resumoCfg).forEach(function (nomeMetrica) {
        var def = resumoCfg[nomeMetrica];
        if (!def) return;
        var indicador = (def.indicador != null) ? String(def.indicador).trim() : '';
        var modo = (def.modo != null) ? String(def.modo).toUpperCase().trim() : '';
        var colValor = def.colunaValor || null;
        var destino = def.celulaDestino || null;
        if (!indicador || !modo || !destino) return;

        metricas.push({
          indicador: indicador,
          modo: modo,
          colunaValor: colValor,
          celulaDestino: destino,
        });
      });
    });

    if (metricas.length === 0) return;
    var numRows = linhaFinal - linhaInicial + 1;
    if (numRows <= 0) return;

    var colIndicadorIndex = this._colLetterToIndex('A');
    var colFimLetter = (cfg.DESTINO && cfg.DESTINO.COLUNA_FIM) || 'J';
    var colFimIndex = this._colLetterToIndex(colFimLetter);
    var range = sheetDestino.getRange(linhaInicial, 1, numRows, colFimIndex);
    var values = range.getValues();

    metricas.forEach(function (m) {
      if (m.modo === 'VALOR' && m.colunaValor) {
        m.colunaValorIndex = this._colLetterToIndex(m.colunaValor);
      } else {
        m.colunaValorIndex = null;
      }
    }, this);

    var resultados = metricas.map(function (m) {
      return { metrica: m, acumulado: 0 };
    });

    for (var i = 0; i < numRows; i++) {
      var linha = values[i];
      var indicadorLinha = String(linha[colIndicadorIndex - 1] || '').trim();
      if (!indicadorLinha) continue;

      for (var r = 0; r < resultados.length; r++) {
        var m = resultados[r].metrica;
        if (indicadorLinha !== m.indicador) continue;

        if (m.modo === 'VALOR') {
          if (!m.colunaValorIndex) continue;
          var v = linha[m.colunaValorIndex - 1];
          if (typeof v === 'number') {
            resultados[r].acumulado += (v || 0);
          }
        } else if (m.modo === 'LINHAS') {
          resultados[r].acumulado += 1;
        }
      }
    }

    resultados.forEach(function (res) {
      var m = res.metrica;
      var valorFinal = res.acumulado || 0;
      var rangeDestino = sheetDestino.getRange(m.celulaDestino);
      rangeDestino.setValue(valorFinal);
      if (m.modo === 'VALOR') {
        rangeDestino.setNumberFormat('"R$" #,##0.00');
      } else if (m.modo === 'LINHAS') {
        rangeDestino.setNumberFormat('0');
      }
    });
  },

  // --- MÉTODOS DE PROCESSAMENTO DE ORIGENS (Mantidos idênticos ao original) ---
  
  _processarOrigemContratado_: function (shOrigem, shDestino, cfgOrigem, cfgDestino, proximaLinha, numColsDestino) {
    const lastRow = shOrigem.getLastRow();
    if (lastRow <= 1) return proximaLinha;
    const linhaInicio = cfgOrigem.LINHA_INICIO || 4;
    if (linhaInicio > lastRow) return proximaLinha;
    const lastCol = shOrigem.getLastColumn();
    const values = shOrigem.getRange(linhaInicio, 1, lastRow - linhaInicio + 1, lastCol).getValues();
    const idxL = this._colLetterToIndex(cfgOrigem.FILTRO.COLUNA_ELEGIVEL) - 1;
    const idxGrupo = this._colLetterToIndex(cfgOrigem.COLUNA_GRUPO) - 1;
    const linhasBase = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const valorL = row[idxL];
      if (valorL === cfgOrigem.FILTRO.VALOR_ELEGIVEL) {
        linhasBase.push({ rowIndex: linhaInicio + i, data: row });
      }
    }
    if (!linhasBase.length) return proximaLinha;
    const grupos = cfgOrigem.GRUPOS || {};
    const map = cfgOrigem.MAPEAMENTO;
    Object.keys(grupos).forEach((tipoEixo) => {
      const produtosGrupo = grupos[tipoEixo];
      const linhasGrupo = linhasBase.filter(
        (entry) => produtosGrupo.indexOf(String(entry.data[idxGrupo]).trim()) !== -1
      );
      if (!linhasGrupo.length) return;
      const bloco = [];
      for (const entry of linhasGrupo) {
        const row = entry.data;
        const linhaDestino = this._montarLinhaDestino_Contratado(row, tipoEixo, map);
        bloco.push(linhaDestino);
      }
      const colInicioIndex = this._colLetterToIndex(cfgDestino.COLUNA_INICIO);
      const range = shDestino.getRange(proximaLinha, colInicioIndex, bloco.length, numColsDestino);
      range.setValues(bloco);
      proximaLinha += bloco.length;
    });
    return proximaLinha;
  },

  _processarOrigemSeguros_: function (shOrigem, shDestino, cfgOrigem, cfgDestino, proximaLinha, numColsDestino) {
    const lastRow = shOrigem.getLastRow();
    if (lastRow <= 1) return proximaLinha;
    const linhaInicio = cfgOrigem.LINHA_INICIO || 4;
    if (linhaInicio > lastRow) return proximaLinha;
    const lastCol = shOrigem.getLastColumn();
    const values = shOrigem.getRange(linhaInicio, 1, lastRow - linhaInicio + 1, lastCol).getValues();
    const idxL = this._colLetterToIndex(cfgOrigem.FILTRO.COLUNA_ELEGIVEL) - 1;
    const idxStatus = this._colLetterToIndex(cfgOrigem.FILTRO.COLUNA_STATUS) - 1;
    const idxGrupo = this._colLetterToIndex(cfgOrigem.COLUNA_GRUPO) - 1;
    const linhasBase = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const valorL = row[idxL];
      const valorStatus = String(row[idxStatus] || '').trim();
      if (valorL === cfgOrigem.FILTRO.VALOR_ELEGIVEL && valorStatus === cfgOrigem.FILTRO.STATUS_CONTRATADO) {
        linhasBase.push({ rowIndex: linhaInicio + i, data: row });
      }
    }
    if (!linhasBase.length) return proximaLinha;
    const grupos = cfgOrigem.GRUPOS || {};
    const map = cfgOrigem.MAPEAMENTO;
    Object.keys(grupos).forEach((tipoEixo) => {
      const produtosGrupo = grupos[tipoEixo];
      const linhasGrupo = linhasBase.filter(
        (entry) => produtosGrupo.indexOf(String(entry.data[idxGrupo]).trim()) !== -1
      );
      if (!linhasGrupo.length) return;
      const bloco = [];
      for (const entry of linhasGrupo) {
        const row = entry.data;
        const linhaDestino = this._montarLinhaDestino_Seguros(row, tipoEixo, map);
        bloco.push(linhaDestino);
      }
      const colInicioIndex = this._colLetterToIndex(cfgDestino.COLUNA_INICIO);
      const range = shDestino.getRange(proximaLinha, colInicioIndex, bloco.length, numColsDestino);
      range.setValues(bloco);
      proximaLinha += bloco.length;
    });
    return proximaLinha;
  },

  _processarOrigemEncarteiramento_: function (shOrigem, shDestino, cfgOrigem, cfgDestino, proximaLinha, numColsDestino) {
    const lastRow = shOrigem.getLastRow();
    if (lastRow <= 1) return proximaLinha;
    const linhaInicio = cfgOrigem.LINHA_INICIO || 4;
    if (linhaInicio > lastRow) return proximaLinha;
    const lastCol = shOrigem.getLastColumn();
    const values = shOrigem.getRange(linhaInicio, 1, lastRow - linhaInicio + 1, lastCol).getValues();
    const filtro = cfgOrigem.FILTRO || {};
    const colElegivel = filtro.COLUNA_ELEGIVEL;
    const valorElegivel = filtro.VALOR_ELEGIVEL;
    const colStatus = filtro.COLUNA_STATUS;
    const statusValores = (filtro.STATUS_VALORES || []).map(function (s) { return String(s || '').trim(); });
    const idxElegivel = colElegivel ? this._colLetterToIndex(colElegivel) - 1 : null;
    const idxStatus = colStatus ? this._colLetterToIndex(colStatus) - 1 : null;
    const linhasBase = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const isEmpty = row.every((v, idx) => {
        if (idxElegivel !== null && idx === idxElegivel) return v === '' || v === null || v === false;
        return v === '' || v === null;
      });
      if (isEmpty) continue;
      let ok = true;
      if (idxElegivel !== null) ok = ok && (row[idxElegivel] === valorElegivel);
      if (idxStatus !== null && statusValores.length > 0) {
        const statusLinha = String(row[idxStatus] || '').trim();
        ok = ok && (statusValores.indexOf(statusLinha) !== -1);
      }
      if (!ok) continue;
      linhasBase.push({ rowIndex: linhaInicio + i, data: row });
    }
    if (!linhasBase.length) return proximaLinha;
    const map = cfgOrigem.MAPEAMENTO;
    const indicador = 'ENCARTEIRAMENTO';
    const bloco = [];
    for (const entry of linhasBase) {
      const row = entry.data;
      const linhaDestino = this._montarLinhaDestino_Encarteiramento(row, indicador, map);
      bloco.push(linhaDestino);
    }
    if (!bloco.length) return proximaLinha;
    const colInicioIndex = this._colLetterToIndex(cfgDestino.COLUNA_INICIO);
    const range = shDestino.getRange(proximaLinha, colInicioIndex, bloco.length, numColsDestino);
    range.setValues(bloco);
    return proximaLinha + bloco.length;
  },

  _montarLinhaDestino_Contratado: function (rowOrigem, tipoEixo, map) {
    const linha = [];
    linha.push(tipoEixo);
    linha.push(this._getCellByLetter(rowOrigem, map.cnpj.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.cliente.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.produto.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.numeroOperacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.docsflow.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.valor.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.dataInternacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.dataContratacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.comentario.origem));
    return linha;
  },

  _montarLinhaDestino_Encarteiramento: function (rowOrigem, tipoEixo, map) {
    const linha = [];
    linha.push(tipoEixo);
    linha.push(this._getCellByLetter(rowOrigem, map.cnpj.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.cliente.origem));
    linha.push(null);
    linha.push(null);
    linha.push(null);
    linha.push(null);
    linha.push(this._getCellByLetter(rowOrigem, map.dataInternacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.dataContratacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.comentario.origem));
    return linha;
  },

  _montarLinhaDestino_Seguros: function (rowOrigem, tipoEixo, map) {
    const linha = [];
    linha.push(tipoEixo);
    linha.push(this._getCellByLetter(rowOrigem, map.cnpj.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.cliente.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.produto.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.numeroOperacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.docsflow.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.valor.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.dataInternacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.dataContratacao.origem));
    linha.push(this._getCellByLetter(rowOrigem, map.comentario.origem));
    return linha;
  },

  _colLetterToIndex: function (letter) {
    if (!letter) return 0;
    let col = 0;
    const str = String(letter).toUpperCase();
    for (let i = 0; i < str.length; i++) {
      col = col * 26 + (str.charCodeAt(i) - 64);
    }
    return col;
  },

  _numCols: function (colInicio, colFim) {
    const ci = this._colLetterToIndex(colInicio);
    const cf = this._colLetterToIndex(colFim);
    return cf - ci + 1;
  },

  _getCellByLetter: function (row, colLetter) {
    if (!colLetter) return '';
    const idx = this._colLetterToIndex(colLetter) - 1;
    return row[idx];
  },
};

/**
 * Função de entrada para o botão 1: “Analisar operações fora do bora”.
 * (Apenas IMPORTA)
 */
function onClick_ImportarOperacoes() {
  MetasOperacoesManager.importarDados();
}

/**
 * Função de entrada para o botão 2: “Calcular Resumos”.
 * (Apenas CALCULA usando a tabela existente)
 */
function onClick_CalcularResumos() {
  MetasOperacoesManager.calcularResumos();
}

/**
 * Função de entrada para o botão “Limpar operações fora do bora”.
 */
function onClick_LimparOperacoesForaBora() {
  MetasOperacoesManager.limparApenasTabelaEResumos();
}