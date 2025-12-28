/**
 * MetasOperacoesManager
 * Reconstrói a tabela OPERACOES_FORA_BORA na aba METAS🎯
 * a partir das abas CONTRATADO/LIBERACAO✅ e SEGUROS🛡️.
 *
 * Regras:
 * - Executado manualmente via botão.
 * - Limpa a tabela destino (mantendo cabeçalho na linha 50).
 * - Lê origens em memória, filtra por L = FALSE (e status, no caso de SEGUROS).
 * - Classifica linhas por grupo (produto).
 * - Para cada linha elegível, gera 1 linha na tabela contínua.
 */
const MetasOperacoesManager = {
  executar: function () {
    const cfg = CONFIG?.METAS_OPERACOES_FORA_BORA;
    if (!cfg || cfg.ENABLED === false) return;

    const ss = SpreadsheetApp.getActive();
    const shDestino = ss.getSheetByName(cfg.DESTINO.ABA);
    if (!shDestino) return;

    // 0) Limpar células de resumo (SOMA_VALORES_NAO_COMP_BORA)
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

    // 4) Aplicar formato de moeda BRL na coluna G (Valor), para todas as linhas preenchidas
    const primeiraLinhaDados = cfg.DESTINO.PRIMEIRA_LINHA_DADOS;
    const ultimaLinhaPreenchida = proximaLinhaDestino - 1;

    if (ultimaLinhaPreenchida >= primeiraLinhaDados) {
      this._aplicarFormatoMoedaBRL_(
        shDestino,
        primeiraLinhaDados,
        ultimaLinhaPreenchida
      );

      // Após formatar os valores, aplica os resumos configurados em SOMA_VALORES_NAO_COMP_BORA
      this._aplicarResumosValoresNaoComputadosBora_(
        shDestino,
        cfg,
        primeiraLinhaDados,
        ultimaLinhaPreenchida
      );
    }
  },

  /**
   * Limpa apenas a tabela OPERACOES_FORA_BORA e os resumos
   * de SOMA_VALORES_NAO_COMP_BORA, sem recálculo.
   * Usado por um botão exclusivo de limpeza.
   */
  limparApenasTabelaEResumos: function () {
    const cfg = CONFIG?.METAS_OPERACOES_FORA_BORA;
    if (!cfg || cfg.ENABLED === false) return;

    const ss = SpreadsheetApp.getActive();
    const shDestino = ss.getSheetByName(cfg.DESTINO.ABA);
    if (!shDestino) return;

    // Limpa células de somatório (resumos)
    this._limparResumoValoresNaoComputadosBora_(shDestino, cfg);

    // Limpa apenas as linhas de dados da tabela OPERACOES_FORA_BORA (mantém cabeçalho)
    this._limparTabelaDestino(shDestino, cfg.DESTINO);
  },

  /**
   * Limpa as células de destino configuradas em SOMA_VALORES_NAO_COMP_BORA
   * nas origens configuradas (ex.: CONTRATADO_LIBERACAO, SEGUROS),
   * antes de recalcular os resumos.
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
        if (cel && celulas.indexOf(cel) === -1) {
          celulas.push(cel);
        }
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
      // Se quiser, pode usar .clearFormat() ou não, depende do layout.
    }
  },

  /**
   * Aplica formato de moeda BRL na coluna G (Valor)
   * entre linhaInicial e linhaFinal (inclusive) da aba de destino.
   */
  _aplicarFormatoMoedaBRL_: function (sheetDestino, linhaInicial, linhaFinal) {
    const numRows = linhaFinal - linhaInicial + 1;
    if (numRows <= 0) return;

    const colValorIndex = this._colLetterToIndex('G'); // coluna G em METAS
    const rangeValor = sheetDestino.getRange(linhaInicial, colValorIndex, numRows, 1);

    // Formato padrão de moeda BRL
    rangeValor.setNumberFormat('"R$" #,##0.00');
  },

  /**
   * Calcula os resumos de "OPERAÇÕES/VALORES NÃO COMPUTADOS NO BORA"
   * usando a tabela OPERACOES_FORA_BORA já preenchida em METAS🎯.
   *
   * Para cada métrica em SOMA_VALORES_NAO_COMP_BORA:
   * - modo = 'VALOR'  → soma os valores da coluna configurada;
   * - modo = 'LINHAS' → conta quantas linhas pertencem ao indicador;
   * e grava o resultado na celulaDestino.
   */
  _aplicarResumosValoresNaoComputadosBora_: function (sheetDestino, cfg, linhaInicial, linhaFinal) {
    var origens = cfg.ORIGENS || {};
    var metricas = [];

    // 1) Montar a lista de métricas a partir de todas as origens
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

        // Só considera métricas com indicador, modo e célula de destino válidos
        if (!indicador || !modo || !destino) return;

        metricas.push({
          indicador: indicador,
          modo: modo,              // 'VALOR' ou 'LINHAS'
          colunaValor: colValor,   // ex.: 'G' ou null
          celulaDestino: destino,  // ex.: 'E11'
        });
      });
    });

    if (metricas.length === 0) return;

    var numRows = linhaFinal - linhaInicial + 1;
    if (numRows <= 0) return;

    // Indicador SEMPRE está na coluna A da tabela OPERACOES_FORA_BORA
    var colIndicadorIndex = this._colLetterToIndex('A');

    // Ler até COLUNA_FIM definida no DESTINO (por padrão, 'J')
    var colFimLetter = (cfg.DESTINO && cfg.DESTINO.COLUNA_FIM) || 'J';
    var colFimIndex = this._colLetterToIndex(colFimLetter);

    var range = sheetDestino.getRange(linhaInicial, 1, numRows, colFimIndex);
    var values = range.getValues();

    // Pré-calcula índice numérico da coluna de valor para cada métrica de VALOR
    metricas.forEach(function (m) {
      if (m.modo === 'VALOR' && m.colunaValor) {
        m.colunaValorIndex = this._colLetterToIndex(m.colunaValor);
      } else {
        m.colunaValorIndex = null;
      }
    }, this);

    // Inicializa acumuladores
    var resultados = metricas.map(function (m) {
      return {
        metrica: m,
        acumulado: 0,
      };
    });

    // 3) Percorrer as linhas e acumular resultados
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

    // 4) Escrever resultados nas células destino, com formatação adequada
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

  /**
   * Processa a origem CONTRATADO/LIBERACAO✅
   * - Filtro: L = FALSE
   * - Agrupamento: coluna E (COLUNA_GRUPO)
   * - Grupos: FNO, RPL, OUTRAS_FONTES (cfg.GRUPOS)
   */
  _processarOrigemContratado_: function (
    shOrigem,
    shDestino,
    cfgOrigem,
    cfgDestino,
    proximaLinha,
    numColsDestino
  ) {
    const lastRow = shOrigem.getLastRow();
    if (lastRow <= 1) return proximaLinha;

    const linhaInicio = cfgOrigem.LINHA_INICIO || 4; // ajustar se necessário
    if (linhaInicio > lastRow) return proximaLinha;

    const lastCol = shOrigem.getLastColumn();
    const values = shOrigem.getRange(linhaInicio, 1, lastRow - linhaInicio + 1, lastCol).getValues();

    const idxL = this._colLetterToIndex(cfgOrigem.FILTRO.COLUNA_ELEGIVEL) - 1;
    const idxGrupo = this._colLetterToIndex(cfgOrigem.COLUNA_GRUPO) - 1;

    // 1) Filtrar base por L = FALSE
    const linhasBase = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const valorL = row[idxL];
      if (valorL === cfgOrigem.FILTRO.VALOR_ELEGIVEL) {
        linhasBase.push({ rowIndex: linhaInicio + i, data: row });
      }
    }

    if (!linhasBase.length) return proximaLinha;

    // 2) Para cada grupo de produtos (FNO, RPL, OUTRAS_FONTES)
    const grupos = cfgOrigem.GRUPOS || {};
    const map = cfgOrigem.MAPEAMENTO;

    Object.keys(grupos).forEach((tipoEixo) => {
      const produtosGrupo = grupos[tipoEixo]; // array de strings

      // Seleciona linhas cuja coluna grupo (E) esteja na lista de produtos
      const linhasGrupo = linhasBase.filter(
        (entry) => produtosGrupo.indexOf(String(entry.data[idxGrupo]).trim()) !== -1
      );

      if (!linhasGrupo.length) return;

      // Montar bloco de saída
      const bloco = [];
      for (const entry of linhasGrupo) {
        const row = entry.data;
        const linhaDestino = this._montarLinhaDestino_Contratado(row, tipoEixo, map);
        bloco.push(linhaDestino);
      }

      // Escrever no destino em bloco
      const colInicioIndex = this._colLetterToIndex(cfgDestino.COLUNA_INICIO);
      const range = shDestino.getRange(proximaLinha, colInicioIndex, bloco.length, numColsDestino);
      range.setValues(bloco);
      proximaLinha += bloco.length;
    });

    return proximaLinha;
  },

  /**
   * Processa a origem SEGUROS🛡️
   * - Filtro: L = FALSE e I = 'PAGO/CONTRATADO'
   * - Agrupamento: coluna C
   * - Grupos: CAPITALIZACAO, RISCOS_EMPRESARIAIS, etc.
   */
  _processarOrigemSeguros_: function (
    shOrigem,
    shDestino,
    cfgOrigem,
    cfgDestino,
    proximaLinha,
    numColsDestino
  ) {
    const lastRow = shOrigem.getLastRow();
    if (lastRow <= 1) return proximaLinha;

    const linhaInicio = cfgOrigem.LINHA_INICIO || 4; // ajustar se necessário
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
      if (
        valorL === cfgOrigem.FILTRO.VALOR_ELEGIVEL &&
        valorStatus === cfgOrigem.FILTRO.STATUS_CONTRATADO
      ) {
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

  /**
   * Processa a origem ENCARTEIRAMENTO🧑‍💻
   * - Filtro: COLUNA_ELEGIVEL = VALOR_ELEGIVEL (ex.: L = FALSE)
   * - Todas as linhas elegíveis vão para o indicador único "ENCARTEIRAMENTO".
   */
 _processarOrigemEncarteiramento_: function (
    shOrigem,
    shDestino,
    cfgOrigem,
    cfgDestino,
    proximaLinha,
    numColsDestino
  ) {
    const lastRow = shOrigem.getLastRow();
    if (lastRow <= 1) return proximaLinha;

    const linhaInicio = cfgOrigem.LINHA_INICIO || 4; // segue padrão das outras origens
    if (linhaInicio > lastRow) return proximaLinha;

    const lastCol = shOrigem.getLastColumn();
    const values = shOrigem
      .getRange(linhaInicio, 1, lastRow - linhaInicio + 1, lastCol)
      .getValues();

    const filtro = cfgOrigem.FILTRO || {};

    // CONFIG:
    // FILTRO.COLUNA_ELEGIVEL = 'G' (NO BORA?)
    // FILTRO.VALOR_ELEGIVEL  = false
    // FILTRO.COLUNA_STATUS   = 'C' (SOLICITAÇÃO)
    // FILTRO.STATUS_VALORES  = ['SOL. ENCARTEIRAMENTO', 'ENCARTEIRADO']
    const colElegivel = filtro.COLUNA_ELEGIVEL;   // 'G'
    const valorElegivel = filtro.VALOR_ELEGIVEL;  // false
    const colStatus = filtro.COLUNA_STATUS;       // 'C'
    const statusValores = (filtro.STATUS_VALORES || []).map(function (s) {
      return String(s || '').trim();
    });

    const idxElegivel = colElegivel
      ? this._colLetterToIndex(colElegivel) - 1
      : null;
    const idxStatus = colStatus
      ? this._colLetterToIndex(colStatus) - 1
      : null;

    const linhasBase = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      // 0) Ignora linhas “vazias”: tudo vazio e, no máximo, o checkbox em G = false
      const isEmpty = row.every((v, idx) => {
        // coluna do checkbox (G): considera false como "vazio" para esse teste
        if (idxElegivel !== null && idx === idxElegivel) {
          return v === '' || v === null || v === false;
        }
        return v === '' || v === null;
      });
      if (isEmpty) continue;

      let ok = true;

      // 1) G == false (ou VALOR_ELEGIVEL)
      if (idxElegivel !== null) {
        ok = ok && (row[idxElegivel] === valorElegivel);
      }

      // 2) C ∈ STATUS_VALORES
      if (idxStatus !== null && statusValores.length > 0) {
        const statusLinha = String(row[idxStatus] || '').trim();
        ok = ok && (statusValores.indexOf(statusLinha) !== -1);
      }

      // Só entra se TODAS as condições forem verdadeiras
      if (!ok) continue;

      linhasBase.push({ rowIndex: linhaInicio + i, data: row });
    }

    if (!linhasBase.length) return proximaLinha;

    const map = cfgOrigem.MAPEAMENTO;
    const indicador = 'ENCARTEIRAMENTO';
    const bloco = [];

    for (const entry of linhasBase) {
      const row = entry.data;
      const linhaDestino = this._montarLinhaDestino_Encarteiramento(
        row,
        indicador,
        map
      );
      bloco.push(linhaDestino);
    }

    if (!bloco.length) return proximaLinha;

    const colInicioIndex = this._colLetterToIndex(cfgDestino.COLUNA_INICIO);
    const range = shDestino.getRange(
      proximaLinha,
      colInicioIndex,
      bloco.length,
      numColsDestino
    );
    range.setValues(bloco);

    return proximaLinha + bloco.length;
  },

  /**
   * Monta uma linha de saída (CONTRATADO/LIBERACAO✅ → METAS🎯)
   * de acordo com o mapeamento configurado.
   */
  _montarLinhaDestino_Contratado: function (rowOrigem, tipoEixo, map) {
    const linha = [];

    // Coluna A: Tipo/Eixo
    linha.push(tipoEixo); // A

    // Colunas B…J conforme mapeamento (A→B, B→C, E→D, F→E, G→F, H→G, I→H, J→I, K→J)
    linha.push(this._getCellByLetter(rowOrigem, map.cnpj.origem)); // B
    linha.push(this._getCellByLetter(rowOrigem, map.cliente.origem)); // C
    linha.push(this._getCellByLetter(rowOrigem, map.produto.origem)); // D
    linha.push(this._getCellByLetter(rowOrigem, map.numeroOperacao.origem)); // E
    linha.push(this._getCellByLetter(rowOrigem, map.docsflow.origem)); // F
    linha.push(this._getCellByLetter(rowOrigem, map.valor.origem)); // G
    linha.push(this._getCellByLetter(rowOrigem, map.dataInternacao.origem)); // H
    linha.push(this._getCellByLetter(rowOrigem, map.dataContratacao.origem)); // I
    linha.push(this._getCellByLetter(rowOrigem, map.comentario.origem)); // J

    return linha;
  },

  /**
   * Monta uma linha de saída (ENCARTEIRAMENTO🧑‍💻 → METAS🎯),
   * seguindo o mapeamento A:B, B:C, E:H, F:I, D:J.
   */
  _montarLinhaDestino_Encarteiramento: function (rowOrigem, tipoEixo, map) {
    const linha = [];

    // Coluna A: Indicador fixo (ENCARTEIRAMENTO)
    linha.push(tipoEixo); // A

    // B: CNPJ (A → B)
    linha.push(this._getCellByLetter(rowOrigem, map.cnpj.origem)); // B

    // C: Cliente (B → C)
    linha.push(this._getCellByLetter(rowOrigem, map.cliente.origem)); // C

    // D, E, F, G não são usados neste fluxo de ENCARTEIRAMENTO
    linha.push(null); // D
    linha.push(null); // E
    linha.push(null); // F
    linha.push(null); // G

    // H: Data internação (E → H)
    linha.push(this._getCellByLetter(rowOrigem, map.dataInternacao.origem)); // H

    // I: Data contratação (F → I)
    linha.push(this._getCellByLetter(rowOrigem, map.dataContratacao.origem)); // I

    // J: Comentário (D → J)
    linha.push(this._getCellByLetter(rowOrigem, map.comentario.origem)); // J

    return linha;
  },

  /**
   * Monta uma linha de saída (SEGUROS🛡️ → METAS🎯),
   * respeitando o mapeamento funcional atual.
   */
  _montarLinhaDestino_Seguros: function (rowOrigem, tipoEixo, map) {
    const linha = [];

    // Coluna A: Indicador/Grupo (ex.: CAPITALIZACAO, RISCOS_EMPRESARIAIS, CONSORCIO)
    linha.push(tipoEixo); // A

    // B…J conforme mapeamento A:B, B:C, C:D, F:E, G:F, E:G, H:H, K:I, M:J
    linha.push(this._getCellByLetter(rowOrigem, map.cnpj.origem));            // B
    linha.push(this._getCellByLetter(rowOrigem, map.cliente.origem));         // C
    linha.push(this._getCellByLetter(rowOrigem, map.produto.origem));         // D
    linha.push(this._getCellByLetter(rowOrigem, map.numeroOperacao.origem));  // E
    linha.push(this._getCellByLetter(rowOrigem, map.docsflow.origem));        // F
    linha.push(this._getCellByLetter(rowOrigem, map.valor.origem));           // G
    linha.push(this._getCellByLetter(rowOrigem, map.dataInternacao.origem));  // H
    linha.push(this._getCellByLetter(rowOrigem, map.dataContratacao.origem)); // I
    linha.push(this._getCellByLetter(rowOrigem, map.comentario.origem));      // J

    return linha;
  },

  /**
   * Converte letra de coluna (ex.: 'A') em índice (1-based).
   */
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

  /**
   * Lê valor de rowOrigem pela letra de coluna (ex.: 'A' → row[0]).
   */
  _getCellByLetter: function (row, colLetter) {
    if (!colLetter) return '';
    const idx = this._colLetterToIndex(colLetter) - 1;
    return row[idx];
  },
};

/**
 * Função de entrada para o botão “Analisar operações fora do bora”.
 */
function onClick_AnalisarOperacoesForaBora() {
  MetasOperacoesManager.executar();
}

/**
 * Função de entrada para o botão “Limpar operações fora do bora”.
 * Limpa a tabela OPERACOES_FORA_BORA e os somatórios configurados,
 * sem recálculo.
 */
function onClick_LimparOperacoesForaBora() {
  MetasOperacoesManager.limparApenasTabelaEResumos();
}
