const onEditHandler = {
  gerenciarEdicoes: (e) => {
    // Validações iniciais de segurança
    if (!e || !e.range) return;
    
    const sheet = e.source.getActiveSheet();
    const aba = sheet.getName();
    const linha = e.range.getRow();
    const coluna = e.range.getColumn();
    
    // Carrega configurações
    const configAba = Utils.obterConfigAba(aba);
    const configParcelas = CONFIG.ACOMPANHAMENTO_PARCELAS;

    // -----------------------------------------------------------------------
    // ███ Etapa 1: ACOMPANHAMENTO DE PARCELAS (Aba SEGUROS)
    // -----------------------------------------------------------------------
    if (configParcelas.ENABLED === true && aba === configParcelas.ABA && linha >= configParcelas.LINHA_INICIO) {
       if (typeof ParcelasManager !== 'undefined' && ParcelasManager.monitorarEdicao) {
         // Se houver lógica de monitoramento em tempo real (ex: check de checkbox)
         ParcelasManager.monitorarEdicao(e);
       }
    }

    // -----------------------------------------------------------------------
    // ███ Etapa 2: MOVIMENTAÇÃO DE LINHAS (RowManager)
    // Usa LockService para evitar conflitos em edições rápidas
    // -----------------------------------------------------------------------
    if (configAba?.MOVER_LINHA_CONFIG) {
      let lockRow;
      try {
        lockRow = LockService.getScriptLock();
        // Aguarda até 1s para garantir integridade se houver muitas edições simultâneas
        if (lockRow.tryLock(1000)) { 
          if (typeof RowManager !== 'undefined' && RowManager.mudarStatusLinha) {
            RowManager.mudarStatusLinha(sheet, linha, coluna);
          }
        }
      } catch (err) {
        Logger.registrarLogBatch([["ERRO", "RowManager Lock", err.message]]);
      } finally {
        if (lockRow) lockRow.releaseLock();
      }
    }

    // -----------------------------------------------------------------------
    // ███ Etapa 3: REGRAS DE NEGÓCIO (Validações e Automatismos)
    // -----------------------------------------------------------------------
    try {
      if (typeof RegrasNegocioV2 !== 'undefined') {
        RegrasNegocioV2.processar(e);
      }
    } catch (err) {
      Logger.registrarLogBatch([["AVISO", "RegrasNegocioV2", err.message]]);
    }

    // -----------------------------------------------------------------------
    // ███ Etapa 4: CRIAÇÃO DE DOCUMENTOS (Gatilho +DOC)
    // Verifica se a coluna editada é a coluna de gatilho configurada
    // -----------------------------------------------------------------------
    if (configAba?.GOOGLE_DOC_CONFIG && coluna === configAba.GOOGLE_DOC_CONFIG.TRIGGER_COL) {
       if (typeof DocumentGenerator !== 'undefined' && DocumentGenerator.criarDocumento) {
         // O DocumentGenerator valida internamente se o texto é "+DOC"
         DocumentGenerator.criarDocumento(e);
       }
    }

    // -----------------------------------------------------------------------
    // ███ Etapa 5: INTEGRAÇÃO PRODUTOS -> SEGUROS (Fila Assíncrona)
    // -----------------------------------------------------------------------
    try {
      const cfgProd = (CONFIG.PRODUTOS_SEGUROS || {});
      if (cfgProd.enabled && aba === cfgProd.origem) {
         if (typeof ProdutosSegurosSimples !== 'undefined' && ProdutosSegurosSimples.processar) {
           ProdutosSegurosSimples.processar(e);
         }
      }
    } catch (erroPSS) {
      Logger.registrarLogBatch([["ERRO", "ProdutosSeguros", erroPSS.message]]);
    }

    // -----------------------------------------------------------------------
    // ███ Etapa 6: TIMESTAMP (Registro de Data/Hora)
    // -----------------------------------------------------------------------
    if (configAba?.COLUNAS_MONITORADAS?.includes(Utils.colunaParaLetra(coluna))) {
      TimestampManager.registrarTimestamp(sheet, e.range);
    }

    // -----------------------------------------------------------------------
    // ███ Etapa Final: EXECUÇÃO EM LOTE (BatchOperations)
    // Executa todas as escritas pendentes acumuladas (incluindo o link do Doc)
    // -----------------------------------------------------------------------
    if (CONFIG.BATCH_OPS?.ENABLED && typeof BatchOperations !== 'undefined') {
      BatchOperations.execute('onEdit');
    }
  }
};