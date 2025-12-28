// RowManager_v3.0 — Alta Performance com Suporte a Funções e Valores Híbridos
const RowManager = {
  mudarStatusLinha: (sheet, linha, colunaAlterada) => {
    try {
      const nomeAba = sheet.getName();
      const cfgAba = CONFIG?.SHEETS?.[nomeAba];
      const regrasColuna = cfgAba?.MOVER_LINHA_CONFIG?.[colunaAlterada];
      if (!regrasColuna) return;

      const valor = String(sheet.getRange(linha, colunaAlterada).getValue() ?? '').trim();
      const regra = regrasColuna[valor];
      if (!regra || typeof regra === 'string') return;

      const destinoNome = String(regra.destino || '').trim();
      const abaDestino = Utils.getCachedSheet(destinoNome);
      if (!abaDestino) return;

      const map = regra.colunas;
      if (!Array.isArray(map?.origem) || !Array.isArray(map?.destino)) return;

      // Define a linha de destino
      const fnLinha = (typeof regra.linhaDestino === 'function')
        ? regra.linhaDestino
        : (typeof cfgAba?.linhaDestino === 'function') ? cfgAba.linhaDestino
        : (aba => aba.getLastRow() + 1);
      
      const novaLinha = fnLinha(abaDestino);

      let lock;
      try {
        lock = LockService?.getDocumentLock?.();
        lock?.tryLock?.(500);

        // 1. Coleta todos os dados em memória (Lista de {col, val})
        const dadosParaGravar = [];

        for (let i = 0; i < map.origem.length; i++) {
            const origemDef = map.origem[i];
            const destinoColIdx = map.destino[i];
            let valorFinal;

            if (typeof origemDef === 'number') {
                // Modo Clássico: Copiar da origem
                valorFinal = sheet.getRange(linha, origemDef).getValue();
            } else if (typeof origemDef === 'function') {
                // Modo Função: Executar lógica (Data, ID, Calculo)
                try {
                    valorFinal = origemDef(sheet, linha);
                } catch (e) {
                    valorFinal = `Erro: ${e.message}`;
                }
            } else {
                // Modo Valor Fixo
                valorFinal = origemDef;
            }

            // Tratamento de Data para String (se necessário)
            if (valorFinal instanceof Date) {
               valorFinal = Utilities.formatDate(valorFinal, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }

            dadosParaGravar.push({ col: destinoColIdx, val: valorFinal });
        }

        // 2. Agrupamento Inteligente (Smart Batching)
        // Ordena por coluna destino para identificar blocos contíguos (Ex: Col 1, 2, 3, 4...)
        dadosParaGravar.sort((a, b) => a.col - b.col);

        if (dadosParaGravar.length > 0) {
            let blocoAtual = [dadosParaGravar[0].val];
            let colInicial = dadosParaGravar[0].col;
            let colAnterior = colInicial;

            for (let i = 1; i < dadosParaGravar.length; i++) {
                const item = dadosParaGravar[i];
                
                // Se a coluna é vizinha (contígua), adiciona ao mesmo bloco
                if (item.col === colAnterior + 1) {
                    blocoAtual.push(item.val);
                    colAnterior = item.col;
                } else {
                    // Se houve salto (ex: Col 4 p/ Col 8), despacha o bloco anterior e inicia novo
                    const letraCol = RowManager._numToCol(colInicial);
                    BatchOperations.add('setValues', `'${destinoNome}'!${letraCol}${novaLinha}`, [blocoAtual]);
                    
                    blocoAtual = [item.val];
                    colInicial = item.col;
                    colAnterior = colInicial;
                }
            }
            // Despacha o último bloco
            const letraCol = RowManager._numToCol(colInicial);
            BatchOperations.add('setValues', `'${destinoNome}'!${letraCol}${novaLinha}`, [blocoAtual]);
        }

        // 3. Limpeza de colunas vazias configuradas
        (map.destinoVazias || []).forEach(dstIdx => {
          const colLetter = RowManager._numToCol(dstIdx);
          BatchOperations.add('setValue', `'${destinoNome}'!${colLetter}${novaLinha}`, '');
        });

        // 4. Apagar origem se configurado
        if (regra.APAGAR_LINHA_ORIGEM === true) {
          BatchOperations.add('deleteRow', `'${sheet.getName()}'!A${linha}`);
        }

        Logger?.registrarLogBatch?.([['INFO','RowManager',`Linha movida para ${destinoNome} (Smart Batch)`]]);

      } catch (e) {
        Logger?.registrarLogBatch?.([['ERRO','RowManager', e.message]]);
      } finally {
        try { lock?.releaseLock?.(); } catch(_) {}
      }
    } catch (erro) {
      Logger?.registrarLogBatch?.([['ERRO','RowManager', erro.message]]);
    }
  },

  _numToCol: (num) => {
    let s = '';
    while (num > 0) {
      const mod = (num - 1) % 26;
      s = String.fromCharCode(65 + mod) + s;
      num = Math.floor((num - 1) / 26);
    }
    return s || 'A';
  }
};