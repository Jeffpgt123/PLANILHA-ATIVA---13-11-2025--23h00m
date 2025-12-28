// Arquivo: ProdutosSeguros.js
// VERSÃO 6.7: DUPLICAÇÃO DE PRODUTO EM SEGUROS PERMITIDA
//
// Changelog v6.7:
// - REMOVIDO: Regra que impedia a criação de linhas duplicadas de produto para o mesmo ID de Cliente na aba SEGUROS.
// - MANTIDO: O lock de performance (ScriptLock/DocumentLock) e a verificação de duplicidade dentro do LOTE atual (cacheLote).

const ProdutosSegurosSimples = (function () {
  const CFG = (typeof CONFIG !== 'undefined' && CONFIG.PRODUTOS_SEGUROS) || {};
  
  // --- Configurações ---
  const ABA_ORIGEM = CFG.origem || 'EM ANALISE📊';
  const ABA_SEG = CFG.destino || 'SEGUROS🛡️';
  const COL_ORIGEM_LETRA = (CFG.colunaOrigem || 'J').toUpperCase();
  const COL_SEG_PROD_NUM = CFG.colunaProdutoDestinoNum || 3; 
  const COL_SEG_ID_NUM = 2; 
  const START_ROW_SEG = CFG.startRow || 4; 
  const MAPA_COPIA = CFG.mapeamento || {};

  const KEY_QUEUE = 'PRODUTOS_SEGUROS_QUEUE';

  // --- Helpers de Infraestrutura ---
  
  function _getSegSheet() {
    if (typeof Utils !== 'undefined' && Utils.getCachedSheet) {
      return Utils.getCachedSheet(ABA_SEG);
    }
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_SEG);
  }

  function _colLetraParaNumero(letra) {
    if (!letra) return 0;
    letra = letra.toString().toUpperCase().replace(/[^A-Z]/g, '');
    let soma = 0;
    for (let i = 0; i < letra.length; i++) {
      soma = soma * 26 + (letra.charCodeAt(i) - 64);
    }
    return soma;
  }

  function _norm(s) {
    return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  }

  function _parseList(text) {
    if (!text) return [];
    return text.toString().split(/\r?\n|;|,/).map(s => s.trim()).filter(s => s !== '');
  }

  function _getProximaLinhaVaziaPeloID(sheet) {
    const lastRowGeral = sheet.getLastRow();
    if (lastRowGeral < START_ROW_SEG) return START_ROW_SEG;
    const rangeIDs = sheet.getRange(1, COL_SEG_ID_NUM, lastRowGeral, 1).getValues();
    for (let i = lastRowGeral - 1; i >= START_ROW_SEG - 1; i--) {
      if (rangeIDs[i][0] !== "" && rangeIDs[i][0] !== null) {
        return (i + 1) + 1;
      }
    }
    return START_ROW_SEG;
  }

  // --- Gerenciamento da Fila (ScriptProperties + ScriptLock) ---
  // Usa ScriptLock: É global para o script, ideal para proteger a propriedade JSON.

  function _enqueue(item) {
    try {
      // Lock específico para a Fila (RÁPIDO)
      const lock = LockService.getScriptLock();
      
      // Tenta pegar lock por 3s. Como só lemos/escrevemos JSON, é muito rápido.
      if (lock.tryLock(3000)) { 
        const props = PropertiesService.getScriptProperties();
        const currentQueueJson = props.getProperty(KEY_QUEUE);
        let queue = currentQueueJson ? JSON.parse(currentQueueJson) : [];
        queue.push(item);
        props.setProperty(KEY_QUEUE, JSON.stringify(queue));
        lock.releaseLock();
        return true;
      } else {
        console.error('Timeout ao tentar enfileirar (ScriptLock ocupado).');
      }
    } catch (e) {
      console.error('Erro ao enfileirar:', e);
    }
    return false;
  }

  function _dequeueAll() {
    try {
      // O processador chama isso, mas precisamos garantir atomicidade com o _enqueue
      const lock = LockService.getScriptLock();
      if (lock.tryLock(2000)) {
        const props = PropertiesService.getScriptProperties();
        const currentQueueJson = props.getProperty(KEY_QUEUE);
        let queue = [];
        
        if (currentQueueJson) {
          queue = JSON.parse(currentQueueJson);
          if (queue.length > 0) {
            props.deleteProperty(KEY_QUEUE);
          }
        }
        lock.releaseLock();
        return queue;
      }
    } catch (e) {
      console.error('Erro ao desenfileirar:', e);
    }
    return [];
  }

  // --- Função Principal (Gatilho) ---
  function processar(e) {
    try {
      if (!e || !e.range) return;
      const sheet = e.range.getSheet();
      
      if (sheet.getName() !== ABA_ORIGEM) return;
      if (e.range.getColumn() !== _colLetraParaNumero(COL_ORIGEM_LETRA)) return;
      if (e.range.getRow() < 4) return;
      
      const idCliente = sheet.getRange(e.range.getRow(), 2).getValue(); 
      if (!idCliente) return; 

      const produtosTexto = e.value || e.range.getValue();
      const lastCol = sheet.getLastColumn();
      const dadosOrigem = sheet.getRange(e.range.getRow(), 1, 1, lastCol).getValues()[0];

      const pedido = {
        idCliente: idCliente,
        produtosTexto: produtosTexto,
        dadosOrigem: dadosOrigem,
        timestamp: new Date().getTime()
      };
      
      // 1. Enfileira (Usa ScriptLock - Rápido)
      // Isso não briga mais com o processador lento.
      _enqueue(pedido);
      
      // 2. Tenta Processar (Usa DocumentLock - Lento/Singleton)
      _processarFila();

    } catch (err) {
      console.error(`ERRO em ProdutosSeguros: ${err.message}`);
    }
  }

  // --- O Processador de Fila (Stateful + DocumentLock) ---
  function _processarFila() {
    // Usa DocumentLock: Garante que apenas UM processador rode por vez neste documento.
    // Isso é separado do ScriptLock da fila.
    const lock = LockService.getDocumentLock();
    
    // Tenta pegar o lock. Se não conseguir (0ms), significa que já tem um processador rodando.
    // Podemos sair tranquilos, pois o item já está na fila e o processador ativo vai pegar.
    // Coloquei 500ms de tolerância para casos de troca de turno (um acabando, outro começando).
    if (!lock.tryLock(500)) return; 

    try {
      let fila = _dequeueAll(); // Usa ScriptLock internamente por breves milissegundos
      const startTime = new Date().getTime();
      
      const seg = _getSegSheet();
      if (!seg) return;

      // Inicializa Cursor
      let nextRowCursor = _getProximaLinhaVaziaPeloID(seg);
      const cacheLoteAdicionados = new Set(); 

      while (fila.length > 0) {
        
        fila.forEach(pedido => {
           nextRowCursor = _executarPedidoStateful(seg, pedido, nextRowCursor, cacheLoteAdicionados);
        });

        SpreadsheetApp.flush(); // Salva o lote na planilha

        // Timeout de segurança do script (25s)
        if (new Date().getTime() - startTime > 25000) break;
        
        // Busca mais itens que podem ter chegado enquanto processávamos
        fila = _dequeueAll();
        
        // Se chegaram novos itens, atualizamos o cursor pois o flush consolidou os anteriores
        if (fila.length > 0) {
          nextRowCursor = _getProximaLinhaVaziaPeloID(seg);
        }
      }

    } catch (e) {
      console.error('Erro no processador:', e);
    } finally {
      lock.releaseLock();
    }
  }

  // --- Lógica de Negócio (Stateful) ---
  function _executarPedidoStateful(seg, pedido, cursorLinha, cacheLote) {
    const { idCliente, produtosTexto, dadosOrigem } = pedido;
    
    const listaDesejada = _parseList(produtosTexto);
    if (listaDesejada.length === 0) return cursorLinha;

    // A lógica de verificação de produtos existentes na planilha foi removida
    // para permitir a criação de linhas duplicadas (Produto+Cliente).
    
    const idClienteNorm = _norm(idCliente); 

    const paraAdicionar = listaDesejada.filter(prod => {
      const prodNorm = _norm(prod);
      const chaveUnica = idClienteNorm + "|" + prodNorm;
      
      // Mantido: Bloqueia a adição se o produto já foi adicionado NESTE lote de processamento
      if (cacheLote.has(chaveUnica)) return false; 
      
      return true;
    });

    if (paraAdicionar.length === 0) return cursorLinha;

    // Escrita usando Cursor
    paraAdicionar.forEach(prodName => {
      const novaLinhaArr = new Array(seg.getLastColumn()).fill('');
      
      if (COL_SEG_PROD_NUM <= novaLinhaArr.length) novaLinhaArr[COL_SEG_PROD_NUM - 1] = prodName;
      
      if (Object.keys(MAPA_COPIA).length === 0) {
          novaLinhaArr[0] = dadosOrigem[0]; 
          novaLinhaArr[1] = dadosOrigem[1];
      } else {
          for (const [colOrig, colDest] of Object.entries(MAPA_COPIA)) {
            const idxOrig = _colLetraParaNumero(colOrig) - 1;
            const idxDest = _colLetraParaNumero(colDest) - 1;
            if (idxOrig >= 0 && idxOrig < dadosOrigem.length && idxDest >= 0) {
               novaLinhaArr[idxDest] = dadosOrigem[idxOrig];
            }
          }
      }

      seg.getRange(cursorLinha, 1, 1, novaLinhaArr.length).setValues([novaLinhaArr]);
      
      cacheLote.add(idClienteNorm + "|" + _norm(prodName));
      cursorLinha++;
    });
    
    if (typeof Logger !== 'undefined' && Logger.registrarLogBatch) {
       Logger.registrarLogBatch([['INFO', 'ProdutosSeguros', `Lote: +${paraAdicionar.length} para ${idCliente}`]]);
    }

    return cursorLinha;
  }

  return { processar };
})();