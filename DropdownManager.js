/**
 * ████████████████████████████████████████████████████████████████████████
 * ▶ DROPDOWNMANAGER - VERSÃO ORIGINAL MODIFICADA
 * ▶ Estrutura original mantida (Cache, Utils, Lógica Verbosa)
 * ▶ MODIFICAÇÃO APLICADA: Limpeza estrita de formatação em subcategorias vazias
 * ████████████████████████████████████████████████████████████████████████
 */

const DropdownManager = {
  
  _cache: { dados: new Map(), subcategorias: new Map(), timestamps: new Map(), TTL: 300000 },
  _norm: (s) => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase(),
  _decodeKey: (k) => (typeof k==='string' && k.endsWith('_SAFE')) ? k.replace('_SAFE','') : k,

  // === INFRAESTRUTURA ===
  _lerConfiguracaoPersistida: () => {
    try {
      const docProps = PropertiesService.getDocumentProperties();
      const scriptProps = PropertiesService.getScriptProperties();
      const KEY = 'DROPDOWN_CORES_OVERRIDE';

      let json = docProps.getProperty(KEY);
      if (!json) {
        json = scriptProps.getProperty(KEY);
      }
      return json;
    } catch (e) {
      console.error("Erro ao ler propriedades:", e);
      return null;
    }
  },

  _gerarIdLista: (nomeAba, nomeColuna) => {
    const norm = (str) => (str || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    return `${norm(nomeAba)}_${norm(nomeColuna)}`;
  },

  _colParaIdx: (c) => {
    if(typeof c === 'number') return c;
    // Tenta Utils primeiro
    if(typeof Utils !== 'undefined' && Utils.colunaParaIndice) {
      try {
        return Utils.colunaParaIndice(c);
      } catch(e) {
        // Fallback se Utils falhar
      }
    }
    // Fallback nativo robusto
    let col = 0, s = String(c).toUpperCase(); 
    for(let i=0; i < s.length; i++) col = col * 26 + (s.charCodeAt(i) - 64);
    return col;
  },

  _mapCor: (coresCfg, valor) => {
    if (!coresCfg) return null;
    if (valor in coresCfg) return coresCfg[valor];
    const k = DropdownManager._norm(valor);
    for (const chave in coresCfg) {
      const chaveReal = DropdownManager._decodeKey(chave);
      if (DropdownManager._norm(chaveReal) === k) return coresCfg[chave];
    }
    return null;
  },

  /**
   * Obtém lista de configurações mesclando HARDCODED + OVERRIDE
   */
  _obterListasConfig: (nomeAba = null) => {
    const listasPorAba = [];
    const sheetsCfg = CONFIG?.SHEETS || {};
    
    // 1. Lê do CONFIG
    Object.keys(sheetsCfg).forEach(aba => {
      const arr = sheetsCfg[aba]?.LISTAS_SUSPENSAS || [];
      arr.forEach(cfg => listasPorAba.push({ nomeAba: aba, ...cfg }));
    });
    const globais = (CONFIG?.LISTAS_SUSPENSAS || []).filter(g => !nomeAba || g.nomeAba === nomeAba);
    let todasListas = [...listasPorAba, ...globais];

    if (nomeAba) todasListas = todasListas.filter(l => l.nomeAba === nomeAba);

    // 2. Merge com Override Salvo
    try {
      const overridesJSON = DropdownManager._lerConfiguracaoPersistida();
      if (overridesJSON) {
        const overrides = JSON.parse(overridesJSON);

        const sanitizeMap = (m) => {
          const out = {};
          if (!m || typeof m !== 'object') return out;
          for (const k in m) {
            const v = m[k];
            if (!v || typeof v !== 'object') continue;
            const bg = (v.fundo || '').toString().toLowerCase();
            const tx = (v.texto || '').toString().toLowerCase();
            if (!bg || !tx) continue;
            if (bg === '#ffffff' && tx === '#000000') continue; // Ignora neutro
            out[k] = v;
          }
          return out;
        };

        todasListas = todasListas.map(lista => {
          const idListaCalculado = DropdownManager._gerarIdLista(lista.nomeAba, lista.colunaCategoria);
          
          let ov = overrides[idListaCalculado];
          if (!ov) {
              const idNorm = idListaCalculado.toUpperCase();
              const keyFound = Object.keys(overrides).find(k => k.toUpperCase() === idNorm);
              if (keyFound) ov = overrides[keyFound];
          }

          if (ov) {
            const coresOv = sanitizeMap(ov.cores);
            const subOv = sanitizeMap(ov.coresSubcategoria);
            lista.cores = { ...(lista.cores || {}), ...coresOv };
            lista.coresSubcategoria = { ...(lista.coresSubcategoria || {}), ...subOv };
          }
          return lista;
        });
      }
    } catch (e) {
      console.error('Erro ao ler overrides de cores:', e);
    }

    return todasListas;
  },

  /**
   * Obtém configuração de cores dinâmica (OnEdit)
   */
  _obterConfigCoresDinamica: (configBase, listaId) => {
    try {
      const json = DropdownManager._lerConfiguracaoPersistida();
      const savedConfig = JSON.parse(json || '{}');

      const getSaved = (idTarget) => {
        if (savedConfig[idTarget]) return savedConfig[idTarget];
        const normTarget = idTarget.toUpperCase(); 
        const foundKey = Object.keys(savedConfig).find(k => k.toUpperCase() === normTarget);
        return foundKey ? savedConfig[foundKey] : null;
      };

      const sanitizeMap = (m) => {
        const out = {};
        if (!m || typeof m !== 'object') return out;
        for (const k in m) {
          const v = m[k];
          if (!v || typeof v !== 'object') continue;
          const bg = (v.fundo || '').toString().toLowerCase();
          const tx = (v.texto || '').toString().toLowerCase();
          if (!bg || !tx) continue;
          if (bg === '#ffffff' && tx === '#000000') continue;
          out[k] = v;
        }
        return out;
      };

      const idFinal = listaId || DropdownManager._gerarIdLista(configBase.nomeAba, configBase.colunaCategoria);
      const saved = getSaved(idFinal);
      
      if (saved && (Object.keys(saved.cores || {}).length > 0 || Object.keys(saved.coresSubcategoria || {}).length > 0)) {
        return {
          ...configBase,
          cores: sanitizeMap(saved.cores),
          coresSubcategoria: sanitizeMap(saved.coresSubcategoria)
        };
      }
      
      return {
        ...configBase,
        cores: configBase.cores || {},
        coresSubcategoria: configBase.coresSubcategoria || {}
      };
    } catch (erro) {
      console.error('Erro config dinâmica:', erro.message);
      return configBase;
    }
  },

  /**
   * FUNÇÃO PRINCIPAL: CRIA AS LISTAS
   */
  criarListasSuspensas: () => {
    try {
      console.log('🚀 DropdownManager - Iniciando...');
      const _todas = DropdownManager._obterListasConfig();
      if (!_todas?.length) return { sucesso: false, erro: 'Sem listas configuradas' };
      
      DropdownManager._limparCache();
      
      const resultados = [];
      let sucessos = 0;
      
      for (let i = 0; i < _todas.length; i++) {
        const config = _todas[i];
        try {
          const res = DropdownManager._processarListaIndividual(config, i+1);
          resultados.push(res);
          if (res.sucesso) sucessos++;
        } catch (e) {
          resultados.push({ numero: i+1, sucesso: false, erro: e.message });
        }
      }
      
      return {
        sucesso: sucessos > 0,
        processadas: sucessos,
        total: _todas.length,
        detalhes: resultados
      };
    } catch (erro) {
      console.error('Erro geral:', erro.message);
      return { sucesso: false, erro: erro.message };
    }
  },

  _processarListaIndividual: (config, numero) => {
    if (!config.nomeAba || !config.colunaCategoria || !config.fonte) {
      return { sucesso: false, erro: 'Config incompleta' };
    }
    
    const aba = (typeof Utils !== 'undefined') ? Utils.getCachedSheet(config.nomeAba) : SpreadsheetApp.getActive().getSheetByName(config.nomeAba);
    if (!aba) return { sucesso: false, erro: `Aba ${config.nomeAba} não encontrada` };
    
    const categorias = DropdownManager._lerCategoriasDaFonte(config);
    if (!categorias.length) return { sucesso: false, erro: 'Nenhuma categoria na fonte' };
    
    DropdownManager._aplicarValidacaoRobusta(aba, config, categorias, numero);
    
    return { 
      sucesso: true, 
      categorias: categorias.length 
    };
  },

  // === LEITURA DE DADOS (RESTAURADO) ===

  _lerCategoriasDaFonte: (config) => {
    const fonte = config.fonte || config;
    const chaveCache = `${fonte.nomeAba}_${fonte.colunaCategoria}`;
    const agora = Date.now();
    
    if (DropdownManager._cache.dados.has(chaveCache)) {
      const ts = DropdownManager._cache.timestamps.get(chaveCache);
      if (ts && (agora - ts) < DropdownManager._cache.TTL) {
        return DropdownManager._cache.dados.get(chaveCache);
      }
    }
    
    const categorias = DropdownManager._lerDadosDaFonte(fonte);
    
    if (categorias.length > 0) {
      DropdownManager._cache.dados.set(chaveCache, categorias);
      DropdownManager._cache.timestamps.set(chaveCache, agora);
    }
    return categorias;
  },

  _lerDadosDaFonte: (fonte) => {
    try {
      const aba = (typeof Utils !== 'undefined') ? Utils.getCachedSheet(fonte.nomeAba) : SpreadsheetApp.getActive().getSheetByName(fonte.nomeAba);
      if (!aba) return [];
      
      const coluna = DropdownManager._colParaIdx(fonte.colunaCategoria);
      const linhaInicial = fonte.linhaInicial || 4;
      const ultimaLinha = Math.min(aba.getLastRow(), linhaInicial + 3000);
      
      if (ultimaLinha < linhaInicial) return [];
      
      const valores = aba.getRange(linhaInicial, coluna, ultimaLinha - linhaInicial + 1, 1).getValues();
      
      return valores
          .flat()
          .map(v => v?.toString().trim())
          .filter(v => v && v !== '')
          .filter((v, i, a) => a.indexOf(v) === i) // Unique
          .sort();
    } catch (e) {
      console.error('Erro ler dados fonte:', e);
      return [];
    }
  },

  _aplicarValidacaoRobusta: (aba, config, categorias, num) => {
    try {
      const coluna = DropdownManager._colParaIdx(config.colunaCategoria);
      const linhaInicial = config.fonte.linhaInicial || 4;
      
      const validacao = SpreadsheetApp.newDataValidation()
        .requireValueInList(categorias)
        .setAllowInvalid(false)
        .setHelpText(`Lista ${num}`)
        .build();
      
      try {
        aba.getRange(linhaInicial, coluna, aba.getMaxRows() - linhaInicial + 1, 1).setDataValidation(validacao);
      } catch (e) {
        aba.getRange(linhaInicial, coluna, 500, 1).setDataValidation(validacao);
      }
      return true;
    } catch (e) {
      return false;
    }
  },

  // === ON EDIT (Processamento) ===

  processarEdicao: (e) => {
    try {
      if (!e || !e.range) return;
      const sheet = e.range.getSheet();
      const aba = sheet.getName();
      const linha = e.range.getRow();
      const coluna = e.range.getColumn();

      // Recarrega config fresca
      const listas = DropdownManager._obterListasConfig(aba);
      if (!listas || !listas.length) return;

      // 1. Categoria Alterada
      const configCat = listas.find(l => DropdownManager._colParaIdx(l.colunaCategoria) === coluna);
      if (configCat) {
        DropdownManager._criarSubcategoria(sheet, linha, configCat, e.value);
        DropdownManager._aplicarCor(sheet, linha, configCat, e.value);
        return;
      }

      // 2. Subcategoria Alterada
      const configSub = listas.find(l => DropdownManager._colParaIdx(l.colunaSubcategoria) === coluna);
      if (configSub) {
        const colCatPai = DropdownManager._colParaIdx(configSub.colunaCategoria);
        const valPai = sheet.getRange(linha, colCatPai).getValue();
        DropdownManager._aplicarCor(sheet, linha, configSub, valPai);
      }
    } catch (e) {
      console.error(e);
    }
  },

  // === SUBCATEGORIAS ===

  _criarSubcategoria: (sheet, linha, config, categoria) => {
    try {
      const colSub = DropdownManager._colParaIdx(config.colunaSubcategoria);
      const rangeSub = sheet.getRange(linha, colSub);
      
      // Busca Subcategorias Reais
      const subcategorias = DropdownManager._lerSubcategorias(config, categoria);

      // Limpa conteúdo E FORMATAÇÃO (MODIFICAÇÃO SOLICITADA)
      rangeSub.clearContent();
      rangeSub.setBackground(null).setFontColor(null).setFontWeight(null); 

      if (subcategorias.length > 0) {
        const regra = SpreadsheetApp.newDataValidation()
          .requireValueInList(subcategorias, true)
          .setAllowInvalid(true)
          .build();
        rangeSub.setDataValidation(regra);
      } else {
        rangeSub.clearDataValidations();
        // A formatação já foi limpa acima
      }
      return subcategorias;
    } catch (e) {
      console.error('Erro criar sub:', e);
      return [];
    }
  },

  _lerSubcategorias: (config, categoria) => {
    try {
      const chaveCache = `${config.fonte.nomeAba}_${categoria}`;
      const agora = Date.now();
      
      if (DropdownManager._cache.subcategorias.has(chaveCache)) {
        const ts = DropdownManager._cache.timestamps.get(chaveCache);
        if (ts && (agora - ts) < DropdownManager._cache.TTL) {
          return DropdownManager._cache.subcategorias.get(chaveCache);
        }
      }
      
      const aba = (typeof Utils !== 'undefined') ? Utils.getCachedSheet(config.fonte.nomeAba) : SpreadsheetApp.getActive().getSheetByName(config.fonte.nomeAba);
      if (!aba) return [];
      
      const colCat = DropdownManager._colParaIdx(config.fonte.colunaCategoria);
      const colSub = DropdownManager._colParaIdx(config.fonte.colunaSubcategoria);
      const row = config.fonte.linhaInicial || 4;
      const lastRow = Math.min(aba.getLastRow(), row + 3000);
      
      const subcategorias = [];
      const catNorm = DropdownManager._norm(categoria);
      
      const range = aba.getRange(row, 1, lastRow - row + 1, Math.max(colCat, colSub));
      const valores = range.getValues();

      valores.forEach(r => {
        const vCat = r[colCat-1];
        if (DropdownManager._norm(vCat) === catNorm) {
          const vSub = r[colSub-1];
          if (vSub && String(vSub).trim() !== '' && !subcategorias.includes(String(vSub).trim())) {
            subcategorias.push(String(vSub).trim());
          }
        }
      });
      
      if (subcategorias.length > 0) {
        DropdownManager._cache.subcategorias.set(chaveCache, subcategorias);
        DropdownManager._cache.timestamps.set(chaveCache, agora);
      }
      return subcategorias;
    } catch (e) {
      return [];
    }
  },

  // === APLICAÇÃO DE CORES (MODIFICADA) ===
  
  _aplicarCor: (sheet, linha, configBase, categoria) => {
    try {
      const listaId = DropdownManager._gerarIdLista(configBase.nomeAba, configBase.colunaCategoria);
      const config = DropdownManager._obterConfigCoresDinamica(configBase, listaId);

      const colCat = DropdownManager._colParaIdx(config.colunaCategoria);
      const colSub = config.colunaSubcategoria ? DropdownManager._colParaIdx(config.colunaSubcategoria) : null;

      const rCat = sheet.getRange(linha, colCat);
      const rSub = colSub ? sheet.getRange(linha, colSub) : null;

      const valCat = String(categoria || '').trim();
      const valSub = rSub ? String(rSub.getValue() || '').trim() : '';
      const temSub = valSub !== '';

      const corCat = (config.cores && valCat) ? DropdownManager._mapCor(config.cores, valCat) : null;
      const corSub = (config.coresSubcategoria && temSub) ? DropdownManager._mapCor(config.coresSubcategoria, valSub) : null;

      const ehNeutro = (c) => {
        if (!c || typeof c !== 'object') return true;
        if (!c.fundo || !c.texto) return true;
        const bg = String(c.fundo).toLowerCase();
        const tx = String(c.texto).toLowerCase();
        return (bg === '#ffffff' && tx === '#000000');
      };
      
      const valido = (c) => c && c.fundo && c.texto && !ehNeutro(c);

      // Herança: Subcategoria > Categoria
      let final = null;
      if (valido(corSub)) final = corSub;
      else if (valido(corCat)) final = corCat;

      if (final) {
        // Pinta a Categoria
        rCat.setBackground(final.fundo).setFontColor(final.texto).setFontWeight(final.negrito ? 'bold' : 'normal');
        
        // MODIFICAÇÃO: Só pinta a Subcategoria se houver valor selecionado
        if (rSub && temSub) {
           rSub.setBackground(final.fundo).setFontColor(final.texto).setFontWeight(final.negrito ? 'bold' : 'normal');
        } else if (rSub) {
           // Se a subcategoria estiver vazia, garante limpeza
           rSub.setBackground(null).setFontColor(null).setFontWeight(null);
        }
      } else {
        // Cor neutra: limpa formatação
        rCat.setBackground(null).setFontColor(null).setFontWeight(null);
        if (rSub) rSub.setBackground(null).setFontColor(null).setFontWeight(null);
      }
    } catch (e) {
      console.error('Erro _aplicarCor:', e);
    }
  },

  _limparCache: () => {
    DropdownManager._cache.dados.clear();
    DropdownManager._cache.subcategorias.clear();
    DropdownManager._cache.timestamps.clear();
    console.log('🧹 Cache limpo');
  },

  getStatus: () => {
    return {
      cacheDados: DropdownManager._cache.dados.size,
      listas: DropdownManager._obterListasConfig().length
    };
  }
};

/**
 * Funções Globais (Pontos de Entrada)
 */
function executarDropdownCompleto() {
  return DropdownManager.criarListasSuspensas().sucesso;
}

function diagnosticoDropdownCompleto() {
  const s = DropdownManager.getStatus();
  console.log(s);
  return s;
}

function testeRapidoDropdown() {
  return DropdownManager.criarListasSuspensas();
}

function resetDropdownCompleto() {
  DropdownManager._limparCache();
  if (typeof Utils !== 'undefined') Utils.invalidateCache('all');
  return true;
}

function resetDropdownCompletoDetalhes() {
  const res = resetDropdownCompleto();
  if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert("Reset completo realizado com sucesso (Detalhes).");
  }
  return res;
}