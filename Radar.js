/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RADAR → PLANNER - FASE 2: CONFIGURÁVEL VIA PLANILHA
 * ═══════════════════════════════════════════════════════════════════════════
 * VERSÃO: 3.0 (Chaves por Letra + Sem Timestamp)
 * DATA: 31/12/2024
 * AUTOR: Sistema RADAR
 * 
 * ATUALIZAÇÕES v3.0:
 * ✅ REMOÇÃO TOTAL DE TIMESTAMP: Script não calcula mais data de vencimento.
 *    (O Power Automate deve calcular: Hoje + sla_dias).
 * ✅ KEY_STRATEGY POR LETRA: Configura-se ['F', 'ABA_LINHA'] em vez de 'DOCSFLOW'.
 * ✅ COLUNAS_LOOKUP LIMPO: Apenas 'categoria' e 'subcategoria'.
 * ✅ MODO CONFERÊNCIA: Filtra apenas por prioridade (já que não há data calculada).
 * ✅ LEITURA TABELAS A/B: Configuração por letras de coluna (sem dependência de headers)
 * 
 * DEPENDÊNCIAS: SpreadsheetApp + MailApp (nativo Google Apps Script)
 * INTEGRAÇÃO: Power Automate lê e-mail → Cria tarefas no Planner
 * ═══════════════════════════════════════════════════════════════════════════
 */

const Radar = (() => {
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CONFIGURAÇÃO DO RADAR (ajustável)
  // ═══════════════════════════════════════════════════════════════════════════
  
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FUNÇÕES AUXILIARES PARA RADAR_CFG DINÂMICO
  // ═══════════════════════════════════════════════════════════════════════════
  
  /**
   * Lê o template de uma aba específica salvo no Properties Service
   * @param {string} nomeAba - Nome da aba (ex: "DEMANDAS DIVERSAS🔧")
   * @returns {Object} - {titulo: string, description: string, assignedTo: string}
   */
  function _getTemplateAba_(nomeAba) {
    const props = PropertiesService.getScriptProperties();
    const templatesJSON = props.getProperty('TEMPLATES_ABAS');
    
    if (templatesJSON) {
      try {
        const allTemplates = JSON.parse(templatesJSON);
        return allTemplates[nomeAba] || {};
      } catch (e) {
        // Se houver erro no parse, retorna vazio
        return {};
      }
    }
    
    return {};
  }
  
  /**
   * Mescla templates salvos com configuração padrão
   * @param {string} nomeAba - Nome da aba
   * @param {Object} defaultConfig - Configuração padrão da aba
   * @returns {Object} - Configuração mesclada (templates do sidebar têm prioridade)
   */
  function _getConfigAba_(nomeAba, defaultConfig) {
    const templates = _getTemplateAba_(nomeAba);
    
    return {
      ...defaultConfig, // Spread da config padrão
      TEMPLATES: {
        titulo: templates.titulo || defaultConfig.TEMPLATES.titulo,
        description: templates.description || defaultConfig.TEMPLATES.description
      },
      valores_fixos: {
        assignedTo: templates.assignedTo || defaultConfig.valores_fixos?.assignedTo || null
      }
    };
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FIM DAS FUNÇÕES AUXILIARES
  // ═══════════════════════════════════════════════════════════════════════════
  
  const RADAR_CFG = {
    
    // ─────────────────────────────────────────────────────────────────────────
    // ENABLED - lê do Properties Service
    // ─────────────────────────────────────────────────────────────────────────
    get ENABLED() {
      const props = PropertiesService.getScriptProperties();
      return props.getProperty('RADAR_ENABLED') !== 'false';
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // CAMPOS DE NEGÓCIO (Tabelas A e B)
    // ─────────────────────────────────────────────────────────────────────────
    CAMPOS_NEGOCIO: {
      tabela_a: ['prioridade', 'sla_dias', 'bucket_override'],
      tabela_b: ['checklist']
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // TRIGGERS - lê do Properties Service
    // ─────────────────────────────────────────────────────────────────────────
    TRIGGERS: {
      get HORA_IMEDIATO() {
        const props = PropertiesService.getScriptProperties();
        return parseInt(props.getProperty('HORA_IMEDIATO') || '20');
      },
      get HORA_CONFERENCIA() {
        const props = PropertiesService.getScriptProperties();
        return parseInt(props.getProperty('HORA_CONFERENCIA') || '6');
      }
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // TABELAS DE CONFIGURAÇÃO (Aba DADOS) - CONFIGURAÇÃO POR LETRAS
    // ─────────────────────────────────────────────────────────────────────────
    DADOS: {
      ABA: 'DADOS',
      
      TABELA_A: {
        RANGE: 'AG1:AN100',
        COLUNAS: {
          ABA: 'AG',
          Categoria: 'AH',
          Subcategoria: 'AI',
          Prioridade: 'AJ',
          CriaTarefa: 'AK',
          SLA_dias: 'AL',
          ChecklistTemplate: 'AM',
          BucketOverride: 'AN'
        }
      },
      
      TABELA_B: {
        RANGE: 'AP1:AS100',
        COLUNAS: {
          Template: 'AP',
          Ordem: 'AQ',
          Item: 'AR',
          Ativo: 'AS'
        }
      }
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // CONFIGURAÇÃO POR ABA (Dinâmica) - agora com getters que leem templates do sidebar
    // ─────────────────────────────────────────────────────────────────────────
    ABAS: {
      get 'DEMANDAS DIVERSAS🔧'() {
        return _getConfigAba_('DEMANDAS DIVERSAS🔧', {
          ABA: 'DEMANDAS DIVERSAS🔧',
          START_ROW: 4,
          
          // Apenas colunas essenciais para achar a regra de negócio na Tabela A
          COLUNAS_LOOKUP: { 
            categoria: 'G', 
            subcategoria: 'H'
          },
          
          // Estratégia de Chave: Tenta coluna F, se vazio, usa Linha
          KEY_STRATEGY: ['F', 'ABA_LINHA'], 
          TEMPLATES: {
            titulo: '{B} | {G} / {H}',
            description: 'DOCSFLOW: {F}\nComentário: {I}'
          },
          valores_fixos: { assignedTo: 'jefferson.santos@basa.com.br' }
        });
      },
      
      get 'SEGUROS🛡️'() {
        return _getConfigAba_('SEGUROS🛡️', {
          ABA: 'SEGUROS🛡️',
          START_ROW: 4,
          COLUNAS_LOOKUP: { categoria: 'I', subcategoria: 'J' },
          
          // Ajuste a letra conforme a coluna real do identificador nesta aba
          KEY_STRATEGY: ['G', 'ABA_LINHA'], 
          TEMPLATES: {
            titulo: '{B} / {C} - OP: {D}',
            description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}'
          },
          valores_fixos: {}
        });
      },
      
      get 'INTERNALIZADO🎯'() {
        return _getConfigAba_('INTERNALIZADO🎯', {
          ABA: 'INTERNALIZADO🎯',
          START_ROW: 4,
          COLUNAS_LOOKUP: { categoria: 'K', subcategoria: 'L' },
          // Exemplo: Usa coluna G como chave, ou fallback para linha
          KEY_STRATEGY: ['G', 'ABA_LINHA'],
          TEMPLATES: {
            titulo: '{B} - {E} - {K}-{L}',
            description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}'
          },
          valores_fixos: {}
        });
      },
      
      get 'EM ANALISE📊'() {
        return _getConfigAba_('EM ANALISE📊', {
          ABA: 'EM ANALISE📊',
          START_ROW: 4,
          COLUNAS_LOOKUP: { categoria: 'J', subcategoria: 'K' },
          KEY_STRATEGY: ['G', 'ABA_LINHA'],
          TEMPLATES: {
            titulo: '{B} - {E} - {J}-{K} - {G}',
            description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}'
          },
          valores_fixos: {}
        });
      },
      
      get 'CADASTROS🧑‍💻'() {
        return _getConfigAba_('CADASTROS🧑‍💻', {
          ABA: 'CADASTROS🧑‍💻',
          START_ROW: 4,
          COLUNAS_LOOKUP: { categoria: 'F', subcategoria: 'G' },
          KEY_STRATEGY: ['E', 'ABA_LINHA'],
          TEMPLATES: {
            titulo: '{B} - Atualização CAD - {C} - {E}',
            description: 'DOCSFLOW: {E}\nResponsável: {D}'
          },
          valores_fixos: {}
        });
      }
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // EMAIL - lê do Properties Service
    // ─────────────────────────────────────────────────────────────────────────
    EMAIL: {
      get DESTINATARIO() {
        const props = PropertiesService.getScriptProperties();
        return props.getProperty('EMAIL_DEST') || 'jefferson.santos@basa.com.br';
      },
      get ASSUNTO() {
        const props = PropertiesService.getScriptProperties();
        return props.getProperty('EMAIL_ASSUNTO') || 'RADAR_PLANNER';
      },
      FORMATO: 'JSON'
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // SAIDA - lê do Properties Service
    // ─────────────────────────────────────────────────────────────────────────
    SAIDA: {
      get MAX_ITENS() {
        const props = PropertiesService.getScriptProperties();
        return parseInt(props.getProperty('MAX_ITENS') || '50');
      },
      LOG_PERFORMANCE: true
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // CONFERENCIA - lê do Properties Service
    // ─────────────────────────────────────────────────────────────────────────
    CONFERENCIA: {
      get MIN_PRIORIDADE() {
        const props = PropertiesService.getScriptProperties();
        return parseInt(props.getProperty('MIN_PRIORIDADE') || '2');
      },
      get MAX_ITENS() {
        const props = PropertiesService.getScriptProperties();
        return parseInt(props.getProperty('MAX_ITENS_CONFERENCIA') || '25');
      }
    },
    
    PRIORITY_MAP: {
      0: 1, 1: 1,
      2: 3, 3: 3, 4: 3,
      5: 5, 6: 5, 7: 5,
      8: 9, 9: 9, 10: 9
    },
    
    // ─────────────────────────────────────────────────────────────────────────
    // BUCKET_MAP - lê do Properties Service
    // ─────────────────────────────────────────────────────────────────────────
    BUCKET_MAP: {
      get valores_validos() {
        const props = PropertiesService.getScriptProperties();
        const bucketsJSON = props.getProperty('BUCKETS_VALIDOS');
        
        if (bucketsJSON) {
          try {
            return JSON.parse(bucketsJSON);
          } catch (e) {
            return ['LIBERAÇÕES', 'CADASTROS', 'PROJETOS', 'TAREFAS PENDENTES'];
          }
        }
        
        return ['LIBERAÇÕES', 'CADASTROS', 'PROJETOS', 'TAREFAS PENDENTES'];
      }
    }
  };
  
  // ═══════════════════════════════════════════════════════════════════════════
  // API PÚBLICA
  // ═══════════════════════════════════════════════════════════════════════════
  
  function runImediato() { 
    return _run({ mode: 'IMEDIATO' }); 
  }
  
  function runConferencia() { 
    return _run({ mode: 'CONFERENCIA' }); 
  }
  
  function setupTriggers() {
    _clearRadarTriggers_();
    ScriptApp.newTrigger('RADAR_runImediato')
      .timeBased()
      .everyDays(1)
      .atHour(RADAR_CFG.TRIGGERS.HORA_IMEDIATO)
      .create();
    ScriptApp.newTrigger('RADAR_runConferencia')
      .timeBased()
      .everyDays(1)
      .atHour(RADAR_CFG.TRIGGERS.HORA_CONFERENCIA)
      .create();
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CORE ENGINE
  // ═══════════════════════════════════════════════════════════════════════════
  
  function _run({ mode }) {
    if (!RADAR_CFG.ENABLED) {
      return { success: true, data: { skipped: true } };
    }
    
    const ss = SpreadsheetApp.getActive();
    const { configsA, templatesChecklist } = _loadConfigsETemplates_(ss);
    const abasNomes = Object.keys(RADAR_CFG.ABAS || {});
    
    let todosItens = [];
    const logResumo = {};
    
    for (const abaKey of abasNomes) {
      const configAba = RADAR_CFG.ABAS[abaKey];
      if (!configAba || !configAba.ABA) continue;
      
      const abaOrigem = ss.getSheetByName(configAba.ABA);
      if (!abaOrigem) {
        logResumo[configAba.ABA] = 'NAO_ENCONTRADA';
        continue;
      }
      
      // Leitura dinâmica
      const rows = _readAba_(abaOrigem, configAba);
      const candidatos = _buildCandidatos_(rows, configsA, templatesChecklist, configAba);
      
      todosItens = todosItens.concat(candidatos);
      logResumo[configAba.ABA] = `${candidatos.length} itens gerados`;
    }
    
    // Filtragem e Ordenação Global
    let selecionados = (mode === 'CONFERENCIA') 
      ? _filtrarConferencia_(todosItens) 
      : todosItens;
      
    selecionados.sort((a, b) => {
      const pa = (a.prioridade ?? 999);
      const pb = (b.prioridade ?? 999);
      // Ordena apenas por prioridade (sem data de vencimento calculada)
      return pa - pb;
    });
    
    if (selecionados.length === 0) {
      return { success: true, data: { status: 'SEM_ITENS', log: logResumo } };
    }
    
    const payload = _montarPayloadEnvio_(selecionados, mode);
    _enviarEmailJSON_(payload, mode);
    
    return { 
      success: true, 
      data: { 
        enviados: payload.itens.length, 
        logPorAba: logResumo 
      } 
    };
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // LEITURA DE DADOS (100% DINÂMICA)
  // ═══════════════════════════════════════════════════════════════════════════
  
  function _readAba_(sheet, configAba) {
    const startRow = configAba.START_ROW;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < startRow) return [];
    
    const maxCol = _resolveMaxCol_(configAba);
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, maxCol);
    const values = range.getValues();
    
    // Resolve índices de negócio
    const colCat = _colLetterToNumber_(configAba.COLUNAS_LOOKUP?.categoria);
    const colSub = _colLetterToNumber_(configAba.COLUNAS_LOOKUP?.subcategoria);
    
    return values.map((r, idx) => {
      const linha = startRow + idx;
      
      // Objeto base
      const row = { linha };
      
      // Mapeia colunas dinâmicas (A_col, B_col...)
      for (let c = 1; c <= maxCol; c++) {
        const letra = _numberToColLetter_(c);
        row[`${letra}_col`] = r[c - 1];
      }
      
      // Aliases para lookup de Configuração (sempre necessários)
      row.G_situacao = colCat ? row[`${_numberToColLetter_(colCat)}_col`] : null;
      row.H_detalhe  = colSub ? row[`${_numberToColLetter_(colSub)}_col`] : null;
      
      return row;
    });
  }
  
  function _buildCandidatos_(rows, configsA, templatesChecklist, configAba) {
    const itens = [];
    // Estratégia de chave: Lista de letras (ex: ['F', 'G']) ou 'ABA_LINHA'
    const strategies = configAba.KEY_STRATEGY || ['ABA_LINHA']; 
    
    for (const row of rows) {
      if (!row.G_situacao || !row.H_detalhe) continue;
      
      const categoria = String(row.G_situacao).trim();
      const subcategoria = String(row.H_detalhe).trim();
      const config = _lookupConfig_(configsA, configAba, categoria, subcategoria);
      
      // Filtro CriaTarefa
      if (!config || !config.CriaTarefa) continue;
      
      const prioridade = config.Prioridade;
      const slaDias = _toIntOrNull_(config.SLA_dias);
      const checklistTemplate = (config.ChecklistTemplate ?? '').toString().trim();
      
      // ⚠️ Sem cálculo de Data no GAS (PA calcula: Hoje + sla_dias)
      const priorityPlanner = RADAR_CFG.PRIORITY_MAP[prioridade] ?? 5;
      const checklistItens = checklistTemplate ? (templatesChecklist[checklistTemplate] || []) : [];
      
      // Gera chaves dinâmicas
      const keyLinha = String(row.linha);
      const key = _buildKey_(strategies, { aba: configAba.ABA, linha: row.linha, row: row });
      
      // Description para PA localizar a tarefa
      const keyDescription = _buildKeyDescription_({ 
        aba: configAba.ABA, 
        linha: row.linha, 
        keyUsed: key 
      });
      
      const titulo = _buildTitulo_(configAba, row);
      const description = _buildDescription_(configAba, row, keyDescription);
      const bucket = config.BucketOverride || 'TAREFAS PENDENTES';
      const assignedTo = configAba.valores_fixos?.assignedTo || null;
      
      // Limpeza do row para envio no JSON
      const dadosPlanilha = { ...row };
      delete dadosPlanilha.linha;
      delete dadosPlanilha.G_situacao; // remove alias interno
      delete dadosPlanilha.H_detalhe;  // remove alias interno
      
      itens.push({
        keyLinha,
        keyDescription,
        description,
        origem: { aba: configAba.ABA, linha: row.linha },
        titulo, 
        bucket, 
        assignedTo, 
        prioridade, 
        priorityPlanner, 
        sla_dias: slaDias, // Enviamos apenas os dias
        checklistTemplate: checklistTemplate || null, 
        checklistItens,
        dadosPlanilha // Contém { A_col: '...', B_col: '...' }
      });
    }
    
    return itens;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FUNÇÕES AUXILIARES
  // ═══════════════════════════════════════════════════════════════════════════
  
  function _filtrarConferencia_(itens) {
    // ⚠️ Conferência agora só olha Prioridade (Data é calculada no PA)
    return itens.filter(it => {
      const prioOk = (it.prioridade ?? 999) <= RADAR_CFG.CONFERENCIA.MIN_PRIORIDADE;
      return prioOk;
    });
  }
  
  function _montarPayloadEnvio_(itens, mode) {
    const max = (mode === 'CONFERENCIA') 
      ? RADAR_CFG.CONFERENCIA.MAX_ITENS 
      : RADAR_CFG.SAIDA.MAX_ITENS;
    const selecionados = itens.slice(0, max);
    
    return {
      meta: {
        mode,
        geradoEm: new Date().toISOString(),
        origem: 'CONSOLIDADO',
        totalSelecionado: selecionados.length
      },
      itens: selecionados
    };
  }
  
  function _enviarEmailJSON_(payload, mode) {
    MailApp.sendEmail({
      to: RADAR_CFG.EMAIL.DESTINATARIO,
      subject: RADAR_CFG.EMAIL.ASSUNTO + (mode === 'CONFERENCIA' ? ' [CONFERENCIA]' : ''),
      body: JSON.stringify(payload, null, 2)
    });
  }
  
  function _loadConfigsETemplates_(ss) {
    const sheetDados = ss.getSheetByName(RADAR_CFG.DADOS.ABA);
    if (!sheetDados) {
      throw new Error(`RADAR: aba DADOS não encontrada`);
    }
    
    return { 
      configsA: _readTabelaA_(sheetDados, RADAR_CFG.DADOS.TABELA_A),
      templatesChecklist: _readTabelaB_(sheetDados, RADAR_CFG.DADOS.TABELA_B)
    };
  }
  
  /**
   * Lê Tabela A usando configuração de colunas por letras
   * Não depende de headers - totalmente configurável
   */
  function _readTabelaA_(sheet, config) {
    const range = sheet.getRange(config.RANGE);
    const vals = range.getValues();
    const rangeStartCol = range.getColumn();
    const out = [];
    
    // Calcula índices relativos uma vez
    const getIdx = (letra) => _colLetterToNumber_(letra) - rangeStartCol;
    
    const idxABA = getIdx(config.COLUNAS.ABA);
    const idxCat = getIdx(config.COLUNAS.Categoria);
    const idxSub = getIdx(config.COLUNAS.Subcategoria);
    const idxPrio = getIdx(config.COLUNAS.Prioridade);
    const idxCria = getIdx(config.COLUNAS.CriaTarefa);
    const idxSLA = getIdx(config.COLUNAS.SLA_dias);
    const idxCheck = getIdx(config.COLUNAS.ChecklistTemplate);
    const idxBucket = getIdx(config.COLUNAS.BucketOverride);
    
    // Pula linha 1 (header) e processa dados
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      
      const aba = r[idxABA];
      const cat = r[idxCat];
      
      if (!aba || !cat) continue; // Pula linha vazia
      
      out.push({
        ABA: String(aba).trim(),
        Categoria: String(cat).trim(),
        Subcategoria: String(r[idxSub] || '').trim(),
        Prioridade: _toIntOrNull_(r[idxPrio]),
        CriaTarefa: String(r[idxCria] || '').toUpperCase() === 'SIM',
        SLA_dias: r[idxSLA],
        ChecklistTemplate: String(r[idxCheck] || '').trim() || null,
        BucketOverride: String(r[idxBucket] || '').trim() || null
      });
    }
    
    return out;
  }
  
  /**
   * Lê Tabela B usando configuração de colunas por letras
   * Não depende de headers - totalmente configurável
   */
  function _readTabelaB_(sheet, config) {
    const range = sheet.getRange(config.RANGE);
    const vals = range.getValues();
    const rangeStartCol = range.getColumn();
    const map = {};
    
    if (vals.length < 2) return map;
    
    // Calcula índices relativos
    const getIdx = (letra) => _colLetterToNumber_(letra) - rangeStartCol;
    
    const idxTemplate = getIdx(config.COLUNAS.Template);
    const idxOrdem = getIdx(config.COLUNAS.Ordem);
    const idxItem = getIdx(config.COLUNAS.Item);
    const idxAtivo = getIdx(config.COLUNAS.Ativo);
    
    // Pula linha 1 (header) e processa dados
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      
      const t = String(r[idxTemplate] || '').trim();
      const ativo = String(r[idxAtivo] || '').toUpperCase();
      
      if (!t || ativo === 'FALSE' || ativo === 'NÃO' || ativo === 'NAO') continue;
      
      if (!map[t]) map[t] = [];
      map[t].push({ 
        ordem: _toIntOrNull_(r[idxOrdem]) ?? 999, 
        item: String(r[idxItem]) 
      });
    }
    
    // Ordena itens por ordem e extrai só o texto
    Object.keys(map).forEach(k => {
      map[k] = map[k].sort((a, b) => a.ordem - b.ordem).map(x => x.item);
    });
    
    return map;
  }
  
  function _lookupConfig_(configsA, configAba, cat, sub) {
    // Busca exata
    let found = configsA.find(c => 
      c.ABA === configAba.ABA && 
      c.Categoria === cat && 
      c.Subcategoria === sub
    );
    
    // Busca com wildcard (*)
    if (!found) {
      found = configsA.find(c => 
        c.ABA === configAba.ABA && 
        c.Categoria === cat && 
        c.Subcategoria === '*'
      );
    }
    
    return found;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // HELPERS E PARSERS
  // ═══════════════════════════════════════════════════════════════════════════
  
  function _indexHeaders_(h, req) {
    const idx = {};
    req.forEach(n => { 
      if(h.indexOf(n) === -1) throw new Error(`Header ausente: ${n}`); 
      idx[n] = h.indexOf(n); 
    });
    return idx;
  }
  
  function _colLetterToNumber_(l) {
    if (!l) return 0;
    let n = 0, s = String(l).toUpperCase();
    for(let i = 0; i < s.length; i++) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n;
  }
  
  function _numberToColLetter_(n) {
    let s = '', x = n;
    while(x > 0) { 
      s = String.fromCharCode(65 + (x - 1) % 26) + s; 
      x = Math.floor((x - 1) / 26); 
    }
    return s || 'A';
  }
  
  function _resolveMaxCol_(cfg) {
    let max = 11;
    const p = (l) => { 
      const n = _colLetterToNumber_(l); 
      if(n > max) max = n; 
    };
    
    // Calcula max col baseado nas Lookups e na Estratégia de Chave
    p(cfg.COLUNAS_LOOKUP?.categoria); 
    p(cfg.COLUNAS_LOOKUP?.subcategoria);
    
    // Adiciona colunas usadas na KEY_STRATEGY
    if (Array.isArray(cfg.KEY_STRATEGY)) {
      cfg.KEY_STRATEGY.forEach(k => {
        if (k !== 'ABA_LINHA') p(k);
      });
    }
    
    // Escaneia templates para encontrar referências {LETRA}
    const txt = (cfg.TEMPLATES?.titulo || '') + (cfg.TEMPLATES?.description || '');
    let m; 
    const re = /\{([A-Z]+)\}/g;
    while((m = re.exec(txt)) !== null) p(m[1]);
    
    return max;
  }
  
  function _toIntOrNull_(v) { 
    return (v == null || v === '') ? null : Math.trunc(Number(v)); 
  }
  
  /**
   * Constrói a chave única iterando sobre as estratégias (Letras ou ABA_LINHA)
   */
  function _buildKey_(strategies, {aba, linha, row}) {
    for (const s of strategies) {
      if (s === 'ABA_LINHA') {
        return `RADAR:${aba}:LINHA:${linha}`;
      }
      
      // Assume que 's' é uma letra de coluna (Ex: 'F')
      // Verifica se existe valor na coluna correspondente
      const colKey = `${s}_col`; // ex: F_col
      const val = row[colKey];
      
      if (val && String(val).trim() !== '') {
        // Retorna chave baseada no valor encontrado na coluna
        return `RADAR:${aba}:KEY:${String(val).trim()}`;
      }
    }
    
    // Fallback final de segurança
    return `RADAR:${aba}:LINHA:${linha}`;
  }
  
  function _buildKeyDescription_({aba, linha, keyUsed}) {
    // Removemos Timestamp, colocamos a chave usada para referencia
    return `[RADAR_KEY]\nABA: ${aba}\nLINHA: ${linha}\nKEY: ${keyUsed}\n[/RADAR_KEY]`;
  }
  
  function _buildTitulo_(cfg, row) {
    let t = _parseTemplate_(cfg.TEMPLATES.titulo, row);
    return t.length > 255 ? t.substring(0, 252) + '...' : t;
  }
  
  function _buildDescription_(cfg, row, keyDesc) {
    return keyDesc + '\n\n' + _parseTemplate_(cfg.TEMPLATES.description, row);
  }
  
  function _parseTemplate_(tpl, row) {
    return (tpl || '').replace(/\{([A-Z]+)\}/g, (m, l) => {
      const key = l + '_col';
      return row[key] != null ? String(row[key]).trim() : '';
    });
  }
  
  function _clearRadarTriggers_() {
    ScriptApp.getProjectTriggers().forEach(t => {
      if(['RADAR_runImediato','RADAR_runConferencia'].includes(t.getHandlerFunction())) {
        ScriptApp.deleteTrigger(t);
      }
    });
  }
  
  // Retorna API pública
  return { runImediato, runConferencia, setupTriggers };
  
})();

// ═══════════════════════════════════════════════════════════════════════════
// FUNÇÕES GLOBAIS (chamadas por triggers)
// ═══════════════════════════════════════════════════════════════════════════

function RADAR_runImediato() { 
  return Radar.runImediato(); 
}

function RADAR_runConferencia() { 
  return Radar.runConferencia(); 
}

function RADAR_setupTriggers() { 
  return Radar.setupTriggers(); 
}

// ═══════════════════════════════════════════════════════════════════════════
// FUNÇÕES BACKEND PARA SIDEBAR CONFIGURADOR RADAR
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// CARREGA CONFIGURAÇÕES COMPLETAS (incluindo templates e buckets)
// ═══════════════════════════════════════════════════════════════════════════

function getConfigCompleta() {
  const props = PropertiesService.getScriptProperties();
  
  // Configurações básicas
  const config = {
    enabled: props.getProperty('RADAR_ENABLED') !== 'false',
    horaImediato: props.getProperty('HORA_IMEDIATO') || '20',
    horaConferencia: props.getProperty('HORA_CONFERENCIA') || '6',
    emailDestinatario: props.getProperty('EMAIL_DEST') || 'jefferson.santos@basa.com.br',
    assunto: props.getProperty('EMAIL_ASSUNTO') || 'RADAR_PLANNER',
    maxItens: props.getProperty('MAX_ITENS') || '50',
    maxItensConferencia: props.getProperty('MAX_ITENS_CONFERENCIA') || '25',
    minPrioridade: props.getProperty('MIN_PRIORIDADE') || '2'
  };
  
  // Templates por aba (salvos como JSON)
  const templatesJSON = props.getProperty('TEMPLATES_ABAS');
  if (templatesJSON) {
    try {
      config.templates = JSON.parse(templatesJSON);
    } catch (e) {
      config.templates = _getTemplatesPadrao_();
    }
  } else {
    config.templates = _getTemplatesPadrao_();
  }
  
  // Buckets válidos (salvos como JSON)
  const bucketsJSON = props.getProperty('BUCKETS_VALIDOS');
  if (bucketsJSON) {
    try {
      config.buckets = JSON.parse(bucketsJSON);
    } catch (e) {
      config.buckets = _getBucketsPadrao_();
    }
  } else {
    config.buckets = _getBucketsPadrao_();
  }
  
  return config;
}

// ═══════════════════════════════════════════════════════════════════════════
// SALVA CONFIGURAÇÕES COMPLETAS
// ═══════════════════════════════════════════════════════════════════════════

function saveConfigCompleta(config) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    // Salva configurações básicas
    props.setProperty('RADAR_ENABLED', String(config.enabled));
    props.setProperty('HORA_IMEDIATO', String(config.horaImediato));
    props.setProperty('HORA_CONFERENCIA', String(config.horaConferencia));
    props.setProperty('EMAIL_DEST', String(config.emailDestinatario));
    props.setProperty('EMAIL_ASSUNTO', String(config.assunto));
    props.setProperty('MAX_ITENS', String(config.maxItens));
    props.setProperty('MAX_ITENS_CONFERENCIA', String(config.maxItensConferencia));
    props.setProperty('MIN_PRIORIDADE', String(config.minPrioridade));
    
    // Salva templates como JSON
    if (config.templates) {
      props.setProperty('TEMPLATES_ABAS', JSON.stringify(config.templates));
    }
    
    // Salva buckets como JSON
    if (config.buckets) {
      props.setProperty('BUCKETS_VALIDOS', JSON.stringify(config.buckets));
    }
    
    Logger.log('Configurações completas salvas: ' + JSON.stringify(config));
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Erro ao salvar configurações: ' + error);
    return { success: false, error: error.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// RESTAURA PADRÕES COMPLETOS
// ═══════════════════════════════════════════════════════════════════════════

function restaurarPadroesCompletos() {
  return {
    enabled: true,
    horaImediato: '20',
    horaConferencia: '6',
    emailDestinatario: 'jefferson.santos@basa.com.br',
    assunto: 'RADAR_PLANNER',
    maxItens: '50',
    maxItensConferencia: '25',
    minPrioridade: '2',
    templates: _getTemplatesPadrao_(),
    buckets: _getBucketsPadrao_()
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// TEMPLATES PADRÃO POR ABA
// ═══════════════════════════════════════════════════════════════════════════

function _getTemplatesPadrao_() {
  return {
    'DEMANDAS DIVERSAS🔧': {
      titulo: '{B} | {G} / {H}',
      description: 'DOCSFLOW: {F}\nComentário: {I}',
      assignedTo: 'jefferson.santos@basa.com.br'
    },
    'SEGUROS🛡️': {
      titulo: '{B} / {C} - OP: {D}',
      description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}',
      assignedTo: ''
    },
    'INTERNALIZADO🎯': {
      titulo: '{B} - {E} - {K}-{L}',
      description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}',
      assignedTo: ''
    },
    'EM ANALISE📊': {
      titulo: '{B} - {E} - {J}-{K} - {G}',
      description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}',
      assignedTo: ''
    },
    'CADASTROS🧑‍💻': {
      titulo: '{B} - Atualização CAD - {C} - {E}',
      description: 'DOCSFLOW: {E}\nResponsável: {D}',
      assignedTo: ''
    }
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// BUCKETS PADRÃO
// ═══════════════════════════════════════════════════════════════════════════

function _getBucketsPadrao_() {
  return [
    'LIBERAÇÕES',
    'CADASTROS',
    'PROJETOS',
    'TAREFAS PENDENTES'
  ];
}