/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RADAR → PLANNER - FASE 2: CONFIGURÁVEL VIA PLANILHA
 * ═══════════════════════════════════════════════════════════════════════════
 * * VERSÃO: 3.0 (Chaves por Letra + Sem Timestamp)
 * DATA: 31/12/2024
 * AUTOR: Sistema RADAR
 * * ATUALIZAÇÕES:
 * ✅ REMOÇÃO TOTAL DE TIMESTAMP: Script não calcula mais data de vencimento.
 * (O Power Automate deve calcular: Hoje + sla_dias).
 * ✅ KEY_STRATEGY POR LETRA: Configura-se ['F', 'ABA_LINHA'] em vez de 'DOCSFLOW'.
 * ✅ COLUNAS_LOOKUP LIMPO: Apenas 'categoria' e 'subcategoria'.
 * ✅ MODO CONFERÊNCIA: Filtra apenas por prioridade (já que não há data calculada).
 * * DEPENDÊNCIAS: SpreadsheetApp + MailApp (nativo Google Apps Script)
 * INTEGRAÇÃO: Power Automate lê e-mail → Upsert no Planner
 * ═══════════════════════════════════════════════════════════════════════════
 */

const Radar = (() => {
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CONFIGURAÇÃO DO RADAR (ajustável)
  // ═══════════════════════════════════════════════════════════════════════════
  
  const RADAR_CFG = {
    ENABLED: true,

    // ─────────────────────────────────────────────────────────────────────────
    // CAMPOS DE NEGÓCIO (Tabelas A e B)
    // ─────────────────────────────────────────────────────────────────────────
    CAMPOS_NEGOCIO: {
      tabela_a: ['prioridade', 'sla_dias', 'bucket_override'],
      tabela_b: ['checklist']
    },

    // ─────────────────────────────────────────────────────────────────────────
    // AGENDAMENTO
    // ─────────────────────────────────────────────────────────────────────────
    TRIGGERS: {
      HORA_IMEDIATO: 20,
      HORA_CONFERENCIA: 6,
    },

    // ─────────────────────────────────────────────────────────────────────────
    // TABELAS DE CONFIGURAÇÃO (Aba DADOS)
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
    // CONFIGURAÇÃO POR ABA (Dinâmica)
    // ─────────────────────────────────────────────────────────────────────────
    ABAS: {
      'DEMANDAS DIVERSAS🔧': {
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
      },
      
      'SEGUROS🛡️': {
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
      },
      
      'INTERNALIZADO🎯': {
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
      },

      'EM ANALISE📊': {
        ABA: 'EM ANALISE📊',
        START_ROW: 4,
        COLUNAS_LOOKUP: { categoria: 'J', subcategoria: 'K' },
        KEY_STRATEGY: ['G', 'ABA_LINHA'],
        TEMPLATES: {
          titulo: '{B} - {E} - {J}-{K} - {G}',
          description: 'OP Vinculada: {D}\nDOCSFLOW: {G}\nValor: {E}'
        },
        valores_fixos: {}
      },

      'CADASTROS🧑‍💻': {
        ABA: 'CADASTROS🧑‍💻',
        START_ROW: 4,
        COLUNAS_LOOKUP: { categoria: 'F', subcategoria: 'G' },
        KEY_STRATEGY: ['E', 'ABA_LINHA'],
        TEMPLATES: {
          titulo: '{B} - Atualização CAD - {C} - {E}',
          description: 'DOCSFLOW: {E}\nResponsável: {D}'
        },
        valores_fixos: {}
      }
    },

    // ─────────────────────────────────────────────────────────────────────────
    // E-MAIL
    // ─────────────────────────────────────────────────────────────────────────
    EMAIL: {
      DESTINATARIO: 'jefferson.santos@basa.com.br',
      ASSUNTO: 'RADAR_PLANNER',
      FORMATO: 'JSON'
    },

    SAIDA: {
      MAX_ITENS: 50,
      LOG_PERFORMANCE: true
    },

    CONFERENCIA: {
      MIN_PRIORIDADE: 2, // Apenas prioridade, pois não há calculo de data
      MAX_ITENS: 25,
    },

    PRIORITY_MAP: {
      0: 1, 1: 1,
      2: 3, 3: 3, 4: 3,
      5: 5, 6: 5, 7: 5,
      8: 9, 9: 9, 10: 9
    },

    BUCKET_MAP: {
      valores_validos: [
        'LIBERAÇÕES', 'CADASTROS', 'PROJETOS', 'TAREFAS PENDENTES'
      ]
    }
  };

  // ═══════════════════════════════════════════════════════════════════════════
  // API
  // ═══════════════════════════════════════════════════════════════════════════

  function runImediato() { return _run({ mode: 'IMEDIATO' }); }
  function runConferencia() { return _run({ mode: 'CONFERENCIA' }); }

  function setupTriggers() {
    _clearRadarTriggers_();
    ScriptApp.newTrigger('RADAR_runImediato').timeBased().everyDays(1).atHour(RADAR_CFG.TRIGGERS.HORA_IMEDIATO).create();
    ScriptApp.newTrigger('RADAR_runConferencia').timeBased().everyDays(1).atHour(RADAR_CFG.TRIGGERS.HORA_CONFERENCIA).create();
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // CORE
  // ═══════════════════════════════════════════════════════════════════════════

  function _run({ mode }) {
    if (!RADAR_CFG.ENABLED) return { success: true, data: { skipped: true } };

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

    return { success: true, data: { enviados: payload.itens.length, logPorAba: logResumo } };
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
      const keyDescription = _buildKeyDescription_({ aba: configAba.ABA, linha: row.linha, keyUsed: key });
      
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
        titulo, bucket, assignedTo, prioridade, priorityPlanner, 
        sla_dias: slaDias, // Enviamos apenas os dias
        checklistTemplate: checklistTemplate || null, checklistItens,
        dadosPlanilha // Contém { A_col: '...', B_col: '...' }
      });
    }
    return itens;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // AUXILIARES
  // ═══════════════════════════════════════════════════════════════════════════

  function _filtrarConferencia_(itens) {
    // ⚠️ Conferência agora só olha Prioridade (Data é calculada no PA)
    return itens.filter(it => {
      const prioOk = (it.prioridade ?? 999) <= RADAR_CFG.CONFERENCIA.MIN_PRIORIDADE;
      return prioOk;
    });
  }

  function _montarPayloadEnvio_(itens, mode) {
    const max = (mode === 'CONFERENCIA') ? RADAR_CFG.CONFERENCIA.MAX_ITENS : RADAR_CFG.SAIDA.MAX_ITENS;
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
    if (!sheetDados) throw new Error(`RADAR: aba DADOS não encontrada`);
    
    return { 
      configsA: _readTabelaA_(sheetDados, RADAR_CFG.DADOS.TABELA_A),
      templatesChecklist: _readTabelaB_(sheetDados, RADAR_CFG.DADOS.TABELA_B)
    };
  }
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
  
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    
    const aba = r[idxABA];
    const cat = r[idxCat];
    
    if (!aba || !cat) continue;
    
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
      
      Object.keys(map).forEach(k => {
        map[k] = map[k].sort((a, b) => a.ordem - b.ordem).map(x => x.item);
      });
      
      return map;
    }

  function _lookupConfig_(configsA, configAba, cat, sub) {
    return configsA.find(c => c.ABA === configAba.ABA && c.Categoria === cat && c.Subcategoria === sub) ||
           configsA.find(c => c.ABA === configAba.ABA && c.Categoria === cat && c.Subcategoria === '*');
  }

  // --- HELPERS E PARSERS ---
  function _indexHeaders_(h, req) {
    const idx = {};
    req.forEach(n => { if(h.indexOf(n)===-1) throw new Error(`Header ausente: ${n}`); idx[n]=h.indexOf(n); });
    return idx;
  }
  function _colLetterToNumber_(l) {
    if (!l) return 0;
    let n=0, s=String(l).toUpperCase();
    for(let i=0;i<s.length;i++) n=n*26+(s.charCodeAt(i)-64);
    return n;
  }
  function _numberToColLetter_(n) {
    let s='', x=n;
    while(x>0) { s=String.fromCharCode(65+(x-1)%26)+s; x=Math.floor((x-1)/26); }
    return s||'A';
  }
  function _resolveMaxCol_(cfg) {
    let max = 11;
    const p = (l) => { const n=_colLetterToNumber_(l); if(n>max) max=n; };
    
    // Calcula max col baseado nas Lookups e na Estratégia de Chave
    p(cfg.COLUNAS_LOOKUP?.categoria); 
    p(cfg.COLUNAS_LOOKUP?.subcategoria);
    
    // Adiciona colunas usadas na KEY_STRATEGY
    if (Array.isArray(cfg.KEY_STRATEGY)) {
      cfg.KEY_STRATEGY.forEach(k => {
        if (k !== 'ABA_LINHA') p(k);
      });
    }

    const txt = (cfg.TEMPLATES?.titulo||'') + (cfg.TEMPLATES?.description||'');
    let m; const re=/\{([A-Z]+)\}/g;
    while((m=re.exec(txt))!==null) p(m[1]);
    return max;
  }

  function _toIntOrNull_(v) { return (v==null||v==='') ? null : Math.trunc(Number(v)); }

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
    return t.length>255 ? t.substring(0,252)+'...' : t;
  }
  function _buildDescription_(cfg, row, keyDesc) {
    return keyDesc + '\n\n' + _parseTemplate_(cfg.TEMPLATES.description, row);
  }
  function _parseTemplate_(tpl, row) {
    return (tpl||'').replace(/\{([A-Z]+)\}/g, (m,l) => {
      const key = l + '_col';
      return row[key] != null ? String(row[key]).trim() : '';
    });
  }
  function _clearRadarTriggers_() {
    ScriptApp.getProjectTriggers().forEach(t => {
      if(['RADAR_runImediato','RADAR_runConferencia'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
    });
  }

  return { runImediato, runConferencia, setupTriggers };
})();

function RADAR_runImediato() { return Radar.runImediato(); }
function RADAR_runConferencia() { return Radar.runConferencia(); }
function RADAR_setupTriggers() { return Radar.setupTriggers(); }