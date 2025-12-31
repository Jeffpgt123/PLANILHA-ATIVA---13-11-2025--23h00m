/**
 * RADAR → PLANNER (mínimo, sem over-engineering)
 *
 * VERSÃO: FASE 1 - MVP ULTRA-MÍNIMO (SIMPLIFICADO)
 * - E-mail destino configurável
 * - Apenas categorias REAIS da planilha
 * - Buckets configurados: TAREFAS PENDENTES
 *
 * Dependências: nenhuma (só SpreadsheetApp + MailApp).
 * Integração: Power Automate/Flow lê o e-mail e faz upsert no Planner.
 *
 * Premissas do seu ambiente:
 * - Aba de origem: DEMANDAS DIVERSAS🔧
 * - SITUAÇÃO / DETALHE: colunas G/H
 * - DATA_CONCLUSAO: coluna J
 * - TIMESTAMP: coluna K
 * - DOCSFLOW: coluna F
 */

const Radar = (() => {
  
  // ═════════════════════════════════════════════════════════════════════════════
  // ███████╗ CONFIGURAÇÕES - EDITE AQUI ███████╗
  // ═════════════════════════════════════════════════════════════════════════════
  
  const CONFIG = {
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 1. CONFIGURAÇÕES GERAIS                                                 │
    // └─────────────────────────────────────────────────────────────────────────┘
    
    ENABLED: true, // ⚠️ false = desabilita RADAR completamente
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 2. HORÁRIOS DE EXECUÇÃO AUTOMÁTICA                                      │
    // └─────────────────────────────────────────────────────────────────────────┘
    
    TRIGGERS: {
      HORA_IMEDIATO: 20,     // Execução principal: 20h (8 PM)
      HORA_CONFERENCIA: 6,   // Execução crítica: 6h (6 AM)
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 3. ABA E COLUNAS DA PLANILHA                                            │
    // └─────────────────────────────────────────────────────────────────────────┘
    
    ORIGEM: {
      ABA: 'DEMANDAS DIVERSAS🔧', // ⚠️ Nome exato (com emoji)
      START_ROW: 4,                // ⚠️ Linha inicial dos dados
      
      COLS: {
        CNPJ: 'A',           // Coluna A: CNPJ
        CLIENTE: 'B',        // Coluna B: Nome cliente
        TELEFONE: 'C',       // Coluna C: Telefone
        CONTATO: 'D',        // Coluna D: Pessoa contato
        PRODUTO: 'E',        // Coluna E: Produto/serviço
        DOCSFLOW: 'F',       // Coluna F: Número processo (chave única)
        SITUACAO: 'G',       // Coluna G: CATEGORIA ⚠️ OBRIGATÓRIO
        DETALHE: 'H',        // Coluna H: SUBCATEGORIA ⚠️ OBRIGATÓRIO
        COMENTARIO: 'I',     // Coluna I: Comentários
        DATA_CONCLUSAO: 'J', // Coluna J: Data conclusão (vazio=aberto)
        TIMESTAMP: 'K',      // Coluna K: Data criação ⚠️ OBRIGATÓRIO
      }
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 4. E-MAIL DE SAÍDA (Power Automate)                                     │
    // └─────────────────────────────────────────────────────────────────────────┘
    
    SAIDA: {
      EMAIL_TO: 'jefferson.santos@basa.com.br', // ✅ E-mail configurado
      SUBJECT_IMEDIATO: 'RADAR_PLANNER',
      SUBJECT_CONFERENCIA: 'RADAR_PLANNER_CHECK',
      MAX_ITENS_POR_ENVIO: 60,
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 5. CRITÉRIOS CONFERÊNCIA (Execução 6h)                                 │
    // └─────────────────────────────────────────────────────────────────────────┘
    
    CONFERENCIA: {
      MIN_URGENCIA: 7,      // Envia se urgência >= 7
      VENCE_EM_DIAS: 0,     // Envia se vence hoje ou antes
      MAX_ITENS: 25,        // Top 25 mais críticos
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 7. FORMATO DO TÍTULO DA TAREFA NO PLANNER                              │
    // └─────────────────────────────────────────────────────────────────────────┘
    //
    // ⚠️ EDITE AQUI! Defina quais campos aparecerão no título e em que ordem
    //
    // Campos disponíveis:
    // - {key}: Identificador único (ex: RADAR:DEMANDAS DIVERSAS:LINHA:14)
    // - {cnpj}: CNPJ do cliente
    // - {cliente}: Nome do cliente
    // - {telefone}: Telefone
    // - {contato}: Pessoa de contato
    // - {produto}: Produto/serviço
    // - {docsflow}: Número do processo
    // - {categoria}: Situação (ex: A SOLICITAR)
    // - {subcategoria}: Detalhe (ex: AGUARDANDO DOCS)
    // - {comentario}: Comentários
    //
    // EXEMPLOS:
    // '{key} | {cliente} | {categoria} / {subcategoria}'
    // '{docsflow} - {cliente} ({categoria})'
    // '{cliente} | {cnpj} | {categoria}'
    //
    TITULO: {
      FORMATO: '{cliente} - {categoria} / {subcategoria} - DOCS {docsflow}',
      MAX_CLIENTE_CHARS: 80, // Limite de caracteres para nome do cliente
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 8. REGRAS SLA (CATEGORIAS REAIS DA PLANILHA)                           │
    // └─────────────────────────────────────────────────────────────────────────┘
    //
    // ⚠️ EDITE AQUI! Adicione novas categorias conforme aparecem na planilha
    //
    // FORMATO:
    // 'CATEGORIA|SUBCATEGORIA': { SLA_dias: X, Prioridade: Y, Urgencia: Z },
    //
    // VALORES:
    // - SLA_dias: Prazo em dias (1, 2, 3, 5, 7, 10...)
    // - Prioridade: 1 (alta/vermelha), 5 (média/amarela), 9 (baixa/sem cor)
    // - Urgencia: 0-10 (quanto maior, mais urgente)
    //
    // ═════════════════════════════════════════════════════════════════════════
    
    REGRAS_SLA: {
      
      // ═══ CATEGORIA: A SOLICITAR ═══
      
      'A SOLICITAR|CONFEC AUTORIZAÇÃO': { 
        SLA_dias: 5,       // Vence em 5 dias
        Prioridade: 1,     // Alta (vermelha)
        Urgencia: 8        // Muito urgente
      },
      
      // Fallback: qualquer outra subcategoria de "A SOLICITAR"
      'A SOLICITAR|*': { 
        SLA_dias: 7,       // Vence em 7 dias (padrão)
        Prioridade: 5,     // Média (amarela)
        Urgencia: 5        // Média
      },
      
      // ═══ ADICIONE NOVAS CATEGORIAS AQUI ⬇️ ═══
      //
      // Quando aparecer nova categoria/subcategoria na planilha, adicione assim:
      //
      // 'NOVA_CATEGORIA|NOVA_SUBCATEGORIA': { 
      //   SLA_dias: 7, 
      //   Prioridade: 5, 
      //   Urgencia: 6 
      // },
      //
      // 'NOVA_CATEGORIA|*': { // Fallback para categoria inteira
      //   SLA_dias: 10, 
      //   Prioridade: 5, 
      //   Urgencia: 5 
      // },
      
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 8. MAPEAMENTO DE BUCKETS DO PLANNER                                    │
    // └─────────────────────────────────────────────────────────────────────────┘
    //
    // ⚠️ IMPORTANTE: Nomes devem ser EXATAMENTE iguais aos buckets no Planner
    //
    // BUCKETS ATUAIS NO PLANNER:
    // - TAREFAS PENDENTES (único configurado atualmente)
    //
    // ═════════════════════════════════════════════════════════════════════════
    
    BUCKET_MAP: {
      
      // ─── POR SUBCATEGORIA (prioridade ALTA) ───
      subcategoria: {
        'CONFEC AUTORIZAÇÃO': 'TAREFAS PENDENTES',
        
        // ═══ ADICIONE NOVAS SUBCATEGORIAS AQUI ⬇️ ═══
        // 'NOVA_SUBCATEGORIA': 'TAREFAS PENDENTES',
      },
      
      // ─── POR CATEGORIA (prioridade MÉDIA - fallback) ───
      categoria: {
        'A SOLICITAR': 'TAREFAS PENDENTES',
        
        // ═══ ADICIONE NOVAS CATEGORIAS AQUI ⬇️ ═══
        // 'NOVA_CATEGORIA': 'TAREFAS PENDENTES',
      },
      
      // ─── BUCKET DEFAULT (se não encontrar em nenhum mapa) ───
      default: 'TAREFAS PENDENTES'
    },
    
    
    // ┌─────────────────────────────────────────────────────────────────────────┐
    // │ 8. CONFIGURAÇÕES AVANÇADAS (Não mexer)                                 │
    // └─────────────────────────────────────────────────────────────────────────┘
    
    DADOS: {
      ABA: 'DADOS',
      TABELA_A_RANGE: 'A1:G1', // FASE 2
      TABELA_B_RANGE: 'A1:D1', // FASE 2
    },
    
    KEY_STRATEGY: ['DOCSFLOW', 'ABA_LINHA'],
    
  };
  
  // ═════════════════════════════════════════════════════════════════════════════
  // ███████╗ FIM DAS CONFIGURAÇÕES ███████╗
  // ═════════════════════════════════════════════════════════════════════════════
  //
  // Daqui para baixo é CÓDIGO FUNCIONAL - não precisa mexer
  //
  // ═════════════════════════════════════════════════════════════════════════════

  // =========================
  // FUNÇÕES PÚBLICAS (TRIGGERS)
  // =========================
  
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
      .atHour(CONFIG.TRIGGERS.HORA_IMEDIATO)
      .create();

    ScriptApp.newTrigger('RADAR_runConferencia')
      .timeBased()
      .everyDays(1)
      .atHour(CONFIG.TRIGGERS.HORA_CONFERENCIA)
      .create();
  }

  // =========================
  // CORE - LÓGICA PRINCIPAL
  // =========================
  
  function _run({ mode }) {
    if (!CONFIG.ENABLED) return { success: true, data: { skipped: true, reason: 'RADAR_DISABLED' } };

    const ss = SpreadsheetApp.getActive();
    const abaOrigem = ss.getSheetByName(CONFIG.ORIGEM.ABA);
    if (!abaOrigem) throw new Error(`RADAR: aba origem não encontrada: ${CONFIG.ORIGEM.ABA}`);

    const { regrasA, itensChecklistPorTemplate } = _loadRegrasEApoios_(ss);

    const rows = _readOrigem_(abaOrigem);
    const candidatos = _buildCandidatos_(rows, regrasA, itensChecklistPorTemplate);

    const selecionados = (mode === 'CONFERENCIA')
      ? _filtrarConferencia_(candidatos)
      : candidatos;

    const payload = _montarPayloadEnvio_(selecionados, mode);

    if (!payload.itens.length) {
      console.log(`RADAR [${mode}]: SEM ITENS para enviar`);
      return { success: true, data: { mode, enviados: 0, motivo: 'SEM_ITENS' } };
    }

    _enviarEmailJSON_(payload, mode);

    console.log(`RADAR [${mode}]: E-mail enviado com ${payload.itens.length} itens`);
    return { success: true, data: { mode, enviados: payload.itens.length } };
  }

  function _readOrigem_(sheet) {
    const startRow = CONFIG.ORIGEM.START_ROW;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return [];

    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 11);
    const values = range.getValues();

    return values.map((r, idx) => ({
      linha: startRow + idx,
      A_cnpj: r[0],
      B_cliente: r[1],
      C_telefone: r[2],
      D_contato: r[3],
      E_produto: r[4],
      F_docsflow: r[5],
      G_situacao: r[6],
      H_detalhe: r[7],
      I_comentario: r[8],
      J_dataConclusao: r[9],
      K_timestamp: r[10],
    }));
  }

  function _buildCandidatos_(rows, regrasA, itensChecklistPorTemplate) {
    const itens = [];

    for (const row of rows) {
      // Filtro: sem data conclusão + com situacao e detalhe
      if (row.J_dataConclusao) continue;
      if (!row.G_situacao || !row.H_detalhe) continue;

      const categoria = String(row.G_situacao).trim();
      const subcategoria = String(row.H_detalhe).trim();

      const regra = _lookupRegra_(regrasA, categoria, subcategoria);

      const slaDias = _toIntOrNull_(regra?.SLA_dias);
      const prioridade = _toIntOrNull_(regra?.Prioridade);
      const urgencia = _toIntOrNull_(regra?.Urgencia);
      const bucketOverride = (regra?.BucketOverride ?? '').toString().trim();
      const checklistTemplate = (regra?.ChecklistTemplate ?? '').toString().trim();

      const ts = _parseDate_(row.K_timestamp);
      const dueDate = _calcDueDate_(ts, slaDias, row.K_timestamp);

      const missingParams = [];
      if (slaDias === null) missingParams.push('SLA_dias');
      if (prioridade === null) missingParams.push('Prioridade');
      if (urgencia === null) missingParams.push('Urgencia');

      const bucket = bucketOverride || _resolveBucket_(categoria, subcategoria);

      const checklistItens = checklistTemplate
        ? (itensChecklistPorTemplate[checklistTemplate] || [])
        : [];

      const key = _buildKey_({
        aba: CONFIG.ORIGEM.ABA,
        docsflow: row.F_docsflow,
        linha: row.linha
      });

      const titulo = _buildTitulo_(row, key);

      const descricao = _buildDescricao_({
        row,
        key,
        bucket,
        slaDias,
        prioridade,
        urgencia,
        dueDate,
        missingParams,
        checklistTemplate
      });

      itens.push({
        key,
        origem: { 
          aba: _sanitizeForJSON_(CONFIG.ORIGEM.ABA) || 'ABA',
          linha: row.linha 
        },
        titulo,
        descricao,
        bucket,
        prioridade,
        urgencia,
        sla_dias: slaDias,
        dueDateISO: dueDate ? dueDate.toISOString() : null,
        checklistTemplate: checklistTemplate || null,
        checklistItens,
        dadosPlanilha: {
          cnpj: _sanitizeForJSON_(row.A_cnpj) || null,
          cliente: _sanitizeForJSON_(row.B_cliente) || null,
          telefone: _sanitizeForJSON_(row.C_telefone) || null,
          contato: _sanitizeForJSON_(row.D_contato) || null,
          produto: _sanitizeForJSON_(row.E_produto) || null,
          docsflow: _sanitizeForJSON_(row.F_docsflow) || null,
          situacao: _sanitizeForJSON_(categoria),
          detalhe: _sanitizeForJSON_(subcategoria),
          comentario: _sanitizeForJSON_(row.I_comentario) || null,
          timestamp: row.K_timestamp ? row.K_timestamp.toString() : null,
        },
        flags: {
          missingParams: missingParams.length ? missingParams : null,
          fonteRegra: regra ? 'TABELA_A' : 'HARDCODED',
        }
      });
    }

    // Ordenação: urgência desc, prioridade asc, dueDate asc
    itens.sort((a, b) => {
      const ua = (a.urgencia ?? -1), ub = (b.urgencia ?? -1);
      if (ua !== ub) return ub - ua;

      const pa = (a.prioridade ?? 999), pb = (b.prioridade ?? 999);
      if (pa !== pb) return pa - pb;

      const da = a.dueDateISO ? Date.parse(a.dueDateISO) : Infinity;
      const db = b.dueDateISO ? Date.parse(b.dueDateISO) : Infinity;
      return da - db;
    });

    return itens;
  }

  function _filtrarConferencia_(itens) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const limite = new Date(hoje);
    limite.setDate(limite.getDate() + CONFIG.CONFERENCIA.VENCE_EM_DIAS);

    const filtrados = itens.filter(item => {
      const urg = item.urgencia ?? 0;
      if (urg < CONFIG.CONFERENCIA.MIN_URGENCIA) return false;

      if (item.dueDateISO) {
        const due = new Date(item.dueDateISO);
        due.setHours(0, 0, 0, 0);
        return due <= limite;
      }
      return true;
    });

    return filtrados.slice(0, CONFIG.CONFERENCIA.MAX_ITENS);
  }

  function _montarPayloadEnvio_(itens, mode) {
    const max = (mode === 'CONFERENCIA')
      ? CONFIG.CONFERENCIA.MAX_ITENS
      : CONFIG.SAIDA.MAX_ITENS_POR_ENVIO;

    const selecionados = itens.slice(0, max);

    return {
      meta: {
        mode,
        geradoEm: new Date().toISOString(),
        origem: _sanitizeForJSON_(CONFIG.ORIGEM.ABA) || 'ABA',
        totalSelecionado: selecionados.length,
        keyStrategy: CONFIG.KEY_STRATEGY
      },
      itens: selecionados
    };
  }

  function _enviarEmailJSON_(payload, mode) {
    const subject = (mode === 'CONFERENCIA')
      ? CONFIG.SAIDA.SUBJECT_CONFERENCIA
      : CONFIG.SAIDA.SUBJECT_IMEDIATO;

    const body = JSON.stringify(payload, null, 2);

    // ⚠️ CRÍTICO: enviar como texto plano (não HTML)
    MailApp.sendEmail({
      to: CONFIG.SAIDA.EMAIL_TO,
      subject: subject,
      body: body,
      noReply: true,
      // ADICIONAR ESTAS LINHAS:
      htmlBody: '', // Força body vazio em HTML
      mimeType: 'text/plain' // Força texto plano
    });
  }

  // =========================
  // REGRAS (TABELA A/B)
  // =========================
  
  function _loadRegrasEApoios_(ss) {
    // FASE 1: Tabelas A/B desabilitadas - usa fallbacks hardcoded
    console.log('RADAR [MVP]: Tabelas A/B desabilitadas - usando CONFIG.REGRAS_SLA');
    
    return {
      regrasA: [],
      itensChecklistPorTemplate: {}
    };
  }

  function _lookupRegra_(regrasA, categoria, subcategoria) {
    const encontrado = regrasA.find(r => r.Categoria === categoria && r.Subcategoria === subcategoria);
    
    if (!encontrado) {
      return _getRegraHardcoded_(categoria, subcategoria);
    }
    
    return encontrado;
  }
  
  function _getRegraHardcoded_(categoria, subcategoria) {
    // Usa CONFIG.REGRAS_SLA configurado no topo
    const chaveExata = `${categoria}|${subcategoria}`;
    if (CONFIG.REGRAS_SLA[chaveExata]) {
      return {
        Categoria: categoria,
        Subcategoria: subcategoria,
        ...CONFIG.REGRAS_SLA[chaveExata],
        BucketOverride: '',
        ChecklistTemplate: ''
      };
    }
    
    // Fallback por categoria
    const chaveFallback = `${categoria}|*`;
    if (CONFIG.REGRAS_SLA[chaveFallback]) {
      return {
        Categoria: categoria,
        Subcategoria: subcategoria,
        ...CONFIG.REGRAS_SLA[chaveFallback],
        BucketOverride: '',
        ChecklistTemplate: ''
      };
    }
    
    // Fallback geral
    return {
      Categoria: categoria,
      Subcategoria: subcategoria,
      SLA_dias: 7,
      Prioridade: 5,
      Urgencia: 5,
      BucketOverride: '',
      ChecklistTemplate: ''
    };
  }

  // =========================
  // HELPERS
  // =========================
  
  function _indexHeaders_(headers, required) {
    const idx = {};
    for (const name of required) {
      const pos = headers.indexOf(name);
      if (pos === -1) {
        throw new Error(`RADAR: header obrigatório ausente na tabela: ${name}`);
      }
      idx[name] = pos;
    }
    return idx;
  }

  function _sanitizeForJSON_(str) {
    if (!str) return str;
    return String(str)
      .replace(/[\u0000-\u001F\u007F-\u009F]/g, '') // Remove caracteres de controle
      .replace(/[\uD800-\uDFFF]/g, '') // Remove emojis/símbolos especiais
      .trim();
  }

  function _buildKey_({ aba, docsflow, linha }) {
    // Sanitiza nome da aba (remove emojis para JSON seguro)
    const abaSanitizada = _sanitizeForJSON_(aba) || 'ABA';
    
    for (const s of CONFIG.KEY_STRATEGY) {
      if (s === 'DOCSFLOW') {
        const df = _sanitizeForJSON_(docsflow);
        if (df) return `RADAR:${abaSanitizada}:DOCSFLOW:${df}`;
      }
      if (s === 'ABA_LINHA') {
        if (aba && linha) return `RADAR:${abaSanitizada}:LINHA:${linha}`;
      }
    }
    return `RADAR:${abaSanitizada}:LINHA:${linha}`;
  }

  function _buildTitulo_(row, key) {
    // Buscar formato configurado
    const formato = CONFIG.TITULO.FORMATO;
    const maxCliente = CONFIG.TITULO.MAX_CLIENTE_CHARS;
    
    // Mapa de campos disponíveis (sanitizados)
    const campos = {
      key: key,
      cnpj: _sanitizeForJSON_(row.A_cnpj) || '',
      cliente: (_sanitizeForJSON_(row.B_cliente) || '').slice(0, maxCliente),
      telefone: _sanitizeForJSON_(row.C_telefone) || '',
      contato: _sanitizeForJSON_(row.D_contato) || '',
      produto: _sanitizeForJSON_(row.E_produto) || '',
      docsflow: _sanitizeForJSON_(row.F_docsflow) || '',
      categoria: _sanitizeForJSON_(row.G_situacao) || '',
      subcategoria: _sanitizeForJSON_(row.H_detalhe) || '',
      comentario: _sanitizeForJSON_(row.I_comentario) || '',
    };
    
    // Substituir placeholders {campo} pelo valor correspondente
    let titulo = formato;
    for (const [campo, valor] of Object.entries(campos)) {
      titulo = titulo.replace(new RegExp(`\\{${campo}\\}`, 'g'), valor);
    }
    
    return titulo;
  }

  function _buildDescricao_({ row, key, bucket, slaDias, prioridade, urgencia, dueDate, missingParams, checklistTemplate }) {
    const parts = [];
    parts.push(`KEY: ${key}`);
    parts.push(`BUCKET: ${bucket}`);
    parts.push(`PRIORIDADE: ${prioridade ?? 'N/A'}`);
    parts.push(`URGENCIA: ${urgencia ?? 'N/A'}`);
    parts.push(`SLA_dias: ${slaDias ?? 'N/A'}`);
    parts.push(`DUE: ${dueDate ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'N/A'}`);
    if (checklistTemplate) parts.push(`CHECKLIST_TEMPLATE: ${checklistTemplate}`);
    if (missingParams?.length) parts.push(`MISSING_PARAMS: ${missingParams.join(', ')}`);

    parts.push('---');
    parts.push(`DOCSFLOW: ${row.F_docsflow || ''}`);
    parts.push(`CNPJ: ${row.A_cnpj || ''}`);
    parts.push(`CLIENTE: ${row.B_cliente || ''}`);
    parts.push(`SIT/DET: ${row.G_situacao || ''} / ${row.H_detalhe || ''}`);
    parts.push(`COMENTARIO: ${row.I_comentario || ''}`);

    return parts.join('\n');
  }

  function _calcDueDate_(ts, slaDias, rawTs) {
    // Regra principal: timestamp + SLA_dias
    if (ts && typeof slaDias === 'number') {
      const d = new Date(ts);
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() + slaDias);
      return d;
    }

    // Fallback: por idade (se faltar SLA)
    const t = ts || _parseDate_(rawTs);
    if (!t) return null;

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const base = new Date(t);
    base.setHours(0, 0, 0, 0);

    const idadeDias = Math.floor((hoje - base) / (1000 * 60 * 60 * 24));

    const due = new Date(hoje);
    if (idadeDias >= 15) due.setDate(due.getDate() + 0);
    else if (idadeDias >= 7) due.setDate(due.getDate() + 2);
    else if (idadeDias >= 3) due.setDate(due.getDate() + 5);
    else due.setDate(due.getDate() + 7);

    return due;
  }

  function _parseDate_(v) {
    if (!v) return null;
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;

    const s = String(v).trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const dd = Number(m[1]), mm = Number(m[2]) - 1, yyyy = Number(m[3]);
      const d = new Date(yyyy, mm, dd);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  }

  function _toIntOrNull_(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return Number.isFinite(n) ? Math.trunc(n) : null;
  }

  function _resolveBucket_(categoria, subcategoria) {
    // Usa CONFIG.BUCKET_MAP configurado no topo
    const subNorm = _normalize_(subcategoria);
    if (CONFIG.BUCKET_MAP.subcategoria[subNorm]) {
      return CONFIG.BUCKET_MAP.subcategoria[subNorm];
    }
    
    const catNorm = _normalize_(categoria);
    if (CONFIG.BUCKET_MAP.categoria[catNorm]) {
      return CONFIG.BUCKET_MAP.categoria[catNorm];
    }
    
    return CONFIG.BUCKET_MAP.default;
  }

  function _normalize_(s) {
    return String(s || '')
      .trim()
      .toUpperCase()
      .replace(/\s+/g, ' ');
  }

  function _clearRadarTriggers_() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      const h = t.getHandlerFunction();
      if (h === 'RADAR_runImediato' || h === 'RADAR_runConferencia') {
        ScriptApp.deleteTrigger(t);
      }
    });
  }

  // =========================
  // API PÚBLICA
  // =========================
  
  return {
    runImediato,
    runConferencia,
    setupTriggers,
  };
})();

// Funções globais (gatilhos chamam por nome)
function RADAR_runImediato() { return Radar.runImediato(); }
function RADAR_runConferencia() { return Radar.runConferencia(); }
function RADAR_setupTriggers() { return Radar.setupTriggers(); }