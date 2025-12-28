/**
 * RegrasNegocioV2 v1.1 — motor unificado de Regras de Negócio (Fase 1.1)
 * ----------------------------------------------------------------------
 * Fonte oficial das regras: REGRAS_DE_NEGOCIO (regrasConfig.js)
 * Sem eval: condições dirigidas pelo evento (aba/coluna/valor) + CondicaoRegistry
 * Ponto de chamada no pipeline: após onEditHandler, antes do Timestamp/Batch
 *
 * Novidades v1.1 (após revisão técnica):
 * - Guard/validação de evento (e/range) com logs leves
 * - CondicaoRegistry em _regraBateNoEvento (valor_igual, valor_diferente, valor_contem, valor_in[], valor_vazio, valor_nao_vazio)
 * - Normalização única do valor novo (ctx.valorNovoNorm)
 * - Logs estruturados por regra/ação e métricas de tempo
 */

const RegrasNegocioV2 = (function () {
  // ===== Utilidades internas =================================================
  const CACHE_TTL_MS = 5 * 60 * 1000; // 5 min
  const cache = { ts: 0, regras: null };

  // ===== Configurações de execução (modo enxuto) ============================
  const LOG_LEVEL = 'WARN'; // 'INFO' | 'WARN' | 'ERROR' | 'OFF'
  const PERF_LOG_THRESHOLD_MS = 80; // só loga INFO se a ação exceder esse tempo (ou se houver alteração)
  const COND_OPERATORS_ENABLED = {
    valor_igual: true,
    valor_in: true,
    valor_contem: false,
    valor_diferente: false,
    valor_vazio: false,
    valor_nao_vazio: false
  };

  const _now = () => Date.now();
  const _norm = (s) => (s || "").toString()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .trim().toUpperCase();

  const _colunaParaIndiceFallback = (letra) => {
    if (!letra) return null;
    if (typeof letra === 'number') return letra;
    letra = letra.toString().trim().toUpperCase();
    let num = 0;
    for (let i = 0; i < letra.length; i++) { num = num * 26 + (letra.charCodeAt(i) - 64); }
    return num;
  };
  const _c2i = (c) => {
    try { if (typeof Utils !== 'undefined' && typeof Utils.colunaParaIndice === 'function') return Utils.colunaParaIndice(c); } catch(_) {}
    return _colunaParaIndiceFallback(c);
  };
  // (Função _getTZ omitida para brevidade - sem alteração)
  const _getTZ = () => {
    try { if (typeof Utils !== 'undefined' && typeof Utils.getTimezone === 'function') return Utils.getTimezone(); } catch(_) {}
    try { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(_) { return 'America/Sao_Paulo'; }
  };
  const _logBatch = (payload) => {
    try { Logger?.registrarLogBatch?.(payload); } catch(_) { try{ console.log(JSON.stringify(payload)); }catch(_){} }
  };
  // Logging com nível (INFO/WARN/ERROR) e controle de volume
  const _levelOrder = { OFF:0, ERROR:1, WARN:2, INFO:3 };
  const _canLog = (level) => _levelOrder[LOG_LEVEL] >= _levelOrder[level];
  const _logIf = (level, payload) => { if (_canLog(level)) _logBatch(payload); };

  // ===== Carregamento de regras (fonte JS/DSL) ===============================
  function _getRegrasAtivas() {
    if (cache.regras && (_now() - cache.ts) < CACHE_TTL_MS) return cache.regras;
    let regras = [];
    try {
      if (typeof REGRAS_DE_NEGOCIO !== 'undefined' && Array.isArray(REGRAS_DE_NEGOCIO)) {
        regras = REGRAS_DE_NEGOCIO.filter(r => _norm(r?.status) === 'ATIVO');
      }
    } catch (err) {
      _logIf('ERROR', [["ERRO","RegrasNegocioV2/_getRegrasAtivas", err && err.message]]);
    }
    regras.sort((a, b) => (a?.prioridade ?? 9999) - (b?.prioridade ?? 9999));
    cache.regras = regras; cache.ts = _now();
    return regras;
  }

  // ===== Contexto do evento ==================================================
  function _montarCtx(e) {
    const sheet = e.range.getSheet();
    const aba = sheet.getName();
    const linha = e.range.getRow();
    const coluna = e.range.getColumn();
    const valorNovo = (typeof e.value !== 'undefined' ? e.value : (e.range.getDisplayValue() || '')).toString();
    return { e, sheet, aba, linha, coluna, valorNovo, valorNovoNorm: _norm(valorNovo) };
  }

  // ===== CondicaoRegistry (v1.1) ============================================
  function _condicoesBatem(ctx, p) {
    if (!p) return true;
    const v = ctx.valorNovoNorm;
    // valor_vazio / valor_nao_vazio
    if (COND_OPERATORS_ENABLED.valor_vazio && typeof p.valor_vazio !== 'undefined') {
      if (!!p.valor_vazio !== (v === '')) return false;
    }
    if (COND_OPERATORS_ENABLED.valor_nao_vazio && typeof p.valor_nao_vazio !== 'undefined') {
      if (!!p.valor_nao_vazio !== (v !== '')) return false;
    }
    // valor_igual / valor_diferente
    if (COND_OPERATORS_ENABLED.valor_igual && typeof p.valor_igual !== 'undefined') {
      if (_norm(p.valor_igual) !== v) return false;
    }
    if (COND_OPERATORS_ENABLED.valor_diferente && typeof p.valor_diferente !== 'undefined') {
      if (_norm(p.valor_diferente) === v) return false;
    }
    // valor_contem (substring)
    if (COND_OPERATORS_ENABLED.valor_contem && typeof p.valor_contem !== 'undefined') {
      if (!v.includes(_norm(p.valor_contem))) return false;
    }
    // valor_in (array de valores equivalentes)
    if (COND_OPERATORS_ENABLED.valor_in && Array.isArray(p.valor_in) && p.valor_in.length > 0) {
      const arr = p.valor_in.map(_norm);
      if (!arr.includes(v)) return false;
    }
    return true; // todas condições atendidas
  }

  // ===== Avaliação simples da regra =========================================
  function _regraBateNoEvento(ctx, regra) {
    if (!regra) return false;
    if (regra.abaAlvo && regra.abaAlvo !== ctx.aba) return false;
    if (Array.isArray(regra.colunasAlvo) && regra.colunasAlvo.length) {
      const idxs = regra.colunasAlvo.map(c => _c2i(c));
      if (!idxs.includes(ctx.coluna)) return false;
    }
    return _condicoesBatem(ctx, regra.parametros || {});
  }

  // ===== Action Registry =====================================================
  const ActionRegistry = {
    _map: Object.create(null),
    register: function (nome, fn) { this._map[nome] = fn; },
    get: function (nome) { return this._map[nome]; }
  };

  // -- Ação: setar_data_se_vazio ---------------------------------------------
  ActionRegistry.register('setar_data_se_vazio', function (ctx, regra) {
    const p = regra.parametros || {};
    const colDest = p.coluna_destino; if (!colDest) return { alterou: false };
    const colIdx = _c2i(colDest);
    const cel = ctx.sheet.getRange(ctx.linha, colIdx);
    const vazio = ((cel.getDisplayValue() || '').toString().trim() === '');
    if (!vazio) return { alterou: false };
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    cel.setValue(hoje);
    const formato = p.formato || (CONFIG?.SHEETS?.[ctx.aba]?.FORMATO) || 'dd/MM/yyyy';
    try { cel.setNumberFormat(formato); } catch (_) { }
    return { alterou: true };
  });

  // -- Ação: copiar_valor -----------------------------------------------------
  ActionRegistry.register('copiar_valor', function (ctx, regra) {
    const p = regra.parametros || {};
    const colDest = p.coluna_destino; if (!colDest) return { alterou: false };
    const destIdx = _c2i(colDest);
    const cel = ctx.sheet.getRange(ctx.linha, destIdx);

    if (typeof p.valor !== 'undefined') {
      if (p.valor === 'TODAY()') {
        const d = new Date(); d.setHours(0, 0, 0, 0);
        cel.setValue(d);
      } else {
        cel.setValue(p.valor);
      }
    } else if (p.coluna_origem) {
      const srcIdx = _c2i(p.coluna_origem);
      const v = ctx.sheet.getRange(ctx.linha, srcIdx).getValue();
      cel.setValue(v);
    }
    if (p.formato) { try { cel.setNumberFormat(p.formato); } catch (_) { } }
    return { alterou: true };
  });

  // -- Ação: limpar_celula ----------------------------------------------------
  ActionRegistry.register('limpar_celula', function (ctx, regra) {
    const p = regra.parametros || {};
    const destIdx = _c2i(p.coluna_destino);
    if (!destIdx) return { alterou: false };
    ctx.sheet.getRange(ctx.linha, destIdx).clearContent();
    return { alterou: true };
  });

  // -- Ação: aplicar_formato --------------------------------------------------
  ActionRegistry.register('aplicar_formato', function (ctx, regra) {
    const p = regra.parametros || {};
    const destIdx = _c2i(p.coluna_destino);
    if (!destIdx || !p.formato) return { alterou: false };
    ctx.sheet.getRange(ctx.linha, destIdx).setNumberFormat(p.formato);
    return { alterou: true };
  });

  // -- Ação: mostrar_mensagem -------------------------------------------------
  ActionRegistry.register('mostrar_mensagem', function (ctx, regra) {
    const texto = (regra.parametros && regra.parametros.texto) || regra.descricao || regra.id;
    _logIf('INFO', [["INFO", "RegrasNegocioV2/mensagem", `${ctx.aba}:${ctx.linha} → ${texto}`]]);
    return { alterou: false };
  });

  // -- AÇÃO NOVA: sincronizar_data_por_situacao (regra única I⟷M) ------------
  // █████████████████████████████████████████████████████████████████████████
  // ▶▶▶ INÍCIO DA MODIFICAÇÃO (ADIÇÃO DE LOGS)
  // █████████████████████████████████████████████████████████████████████████
  ActionRegistry.register('sincronizar_data_por_situacao', function (ctx, regra) {
    const p = regra.parametros || {};
    const colDest = p.coluna_destino; if (!colDest) return { alterou: false };
    const destIdx = _c2i(colDest);
    const cel = ctx.sheet.getRange(ctx.linha, destIdx);

    const ativadorNorm = _norm(p.valor_ativador || '');
    const formato = p.formato || (CONFIG?.SHEETS?.[ctx.aba]?.FORMATO) || 'dd/MM/yyyy';
    
    // --- [LOG 1: INÍCIO] ---
    // Prepara um payload de log para rastrear a execução
    const logPayload = {
      regraId: regra.id,
      gatilho: `${ctx.aba}!${Utils.colunaParaLetra(ctx.coluna)}${ctx.linha}`,
      valorNovo: ctx.valorNovo,
      valorAtivadorEsperado: p.valor_ativador,
      destino: `${p.coluna_destino}${ctx.linha}`,
      etapa: ""
    };
    _logIf('WARN', [["DEPURACAO", "sync_data", { ...logPayload, etapa: "Iniciando verificação da regra" }]]);
    // --- [FIM LOG 1] ---

    if (ctx.valorNovoNorm === ativadorNorm) {
      const policy = _norm(p.politica_quando_ativado || 'SETAR_HOJE_SE_VAZIO');

      // --- [LOG 2: VALOR ATIVADO] ---
      _logIf('WARN', [["DEPURACAO", "sync_data", { ...logPayload, etapa: `Valor ATIVADO. Política: ${policy}` }]]);
      // --- [FIM LOG 2] ---

      if (policy === 'PRESERVAR') return { alterou: false };
      if (policy === 'SETAR_HOJE') {
        const d = new Date(); d.setHours(0,0,0,0);
        cel.setValue(d); try { cel.setNumberFormat(formato); } catch(_){}
        return { alterou: true };
      }
      
      // SETAR_HOJE_SE_VAZIO (padrão)
      // --- [LOG 3: VERIFICAÇÃO CRÍTICA DE CÉLULA VAZIA] ---
      const valorDestinoRaw = cel.getDisplayValue();
      const valorDestinoStr = (valorDestinoRaw || '').toString();
      const valorDestinoTrim = valorDestinoStr.trim();
      const vazio = (valorDestinoTrim === '');

      const debugCheck = {
        valorNaCelulaK_Raw: valorDestinoRaw,
        valorNaCelulaK_Str: valorDestinoStr,
        valorNaCelulaK_Trim: valorDestinoTrim,
        // Informação crucial: código ASCII do primeiro caractere (se houver)
        asciiPrimeiroChar: valorDestinoStr.length > 0 ? valorDestinoStr.charCodeAt(0) : null,
        tamanhoAntesTrim: valorDestinoStr.length,
        tamanhoDepoisTrim: valorDestinoTrim.length,
        resultadoIsEmpty: vazio
      };
      _logIf('WARN', [["DEPURACAO_CHECK", "sync_data_CHECK", { ...logPayload, etapa: "Verificando política SETAR_HOJE_SE_VAZIO", check: debugCheck }]]);
      // --- [FIM LOG 3] ---
      
      if (!vazio) {
        // --- [LOG 4: FALHA] ---
        _logIf('WARN', [["DEPURACAO", "sync_data", { ...logPayload, etapa: "FALHA: Célula K não está vazia. Abortando." }]]);
        // --- [FIM LOG 4] ---
        return { alterou: false };
      }

      const d = new Date(); d.setHours(0,0,0,0);
      cel.setValue(d); try { cel.setNumberFormat(formato); } catch(_){}
      
      // --- [LOG 5: SUCESSO] ---
      _logIf('WARN', [["DEPURACAO", "sync_data", { ...logPayload, etapa: "SUCESSO: Data inserida em K." }]]);
      // --- [FIM LOG 5] ---
      
      return { alterou: true };
      
    } else {
      // --- [LOG 6: VALOR DESATIVADO] ---
      _logIf('WARN', [["DEPURACAO", "sync_data", { ...logPayload, etapa: "Valor DESATIVADO. Verificando política 'outros'." }]]);
      // --- [FIM LOG 6] ---
      
      const policy2 = _norm(p.politica_outros || 'LIMPAR');
      if (policy2 === 'PRESERVAR') return { alterou: false };
      cel.clearContent();
      return { alterou: true };
    }
  });
  // █████████████████████████████████████████████████████████████████████████
  // ▶▶▶ FIM DA MODIFICAÇÃO
  // █████████████████████████████████████████████████████████████████████████


// -- AÇÃO NOVA: carimbar_data_por_preenchimento ------------------------------
ActionRegistry.register('carimbar_data_por_preenchimento', function (ctx, regra) {
  const p = regra.parametros || {};
  const dest = p.coluna_destino; if (!dest) return { alterou: false };
  const destIdx = _c2i(dest); if (!destIdx || destIdx < 1) return { alterou: false };
  const cols = Array.isArray(p.colunas_monitoradas) ? p.colunas_monitoradas : [];
  if (cols.length === 0) return { alterou: false };
  const criterio = (p.criterio || 'QUALQUER').toString().trim().toUpperCase();
  const formato = p.formato || (CONFIG?.SHEETS?.[ctx.aba]?.FORMATO) || 'dd/MM/yyyy';
  const celDest = ctx.sheet.getRange(ctx.linha, destIdx);

  // coleta valores monitorados na linha
  const valores = cols.map(c => {
    const i = _c2i(c); if (!i || i < 1) return '';
    return (ctx.sheet.getRange(ctx.linha, i).getDisplayValue() || '').toString().trim();
  });

  const preenchidos = valores.filter(v => v !== '');
  const isPreenchido = (criterio === 'TODAS') ? (preenchidos.length === cols.length)
                                              : (preenchidos.length > 0);

  if (isPreenchido) {
    const pol = (p.politica_quando_preenchido || 'SETAR_HOJE_SE_VAZIO').toString().toUpperCase();
    if (pol === 'PRESERVAR') return { alterou: false };
    if (pol === 'SETAR_HOJE') {
      const d = new Date(); d.setHours(0,0,0,0);
      celDest.setValue(d); try { celDest.setNumberFormat(formato); } catch(_){}
      return { alterou: true };
    }
    // SETAR_HOJE_SE_VAZIO
    const vazio = ((celDest.getDisplayValue() || '').toString().trim() === '');
    if (!vazio) return { alterou: false };
    const d = new Date(); d.setHours(0,0,0,0);
    celDest.setValue(d); try { celDest.setNumberFormat(formato); } catch(_){}
    return { alterou: true };
  } else {
    const pol2 = (p.politica_quando_vazio || 'PRESERVAR').toString().toUpperCase();
    if (pol2 === 'LIMPAR') { celDest.clearContent(); return { alterou: true }; }
    return { alterou: false };
  }
});

  // -- AÇÃO NOVA: R2 — restaurar_produto_seguros ----------------------------
  ActionRegistry.register('restaurar_produto_seguros', function (ctx, regra) {
    const p = regra.parametros || {};
    const startRow = p.startRow || 4;
    const colProd = _c2i(p.coluna_produto || 'C');
    if (ctx.linha < startRow) return { alterou: false };
    if (ctx.coluna !== colProd) return { alterou: false };

    // Só atua se houve limpeza (valor anterior existia e o atual ficou vazio)
    const oldV = (typeof ctx.e.oldValue !== 'undefined') ? String(ctx.e.oldValue || '') : '';
    const newV = (typeof ctx.e.value !== 'undefined') ? String(ctx.e.value || '') : ctx.valorNovo;
    if (!oldV || newV) return { alterou: false };

    // Nota deve conter origem=ABA!linha
    const cel = ctx.sheet.getRange(ctx.linha, colProd);
    const nota = cel.getNote() || '';
    const m = nota.match(/(?:^|\b)origem=([^!]+)!([0-9]+)\b/i);
    if (!m) return { alterou: false };
    const origemAba = m[1], origemRow = parseInt(m[2], 10);

    const origemSh = Utils.getCachedSheet(origemAba);
    if (!origemSh) return { alterou: false };
    const colOrigemProdutos = _c2i(p.origem_coluna_produtos || 'J');
    const origemStr = String(origemSh.getRange(origemRow, colOrigemProdutos).getDisplayValue() || '');
    const lista = origemStr.split(/[;,\\n]/).map(s => s.trim()).filter(Boolean);
    const existe = lista.some(v => _norm(v) === _norm(oldV));
    if (!existe) return { alterou: false };

    // Restaura valor + realça + reanota
    BatchOperations.add('setValue', `'${ctx.aba}'!${cel.getA1Notation()}`, oldV);
    try { cel.setNote(`${p.nota_restauracao || ''}${p.nota_restauracao ? ' | ' : ''}${nota}`); } catch(_) {}
    try { ctx.sheet.getRange(ctx.linha, 1, 1, ctx.sheet.getLastColumn()).setBackground(p.cor_pendencia || '#FFF2CC'); } catch(_) {}
    return { alterou: true };
  });

// -- AÇÃO: R3 — proteger_edicao_sensivel (2 toques; cor garantida; restaura formato/visuais)
ActionRegistry.register('proteger_edicao_sensivel', function (ctx, regra) {
  const p = regra.parametros || {};
  const startRow = p.startRow || 4;
  const colSens = _c2i(p.coluna_sensivel || 'E');
  if (ctx.linha < startRow || ctx.coluna !== colSens) return { alterou: false };

  // Linha "preenchida"?
  const letras = Array.isArray(p.cols_linha_preenchida) ? p.cols_linha_preenchida : ['A','B','C','D','E','F'];
  const maxCol = Math.max.apply(null, letras.map(_c2i));
  const sheet = ctx.sheet;
  const rowRange = sheet.getRange(ctx.linha, 1, 1, sheet.getLastColumn());
  const vals = sheet.getRange(ctx.linha, 1, 1, maxCol).getValues()[0];
  const preenchida = letras.every(L => String(vals[_c2i(L)-1]).trim() !== '');
  if (!preenchida) return { alterou: false };

  const cel = sheet.getRange(ctx.linha, colSens);
  const cor = p.cor_pendencia || '#FFF2CC';

  // Estado persistente por linha/célula (usa Properties; Cache é volátil)
  const props = (typeof PropertiesService !== 'undefined') ? PropertiesService.getDocumentProperties() : null; // escopo do arquivo
  const key = ['psens', sheet.getSheetId(), ctx.linha, colSens].join(':');

  // Coerção pt/EN robusta: retorna número quando possível, senão null
  function coerceToNumberLocaleAware(raw, fmt) {
    if (raw == null) return null;
    const s = String(raw).trim();
    if (s === '') return null;
    // Se a célula tem formato numérico/moeda, tentar sempre converter
    const looksNumericFmt = /0[,.]00|#,#|R\$/i.test(String(fmt || ''));
    if (!looksNumericFmt && !/[0-9]/.test(s)) return null;

    // 1) normalizar: remover símbolos, preservar apenas dígitos e separadores
    let t = s.replace(/[^0-9,.\-]/g, '');
    // 2) se tem vírgula e ponto, assumir último como decimal
    if (t.includes(',') && t.includes('.')) {
      const lastSep = Math.max(t.lastIndexOf(','), t.lastIndexOf('.'));
      const dec = t[lastSep];
      t = t.replace(/[.,]/g, m => (m === dec ? '.' : ''));
      const n = Number(t);
      return isNaN(n) ? null : n;
    }
    // 3) se só vírgula, vírgula = decimal (pt-BR)
    if (t.includes(',') && !t.includes('.')) {
      t = t.replace(/\./g, '').replace(',', '.');
      const n = Number(t);
      return isNaN(n) ? null : n;
    }
    // 4) se só ponto, ponto = decimal
    if (!t.includes(',') && t.includes('.')) {
      // remover pontos de milhar múltiplos
      t = t.replace(/\.(?=.*\.)/g, '');
      const n = Number(t);
      return isNaN(n) ? null : n;
    }
    // 5) apenas dígitos
    const n = Number(t);
    return isNaN(n) ? null : n;
  }

  // Limpeza de pendência expirada: se existe estado e já passou o TTL, restaurar visuais no próximo onEdit
  try {
    const j = props && props.getProperty(key);
    if (j) {
      const st = JSON.parse(j);
      if (st && st.expireAt && Date.now() > st.expireAt) {
        if (Array.isArray(st.bg2D) && st.bg2D.length) {
          try { rowRange.setBackgrounds(st.bg2D); } catch (_) { rowRange.setBackground(null); }
        } else {
          try { rowRange.setBackground(null); } catch (_) {}
        }
        if (st.fmt) { try { cel.setNumberFormat(st.fmt); } catch (_) {} } // reaplica formato anterior
        try { cel.setNote(''); } catch (_) {}
        try { props.deleteProperty(key); } catch (_) {}
        // segue para lógica normal deste onEdit
      }
    }
  } catch (_) {}

  // 2 toques em 30s
  const ttl = (p.confirmacao_2_toques && Number(p.confirmacao_2_toques.ttlSegundos)) || 30;
  const state = (() => { try { return props && props.getProperty(key) ? JSON.parse(props.getProperty(key)) : null; } catch(_) { return null; } })();

  if (state && Date.now() <= state.expireAt) {
    // 2º toque dentro da janela: aceitar edição e restaurar visuais/formatos
    try { cel.setNote(''); } catch (_) {}
    try {
      if (Array.isArray(state.bg2D) && state.bg2D.length) rowRange.setBackgrounds(state.bg2D);
      else rowRange.setBackground(null);
    } catch (_) {}
    if (state.fmt) { try { cel.setNumberFormat(state.fmt); } catch (_) {} }
    try { props.deleteProperty(key); } catch (_) {}
    return { alterou: false };
  }

  // 1º toque: marcar pendência, guardar visuais e reverter valor se possível
  const fmtAtual = String(cel.getNumberFormat() || '');
  try {
    const bg2D = rowRange.getBackgrounds(); // 2D para restauração fiel
    props && props.setProperty(key, JSON.stringify({
      bg2D,
      fmt: fmtAtual,
      expireAt: Date.now() + Math.max(5, Math.min(300, ttl)) * 1000
    }));
  } catch (_) {}

  // Feedback visual e nota SEMPRE
  try { rowRange.setBackground(cor); } catch (_) {}
  const aviso = p.msg_aviso || 'Edição sensível.';
  try { cel.setNote(`${aviso} Edite novamente em até ${ttl}s para confirmar.`); } catch (_) {}

  // Reversão: só se oldValue existir; senão mantém valor digitado
  if (Object.prototype.hasOwnProperty.call(ctx.e, 'oldValue') && ctx.e.oldValue != null) {
    const old = ctx.e.oldValue;
    const coerced = coerceToNumberLocaleAware(old, fmtAtual);
    const back = (coerced == null ? String(old) : coerced);
    BatchOperations.add('setValue', `'${ctx.aba}'!${cel.getA1Notation()}`, back);
    // Reaplicar formato anterior (não força máscara nova)
    if (fmtAtual) { try { cel.setNumberFormat(fmtAtual); } catch (_) {} } // formatação garantida pela API
  }

  return { alterou: true };
});

// -- AÇÃO NOVA: validar_seguro_e_vigencia --------------------------------------
// Finalidade: Realizar validações complexas que dependem de múltiplas células (C, I, K)
// e garantir a precisão no cálculo de vigência temporal (L - K).
ActionRegistry.register('validar_seguro_e_vigencia', function (ctx, regra) {
  const p = regra.parametros || {};
  let alterou = false;
  const { sheet, linha } = ctx;
  const colIdx = (c) => _c2i(c); 
  const norm = (s) => _norm(s);

  // Helper visual de erro
  const _dispararErroVisual = (colLetra, msg) => {
    if (!colLetra) return false;
    const range = sheet.getRange(linha, colIdx(colLetra));
    range.setBackground('#FFE6E6');
    range.setNote(`${msg}`); 
    return true;
  };

  const _limparErroVisual = (colLetra) => {
    if (!colLetra) return false;
    const range = sheet.getRange(linha, colIdx(colLetra));
    if ((range.getBackground() || '').toUpperCase() === '#FFE6E6') range.setBackground(null);
    const nota = (range.getNote() || '').trim();
    // Limpa notas de erro padrão ou específicas de vigência
    if (nota.includes('Vigência Inválida') || nota.startsWith('ERRO:') || nota.startsWith('Inserir')) range.clearNote();
    return false;
  };

  // Helper para parser de data
  const _parseDataSegura = (valor) => {
    if (!valor) return null;
    if (valor instanceof Date) return valor; 
    const match = String(valor).match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
    if (match) {
      return new Date(match[3], match[2] - 1, match[1], 12, 0, 0); 
    }
    return null;
  };

  // Leitura de valores chave
  const prodVal = sheet.getRange(linha, colIdx(p.coluna_produto)).getDisplayValue();
  const exigeD = p.tipos_exigem_D.map(norm).includes(norm(prodVal));
  const statusVal = sheet.getRange(linha, colIdx(p.coluna_status)).getDisplayValue();
  // Regra Ativadora Global: Status PAGO/CONTRATADO E Produto da Lista Branca
  const regraAtivada = norm(statusVal) === norm(p.valor_ativador_L) && exigeD;

  // =========================================================================
  // 1. VALIDAÇÃO DE OBRIGATORIEDADE: OPERAÇÃO (D)
  // (Esta validação ocorre independente do Status, apenas pelo Produto)
  // =========================================================================
  const colOperacao = p.coluna_operacao;
  const operacaoVal = sheet.getRange(linha, colIdx(colOperacao)).getDisplayValue();

  if (exigeD) {
    if (!operacaoVal.trim()) {
      alterou = _dispararErroVisual(colOperacao, p.msg_erro_D) || alterou;
    } else {
      _limparErroVisual(colOperacao);
    }
  } else {
    _limparErroVisual(colOperacao);
  }

  // =========================================================================
  // 2. VALIDAÇÕES DE OBRIGATORIEDADE CONDICIONAL (F, G e L)
  // (Só ocorrem se Status = PAGO/CONTRATADO e Produto Exige)
  // =========================================================================
  
  // Colunas a validar
  const colValorSeg = p.coluna_valor_seguro;   // F
  const colNumProp = p.coluna_numero_proposta; // G
  const colDataFim = p.coluna_data_fim;        // L

  if (regraAtivada) {
    // A. Validação Coluna F (Valor Seguro)
    const valF = sheet.getRange(linha, colIdx(colValorSeg)).getDisplayValue();
    if (!valF.trim()) {
      alterou = _dispararErroVisual(colValorSeg, p.msg_erro_F) || alterou;
    } else {
      _limparErroVisual(colValorSeg);
    }

    // B. Validação Coluna G (Número Proposta)
    const valG = sheet.getRange(linha, colIdx(colNumProp)).getDisplayValue();
    if (!valG.trim()) {
      alterou = _dispararErroVisual(colNumProp, p.msg_erro_G) || alterou;
    } else {
      _limparErroVisual(colNumProp);
    }

    // C. Validação Coluna L (Vigência)
    const txtInicio = sheet.getRange(linha, colIdx(p.coluna_data_inicio)).getDisplayValue();
    const txtFim = sheet.getRange(linha, colIdx(colDataFim)).getDisplayValue();
    
    if (!txtFim.trim()) {
      alterou = _dispararErroVisual(colDataFim, p.msg_erro_L) || alterou;
    } else {
      const dtInicio = _parseDataSegura(txtInicio);
      const dtFim = _parseDataSegura(txtFim);
      let vigenciaInvalida = false;
      
      if (dtInicio && dtFim) {
        const umDia = 1000 * 60 * 60 * 24;
        const diffDias = Math.floor((dtFim - dtInicio) / umDia);
        const tolerancia = p.tolerancia_dias || 30;
        
        if (diffDias < (365 - tolerancia)) {
          vigenciaInvalida = true; 
        } else {
          const anosEstimados = Math.round(diffDias / 365);
          const diasAlvo = anosEstimados * 365;
          const desvio = Math.abs(diffDias - diasAlvo);
          
          if (desvio > tolerancia) {
            vigenciaInvalida = true;
          }
        }
      } else {
        vigenciaInvalida = true;
      }
      
      if (vigenciaInvalida) {
        alterou = _dispararErroVisual(colDataFim, p.msg_erro_L_vigencia) || alterou;
      } else {
        _limparErroVisual(colDataFim);
      }
    }
  } else {
    // Se a regra não está ativada (ex: Status mudou para "A CONTRATAR"), limpa os erros
    _limparErroVisual(colValorSeg);
    _limparErroVisual(colNumProp);
    _limparErroVisual(colDataFim);
  }

  return { alterou: alterou };
});


  // ===== Engine principal ====================================================
  function processar(e) {
    const t0 = _now();
    try {
      // Guards de evento
      if (!e || !e.range || typeof e.range.getSheet !== 'function') {
        _logIf('WARN', [["AVISO","RegrasNegocioV2/processar","Evento inválido ou fora de planilha"]]);
        return { alterouAlgo: false, erro: 'evento_invalido' };
      }
      const ctx = _montarCtx(e);
      const candidatas = _getRegrasAtivas().filter(r => !r.abaAlvo || r.abaAlvo === ctx.aba);
      let alterouAlgo = false; let aplicadas = 0;
      for (const r of candidatas) {
        if (!_regraBateNoEvento(ctx, r)) continue;
        const acaoFn = ActionRegistry.get(r.acao);
        const t1 = _now();
        if (typeof acaoFn === 'function') {
          const res = acaoFn(ctx, r) || {};
          const dt = _now() - t1;
          if (res.alterou) { alterouAlgo = true; aplicadas++; }
         if (res.alterou || dt > PERF_LOG_THRESHOLD_MS) {
            _logIf('INFO', [["INFO","RegrasNegocioV2/acao", { id: r.id, acao: r.acao, dt, alterou: !!res.alterou }]]);
          }
        } else {
          _logIf('WARN', [["AVISO","RegrasNegocioV2/acao_nao_registrada", `${r.id}|${r.acao}`]]);
        }
      }
      const total = _now() - t0;
      if (total > PERF_LOG_THRESHOLD_MS) {
        _logIf('INFO', [["INFO","RegrasNegocioV2/processar_fim", { total, aplicadas, avaliadas: candidatas.length }]]);
      }
      return { alterouAlgo };
    } catch (err) {
      _logIf('ERROR', [["ERRO", "RegrasNegocioV2/processar", err && err.message]]);
      return { alterouAlgo: false, erro: err && err.message };
    }
  }

  return { processar, ActionRegistry, _test: { _norm, _c2i, _getRegrasAtivas, _condicoesBatem } };
})();


/**
 * ▶ Integração no pipeline (cole no onEditHandler.js antes do Timestamp)
 * ----------------------------------------------------------------------
 * Posição sugerida: logo antes da seção "// ███ Etapa 7: TIMESTAMP"
 * Exemplo de bloco (resistente a ausência do módulo):
 *
 * try {
 * if (typeof RegrasNegocioV2 !== 'undefined') {
 * RegrasNegocioV2.processar(e);
 * }
 * } catch (err) {
 * try { Logger?.registrarLogBatch?.([["AVISO", "RegrasNegocioV2", err && err.message]]); } catch(_) {}
 * }
 */


/**
 * ▶ DSL — Regra única SEGUROS I⟷M (adicione em regrasConfig.js → REGRAS_DE_NEGOCIO)
 * -------------------------------------------------------------------------------
 * {
 * id: "SEG_I2M_SYNC_001",
 * status: "ATIVO",
 * abaAlvo: "SEGUROS",
 * colunasAlvo: ["I"],
 * prioridade: 1,
 * acao: "sincronizar_data_por_situacao",
 * parametros: {
 * coluna_destino: "M",
 * formato: "dd/MM/yyyy",
 * valor_ativador: "PAGO/CONTRATADO",
 * politica_quando_ativado: "SETAR_HOJE_SE_VAZIO",
 * politica_outros: "LIMPAR"
 * },
 * descricao: "Sincroniza M com a situação de I: preenche hoje se I=PAGO/CONTRATADO (se vazia) e limpa M caso contrário."
 * }
 */