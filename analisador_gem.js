/**
 * ANALISE_PLANILHA — Inventário imediato da estrutura da planilha e do código que atua sobre ela.
 * Saída única apresentável + export “IA-Ready” (CSV/MD/YAML) + seleção de abas e cabeçalho via guia.
 *
 * VERSÃO MODIFICADA (v2.0):
 * - Inclui sistema de PRESETS.
 * - Inclui ANÁLISE DE LÓGICA DA APLICAÇÃO.
 * - Inclui ANÁLISE VISUAL (Estilos e Tabelas/Ilhas de dados).
 * - Painel de controle aprimorado com limites de range.
 *
 * Como usar:
 * 1) Execute prepararSelecaoAbasTemp() no editor GAS → configure o painel na planilha.
 * 2) Execute executarAnalisePeloPainel() no editor GAS → o relatório será gerado.
 */

/** ===================== Configurações ===================== */
const MODO_SAIDA_UNICA = true;
const NOME_SAIDA_UNICA = 'RELATORIO_ANALISE';

const HEADER_SCAN_LIMIT = 20;
const HEADER_ASK_WHEN_UNCERTAIN = false;
const HEADER_SCORE_THRESHOLD = 1.5;

const EXPORT_TO_DRIVE = true;
const EXPORT_FOLDER_BASENAME = 'RELATORIO_ANALISE_EXPORT';
const EXPORT_MD_SPLIT_THRESHOLD = 12000;

/**
 * =======================================================================================
 * PRESETS DE ANÁLISE
 * =======================================================================================
 */
const PRESETS = {
  COMPLETO: {
    documentacaoGeral: true,
    formatoExpandido: true,
    triggers: true,
    logicaDeAplicacao: true,
    validacoes: true,
    monitoramentosCONFIG: true,
    formatacaoCondicional: true,
    filtros: true,
    protecoes: true,
    rangesNomeados: true,
    graficos: true,
    recursosAvancados: true,
    analiseFormatacaoVisual: true // Ativa varredura visual e de tabelas
  },
  FOCO_CODIGO_E_REGRAS: {
    documentacaoGeral: true,
    formatoExpandido: false,
    triggers: true,
    logicaDeAplicacao: true,
    validacoes: true,
    monitoramentosCONFIG: true,
    formatacaoCondicional: true,
    filtros: false,
    protecoes: false,
    rangesNomeados: true,
    graficos: false,
    recursosAvancados: false,
    analiseFormatacaoVisual: false
  },
  ESTRUTURA_RAPIDA: {
    documentacaoGeral: true,
    formatoExpandido: true,
    triggers: false,
    logicaDeAplicacao: false,
    validacoes: true,
    monitoramentosCONFIG: false,
    formatacaoCondicional: false,
    filtros: true,
    protecoes: true,
    rangesNomeados: true,
    graficos: false,
    recursosAvancados: false,
    analiseFormatacaoVisual: false
  },
  // Preset para Fase 1: gerar/atualizar aba ANL_REGRAS_EDITAVEL
  DUAS_FASES_RASCUNHO: {
    documentacaoGeral: false,
    formatoExpandido: false,
    triggers: false,
    logicaDeAplicacao: false,
    validacoes: false,
    monitoramentosCONFIG: false,
    formatacaoCondicional: false,
    filtros: false,
    protecoes: false,
    rangesNomeados: false,
    graficos: false,
    recursosAvancados: false,
    analiseFormatacaoVisual: false,
    fase1RegrasEditavel: true,
    fase2RegrasConsolidadas: false
  },
  // Preset para Fase 2: gerar saída consolidada REGRAS_DE_NEGOCIO_CONSOLIDADAS
  DUAS_FASES_FINAL: {
    documentacaoGeral: false,
    formatoExpandido: false,
    triggers: false,
    logicaDeAplicacao: false,
    validacoes: false,
    monitoramentosCONFIG: false,
    formatacaoCondicional: false,
    filtros: false,
    protecoes: false,
    rangesNomeados: false,
    graficos: false,
    recursosAvancados: false,
    analiseFormatacaoVisual: false,
    fase1RegrasEditavel: false,
    fase2RegrasConsolidadas: true
  }
};


/** Nomes de guias de saída (ignorar na varredura) */
var SAIDAS = [
  'DOCUMENTACAO_GERAL', 'FORMATO_EXPANDIDO', 'TRIGGERS_PROJETO', 'VALIDACOES_POR_COLUNA',
  'MONITORAMENTOS_DECLARADOS', 'FORMATACAO_CONDICIONAL', 'FILTROS_ATIVOS', 'PROTECOES',
  'RANGES_NOMEADOS', 'GRAFICOS', 'FILTER_VIEWS_E_AFINS', 'SELECIONAR_ABAS_TEMP',
  'LOGICA_DA_APLICACAO', 'ANALISE_VISUAL_ESTILOS',
  'ANL_REGRAS_EDITAVEL', 'REGRAS_DE_NEGOCIO_CONSOLIDADAS'
];
if (MODO_SAIDA_UNICA && SAIDAS.indexOf(NOME_SAIDA_UNICA) === -1) SAIDAS.push(NOME_SAIDA_UNICA);


/** ===================== Estado interno ===================== */
var __OUT_SH = null;
var __CUR_ROW = 6;
var __TOC = [];
var __HEADER_INDEX = {};
var __SECTIONS = [];
var __SEC_INDEX = 0;
var __EXPORT = null;
var __HEADER_ROW_GLOBAL = null;

/* ======================================================================
 * EXECUÇÃO GERAL (MODIFICADO PARA MAPA DE CONFIGURAÇÃO)
 * ====================================================================== */
function documentarEstruturaCompleta(presetName, mapaConfig) {
  const presetAtivo = PRESETS[presetName] || PRESETS.COMPLETO;
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var analisadas = [];

  if (MODO_SAIDA_UNICA) beginReport_(ss, presetAtivo.exportarDrive);

  __HEADER_ROW_GLOBAL = readGlobalHeaderRowFromSelection_();

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var name = sh.getName();
    
    // Se temos um mapa de configuração, verificamos se a aba está nele
    if (mapaConfig && !mapaConfig[name]) continue;

    if (SAIDAS.indexOf(name) !== -1) continue;
    if (sh.isSheetHidden && sh.isSheetHidden()) continue; // Opcional: depender da seleção do usuário
    
    var info = documentarAbaCompleta_(sh);
    analisadas.push(info);
    __HEADER_INDEX[name.toUpperCase()] = buildHeaderIndex_(info.headers);
  }

  // Análises Padrão
  if (presetAtivo.documentacaoGeral) salvarDocumentacaoGeral_(ss, analisadas);
  if (presetAtivo.formatoExpandido) salvarFormatoExpandido_(ss, analisadas);
  if (presetAtivo.triggers) { try { salvarTriggersProjeto_(ss, inventariarTriggersProjeto_()); } catch (e1) { log_('TRIGGERS_PROJETO: ' + (e1 && e1.message)); } }
  if (presetAtivo.logicaDeAplicacao) { try { salvarLogicaDeAplicacao_(ss, coletarLogicaDeAplicacao_()); } catch (e_logic) { log_('LOGICA_DA_APLICAÇÃO: ' + (e_logic && e_logic.message)); } }

  // Validações
  if (presetAtivo.validacoes) {
    try {
      var mapaValid = [];
      for (var j = 0; j < sheets.length; j++) {
        var s2 = sheets[j]; var nm = s2.getName();
        if (mapaConfig && !mapaConfig[nm]) continue;
        if (SAIDAS.indexOf(nm) !== -1) continue;
        mapaValid = mapaValid.concat(listarValidacoesPorColuna_(s2));
      }
      salvarValidacoesPorColuna_(ss, mapaValid);
    } catch (e2) { log_('VALIDACOES_POR_COLUNA: ' + (e2 && e2.message)); }
  }

  // Outros módulos
  if (presetAtivo.monitoramentosCONFIG) { try { salvarMonitoramentosDeclarados_(ss, coletarMonitoramentosDeclarados_CONFIG_()); } catch (e3) { log_('MONITORAMENTOS_DECLARADOS: ' + (e3 && e3.message)); } }
  if (presetAtivo.formatacaoCondicional) { try { salvarFormatacaoCondicional_(ss, coletarFormatacaoCondicional_(sheets)); } catch (e4) { log_('FORMATACAO_CONDICIONAL: ' + (e4 && e4.message)); } }
  if (presetAtivo.filtros) { try { salvarFiltrosAtivos_(ss, coletarFiltrosAtivos_(sheets)); } catch (e5) { log_('FILTROS_ATIVOS: ' + (e5 && e5.message)); } }
  
  // Proteções e Ranges (Passando lista de nomes para filtro)
  var listaNomesAbas = mapaConfig ? Object.keys(mapaConfig) : [];
  if (presetAtivo.protecoes) { try { salvarProtecoes_(ss, coletarProtecoes_(ss, listaNomesAbas)); } catch (e6) { log_('PROTECOES: ' + (e6 && e6.message)); } }
  if (presetAtivo.rangesNomeados) { try { salvarRangesNomeados_(ss, coletarNamedRanges_(ss, listaNomesAbas)); } catch (e7) { log_('RANGES_NOMEADOS: ' + (e7 && e7.message)); } }
  
  if (presetAtivo.graficos) { try { salvarGraficos_(ss, coletarGraficos_(sheets)); } catch (e8) { log_('GRAFICOS: ' + (e8 && e8.message)); } }
  if (presetAtivo.recursosAvancados) { try { var fv = coletarFilterViewsESlicers_(ss, listaNomesAbas); if (fv && fv.length) salvarFilterViewsEAfins_(ss, fv); } catch (e9) { log_('FILTER_VIEWS_E_AFINS indisponível.'); } }

  // === ANÁLISE VISUAL E ESTRUTURAL (NOVO MÓDULO) ===
if (presetAtivo.analiseFormatacaoVisual) {
    try {
      var dadosVisuais = [];
      for (var k = 0; k < sheets.length; k++) {
        var sCurrent = sheets[k];
        var sName = sCurrent.getName();

        // Decide se deve analisar a aba e qual limite usar
        var deveAnalisarEstilo = false;
        var limite = null;

        if (!mapaConfig) {
          // Sem mapa de seleção: padrão é analisar todas as abas
          deveAnalisarEstilo = true;
        } else if (mapaConfig[sName]) {
          // Usando painel SELECIONAR_ABAS_TEMP: só analisa se marcado
          deveAnalisarEstilo = mapaConfig[sName].analisarEstilo === true;
          limite = mapaConfig[sName].limite || null;
        }

        if (!deveAnalisarEstilo) {
          continue;
        }

        // 1. Coleta Estilos (Cores/Fontes)
        dadosVisuais = dadosVisuais.concat(coletarFormatacaoVisual_(sCurrent, limite));

        // 2. Coleta Tabelas Distintas (Ilhas de Dados)
        dadosVisuais = dadosVisuais.concat(coletarIlhasDeDados_(sCurrent, limite));
      }

      if (dadosVisuais.length > 0) {
        salvarFormatacaoVisual_(ss, dadosVisuais);
      }
    } catch (eViz) {
      log_('ANALISE_VISUAL_ESTILOS: ' + (eViz && eViz.message));
    }
  }
  // ================================================

 try {
    if (presetAtivo.fase1RegrasEditavel) {
      gerarAbaRegrasEditavel_(ss);
    }
    if (presetAtivo.fase2RegrasConsolidadas) {
      salvarRegrasNegocioConsolidadas_(ss);
    }
  } catch (eFases) {
    log_('REGRAS_NEGOCIO_FASES: ' + (eFases && eFases.message));
  }

  if (MODO_SAIDA_UNICA) finalizeReport_();
}
/* ===================== Núcleo por aba ===================== */
function documentarAbaCompleta_(sheet) {
  var name = sheet.getName();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerRow = obterHeaderRow_(sheet);
  var headers = [];
  if (lastCol > 0 && lastRow >= headerRow) headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  var temFiltro = !!sheet.getFilter();
  var qtdGraficos = sheet.getCharts ? ((sheet.getCharts() || []).length) : 0;
  var amostraValid = false, amostraFormulas = false;
  if (lastRow > headerRow && lastCol > 0) {
    var amRows = Math.min(5, lastRow - headerRow), amCols = Math.min(8, lastCol);
    var rAmostra = sheet.getRange(headerRow + 1, 1, amRows, amCols);
    var dvs = rAmostra.getDataValidations(); var fms = rAmostra.getFormulas();
    amostraValid = _matrizTemAlgum_(dvs); amostraFormulas = _matrizTemAlgumStringNaoVazia_(fms);
  }
  var densidade = calcularDensidade_(sheet, headerRow);
  return {
    aba: name, linhas: lastRow, colunas: lastCol, headerRow: headerRow, headers: headers,
    dadosPercentual: densidade, temFiltro: temFiltro, temValidacaoAmostral: amostraValid,
    temFormulaAmostral: amostraFormulas, graficosQtde: qtdGraficos
  };
}

/* ===================== Saídas Gerais ===================== */
function salvarDocumentacaoGeral_(ss, arr) {
  var header = ['ABA', 'LINHAS', 'COLUNAS', 'DADOS_%', 'FILTROS', 'VALID.', 'FORMUL.', 'GRAFICOS', 'HEADER_ROW'];
  var rows = [header];
  for (var i = 0; i < arr.length; i++) {
    var it = arr[i];
    rows.push([it.aba, it.linhas, it.colunas, _pct(it.dadosPercentual), it.temFiltro ? 'SIM' : 'NAO',
    it.temValidacaoAmostral ? 'SIM' : 'NAO', it.temFormulaAmostral ? 'SIM' : 'NAO',
      it.graficosQtde || 0, it.headerRow]);
  }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'DOCUMENTACAO_GERAL', rows); return; }
  var sh = ensureSheet_(ss, 'DOCUMENTACAO_GERAL'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

function salvarFormatoExpandido_(ss, arr) {
  var header = ['ABA', 'COLUNA_A1', 'HEADER', 'TIPO_INFERIDO'];
  var rows = [header];
  for (var i = 0; i < arr.length; i++) {
    var it = arr[i];
    for (var c = 0; c < it.colunas; c++) {
      var h = it.headers[c] || ''; rows.push([it.aba, columnToLetter_(c + 1), h, inferirTipoPeloHeader_(h)]);
    }
  }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FORMATO_EXPANDIDO', rows); return; }
  var sh = ensureSheet_(ss, 'FORMATO_EXPANDIDO'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

/* ===================== Triggers ===================== */
function inventariarTriggersProjeto_() {
  var triggers = ScriptApp.getProjectTriggers(); var out = [];
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    out.push({
      handler: t.getHandlerFunction ? t.getHandlerFunction() : '',
      eventType: t.getEventType ? String(t.getEventType()) : '',
      triggerSource: t.getTriggerSource ? String(t.getTriggerSource()) : '',
      triggerSourceId: t.getTriggerSourceId ? String(t.getTriggerSourceId()) : ''
    });
  }
  return out;
}
function salvarTriggersProjeto_(ss, registros) {
  var header = ['HANDLER', 'EVENT_TYPE', 'TRIGGER_SOURCE', 'TRIGGER_SOURCE_ID']; var rows = [header];
  for (var i = 0; i < registros.length; i++) { var r = registros[i]; rows.push([r.handler, r.eventType, r.triggerSource, r.triggerSourceId]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'TRIGGERS_PROJETO', rows); return; }
  var sh = ensureSheet_(ss, 'TRIGGERS_PROJETO'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

/* ===================== Validações ===================== */
function listarValidacoesPorColuna_(sheet) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn(); if (lastCol < 1) return [];
  var headerRow = obterHeaderRow_(sheet);
  var headers = lastCol > 0 && lastRow >= headerRow ? sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0] : [];
  if (lastRow <= headerRow) {
    var vazio = []; for (var c = 0; c < lastCol; c++) {
      vazio.push({ aba: sheet.getName(), colunaA1: columnToLetter_(c + 1), header: headers[c] || '', temValidacao: false, criterios: '', criteriosArgs: '' });
    }
    return vazio;
  }
  var range = sheet.getRange(headerRow + 1, 1, Math.max(0, lastRow - headerRow), lastCol);
  var dvs = range.getDataValidations(); var results = [];
  for (var c = 0; c < lastCol; c++) {
    var tipos = {}, argsResumo = {}, tem = false;
    for (var r = 0; r < dvs.length; r++) {
      var dv = dvs[r][c]; if (!dv) continue; tem = true;
      var tipo = 'DESCONHECIDO', args = [];
      try { tipo = String(dv.getCriteriaType()); } catch (e) { }
      try { args = dv.getCriteriaValues() || []; } catch (e2) { }
      tipos[tipo] = true;
      var argStr = (args || []).map(function (a) {
        if (a == null) return '';
        if (a.getA1Notation) return a.getA1Notation();
        if (a.map && a.forEach) return '[array]';
        if (Object.prototype.toString.call(a) === '[object Date]') return Utilities.formatDate(a, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        return String(a);
      }).slice(0, 3).join(' ; ');
      if (argStr) argsResumo[argStr] = true;
    }
    results.push({
      aba: sheet.getName(),
      colunaA1: columnToLetter_(c + 1),
      header: headers[c] || '',
      temValidacao: tem,
      criterios: Object.keys(tipos).sort().join(' | '),
      criteriosArgs: Object.keys(argsResumo).sort().join(' | ')
    });
  }
  return results;
}
function salvarValidacoesPorColuna_(ss, lista) {
  var header = ['ABA', 'COLUNA_A1', 'HEADER', 'TEM_VALIDACAO', 'CRITERIOS', 'ARGS_RESUMO']; var rows = [header];
  for (var i = 0; i < lista.length; i++) {
    var m = lista[i]; rows.push([m.aba, m.colunaA1, m.header, m.temValidacao ? 'SIM' : 'NAO', m.criterios || '', m.criteriosArgs || '']);
  }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'VALIDACOES_POR_COLUNA', rows); return; }
  var sh = ensureSheet_(ss, 'VALIDACOES_POR_COLUNA'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

/* ===================== Lógica de Aplicação (CONFIG) ===================== */
function obterNomeSemantico_(sheetName, indice) {
  const sheetConfig = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) ? CONFIG.SHEETS[sheetName] : null;
  if (!sheetConfig || !sheetConfig.colunas) return null;
  for (const [nome, idx] of Object.entries(sheetConfig.colunas)) {
    if (idx === indice) return nome;
  }
  return null;
}

function formatarComportamentoHandler_(handler, sheetName) {
  const { modulo, options } = handler;
  if (!options) return 'Nenhuma opção configurada.';
  try {
    switch (modulo) {
      case 'Module_VALIDACAO':
        const validacoes = Object.entries(options.colunasValidacao || {})
          .map(([col, tipo]) => `valida a coluna ${columnToLetter_(parseInt(col))} (${obterNomeSemantico_(sheetName, parseInt(col)) || 'N/D'}) como '${tipo}'`)
          .join('; ');
        return `Aplica as seguintes validações: ${validacoes || 'Nenhuma especificada'}.`;
      case 'Module_NUMERADOR':
        const colDestino = options.colunaDestino;
        return `Gera um ID com prefixo "${options.prefixo}" na coluna ${columnToLetter_(colDestino)} (${obterNomeSemantico_(sheetName, colDestino) || 'N/D'}).`;
      case 'Module_RELACIONAMENTO':
        const campos = (options.campos || []).map(c =>
          `ao editar ${columnToLetter_(c.colunaDisparo)}, busca na aba "${c.abaFonte}" pela chave em ${columnToLetter_(c.colunaChaveFonte)}`
        ).join('; ');
        return `Preenche dados automaticamente (PROCV automático): ${campos}.`;
      case 'Module_DROPDOWN':
        if (options.colSituacao) {
          return `Cria um dropdown hierárquico para Etapa (${columnToLetter_(options.colEtapa)}) -> Situação (${columnToLetter_(options.colSituacao)}), com dados da aba "${options.fonte.nomeAba}".`;
        }
        return `Cria um dropdown simples na coluna ${options.colEtapa ? columnToLetter_(options.colEtapa) : '?'} com dados da aba "${options.fonte.nomeAba}".`;
        case 'Module_VISOES':
          const visoes = Object.keys(options.visoes || {}).join(', ');
          return `Permite alternar entre as seguintes visões: ${visoes}.`;

        case 'Module_REGRAS_NEGOCIO':
          if (typeof REGRAS_DE_NEGOCIO !== 'undefined' && Array.isArray(REGRAS_DE_NEGOCIO)) {
            try {
              var regrasDaAba = REGRAS_DE_NEGOCIO.filter(function (r) {
                if (!r) return false;
                var mesmaAba  = (r.abaAlvo === sheetName);
                var statusOk  = !r.status || String(r.status).toUpperCase() === 'ATIVO';
                return mesmaAba && statusOk;
              });

              if (regrasDaAba.length) {
                var descricoes = regrasDaAba.map(function (r) {
                  var id   = r.id || r.codigo || r.nome || '?';
                  var desc = r.descricao || r.descricaoCurta || r.descricaoLonga || '';
                  return '[' + id + '] ' + desc;
                }).join(' | ');

                return 'Regras de negócio ativas para esta aba: ' + descricoes;
              }
            } catch (eRegras) {
              // Se der erro aqui, cai no fallback genérico abaixo
            }
          }
          return 'Executa motor de regras de negócio (ver regrasConfig.js / REGRAS_DE_NEGOCIO).';

        default:
          return `Módulo "${modulo}" não possui tradutor definido.`;
    }
  } catch (e) {
    return `Erro ao traduzir: ${e.message}`;
  }
}

function coletarLogicaDeAplicacao_() {
  const regrasDocumentadas = [];
  if (typeof CONFIG === 'undefined' || !CONFIG.SHEETS) {
    return [{ aba: 'ERRO', gatilho: 'N/A', modulo: 'N/A', comportamento: 'CONFIG inválido.' }];
  }
  Object.entries(CONFIG.SHEETS).forEach(([sheetName, sheetConfig]) => {
    if (!sheetConfig.handlers || sheetConfig.handlers.length === 0) return;
    sheetConfig.handlers.forEach(handler => {
      const colunasGatilho = (handler.colunas || [])
        .map(c => `${columnToLetter_(c)}`)
        .join(', ');
      regrasDocumentadas.push({
        aba: sheetName,
        gatilho: `Colunas: ${colunasGatilho}`,
        modulo: handler.modulo,
        comportamento: formatarComportamentoHandler_(handler, sheetName)
      });
    });
  });
  return regrasDocumentadas;
}

function salvarLogicaDeAplicacao_(ss, lista) {
  var header = ['Aba', 'Gatilho', 'Módulo', 'Comportamento Detalhado'];
  var rows = [header];
  for (var i = 0; i < lista.length; i++) {
    var it = lista[i];
    rows.push([it.aba, it.gatilho, it.modulo, it.comportamento]);
  }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'LOGICA_DA_APLICACAO', rows); return; }
  var sh = ensureSheet_(ss, 'LOGICA_DA_APLICACAO'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

function gerarAbaRegrasEditavel_(ss) {
  if (typeof REGRAS_DE_NEGOCIO === 'undefined' || !Array.isArray(REGRAS_DE_NEGOCIO) || REGRAS_DE_NEGOCIO.length === 0) {
    log_('gerarAbaRegrasEditavel_: REGRAS_DE_NEGOCIO não disponível ou vazio.');
    return;
  }

  var sh = ensureSheet_(ss, 'ANL_REGRAS_EDITAVEL');

  // Mapa de textos já preenchidos pelo usuário (se a aba já existir)
  var textosPorId = {};
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, 13).getValues(); // A:M
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0] || '').trim(); // A
      if (!id) continue;
      textosPorId[id] = {
        descricaoFluxo: data[i][8] || '', // I
        exemplo: data[i][9] || '',        // J
        obs: data[i][10] || '',           // K
        tags: data[i][11] || '',          // L
        statusRevisao: data[i][12] || ''  // M
      };
    }
  }

  var header = [
    'ID_REGRA',                // A
    'ABA_ALVO',                // B
    'TIPO_REGRA',              // C (guard/post)
    'PRIORIDADE',              // D
    'COLUNAS_GATILHO_A1',      // E
    'ACAO_DSL',                // F
    'CAMPOS_AFETADOS_A1',      // G
    'STATUS_REGRA',            // H (ATIVO/INATIVO)
    'DESCRICAO_FLUXO_USUARIO', // I
    'EXEMPLO_PRATICO',         // J
    'OBS_USUARIO',             // K
    'TAGS_NEGOCIO',            // L
    'STATUS_REVISAO'           // M
  ];
  var rows = [header];

  for (var r = 0; r < REGRAS_DE_NEGOCIO.length; r++) {
    var regra = REGRAS_DE_NEGOCIO[r] || {};
    var id = String(regra.id || '').trim();
    if (!id) continue;

    var aba = String(regra.abaAlvo || regra.aba || '').trim();
    var status = String(regra.status || '').toUpperCase();
    var prioridade = (typeof regra.prioridade === 'number' || typeof regra.prioridade === 'string') ? regra.prioridade : '';
    var colunasGatilho = '';
    if (Array.isArray(regra.colunasAlvo)) colunasGatilho = regra.colunasAlvo.join(',');
    else if (regra.colunaAlvo) colunasGatilho = String(regra.colunaAlvo);

    var acao = String(regra.acao || '').trim();
    var camposAfetados = inferirCamposAfetadosA1_(regra);
    var tipoRegra = inferirTipoRegra_(regra);

    var textos = textosPorId[id] || {};
    rows.push([
      id,
      aba,
      tipoRegra,
      prioridade,
      colunasGatilho,
      acao,
      camposAfetados,
      status,
      textos.descricaoFluxo || '',
      textos.exemplo || '',
      textos.obs || '',
      textos.tags || '',
      textos.statusRevisao || ''
    ]);
  }

  sh.clear();
  writeTable_(sh, 1, 1, rows, true);
}

/**
 * Consolida REGRAS_DE_NEGOCIO + textos da aba ANL_REGRAS_EDITAVEL em saída IA-ready.
 * Gera aba/section REGRAS_DE_NEGOCIO_CONSOLIDADAS.
 */
function salvarRegrasNegocioConsolidadas_(ss) {
  if (typeof REGRAS_DE_NEGOCIO === 'undefined' || !Array.isArray(REGRAS_DE_NEGOCIO) || REGRAS_DE_NEGOCIO.length === 0) {
    log_('salvarRegrasNegocioConsolidadas_: REGRAS_DE_NEGOCIO não disponível ou vazio.');
    return;
  }

  var textosPorId = {};
  var sheetEdit = ss.getSheetByName('ANL_REGRAS_EDITAVEL');
  if (sheetEdit) {
    var lastRow = sheetEdit.getLastRow();
    if (lastRow >= 2) {
      var data = sheetEdit.getRange(2, 1, lastRow - 1, 13).getValues(); // A:M
      for (var i = 0; i < data.length; i++) {
        var id = String(data[i][0] || '').trim();
        if (!id) continue;
        textosPorId[id] = {
          descricaoFluxo: data[i][8] || '',
          exemplo: data[i][9] || '',
          obs: data[i][10] || '',
          tags: data[i][11] || '',
          statusRevisao: data[i][12] || ''
        };
      }
    }
  }

  var header = [
    'ID_REGRA',                // A
    'ABA_ALVO',                // B
    'TIPO_REGRA',              // C
    'PRIORIDADE',              // D
    'COLUNAS_GATILHO_A1',      // E
    'ACAO_DSL',                // F
    'CAMPOS_AFETADOS_A1',      // G
    'STATUS_REGRA',            // H
    'DESCRICAO_FLUXO_USUARIO', // I
    'EXEMPLO_PRATICO',         // J
    'OBS_USUARIO',             // K
    'TAGS_NEGOCIO',            // L
    'STATUS_REVISAO'           // M
  ];
  var rows = [header];

  for (var r = 0; r < REGRAS_DE_NEGOCIO.length; r++) {
    var regra = REGRAS_DE_NEGOCIO[r] || {};
    var id = String(regra.id || '').trim();
    if (!id) continue;

    var aba = String(regra.abaAlvo || regra.aba || '').trim();
    var status = String(regra.status || '').toUpperCase();
    var prioridade = (typeof regra.prioridade === 'number' || typeof regra.prioridade === 'string') ? regra.prioridade : '';
    var colunasGatilho = '';
    if (Array.isArray(regra.colunasAlvo)) colunasGatilho = regra.colunasAlvo.join(',');
    else if (regra.colunaAlvo) colunasGatilho = String(regra.colunaAlvo);

    var acao = String(regra.acao || '').trim();
    var camposAfetados = inferirCamposAfetadosA1_(regra);
    var tipoRegra = inferirTipoRegra_(regra);

    var textos = textosPorId[id] || {};
    rows.push([
      id,
      aba,
      tipoRegra,
      prioridade,
      colunasGatilho,
      acao,
      camposAfetados,
      status,
      textos.descricaoFluxo || '',
      textos.exemplo || '',
      textos.obs || '',
      textos.tags || '',
      textos.statusRevisao || ''
    ]);
  }

  if (MODO_SAIDA_UNICA) {
    appendSection_(getOutSheet_(), 'REGRAS_DE_NEGOCIO_CONSOLIDADAS', rows);
  } else {
    var shOut = ensureSheet_(ss, 'REGRAS_DE_NEGOCIO_CONSOLIDADAS');
    shOut.clear();
    writeTable_(shOut, 1, 1, rows, true);
  }
}

/**
 * Heurística simples para inferir tipo de regra (guard/post) com base em campos existentes.
 */
function inferirTipoRegra_(regra) {
  if (!regra) return '';
  var t = String(regra.tipo || '').toLowerCase();
  if (t === 'guard' || t === 'post') return t;
  var acao = String(regra.acao || '').toLowerCase();
  if (acao.indexOf('validar_') === 0) return 'guard';
  return 'post';
}

/**
 * Heurística simples para inferir colunas afetadas em notação A1 a partir de parametros da regra.
 */
function inferirCamposAfetadosA1_(regra) {
  if (!regra || !regra.parametros) return '';
  var cols = [];
  var p = regra.parametros;

  function addCol(v) {
    if (!v) return;
    if (Array.isArray(v)) {
      for (var i = 0; i < v.length; i++) addCol(v[i]);
      return;
    }
    var s = String(v).trim();
    if (!s) return;
    if (cols.indexOf(s) === -1) cols.push(s);
  }

  addCol(p.coluna_destino);
  addCol(p.colunas_destino);
  addCol(p.colunas_monitoradas);
  addCol(p.coluna_produto);
  addCol(p.coluna_operacao);
  addCol(p.coluna_valor_seguro);
  addCol(p.coluna_numero_proposta);
  addCol(p.coluna_status);
  addCol(p.coluna_data_inicio);
  addCol(p.coluna_data_fim);
  addCol(p.coluna_sensivel);

  return cols.join(',');
}

function coletarMonitoramentosDeclarados_CONFIG_() {
  var lista = [];
  try {
    var norm = normalizeConfig_();
    (norm.timestamp || []).forEach(function (x) {
      lista.push({ origem: x.origem || 'CONFIG.TIMESTAMP', aba: x.sheet || '', alvo: x.col || '', handlerOuModulo: 'TimestampManager', detalhe: '' });
    });
    (norm.moveRules || []).forEach(function (x) {
      lista.push({ origem: x.origem || 'CONFIG.MOVE_RULES', aba: x.sheet || '', alvo: x.col || '', handlerOuModulo: 'RowManager', detalhe: JSON.stringify(x.map || {}) });
    });
    (norm.dropdowns || []).forEach(function (x) {
      lista.push({ origem: x.origem || 'CONFIG.DROPDOWNS', aba: x.sheet || '', alvo: x.col || '', handlerOuModulo: 'DropdownManager', detalhe: JSON.stringify({ type: x.type || '', source: x.source || '', parent: x.parentCol || '' }) });
    });
    if (norm.installments && norm.installments.cols) {
      var S = norm.installments.sheet || ''; var cols = norm.installments.cols;
      if (cols.contract) lista.push({ origem: 'CONFIG.ACOMPANHAMENTO_PARCELAS', aba: S, alvo: cols.contract, handlerOuModulo: 'ParcelasManager', detalhe: 'COLUNA_DATA_CONTRATO' });
      if (cols.dueDay) lista.push({ origem: 'CONFIG.ACOMPANHAMENTO_PARCELAS', aba: S, alvo: cols.dueDay, handlerOuModulo: 'ParcelasManager', detalhe: 'COLUNA_DIA_VENCTO' });
      (cols.list || []).forEach(function (letter) {
        lista.push({ origem: 'CONFIG.ACOMPANHAMENTO_PARCELAS.COLUNAS_PARCELAS', aba: S, alvo: letter, handlerOuModulo: 'ParcelasManager', detalhe: '' });
      });
    }
    (norm.handlers || []).forEach(function (x) {
      lista.push({ origem: x.origem || 'CONFIG.HANDLERS', aba: x.sheet || '', alvo: x.target || '', handlerOuModulo: x.handler || '', detalhe: '' });
    });
  } catch (e) { /* silencioso */ }
  return lista;
}

function salvarMonitoramentosDeclarados_(ss, lista) {
  var header = ['ORIGEM', 'ABA', 'ALVO(A1)', 'HANDLER/MODULO', 'DETALHE']; var rows = [header];
  for (var i = 0; i < lista.length; i++) {
    var it = lista[i];
    rows.push([it.origem || '', it.aba || '', it.alvo || '', it.handlerOuModulo || '', it.detalhe || '']);
  }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'MONITORAMENTOS_DECLARADOS', rows); return; }
  var sh = ensureSheet_(ss, 'MONITORAMENTOS_DECLARADOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

function normalizeConfig_() {
  if (typeof CONFIG !== 'object' || !CONFIG) return {};
  var out = { timestamp: [], moveRules: [], dropdowns: [], handlers: [], installments: null };
  function pick(obj, names) { if (!obj) return undefined; var keys = Object.keys(obj); for (var i = 0; i < names.length; i++) { var n = names[i]; for (var k = 0; k < keys.length; k++) { if (String(keys[k]).toLowerCase() === String(n).toLowerCase()) return obj[keys[k]]; } } return undefined; }
  function pickKey(obj, names) { if (!obj) return null; var keys = Object.keys(obj); for (var i = 0; i < names.length; i++) { var n = names[i]; for (var k = 0; k < keys.length; k++) { if (String(keys[k]).toLowerCase() === String(n).toLowerCase()) return keys[k]; } } return null; }
  function sheetNameFrom(item) { if (!item || typeof item !== 'object') return ''; var key = pickKey(item, ['nome', 'name', 'sheet', 'aba', 'title']); return key ? String(item[key]) : ''; }
  function toLetter(sheetName, ref) { return parseColRefToLetter_(sheetName, ref); }
  var sheetsNode = pick(CONFIG, ['SHEETS', 'PLANILHAS', 'ABAS']);
  if (sheetsNode) {
    if (Array.isArray(sheetsNode)) {
      for (var i = 0; i < sheetsNode.length; i++) handleSheetConfig_(sheetNameFrom(sheetsNode[i]), sheetsNode[i]);
    } else if (typeof sheetsNode === 'object') {
      var keys = Object.keys(sheetsNode);
      for (var j = 0; j < keys.length; j++) { var nm = keys[j]; handleSheetConfig_(nm, sheetsNode[nm]); }
    }
  }
  (function () {
    var ddTop = pick(CONFIG, ['LISTAS_SUSPENSAS', 'DROPDOWNS']);
    if (!ddTop) return; if (!Array.isArray(ddTop)) ddTop = [ddTop];
    ddTop.forEach(function (rule) {
      if (!rule) return;
      var sheet = rule.nomeAba || rule.aba || rule.sheet || '';
      var cat = rule.colunaCategoria || rule.categoryCol || rule.categoria;
      var sub = rule.colunaSubcategoria || rule.subcategoryCol || rule.subcategoria;
      if (cat) out.dropdowns.push({ origem: 'CONFIG.LISTAS_SUSPENSAS.categoria', sheet: sheet, col: toLetter(sheet, cat), type: 'categoria', source: (rule.fonte || rule.source || ''), parentCol: null });
      if (sub) out.dropdowns.push({ origem: 'CONFIG.LISTAS_SUSPENSAS.subcategoria', sheet: sheet, col: toLetter(sheet, sub), type: 'subcategoria', source: (rule.fonte || rule.source || ''), parentCol: toLetter(sheet, cat) });
    });
  })();
  (function () {
    var ap = pick(CONFIG, ['ACOMPANHAMENTO_PARCELAS', 'PARCELAS', 'INSTALLMENTS', 'SCHEDULE']); if (!ap || typeof ap !== 'object') return;
    var S = ap.ABA || ap.aba || ap.sheet || 'SEGUROS';
    var cols = {
      contract: toLetter(S, ap.COLUNA_DATA_CONTRATO || ap.contractDateCol || ap.contratoCol),
      dueDay: toLetter(S, ap.COLUNA_DIA_VENCTO || ap.dueDayCol || ap.venctoCol),
      list: []
    };
    var lst = ap.COLUNAS_PARCELAS || ap.parcelasCols || ap.installmentCols || [];
    if (Array.isArray(lst)) lst.forEach(function (x) {
      var col = (x && (x.col || x.c || x.column)) || x; var L = toLetter(S, col); if (L) cols.list.push(L);
    });
    out.installments = { sheet: S, cols: cols };
  })();
  (function () {
    var H = pick(CONFIG, ['HANDLERS', 'ACTIONS', 'EVENTS']); if (!H || typeof H !== 'object') return;
    var keys = Object.keys(H);
    for (var i = 0; i < keys.length; i++) { var key = keys[i]; out.handlers.push({ sheet: '', target: key, handler: String(H[key]) }); }
  })();
  return out;
  function handleSheetConfig_(sheetName, cfg) {
    if (!sheetName) return; cfg = cfg || {};
    var cols = pick(cfg, ['COLUNAS_MONITORADAS', 'MONITORED_COLUMNS', 'WATCHED_COLUMNS', 'TIMESTAMP_COLUMNS']);
    if (Array.isArray(cols)) cols.forEach(function (ref) {
      out.timestamp.push({ origem: 'CONFIG.SHEETS.COLUNAS_MONITORADAS', sheet: sheetName, col: toLetter(sheetName, ref) });
    });
    var mv = pick(cfg, ['MOVER_LINHA_CONFIG', 'MOVE_RULES', 'PIPELINE', 'STATUS_MAP']);
    if (mv) {
      if (Array.isArray(mv)) {
        mv.forEach(function (rule) {
          if (!rule) return; var colKey = rule.col || rule.column || rule.c || ''; var L = toLetter(sheetName, colKey);
          out.moveRules.push({ sheet: sheetName, col: L, map: (rule.map || rule.destinos || {}) });
        });
      } else if (typeof mv === 'object') {
        Object.keys(mv).forEach(function (colKey) {
          var L = toLetter(sheetName, colKey);
          out.moveRules.push({ sheet: sheetName, col: L, map: mv[colKey] });
        });
      }
    }
    var H = pick(cfg, ['HANDLERS', 'ACTIONS', 'EVENTS']);
    if (H && typeof H === 'object') {
      Object.keys(H).forEach(function (target) {
        out.handlers.push({ origem: 'CONFIG.SHEETS.HANDLERS', sheet: sheetName, target: target, handler: String(H[target]) });
      });
    }
    var ddBlocks = pick(cfg, ['LISTAS_SUSPENSAS', 'DROPDOWNS', 'DROPDOWN_DEPENDENTE', 'DEPENDENT_DROPDOWNS', 'VALIDACOES_LISTA']);
    if (Array.isArray(ddBlocks)) ddBlocks.forEach(function (rule) {
      if (!rule) return; var alvo = rule.colunaAlvo || rule.coluna || rule.target || rule.colunaCategoria || '';
      out.dropdowns.push({ origem: 'CONFIG.SHEETS.DROPDOWNS', sheet: sheetName, col: toLetter(sheetName, alvo), type: rule.type || '', source: (rule.fonte || rule.source || ''), parentCol: toLetter(sheetName, rule.colunaCategoria || '') });
    });
  }
}

/* ===================== Coleta de dados nativos da planilha ===================== */
function coletarFormatacaoCondicional_(sheets) {
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i], nm = sh.getName();
    if (SAIDAS.indexOf(nm) !== -1) continue;
    if (sh.isSheetHidden && sh.isSheetHidden()) continue;
    var rules = sh.getConditionalFormatRules() || [];
    for (var r = 0; r < rules.length; r++) {
      var rule = rules[r];
      var ranges = (rule.getRanges() || []).map(function (rg) { return rg.getA1Notation(); }).join(' ; ');
      var tipo = '', valores = '', bc, gc;
      try { bc = rule.getBooleanCondition(); } catch (e) { bc = null; }
      try { gc = rule.getGradientCondition(); } catch (e2) { gc = null; }
      if (bc) { tipo = 'BOOLEAN'; try { valores = String(bc.getCriteriaType()) + ' :: ' + (bc.getValues() || []).join(' ; '); } catch (e3) { valores = 'BOOLEAN'; } }
      else if (gc) { tipo = 'GRADIENT'; valores = 'GRADIENT'; }
      else { tipo = 'DESCONHECIDO'; }
      out.push({ aba: nm, ranges: ranges, tipo: tipo, criterios: valores });
    }
  }
  return out;
}
function salvarFormatacaoCondicional_(ss, lista) {
  var header = ['ABA', 'RANGES', 'TIPO', 'CRITERIOS/VALORES']; var rows = [header];
  for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.aba, it.ranges, it.tipo, it.criterios]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FORMATACAO_CONDICIONAL', rows); return; }
  var sh = ensureSheet_(ss, 'FORMATACAO_CONDICIONAL'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}
function coletarFiltrosAtivos_(sheets) {
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i], nm = sh.getName();
    if (SAIDAS.indexOf(nm) !== -1) continue;
    if (sh.isSheetHidden && sh.isSheetHidden()) continue;
    var filter = sh.getFilter(); if (!filter) continue;
    var fr; try { fr = filter.getRange(); } catch (e) { fr = null; }
    var base = fr ? fr.getA1Notation() : '';
    var lastCol = sh.getLastColumn();
    for (var c = 1; c <= lastCol; c++) {
      var crit = null; try { crit = filter.getColumnFilterCriteria(c); } catch (e2) { crit = null; }
      if (!crit) continue;
      var tipo = '', hidden = []; try { tipo = String(crit.getCriteriaType()); } catch (e3) { } try { hidden = crit.getHiddenValues() || []; } catch (e4) { }
      out.push({ aba: nm, range: base, colunaA1: columnToLetter_(c), criteriaType: tipo, hiddenValues: hidden.join(' | ') });
    }
  }
  return out;
}
function salvarFiltrosAtivos_(ss, lista) {
  var header = ['ABA', 'RANGE', 'COLUNA_A1', 'CRITERIA_TYPE', 'HIDDEN_VALUES']; var rows = [header];
  for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.aba, it.range, it.colunaA1, it.criteriaType, it.hiddenValues]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FILTROS_ATIVOS', rows); return; }
  var sh = ensureSheet_(ss, 'FILTROS_ATIVOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}
function coletarProtecoes_(ss, abasSelecionadas) {
  var out = [];
  var filtroAbas = (abasSelecionadas && abasSelecionadas.length > 0) ? new Set(abasSelecionadas) : null;
  try {
    (ss.getProtections(SpreadsheetApp.ProtectionType.SHEET) || []).forEach(function (p) {
      var abaNome = p.getRange() ? p.getRange().getSheet().getName() : '';
      if (filtroAbas && abaNome && !filtroAbas.has(abaNome)) return; 
      out.push({ tipo: 'SHEET', aba: abaNome, ranges: p.getRange() ? p.getRange().getA1Notation() : '', warningOnly: p.isWarningOnly(), descr: (p.getDescription ? p.getDescription() : '') || '' });
    });
  } catch (e1) { }
  try {
    (ss.getProtections(SpreadsheetApp.ProtectionType.RANGE) || []).forEach(function (pr) {
      var abaNome = pr.getRange() ? pr.getRange().getSheet().getName() : '';
      if (filtroAbas && abaNome && !filtroAbas.has(abaNome)) return;
      out.push({ tipo: 'RANGE', aba: abaNome, ranges: pr.getRange() ? pr.getA1Notation ? pr.getA1Notation() : pr.getRange().getA1Notation() : '', warningOnly: pr.isWarningOnly(), descr: (pr.getDescription ? pr.getDescription() : '') || '' });
    });
  } catch (e2) { }
  return out;
}
function salvarProtecoes_(ss, lista) {
  var header = ['TIPO', 'ABA', 'RANGES', 'WARNING_ONLY', 'DESCRICAO']; var rows = [header];
  for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.tipo, it.aba, it.ranges, it.warningOnly ? 'SIM' : 'NAO', it.descr]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'PROTECOES', rows); return; }
  var sh = ensureSheet_(ss, 'PROTECOES'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}
function coletarNamedRanges_(ss, abasSelecionadas) {
  var out = [];
  var filtroAbas = (abasSelecionadas && abasSelecionadas.length > 0) ? new Set(abasSelecionadas) : null;
  (ss.getNamedRanges() || []).forEach(function (nr) {
    var rg = nr.getRange();
    var abaNome = rg ? rg.getSheet().getName() : '';
    if (filtroAbas && abaNome && !filtroAbas.has(abaNome)) return;
    out.push({ nome: nr.getName(), aba: abaNome, range: rg ? rg.getA1Notation() : '' });
  });
  return out;
}
function salvarRangesNomeados_(ss, lista) {
  var header = ['NOME', 'ABA', 'RANGE']; var rows = [header];
  for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.nome, it.aba, it.range]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'RANGES_NOMEADOS', rows); return; }
  var sh = ensureSheet_(ss, 'RANGES_NOMEADOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}
function coletarGraficos_(sheets) {
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i], nm = sh.getName();
    if (SAIDAS.indexOf(nm) !== -1) continue;
    if (sh.isSheetHidden && sh.isSheetHidden()) continue;
    var charts = sh.getCharts ? (sh.getCharts() || []) : [];
    for (var c = 0; c < charts.length; c++) {
      var ch = charts[c];
      var ranges = (ch.getRanges ? ch.getRanges() : []).map(function (rg) { return rg.getA1Notation(); }).join(' ; ');
      out.push({ aba: nm, tipo: (ch.getChartType ? String(ch.getChartType()) : ''), ranges: ranges, posicao: (ch.getContainerInfo ? JSON.stringify(ch.getContainerInfo()) : '') });
    }
  }
  return out;
}
function salvarGraficos_(ss, lista) {
  var header = ['ABA', 'TIPO', 'RANGES_FONTE', 'CONTAINER_INFO(JSON)']; var rows = [header];
  for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.aba, it.tipo, it.ranges, it.posicao]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'GRAFICOS', rows); return; }
  var sh = ensureSheet_(ss, 'GRAFICOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}
function coletarFilterViewsESlicers_(ss, abasSelecionadas) { //
  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets || !Sheets.Spreadsheets.get) return [];
  var id = ss.getId();
  var resp = Sheets.Spreadsheets.get(id, { fields: 'sheets(properties.title,filterViews,slicers),developerMetadata' });
  var out = []; var arr = (resp && resp.sheets) || [];
  var filtroAbas = (abasSelecionadas && abasSelecionadas.length > 0) ? new Set(abasSelecionadas) : null;
  for (var i = 0; i < arr.length; i++) {
    var s = arr[i]; var title = s.properties && s.properties.title;
    if (filtroAbas && !filtroAbas.has(title)) continue;
    (s.filterViews || []).forEach(function (fv) {
      out.push({ tipo: 'FILTER_VIEW', aba: title, nome: fv.title || '', range: _gridRangeToA1_(fv.range, title), detalhes: JSON.stringify({ criteria: fv.filterSpecs || [] }) });
    });
    (s.slicers || []).forEach(function (sl) {
      out.push({ tipo: 'SLICER', aba: title, nome: (sl.spec && sl.spec.title) || '', range: _gridRangeToA1_(sl.spec && sl.spec.filterCriteriaRange ? sl.spec.filterCriteriaRange : null, title), detalhes: JSON.stringify(sl.spec || {}) });
    });
  }
  (resp.developerMetadata || []).forEach(function (m) {
    out.push({ tipo: 'DEVELOPER_METADATA', aba: '', nome: (m.metadataKey || ''), range: '', detalhes: JSON.stringify(m) });
  });
  return out;
}
function salvarFilterViewsEAfins_(ss, lista) {
  var header = ['TIPO', 'ABA', 'NOME/TITULO', 'RANGE', 'DETALHES_JSON']; var rows = [header];
  for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.tipo, it.aba, it.nome, it.range, it.detalhes]); }
  if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FILTER_VIEWS_E_AFINS', rows); return; }
  var sh = ensureSheet_(ss, 'FILTER_VIEWS_E_AFINS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
}

/* ===================== NOVO: ANÁLISE VISUAL (ESTILOS E TABELAS) ===================== */

function coletarFormatacaoVisual_(sheet, rangeA1Limit) {
  var out = [];
  var name = sheet.getName();
  
  var range;
  if (rangeA1Limit) {
    try { range = sheet.getRange(rangeA1Limit); } 
    catch(e) { return [{ aba: name, tipo: 'ERRO', detalhe: 'Range limite inválido: ' + rangeA1Limit }]; }
  } else {
    range = sheet.getDataRange();
  }
  
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  if (numRows === 0 || numCols === 0) return [];

  // Coleta dados em Lote (Matrizes 2D)
  var bgColors = range.getBackgrounds();
  var fontColors = range.getFontColors();
  var fontFamilies = range.getFontFamilies();
  var fontSizes = range.getFontSizes();
  var fontWeights = range.getFontWeights(); // bold
  var horizontalAligns = range.getHorizontalAlignments();
  var merges = sheet.getMergedRanges(); // Retorna array de Ranges
  
  // A. Inventário de Mesclagens (Simples)
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var endRow = startRow + numRows - 1;
  var endCol = startCol + numCols - 1;

  for (var m = 0; m < merges.length; m++) {
    var mr = merges[m];
    // Verifica intersecção básica com o range limite
    if (mr.getLastRow() < startRow || mr.getRow() > endRow || 
        mr.getLastColumn() < startCol || mr.getColumn() > endCol) {
      continue;
    }
    out.push({
      aba: name,
      tipo: 'MESCLAGEM',
      range: mr.getA1Notation(),
      detalhe: 'Célula principal: ' + mr.getCell(1,1).getValue()
    });
  }

  // B. Inventário de Estilos (Agrupado)
  var styleMap = {};

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var bg = bgColors[r][c];
      var fc = fontColors[r][c];
      var ff = fontFamilies[r][c];
      var fs = fontSizes[r][c];
      var fw = fontWeights[r][c];
      var ha = horizontalAligns[r][c];

      // Ignora células padrão para reduzir ruído
      if (bg === '#ffffff' && fc === '#000000' && fw === 'normal' && ha === 'general') continue;
      
      var key = `BG:${bg}|COR:${fc}|FNT:${ff}|SZ:${fs}|W:${fw}|AL:${ha}`;
      if (!styleMap[key]) {
        styleMap[key] = { count: 0, example: range.getCell(r+1, c+1).getA1Notation() };
      }
      styleMap[key].count++;
    }
  }

  for (var key in styleMap) {
    var item = styleMap[key];
    var desc = key.replace(/BG:/g, 'Fundo: ').replace(/\|COR:/g, ', Texto: ')
                  .replace(/\|FNT:/g, ', Fonte: ').replace(/\|SZ:/g, 'pt, ')
                  .replace(/\|W:/g, ', ').replace(/\|AL:/g, ', Alinh: ');
    
    out.push({
      aba: name,
      tipo: 'ESTILO_AGRUPADO',
      range: 'Qtd: ' + item.count + ' (Ex: ' + item.example + ')',
      detalhe: desc
    });
  }
  
  return out;
}

function coletarIlhasDeDados_(sheet, rangeA1Limit) {
  var out = [];
  var name = sheet.getName();
  
  var rangeTotal;
  try {
    rangeTotal = rangeA1Limit ? sheet.getRange(rangeA1Limit) : sheet.getDataRange();
  } catch(e) { return []; }

  var values = rangeTotal.getValues();
  var numRows = values.length;
  var numCols = values[0].length;
  var visited = new Array(numRows).fill(0).map(function() { return new Array(numCols).fill(false); });

  var startRowOffset = rangeTotal.getRow();
  var startColOffset = rangeTotal.getColumn();

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      if (!visited[r][c] && String(values[r][c]).trim() !== "") {
        var rEnd = r;
        var cEnd = c;
        
        // 1. Expande direita
        while (cEnd < numCols && String(values[r][cEnd]).trim() !== "") {
          cEnd++;
        }
        
        // 2. Expande baixo (tabelas retangulares)
        var isTable = true;
        while (isTable && rEnd < numRows) {
          var hasDataInRow = false;
          for (var k = c; k < cEnd; k++) {
            if (String(values[rEnd][k]).trim() !== "") {
              hasDataInRow = true; break;
            }
          }
          if (hasDataInRow) rEnd++; else isTable = false;
        }
        
        // Marca visitados
        for (var vr = r; vr < rEnd; vr++) {
          for (var vc = c; vc < cEnd; vc++) {
            if (vr < numRows && vc < numCols) visited[vr][vc] = true;
          }
        }

        var a1Start = columnToLetter_(startColOffset + c) + (startRowOffset + r);
        var a1End = columnToLetter_(startColOffset + cEnd - 1) + (startRowOffset + rEnd - 1);
        
        out.push({
          aba: name,
          tipo: 'TABELA_DETECTADA',
          range: a1Start + ':' + a1End,
          detalhe: 'Dimensões: ' + (rEnd - r) + ' linhas x ' + (cEnd - c) + ' colunas'
        });
      }
    }
  }
  return out;
}

function salvarFormatacaoVisual_(ss, lista) {
  var header = ['ABA', 'TIPO', 'RANGE/QTD', 'DETALHES_VISUAIS']; 
  var rows = [header];
  for (var i = 0; i < lista.length; i++) {
    var it = lista[i];
    rows.push([it.aba, it.tipo, it.range, it.detalhe]);
  }
  
  if (MODO_SAIDA_UNICA) { 
    appendSection_(getOutSheet_(), 'ANALISE_VISUAL_ESTILOS', rows); 
    return; 
  }
  
  var sh = ensureSheet_(ss, 'ANALISE_VISUAL_ESTILOS'); 
  sh.clear(); 
  writeTable_(sh, 1, 1, rows, true);
}


/* ===================== Helpers de relatório (visual + export) ===================== */
function beginReport_(ss, exportar) {
  __OUT_SH = ensureSheet_(ss, NOME_SAIDA_UNICA);
  __OUT_SH.clear();

  var head = [[
    'RELATORIO DE ANALISE — ' + (ss.getName() || ''),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
  ]];

  __OUT_SH
    .getRange(1, 1, 1, head[0].length)
    .setValues(head)
    .setFontWeight('bold')
    .setFontSize(12);

  __OUT_SH.getRange(3, 1, 1, 1)
    .setValue('SUMARIO')
    .setFontWeight('bold');

  __OUT_SH.setFrozenRows(5);

  __CUR_ROW = 6;
  __TOC = [];
  __SECTIONS = [];
  __SEC_INDEX = 0;

  // Se não vier nada, cai no fallback global EXPORT_TO_DRIVE
  var deveExportar = (typeof exportar === 'boolean') ? exportar : EXPORT_TO_DRIVE;
  if (deveExportar) {
    exportBegin_(ss);
  }
}

function finalizeReport_(exportar) {
  var gid = __OUT_SH.getSheetId();
  var tocRows = [['Seção', 'Ir']];

  var locale = SpreadsheetApp.getActive().getSpreadsheetLocale() || 'en_US';
  var sep = (locale.indexOf('en_') === 0) ? ',' : ';';

  for (var i = 0; i < __TOC.length; i++) {
    var item = __TOC[i];
    var formula =
      '=HYPERLINK("#gid=' + gid + '&range=A' + item.row + '"' + sep + '"' + item.title + '")';
    tocRows.push([item.title, formula]);
  }

  writeTable_(__OUT_SH, 4, 1, tocRows, true);

  var deveExportar = (typeof exportar === 'boolean') ? exportar : EXPORT_TO_DRIVE;
  if (deveExportar) {
    exportAllSections_();
  }
}
function appendSection_(sheet, titulo, rows) {
  __SEC_INDEX++; __SECTIONS.push({ idx: __SEC_INDEX, title: titulo, rows: rows });
  sheet.getRange(__CUR_ROW, 1).setValue('=== BEGIN:' + titulo + ' ===');
  sheet.getRange(__CUR_ROW + 1, 1, 1, 1).setValue('■ ' + titulo).setFontWeight('bold').setBackground('#efefef');
  var tableStart = __CUR_ROW + 2;
  writeTable_(sheet, tableStart, 1, rows, true);
  sheet.getRange(tableStart + rows.length, 1).setValue('=== END:' + titulo + ' ===');
  __TOC.push({ title: titulo, row: tableStart });
  __CUR_ROW = tableStart + rows.length + 2;
}
function writeTable_(sh, startRow, startCol, rows, withFormat) {
  if (!rows || !rows.length) return; var cols = rows[0].length;
  sh.getRange(startRow, startCol, rows.length, cols).setValues(rows);
  if (withFormat) {
    sh.getRange(startRow, startCol, 1, cols).setFontWeight('bold').setBackground('#f5f5f5');
    sh.getRange(startRow, startCol, rows.length, cols).setWrap(true);
    try { sh.getRange(startRow, startCol, rows.length, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) { }
    try { if (!sh.getFilter()) sh.getRange(startRow, startCol, rows.length, cols).createFilter(); } catch (e) { }
    sh.autoResizeColumns(startCol, cols);
  }
}

/* ===================== Export: YAML + CSV + MD ===================== */
function exportBegin_(ss) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var root = DriveApp.getFoldersByName(EXPORT_FOLDER_BASENAME);
  var baseFolder = root.hasNext() ? root.next() : DriveApp.createFolder(EXPORT_FOLDER_BASENAME);
  var folder = baseFolder.createFolder(EXPORT_FOLDER_BASENAME + '_' + stamp);
  __EXPORT = {
    folder: folder,
    folderId: folder.getId(),
    pathPrefix: EXPORT_FOLDER_BASENAME + '_' + stamp + '/',
    manifest: {
      spreadsheet: { id: ss.getId(), name: ss.getName(), url: ss.getUrl() },
      generatedAt: stamp,
      sections: []
    }
  };
  var readme = ['# README_AI', '', '- Leia 00_MANIFEST.yml primeiro.', '- Depois, carregue apenas os CSV/MD das secções relevantes.', '- No relatório em planilha, use marcadores === BEGIN:/END: para copiar blocos.', ''].join('\n');
  folder.createFile('00_README_AI.md', readme, MimeType.PLAIN_TEXT);
}
function exportAllSections_() {
  if (!__EXPORT) return; var folder = __EXPORT.folder; var man = __EXPORT.manifest;
  for (var i = 0; i < __SECTIONS.length; i++) {
    var S = __SECTIONS[i];
    var slug = sanitizeSlug_(S.title);
    var idx = pad2_(S.idx);
    var csvName = idx + '_' + slug + '.csv';
    var csv = csvEncodeRows_(S.rows);
    folder.createFile(csvName, csv, MimeType.PLAIN_TEXT);
    var mdBase = idx + '_' + slug;
    var mdBase = idx + '_' + slug;
    var md = mdFromRows_(S.title, csv);    var parts = splitText_(md, EXPORT_MD_SPLIT_THRESHOLD);
    var mdFiles = [];
    for (var p = 0; p < parts.length; p++) {
      var nm = mdBase + (parts.length > 1 ? ('.part' + (p + 1)) : '') + '.md';
      folder.createFile(nm, parts[p], MimeType.PLAIN_TEXT);
      mdFiles.push(nm);
    }
    man.sections.push({ idx: S.idx, title: S.title, csv: csvName, md: mdFiles, rows: S.rows.length, cols: (S.rows[0] || []).length });
  }
  var yaml = yamlFromManifest_(man);
  folder.createFile('00_MANIFEST.yml', yaml, MimeType.PLAIN_TEXT);
}
function csvEncodeRows_(rows) {
  return rows.map(function (row) {
    return row.map(function (v) {
      if (v == null) return '';
      var s = String(v);

      // Se tiver aspas, vírgula ou quebra de linha -> precisa de aspas duplas e escape interno
      if (/[",\n\r]/.test(s)) {
        return '"' + s.replace(/"/g, '""') + '"';
      }

      return s;
    }).join(',');
  }).join('\n');
}

function mdFromRows_(title, csvContent) {
  var header = '---\nsection: ' + title + '\n---\n\n';
  var body   = '```csv\n' + csvContent + '\n```\n';
  return header + body;
}
function splitText_(txt, max) { if (txt.length <= max) return [txt]; var out = []; for (var i = 0; i < txt.length; i += max) { out.push(txt.substring(i, Math.min(txt.length, i + max))); } return out; }
function yamlFromManifest_(m) {
  function esc(s) { return String(s).replace(/\n/g, '\\n'); }
  var lines = [];
  lines.push('spreadsheet:');
  lines.push('  id: ' + esc(m.spreadsheet.id));
  lines.push('  name: ' + esc(m.spreadsheet.name));
  lines.push('  url: ' + esc(m.spreadsheet.url));
  lines.push('generatedAt: ' + esc(m.generatedAt));
  lines.push('sections:');
  for (var i = 0; i < m.sections.length; i++) {
    var s = m.sections[i];
    lines.push('  - idx: ' + s.idx);
    lines.push('    title: ' + esc(s.title));
    lines.push('    rows: ' + s.rows);
    lines.push('    cols: ' + s.cols);
    lines.push('    csv: ' + s.csv);
    if (s.md && s.md.length) {
      lines.push('    md:');
      for (var j = 0; j < s.md.length; j++) { lines.push('      - ' + s.md[j]); }
    }
  }
  return lines.join('\n') + '\n';
}
function sanitizeSlug_(s) { return String(s).trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, ''); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/* ===================== Headers ===================== */
function obterHeaderRow_(sheet) {
  var g = readGlobalHeaderRowFromSelection_();
  if (g && g >= 1) return g;
  var perSheet = readPerSheetHeaderRowFromSelection_(sheet.getName());
  if (perSheet && perSheet >= 1) return perSheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 1;
  var maxScan = Math.min(HEADER_SCAN_LIMIT, Math.max(1, lastRow));
  var best = { row: 1, score: -1 };
  for (var r = 1; r <= maxScan; r++) {
    var vals = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    var belowVals = (r < lastRow) ? sheet.getRange(r + 1, 1, 1, Math.min(lastCol, 20)).getDisplayValues()[0] : [];
    var sc = scoreHeaderRow_(vals, belowVals);
    if (sc > best.score) { best = { row: r, score: sc }; }
  }
  if (best.score >= HEADER_SCORE_THRESHOLD) return best.row;
  return best.row || 1;
}
function scoreHeaderRow_(vals, belowVals) {
  var n = vals.length; var nonEmpty = 0, texty = 0, digits = 0; var set = {};
  for (var i = 0; i < n; i++) {
    var v = String(vals[i] || '').trim();
    if (v !== '') {
      nonEmpty++; set[v] = 1;
      if (/^[0-9.,:\-\/]+$/.test(v)) digits++; else texty++;
    }
  }
  var pNon = nonEmpty / n; var pText = texty / Math.max(1, (texty + digits)); var uniq = Object.keys(set).length / Math.max(1, nonEmpty);
  var keywords = 0; var kwRe = /(^|\b)(ID|DATA|DT|VALOR|PRECO|PREÇO|EMAIL|E-MAIL|STATUS|CPF|CNPJ)($|\b)/i;
  for (var i2 = 0; i2 < n; i2++) { if (kwRe.test(String(vals[i2] || ''))) keywords++; }
  var belowNums = 0; for (var j = 0; j < belowVals.length; j++) { if (/^-?\d+[\d.,\/:-]*$/.test(String(belowVals[j] || ''))) belowNums++; }
  var belowRatio = belowVals.length ? (belowNums / belowVals.length) : 0;
  return (pNon * 0.6) + (pText * 0.4) + (uniq * 0.3) + (Math.min(1, keywords / 4) * 0.3) + (belowRatio * 0.2);
}
function buildHeaderIndex_(headers) {
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var h = (headers[c] || '').toString().trim(); if (!h) continue;
    var key = h.toUpperCase(); if (!map[key]) map[key] = columnToLetter_(c + 1);
  }
  return map;
}

/* ===================== Leitura da guia de seleção (header/abas) ===================== */
var NOME_SELECAO_ABAS = 'SELECIONAR_ABAS_TEMP';
function getSelectionSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(NOME_SELECAO_ABAS) || null;
}
function readGlobalHeaderRowFromSelection_() {
  var sh = getSelectionSheet_(); if (!sh) return null;
  var v = Number(sh.getRange('E2').getValue()); return isFinite(v) && v >= 1 ? Math.floor(v) : null;
}
function readPerSheetHeaderRowFromSelection_(sheetName) {
  var sh = getSelectionSheet_(); if (!sh) return null;
  var lastRow = sh.getLastRow(); if (lastRow < 2) return null;
  var vals = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    var nm = String(vals[i][1] || '').trim();
    if (nm === sheetName) {
      var num = Number(vals[i][2]); return isFinite(num) && num >= 1 ? Math.floor(num) : null;
    }
  }
  return null;
}

/* ====================================================================
 * UI SIMPLES (PAINEL DE CONTROLO): Selecionar abas e header na própria guia
 * ==================================================================== */
function prepararSelecaoAbasTemp() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(NOME_SELECAO_ABAS) || ss.insertSheet(NOME_SELECAO_ABAS);
  sh.clear();
  sh.showSheet();

  sh.getRange('A1:E1').setValues([['Incluir?', 'Aba', 'Header Row (Opcional)', 'Limite Leitura (Ex: A1:F50)', 'Analisar Estilo?']])
    .setFontWeight('bold')
    .setBackground('#e0e0e0');

  var sheets = ss.getSheets();
  var rows = [];
  for (var i = 0; i < sheets.length; i++) {
    var nm = sheets[i].getName();
    if (SAIDAS.indexOf(nm) !== -1 || nm === NOME_SELECAO_ABAS) continue;
    var sug = _heuristicaHeaderRowNoPrompt_(sheets[i]);
    // Incluir, Nome, Sugestão Header, Limite Vazio, Checkbox Estilo False
    rows.push([true, nm, sug, "", false]);
  }

  if (rows.length) {
    var rangeAbas = sh.getRange(2, 1, rows.length, 5);
    rangeAbas.setValues(rows);
    
    // Checkbox para "Incluir" (Col A)
    sh.getRange(2, 1, rows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    // Checkbox para "Analisar Estilo" (Col E)
    sh.getRange(2, 5, rows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }

  // Configuração do Preset (Col G/H para não encavalar)
  sh.getRange('G1').setValue('Preset de Análise').setFontWeight('bold');
  const presetCell = sh.getRange('G2');
  presetCell.setBackground('#fff9c4').setNote('Selecione o tipo de análise.');
  const presetNames = Object.keys(PRESETS);
  const presetListRange = sh.getRange(2, 8, presetNames.length, 1); // Lista escondida na H
  presetListRange.setValues(presetNames.map(name => [name]));
  var dvDropdown = SpreadsheetApp.newDataValidation().requireValueInRange(presetListRange, true).setAllowInvalid(false).build();
  presetCell.setDataValidation(dvDropdown);
  presetCell.setValue('COMPLETO');

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 5);
  sh.hideColumns(8); // Esconde a lista auxiliar
  SpreadsheetApp.setActiveSheet(sh);
  SpreadsheetApp.getActive().toast('Painel atualizado. Defina limites (Col D) e ative estilos (Col E) se necessário.');
}


function executarAnalisePeloPainel() {
  const shSel = SpreadsheetApp.getActive().getSheetByName(NOME_SELECAO_ABAS);
  if (!shSel) {
    SpreadsheetApp.getUi().alert('Painel não encontrado. Por favor, execute "prepararSelecaoAbasTemp" primeiro.');
    return;
  }
  const presetSelecionado = shSel.getRange('G2').getValue();
  if (!presetSelecionado || !PRESETS[presetSelecionado]) {
    SpreadsheetApp.getUi().alert(`Preset "${presetSelecionado}" inválido. Selecione uma opção válida.`);
    return;
  }
  SpreadsheetApp.getActive().toast(`Iniciando análise com o preset: "${presetSelecionado}"...`, 'Progresso', 30);
  runAnaliseComSelecao(presetSelecionado);
}

function runAnaliseComSelecao(presetName) {
  var ss = SpreadsheetApp.getActive();
  var mapaConfig = _lerSelecaoAbasTemp_(); // Agora retorna objeto {NomeAba: {incluir, limite, analisarEstilo}}
  var abasSelecionadas = Object.keys(mapaConfig);
  
  var snap = _snapshotVisibilidade_(ss);
  try {
    if (abasSelecionadas.length > 0) {
      _ocultarExceto_(ss, new Set(abasSelecionadas));
    } else {
      SpreadsheetApp.getActive().toast('Nenhuma aba marcada. Verifique o painel.');
      return;
    }
    documentarEstruturaCompleta(presetName, mapaConfig);
  } finally {
    _restaurarVisibilidade_(ss, snap);
  }
}

function _lerSelecaoAbasTemp_() {
  var sh = getSelectionSheet_(); if (!sh) return {};
  var lastRow = sh.getLastRow(); if (lastRow < 2) return {};
  var vals = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  var configMap = {};
  
  for (var i = 0; i < vals.length; i++) {
    // Col A: Incluir?
    if (vals[i][0] === true) {
      var nm = String(vals[i][1] || '').trim();
      if (nm && SAIDAS.indexOf(nm) == -1) {
        configMap[nm] = {
          incluir: true,
          headerRow: vals[i][2],
          limite: String(vals[i][3] || '').trim(), // Col D: Limite
          analisarEstilo: vals[i][4] === true      // Col E: Checkbox Estilo
        };
      }
    }
  }
  return configMap;
}
function _snapshotVisibilidade_(ss) {
  var snap = {}; var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i]; var nm = sh.getName(); snap[nm] = (sh.isSheetHidden && sh.isSheetHidden());
  }
  return snap;
}
function _restaurarVisibilidade_(ss, snap) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i]; var nm = sh.getName(); var wasHidden = !!snap[nm];
    try { if (wasHidden) sh.hideSheet(); else sh.showSheet(); } catch (e) { }
  }
}
function _ocultarExceto_(ss, selectedSet) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i]; var nm = sh.getName();
    if (SAIDAS.indexOf(nm) !== -1 || nm === NOME_SELECAO_ABAS) continue;
    try { if (!selectedSet.has(nm)) sh.hideSheet(); else sh.showSheet(); } catch (e) { }
  }
}

function _heuristicaHeaderRowNoPrompt_(sheet) {
  var lastRow = sheet.getLastRow(); var lastCol = sheet.getLastColumn(); if (lastCol < 1) return 1;
  var maxScan = Math.min(HEADER_SCAN_LIMIT, Math.max(1, lastRow));
  var best = { row: 1, score: -1 };
  for (var r = 1; r <= maxScan; r++) {
    var vals = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    var belowVals = (r < lastRow) ? sheet.getRange(r + 1, 1, 1, Math.min(lastCol, 20)).getDisplayValues()[0] : [];
    var sc = scoreHeaderRow_(vals, belowVals);
    if (sc > best.score) { best = { row: r, score: sc }; }
  }
  return best.row;
}

/* ===================== Utilitários ===================== */
function ensureSheet_(ss, name) { var sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name); return sh; }
function getOutSheet_() { return __OUT_SH || ensureSheet_(SpreadsheetApp.getActive(), NOME_SAIDA_UNICA); }
function columnToLetter_(column) { var temp, letter = ''; while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; } return letter; }
function _pct(n) { return Math.round((n || 0) * 100) + '%'; }
function _matrizTemAlgum_(mat) { if (!mat || !mat.length) return false; for (var r = 0; r < mat.length; r++) { for (var c = 0; c < mat[r].length; c++) { if (mat[r][c]) return true; } } return false; }
function _matrizTemAlgumStringNaoVazia_(mat) {
  if (!mat || !mat.length) return false;
  for (var r = 0; r < mat.length; r++) {
    for (var c = 0; c < mat[r].length; c++) {
      if (String(mat[r][c] || '').trim() !== '') return true;
    }
  }
  return false;
}
function calcularDensidade_(sheet, headerRow) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow <= headerRow || lastCol < 1) return 0;
  var rows = Math.min(lastRow - headerRow, 200), cols = Math.min(lastCol, 50);
  var rg = sheet.getRange(headerRow + 1, 1, rows, cols);
  var vals = rg.getDisplayValues();
  var total = rows * cols, filled = 0;
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[0].length; c++) {
      if (String(vals[r][c]).trim() !== '') filled++;
    }
  }
  return total ? (filled / total) : 0;
}
function inferirTipoPeloHeader_(h) {
  h = (h || '').toString().trim().toUpperCase(); if (!h) return '';
  if (/\bDATA\b|\bDT\b/.test(h)) return 'data';
  if (/VALOR|PRECO|PREÇO|TOTAL|MONTANTE|R\$|\bVLR\b/.test(h)) return 'numero';
  if (/CPF|CNPJ/.test(h)) return 'documento';
  if (/EMAIL|E-MAIL/.test(h)) return 'email';
  if (/ATIVO|APROVADO|SIM|NAO|NÃO|CHECK|STATUS/.test(h)) return 'booleano';
  if (/\bID\b/.test(h)) return 'identificador';
  return 'texto';
}
function parseColRefToLetter_(sheetName, ref) {
  if (ref == null) return '';
  if (typeof ref === 'number') return columnToLetter_(ref);
  var s = String(ref).trim(); if (!s) return '';
  if (/^\d+$/.test(s)) return columnToLetter_(parseInt(s, 10));
  var m = s.match(/^([A-Za-z]+)(:\d.*)?$/); if (m) return m[1].toUpperCase();
  var idx = __HEADER_INDEX[(sheetName || '').toUpperCase()] || {};
  var L = idx[s.toUpperCase()] || ''; return L || '';
}
function _gridRangeToA1_(gr, title) {
  if (!gr) return '';
  function colToA1(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); } return s; }
  var r1 = (gr.startRowIndex || 0) + 1, c1 = (gr.startColumnIndex || 0) + 1;
  var r2 = (gr.endRowIndex == null ? r1 : gr.endRowIndex);
  var c2 = (gr.endColumnIndex == null ? c1 : gr.endColumnIndex);
  return (title ? "'" + title + "'!" : '') + colToA1(c1) + r1 + ':' + colToA1(c2) + r2;
}
function log_(msg) { try { Logger.log(msg); } catch (e) { } }