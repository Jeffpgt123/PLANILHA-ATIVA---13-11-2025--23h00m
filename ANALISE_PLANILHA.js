// /**
//  * ANALISE_PLANILHA — Inventário imediato da estrutura da planilha e do código que atua sobre ela.
//  * Saída única apresentável + export “IA-Ready” (CSV/MD/YAML) + seleção de abas e cabeçalho via guia.
//  *
//  * VERSÃO MODIFICADA: Inclui sistema de PRESETS e ANÁLISE DE LÓGICA DA APLICAÇÃO.
//  *
//  * Como usar:
//  * 1) Execute prepararSelecaoAbasTemp() no editor GAS → configure o painel na planilha.
//  * 2) Execute executarAnalisePeloPainel() no editor GAS → o relatório será gerado.
//  */

// /** ===================== Configurações ===================== */
// const MODO_SAIDA_UNICA = true;
// const NOME_SAIDA_UNICA = 'RELATORIO_ANALISE';

// const HEADER_SCAN_LIMIT = 20;
// const HEADER_ASK_WHEN_UNCERTAIN = false;
// const HEADER_SCORE_THRESHOLD = 1.5;

// const EXPORT_TO_DRIVE = true;
// const EXPORT_FOLDER_BASENAME = 'RELATORIO_ANALISE_EXPORT';
// const EXPORT_MD_SPLIT_THRESHOLD = 12000;

// /**
//  * =======================================================================================
//  * BLOCO DE MODIFICAÇÃO (FASE 1): CONFIGURAÇÃO DE PRESETS DE ANÁLISE
//  * Centraliza todas as configurações de análise num único local, facilitando a
//  * manutenção e a criação de novos modos de análise no futuro.
//  * =======================================================================================
//  */
// const PRESETS = {
//   COMPLETO: {
//     documentacaoGeral: true,
//     formatoExpandido: true,
//     triggers: true,
//     logicaDeAplicacao: true,
//     validacoes: true,
//     monitoramentosCONFIG: true,
//     formatacaoCondicional: true,
//     filtros: true,
//     protecoes: true,
//     rangesNomeados: true,
//     graficos: true,
//     recursosAvancados: true
//   },
//   FOCO_CODIGO_E_REGRAS: {
//     documentacaoGeral: true,
//     formatoExpandido: false,
//     triggers: true,
//     logicaDeAplicacao: true,
//     validacoes: true,
//     monitoramentosCONFIG: true,
//     formatacaoCondicional: true,
//     filtros: false,
//     protecoes: false,
//     rangesNomeados: true,
//     graficos: false,
//     recursosAvancados: false
//   },
//   ESTRUTURA_RAPIDA: {
//     documentacaoGeral: true,
//     formatoExpandido: true,
//     triggers: false,
//     logicaDeAplicacao: false,
//     validacoes: true,
//     monitoramentosCONFIG: false,
//     formatacaoCondicional: false,
//     filtros: true,
//     protecoes: true,
//     rangesNomeados: true,
//     graficos: false,
//     recursosAvancados: false
//   }
// };

// /** Nomes de guias de saída (ignorar na varredura) */
// var SAIDAS = [
//   'DOCUMENTACAO_GERAL', 'FORMATO_EXPANDIDO', 'TRIGGERS_PROJETO', 'VALIDACOES_POR_COLUNA',
//   'MONITORAMENTOS_DECLARADOS', 'FORMATACAO_CONDICIONAL', 'FILTROS_ATIVOS', 'PROTECOES',
//   'RANGES_NOMEADOS', 'GRAFICOS', 'FILTER_VIEWS_E_AFINS', 'SELECIONAR_ABAS_TEMP', 'LOGICA_DA_APLICAÇÃO'
// ];
// if (MODO_SAIDA_UNICA && SAIDAS.indexOf(NOME_SAIDA_UNICA) === -1) SAIDAS.push(NOME_SAIDA_UNICA);

// /** ===================== Estado interno ===================== */
// var __OUT_SH = null;
// var __CUR_ROW = 6;
// var __TOC = [];
// var __HEADER_INDEX = {};
// var __SECTIONS = [];
// var __SEC_INDEX = 0;
// var __EXPORT = null;
// var __HEADER_ROW_GLOBAL = null;

// /* ======================================================================
//  * EXECUÇÃO GERAL (MODIFICADO PARA USAR PRESETS)
//  * ====================================================================== */
// function documentarEstruturaCompleta(presetName, abasSelecionadas) {
//   const presetAtivo = PRESETS[presetName] || PRESETS.COMPLETO;
//   var ss = SpreadsheetApp.getActive();
//   var sheets = ss.getSheets();
//   var analisadas = [];

//   if (MODO_SAIDA_UNICA) beginReport_(ss);

//   __HEADER_ROW_GLOBAL = readGlobalHeaderRowFromSelection_();

//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i];
//     var name = sh.getName();
//     if (SAIDAS.indexOf(name) !== -1) continue;
//     if (sh.isSheetHidden && sh.isSheetHidden()) continue;
//     var info = documentarAbaCompleta_(sh);
//     analisadas.push(info);
//     __HEADER_INDEX[name.toUpperCase()] = buildHeaderIndex_(info.headers);
//   }

//   if (presetAtivo.documentacaoGeral) salvarDocumentacaoGeral_(ss, analisadas);
//   if (presetAtivo.formatoExpandido) salvarFormatoExpandido_(ss, analisadas);
//   if (presetAtivo.triggers) { try { salvarTriggersProjeto_(ss, inventariarTriggersProjeto_()); } catch (e1) { log_('TRIGGERS_PROJETO: ' + (e1 && e1.message)); } }
//   if (presetAtivo.logicaDeAplicacao) { try { salvarLogicaDeAplicacao_(ss, coletarLogicaDeAplicacao_()); } catch (e_logic) { log_('LOGICA_DA_APLICAÇÃO: ' + (e_logic && e_logic.message)); } }

//   if (presetAtivo.validacoes) {
//     try {
//       var mapaValid = [];
//       for (var j = 0; j < sheets.length; j++) {
//         var s2 = sheets[j]; var nm = s2.getName();
//         if (SAIDAS.indexOf(nm) !== -1) continue;
//         if (s2.isSheetHidden && s2.isSheetHidden()) continue;
//         mapaValid = mapaValid.concat(listarValidacoesPorColuna_(s2));
//       }
//       salvarValidacoesPorColuna_(ss, mapaValid);
//     } catch (e2) { log_('VALIDACOES_POR_COLUNA: ' + (e2 && e2.message)); }
//   }

//   if (presetAtivo.monitoramentosCONFIG) { try { salvarMonitoramentosDeclarados_(ss, coletarMonitoramentosDeclarados_CONFIG_()); } catch (e3) { log_('MONITORAMENTOS_DECLARADOS: ' + (e3 && e3.message)); } }
//   if (presetAtivo.formatacaoCondicional) { try { salvarFormatacaoCondicional_(ss, coletarFormatacaoCondicional_(sheets)); } catch (e4) { log_('FORMATACAO_CONDICIONAL: ' + (e4 && e4.message)); } }
//   if (presetAtivo.filtros) { try { salvarFiltrosAtivos_(ss, coletarFiltrosAtivos_(sheets)); } catch (e5) { log_('FILTROS_ATIVOS: ' + (e5 && e5.message)); } }
// if (presetAtivo.protecoes) { try { salvarProtecoes_(ss, coletarProtecoes_(ss, abasSelecionadas)); } catch (e6) { log_('PROTECOES: ' + (e6 && e6.message)); } }
//   if (presetAtivo.rangesNomeados) { try { salvarRangesNomeados_(ss, coletarNamedRanges_(ss, abasSelecionadas)); } catch (e7) { log_('RANGES_NOMEADOS: ' + (e7 && e7.message)); } }
//   if (presetAtivo.graficos) { try { salvarGraficos_(ss, coletarGraficos_(sheets)); } catch (e8) { log_('GRAFICOS: ' + (e8 && e8.message)); } }
//   if (presetAtivo.recursosAvancados) { try { var fv = coletarFilterViewsESlicers_(ss, abasSelecionadas); if (fv && fv.length) salvarFilterViewsEAfins_(ss, fv); } catch (e9) { log_('FILTER_VIEWS_E_AFINS indisponível.'); } }

//   if (MODO_SAIDA_UNICA) finalizeReport_();
// }

// /* ===================== Núcleo por aba ===================== */
// function documentarAbaCompleta_(sheet) {
//   var name = sheet.getName();
//   var lastRow = sheet.getLastRow();
//   var lastCol = sheet.getLastColumn();
//   var headerRow = obterHeaderRow_(sheet);
//   var headers = [];
//   if (lastCol > 0 && lastRow >= headerRow) headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
//   var temFiltro = !!sheet.getFilter();
//   var qtdGraficos = sheet.getCharts ? ((sheet.getCharts() || []).length) : 0;
//   var amostraValid = false, amostraFormulas = false;
//   if (lastRow > headerRow && lastCol > 0) {
//     var amRows = Math.min(5, lastRow - headerRow), amCols = Math.min(8, lastCol);
//     var rAmostra = sheet.getRange(headerRow + 1, 1, amRows, amCols);
//     var dvs = rAmostra.getDataValidations(); var fms = rAmostra.getFormulas();
//     amostraValid = _matrizTemAlgum_(dvs); amostraFormulas = _matrizTemAlgumStringNaoVazia_(fms);
//   }
//   var densidade = calcularDensidade_(sheet, headerRow);
//   return {
//     aba: name, linhas: lastRow, colunas: lastCol, headerRow: headerRow, headers: headers,
//     dadosPercentual: densidade, temFiltro: temFiltro, temValidacaoAmostral: amostraValid,
//     temFormulaAmostral: amostraFormulas, graficosQtde: qtdGraficos
//   };
// }

// /* ===================== Saídas ===================== */
// function salvarDocumentacaoGeral_(ss, arr) {
//   var header = ['ABA', 'LINHAS', 'COLUNAS', 'DADOS_%', 'FILTROS', 'VALID.', 'FORMUL.', 'GRAFICOS', 'HEADER_ROW'];
//   var rows = [header];
//   for (var i = 0; i < arr.length; i++) {
//     var it = arr[i];
//     rows.push([it.aba, it.linhas, it.colunas, _pct(it.dadosPercentual), it.temFiltro ? 'SIM' : 'NAO',
//     it.temValidacaoAmostral ? 'SIM' : 'NAO', it.temFormulaAmostral ? 'SIM' : 'NAO',
//       it.graficosQtde || 0, it.headerRow]);
//   }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'DOCUMENTACAO_GERAL', rows); return; }
//   var sh = ensureSheet_(ss, 'DOCUMENTACAO_GERAL'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// function salvarFormatoExpandido_(ss, arr) {
//   var header = ['ABA', 'COLUNA_A1', 'HEADER', 'TIPO_INFERIDO'];
//   var rows = [header];
//   for (var i = 0; i < arr.length; i++) {
//     var it = arr[i];
//     for (var c = 0; c < it.colunas; c++) {
//       var h = it.headers[c] || ''; rows.push([it.aba, columnToLetter_(c + 1), h, inferirTipoPeloHeader_(h)]);
//     }
//   }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FORMATO_EXPANDIDO', rows); return; }
//   var sh = ensureSheet_(ss, 'FORMATO_EXPANDIDO'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// /* ===================== Triggers ===================== */
// function inventariarTriggersProjeto_() {
//   var triggers = ScriptApp.getProjectTriggers(); var out = [];
//   for (var i = 0; i < triggers.length; i++) {
//     var t = triggers[i];
//     out.push({
//       handler: t.getHandlerFunction ? t.getHandlerFunction() : '',
//       eventType: t.getEventType ? String(t.getEventType()) : '',
//       triggerSource: t.getTriggerSource ? String(t.getTriggerSource()) : '',
//       triggerSourceId: t.getTriggerSourceId ? String(t.getTriggerSourceId()) : ''
//     });
//   }
//   return out;
// }
// function salvarTriggersProjeto_(ss, registros) {
//   var header = ['HANDLER', 'EVENT_TYPE', 'TRIGGER_SOURCE', 'TRIGGER_SOURCE_ID']; var rows = [header];
//   for (var i = 0; i < registros.length; i++) { var r = registros[i]; rows.push([r.handler, r.eventType, r.triggerSource, r.triggerSourceId]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'TRIGGERS_PROJETO', rows); return; }
//   var sh = ensureSheet_(ss, 'TRIGGERS_PROJETO'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// /* ===================== Validações ===================== */
// function listarValidacoesPorColuna_(sheet) {
//   var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn(); if (lastCol < 1) return [];
//   var headerRow = obterHeaderRow_(sheet);
//   var headers = lastCol > 0 && lastRow >= headerRow ? sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0] : [];
//   if (lastRow <= headerRow) {
//     var vazio = []; for (var c = 0; c < lastCol; c++) {
//       vazio.push({ aba: sheet.getName(), colunaA1: columnToLetter_(c + 1), header: headers[c] || '', temValidacao: false, criterios: '', criteriosArgs: '' });
//     }
//     return vazio;
//   }
//   var range = sheet.getRange(headerRow + 1, 1, Math.max(0, lastRow - headerRow), lastCol);
//   var dvs = range.getDataValidations(); var results = [];
//   for (var c = 0; c < lastCol; c++) {
//     var tipos = {}, argsResumo = {}, tem = false;
//     for (var r = 0; r < dvs.length; r++) {
//       var dv = dvs[r][c]; if (!dv) continue; tem = true;
//       var tipo = 'DESCONHECIDO', args = [];
//       try { tipo = String(dv.getCriteriaType()); } catch (e) { }
//       try { args = dv.getCriteriaValues() || []; } catch (e2) { }
//       tipos[tipo] = true;
//       var argStr = (args || []).map(function (a) {
//         if (a == null) return '';
//         if (a.getA1Notation) return a.getA1Notation();
//         if (a.map && a.forEach) return '[array]';
//         if (Object.prototype.toString.call(a) === '[object Date]') return Utilities.formatDate(a, Session.getScriptTimeZone(), 'dd/MM/yyyy');
//         return String(a);
//       }).slice(0, 3).join(' ; ');
//       if (argStr) argsResumo[argStr] = true;
//     }
//     results.push({
//       aba: sheet.getName(),
//       colunaA1: columnToLetter_(c + 1),
//       header: headers[c] || '',
//       temValidacao: tem,
//       criterios: Object.keys(tipos).sort().join(' | '),
//       criteriosArgs: Object.keys(argsResumo).sort().join(' | ')
//     });
//   }
//   return results;
// }
// function salvarValidacoesPorColuna_(ss, lista) {
//   var header = ['ABA', 'COLUNA_A1', 'HEADER', 'TEM_VALIDACAO', 'CRITERIOS', 'ARGS_RESUMO']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) {
//     var m = lista[i]; rows.push([m.aba, m.colunaA1, m.header, m.temValidacao ? 'SIM' : 'NAO', m.criterios || '', m.criteriosArgs || '']);
//   }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'VALIDACOES_POR_COLUNA', rows); return; }
//   var sh = ensureSheet_(ss, 'VALIDACOES_POR_COLUNA'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// /* ===================== Monitoramentos declarados (CONFIG original) ===================== */
// function coletarMonitoramentosDeclarados_CONFIG_() {
//   var lista = [];
//   try {
//     var norm = normalizeConfig_();
//     (norm.timestamp || []).forEach(function (x) {
//       lista.push({ origem: x.origem || 'CONFIG.TIMESTAMP', aba: x.sheet || '', alvo: x.col || '', handlerOuModulo: 'TimestampManager', detalhe: '' });
//     });
//     (norm.moveRules || []).forEach(function (x) {
//       lista.push({ origem: x.origem || 'CONFIG.MOVE_RULES', aba: x.sheet || '', alvo: x.col || '', handlerOuModulo: 'RowManager', detalhe: JSON.stringify(x.map || {}) });
//     });
//     (norm.dropdowns || []).forEach(function (x) {
//       lista.push({ origem: x.origem || 'CONFIG.DROPDOWNS', aba: x.sheet || '', alvo: x.col || '', handlerOuModulo: 'DropdownManager', detalhe: JSON.stringify({ type: x.type || '', source: x.source || '', parent: x.parentCol || '' }) });
//     });
//     if (norm.installments && norm.installments.cols) {
//       var S = norm.installments.sheet || ''; var cols = norm.installments.cols;
//       if (cols.contract) lista.push({ origem: 'CONFIG.ACOMPANHAMENTO_PARCELAS', aba: S, alvo: cols.contract, handlerOuModulo: 'ParcelasManager', detalhe: 'COLUNA_DATA_CONTRATO' });
//       if (cols.dueDay) lista.push({ origem: 'CONFIG.ACOMPANHAMENTO_PARCELAS', aba: S, alvo: cols.dueDay, handlerOuModulo: 'ParcelasManager', detalhe: 'COLUNA_DIA_VENCTO' });
//       (cols.list || []).forEach(function (letter) {
//         lista.push({ origem: 'CONFIG.ACOMPANHAMENTO_PARCELAS.COLUNAS_PARCELAS', aba: S, alvo: letter, handlerOuModulo: 'ParcelasManager', detalhe: '' });
//       });
//     }
//     (norm.handlers || []).forEach(function (x) {
//       lista.push({ origem: x.origem || 'CONFIG.HANDLERS', aba: x.sheet || '', alvo: x.target || '', handlerOuModulo: x.handler || '', detalhe: '' });
//     });
//   } catch (e) { /* silencioso */ }
//   return lista;
// }

// /* =======================================================================================
//  * BLOCO DE MODIFICAÇÃO (FASE 2): O "CÉREBRO" DA ANÁLISE DE LÓGICA
//  * =======================================================================================
//  */

// function obterNomeSemantico_(sheetName, indice) {
//   const sheetConfig = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) ? CONFIG.SHEETS[sheetName] : null;
//   if (!sheetConfig || !sheetConfig.colunas) return null;
//   for (const [nome, idx] of Object.entries(sheetConfig.colunas)) {
//     if (idx === indice) return nome;
//   }
//   return null;
// }

// function formatarComportamentoHandler_(handler, sheetName) {
//   const { modulo, options } = handler;
//   if (!options) return 'Nenhuma opção configurada.';
//   try {
//     switch (modulo) {
//       case 'Module_VALIDACAO':
//         const validacoes = Object.entries(options.colunasValidacao || {})
//           .map(([col, tipo]) => `valida a coluna ${columnToLetter_(parseInt(col))} (${obterNomeSemantico_(sheetName, parseInt(col)) || 'N/D'}) como '${tipo}'`)
//           .join('; ');
//         return `Aplica as seguintes validações: ${validacoes || 'Nenhuma especificada'}.`;
//       case 'Module_NUMERADOR':
//         const colDestino = options.colunaDestino;
//         return `Gera um ID com prefixo "${options.prefixo}" na coluna ${columnToLetter_(colDestino)} (${obterNomeSemantico_(sheetName, colDestino) || 'N/D'}).`;
//       case 'Module_RELACIONAMENTO':
//         const campos = (options.campos || []).map(c =>
//           `ao editar ${columnToLetter_(c.colunaDisparo)}, busca na aba "${c.abaFonte}" pela chave em ${columnToLetter_(c.colunaChaveFonte)}`
//         ).join('; ');
//         return `Preenche dados automaticamente (PROCV automático): ${campos}.`;
//       case 'Module_DROPDOWN':
//         if (options.colSituacao) {
//           return `Cria um dropdown hierárquico para Etapa (${columnToLetter_(options.colEtapa)}) -> Situação (${columnToLetter_(options.colSituacao)}), com dados da aba "${options.fonte.nomeAba}".`;
//         }
//         return `Cria um dropdown simples na coluna ${columnToLetter_(options.colEtapa)} com dados da aba "${options.fonte.nomeAba}".`;
//       case 'Module_VISOES':
//         const visoes = Object.keys(options.visoes || {}).join(', ');
//         return `Permite alternar entre as seguintes visões (mostrar/ocultar colunas): ${visoes}.`;
//       case 'Module_REGRAS_NEGOCIO':
//         return 'Executa regras de negócio específicas quando o gatilho é acionado.';
//       default:
//         return `Módulo "${modulo}" não possui um tradutor de comportamento definido.`;
//     }
//   } catch (e) {
//     return `Erro ao traduzir o comportamento: ${e.message}`;
//   }
// }

// function coletarLogicaDeAplicacao_() {
//   const regrasDocumentadas = [];
//   if (typeof CONFIG === 'undefined' || !CONFIG.SHEETS) {
//     return [{ aba: 'ERRO', gatilho: 'N/A', modulo: 'N/A', comportamento: 'Objeto CONFIG não encontrado ou inválido.' }];
//   }
//   Object.entries(CONFIG.SHEETS).forEach(([sheetName, sheetConfig]) => {
//     if (!sheetConfig.handlers || sheetConfig.handlers.length === 0) return;
//     sheetConfig.handlers.forEach(handler => {
//       const colunasGatilho = (handler.colunas || [])
//         .map(c => `${columnToLetter_(c)} (${obterNomeSemantico_(sheetName, c) || 'N/D'})`)
//         .join(', ');
//       regrasDocumentadas.push({
//         aba: sheetName,
//         gatilho: `Edição na(s) coluna(s): ${colunasGatilho}`,
//         modulo: handler.modulo,
//         comportamento: formatarComportamentoHandler_(handler, sheetName)
//       });
//     });
//   });
//   return regrasDocumentadas;
// }

// /* ===================== Funções de escrita de relatório ===================== */

// function salvarLogicaDeAplicacao_(ss, lista) {
//   var header = ['Aba', 'Gatilho', 'Módulo', 'Comportamento Detalhado'];
//   var rows = [header];
//   for (var i = 0; i < lista.length; i++) {
//     var it = lista[i];
//     rows.push([it.aba, it.gatilho, it.modulo, it.comportamento]);
//   }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'LOGICA_DA_APLICACAO', rows); return; }
//   var sh = ensureSheet_(ss, 'LOGICA_DA_APLICACAO'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// function salvarMonitoramentosDeclarados_(ss, lista) {
//   var header = ['ORIGEM', 'ABA', 'ALVO(A1)', 'HANDLER/MODULO', 'DETALHE']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) {
//     var it = lista[i];
//     rows.push([it.origem || '', it.aba || '', it.alvo || '', it.handlerOuModulo || '', it.detalhe || '']);
//   }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'MONITORAMENTOS_DECLARADOS', rows); return; }
//   var sh = ensureSheet_(ss, 'MONITORAMENTOS_DECLARADOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// function normalizeConfig_() {
//   if (typeof CONFIG !== 'object' || !CONFIG) return {};
//   var out = { timestamp: [], moveRules: [], dropdowns: [], handlers: [], installments: null };
//   function pick(obj, names) { if (!obj) return undefined; var keys = Object.keys(obj); for (var i = 0; i < names.length; i++) { var n = names[i]; for (var k = 0; k < keys.length; k++) { if (String(keys[k]).toLowerCase() === String(n).toLowerCase()) return obj[keys[k]]; } } return undefined; }
//   function pickKey(obj, names) { if (!obj) return null; var keys = Object.keys(obj); for (var i = 0; i < names.length; i++) { var n = names[i]; for (var k = 0; k < keys.length; k++) { if (String(keys[k]).toLowerCase() === String(n).toLowerCase()) return keys[k]; } } return null; }
//   function sheetNameFrom(item) { if (!item || typeof item !== 'object') return ''; var key = pickKey(item, ['nome', 'name', 'sheet', 'aba', 'title']); return key ? String(item[key]) : ''; }
//   function toLetter(sheetName, ref) { return parseColRefToLetter_(sheetName, ref); }
//   var sheetsNode = pick(CONFIG, ['SHEETS', 'PLANILHAS', 'ABAS']);
//   if (sheetsNode) {
//     if (Array.isArray(sheetsNode)) {
//       for (var i = 0; i < sheetsNode.length; i++) handleSheetConfig_(sheetNameFrom(sheetsNode[i]), sheetsNode[i]);
//     } else if (typeof sheetsNode === 'object') {
//       var keys = Object.keys(sheetsNode);
//       for (var j = 0; j < keys.length; j++) { var nm = keys[j]; handleSheetConfig_(nm, sheetsNode[nm]); }
//     }
//   }
//   (function () {
//     var ddTop = pick(CONFIG, ['LISTAS_SUSPENSAS', 'DROPDOWNS']);
//     if (!ddTop) return; if (!Array.isArray(ddTop)) ddTop = [ddTop];
//     ddTop.forEach(function (rule) {
//       if (!rule) return;
//       var sheet = rule.nomeAba || rule.aba || rule.sheet || '';
//       var cat = rule.colunaCategoria || rule.categoryCol || rule.categoria;
//       var sub = rule.colunaSubcategoria || rule.subcategoryCol || rule.subcategoria;
//       if (cat) out.dropdowns.push({ origem: 'CONFIG.LISTAS_SUSPENSAS.categoria', sheet: sheet, col: toLetter(sheet, cat), type: 'categoria', source: (rule.fonte || rule.source || ''), parentCol: null });
//       if (sub) out.dropdowns.push({ origem: 'CONFIG.LISTAS_SUSPENSAS.subcategoria', sheet: sheet, col: toLetter(sheet, sub), type: 'subcategoria', source: (rule.fonte || rule.source || ''), parentCol: toLetter(sheet, cat) });
//     });
//   })();
//   (function () {
//     var ap = pick(CONFIG, ['ACOMPANHAMENTO_PARCELAS', 'PARCELAS', 'INSTALLMENTS', 'SCHEDULE']); if (!ap || typeof ap !== 'object') return;
//     var S = ap.ABA || ap.aba || ap.sheet || 'SEGUROS';
//     var cols = {
//       contract: toLetter(S, ap.COLUNA_DATA_CONTRATO || ap.contractDateCol || ap.contratoCol),
//       dueDay: toLetter(S, ap.COLUNA_DIA_VENCTO || ap.dueDayCol || ap.venctoCol),
//       list: []
//     };
//     var lst = ap.COLUNAS_PARCELAS || ap.parcelasCols || ap.installmentCols || [];
//     if (Array.isArray(lst)) lst.forEach(function (x) {
//       var col = (x && (x.col || x.c || x.column)) || x; var L = toLetter(S, col); if (L) cols.list.push(L);
//     });
//     out.installments = { sheet: S, cols: cols };
//   })();
//   (function () {
//     var H = pick(CONFIG, ['HANDLERS', 'ACTIONS', 'EVENTS']); if (!H || typeof H !== 'object') return;
//     var keys = Object.keys(H);
//     for (var i = 0; i < keys.length; i++) { var key = keys[i]; out.handlers.push({ sheet: '', target: key, handler: String(H[key]) }); }
//   })();
//   return out;
//   function handleSheetConfig_(sheetName, cfg) {
//     if (!sheetName) return; cfg = cfg || {};
//     var cols = pick(cfg, ['COLUNAS_MONITORADAS', 'MONITORED_COLUMNS', 'WATCHED_COLUMNS', 'TIMESTAMP_COLUMNS']);
//     if (Array.isArray(cols)) cols.forEach(function (ref) {
//       out.timestamp.push({ origem: 'CONFIG.SHEETS.COLUNAS_MONITORADAS', sheet: sheetName, col: toLetter(sheetName, ref) });
//     });
//     var mv = pick(cfg, ['MOVER_LINHA_CONFIG', 'MOVE_RULES', 'PIPELINE', 'STATUS_MAP']);
//     if (mv) {
//       if (Array.isArray(mv)) {
//         mv.forEach(function (rule) {
//           if (!rule) return; var colKey = rule.col || rule.column || rule.c || ''; var L = toLetter(sheetName, colKey);
//           out.moveRules.push({ sheet: sheetName, col: L, map: (rule.map || rule.destinos || {}) });
//         });
//       } else if (typeof mv === 'object') {
//         Object.keys(mv).forEach(function (colKey) {
//           var L = toLetter(sheetName, colKey);
//           out.moveRules.push({ sheet: sheetName, col: L, map: mv[colKey] });
//         });
//       }
//     }
//     var H = pick(cfg, ['HANDLERS', 'ACTIONS', 'EVENTS']);
//     if (H && typeof H === 'object') {
//       Object.keys(H).forEach(function (target) {
//         out.handlers.push({ origem: 'CONFIG.SHEETS.HANDLERS', sheet: sheetName, target: target, handler: String(H[target]) });
//       });
//     }
//     var ddBlocks = pick(cfg, ['LISTAS_SUSPENSAS', 'DROPDOWNS', 'DROPDOWN_DEPENDENTE', 'DEPENDENT_DROPDOWNS', 'VALIDACOES_LISTA']);
//     if (Array.isArray(ddBlocks)) ddBlocks.forEach(function (rule) {
//       if (!rule) return; var alvo = rule.colunaAlvo || rule.coluna || rule.target || rule.colunaCategoria || '';
//       out.dropdowns.push({ origem: 'CONFIG.SHEETS.DROPDOWNS', sheet: sheetName, col: toLetter(sheetName, alvo), type: rule.type || '', source: (rule.fonte || rule.source || ''), parentCol: toLetter(sheetName, rule.colunaCategoria || '') });
//     });
//   }
// }

// /* ===================== Coleta de dados nativos da planilha... ===================== */
// // (As funções coletarFormatacaoCondicional_, coletarFiltrosAtivos_, etc. permanecem inalteradas)
// function coletarFormatacaoCondicional_(sheets) {
//   var out = [];
//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i], nm = sh.getName();
//     if (SAIDAS.indexOf(nm) !== -1) continue;
//     if (sh.isSheetHidden && sh.isSheetHidden()) continue;
//     var rules = sh.getConditionalFormatRules() || [];
//     for (var r = 0; r < rules.length; r++) {
//       var rule = rules[r];
//       var ranges = (rule.getRanges() || []).map(function (rg) { return rg.getA1Notation(); }).join(' ; ');
//       var tipo = '', valores = '', bc, gc;
//       try { bc = rule.getBooleanCondition(); } catch (e) { bc = null; }
//       try { gc = rule.getGradientCondition(); } catch (e2) { gc = null; }
//       if (bc) { tipo = 'BOOLEAN'; try { valores = String(bc.getCriteriaType()) + ' :: ' + (bc.getValues() || []).join(' ; '); } catch (e3) { valores = 'BOOLEAN'; } }
//       else if (gc) { tipo = 'GRADIENT'; valores = 'GRADIENT'; }
//       else { tipo = 'DESCONHECIDO'; }
//       out.push({ aba: nm, ranges: ranges, tipo: tipo, criterios: valores });
//     }
//   }
//   return out;
// }
// function salvarFormatacaoCondicional_(ss, lista) {
//   var header = ['ABA', 'RANGES', 'TIPO', 'CRITERIOS/VALORES']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.aba, it.ranges, it.tipo, it.criterios]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FORMATACAO_CONDICIONAL', rows); return; }
//   var sh = ensureSheet_(ss, 'FORMATACAO_CONDICIONAL'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }
// function coletarFiltrosAtivos_(sheets) {
//   var out = [];
//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i], nm = sh.getName();
//     if (SAIDAS.indexOf(nm) !== -1) continue;
//     if (sh.isSheetHidden && sh.isSheetHidden()) continue;
//     var filter = sh.getFilter(); if (!filter) continue;
//     var fr; try { fr = filter.getRange(); } catch (e) { fr = null; }
//     var base = fr ? fr.getA1Notation() : '';
//     var lastCol = sh.getLastColumn();
//     for (var c = 1; c <= lastCol; c++) {
//       var crit = null; try { crit = filter.getColumnFilterCriteria(c); } catch (e2) { crit = null; }
//       if (!crit) continue;
//       var tipo = '', hidden = []; try { tipo = String(crit.getCriteriaType()); } catch (e3) { } try { hidden = crit.getHiddenValues() || []; } catch (e4) { }
//       out.push({ aba: nm, range: base, colunaA1: columnToLetter_(c), criteriaType: tipo, hiddenValues: hidden.join(' | ') });
//     }
//   }
//   return out;
// }
// function salvarFiltrosAtivos_(ss, lista) {
//   var header = ['ABA', 'RANGE', 'COLUNA_A1', 'CRITERIA_TYPE', 'HIDDEN_VALUES']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.aba, it.range, it.colunaA1, it.criteriaType, it.hiddenValues]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FILTROS_ATIVOS', rows); return; }
//   var sh = ensureSheet_(ss, 'FILTROS_ATIVOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }
// function coletarProtecoes_(ss, abasSelecionadas) {
//   var out = [];
  
//   // ▼▼▼ CORREÇÃO ADICIONADA ▼▼▼
//   // Cria um Set (para performance) com as abas permitidas
//   var filtroAbas = (abasSelecionadas && abasSelecionadas.length > 0) ? new Set(abasSelecionadas) : null;
//   // ▲▲▲ FIM DA CORREÇÃO ▲▲▲

//   try {
//     (ss.getProtections(SpreadsheetApp.ProtectionType.SHEET) || []).forEach(function (p) {
//       var abaNome = p.getRange() ? p.getRange().getSheet().getName() : '';
      
//       // ▼▼▼ CORREÇÃO ADICIONADA ▼▼▼
//       // Pula se a proteção pertence a uma aba que não foi selecionada
//       if (filtroAbas && abaNome && !filtroAbas.has(abaNome)) return; 
//       // ▲▲▲ FIM DA CORREÇÃO ▲▲▲
      
//       out.push({ tipo: 'SHEET', aba: abaNome, ranges: p.getRange() ? p.getRange().getA1Notation() : '', warningOnly: p.isWarningOnly(), descr: (p.getDescription ? p.getDescription() : '') || '' });
//     });
//   } catch (e1) { }
//   try {
//     (ss.getProtections(SpreadsheetApp.ProtectionType.RANGE) || []).forEach(function (pr) {
//       var abaNome = pr.getRange() ? pr.getRange().getSheet().getName() : '';
      
//       // ▼▼▼ CORREÇÃO ADICIONADA ▼▼▼
//       // Pula se a proteção pertence a uma aba que não foi selecionada
//       if (filtroAbas && abaNome && !filtroAbas.has(abaNome)) return;
//       // ▲▲▲ FIM DA CORREÇÃO ▲▲▲
      
//       out.push({ tipo: 'RANGE', aba: abaNome, ranges: pr.getRange() ? pr.getA1Notation ? pr.getA1Notation() : pr.getRange().getA1Notation() : '', warningOnly: pr.isWarningOnly(), descr: (pr.getDescription ? pr.getDescription() : '') || '' });
//     });
//   } catch (e2) { }
//   return out;
// }
// function salvarProtecoes_(ss, lista) {
//   var header = ['TIPO', 'ABA', 'RANGES', 'WARNING_ONLY', 'DESCRICAO']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.tipo, it.aba, it.ranges, it.warningOnly ? 'SIM' : 'NAO', it.descr]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'PROTECOES', rows); return; }
//   var sh = ensureSheet_(ss, 'PROTECOES'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }
// function coletarNamedRanges_(ss, abasSelecionadas) {
//   var out = [];

//   // ▼▼▼ CORREÇÃO ADICIONADA ▼▼▼
//   // Cria um Set (para performance) com as abas permitidas
//   var filtroAbas = (abasSelecionadas && abasSelecionadas.length > 0) ? new Set(abasSelecionadas) : null;
//   // ▲▲▲ FIM DA CORREÇÃO ▲▲▲

//   (ss.getNamedRanges() || []).forEach(function (nr) {
//     var rg = nr.getRange();
//     var abaNome = rg ? rg.getSheet().getName() : '';

//     // ▼▼▼ CORREÇÃO ADICIONADA ▼▼▼
//     // Pula se o intervalo nomeado pertence a uma aba que não foi selecionada
//     if (filtroAbas && abaNome && !filtroAbas.has(abaNome)) return;
//     // ▲▲▲ FIM DA CORREÇÃO ▲▲▲

//     out.push({ nome: nr.getName(), aba: abaNome, range: rg ? rg.getA1Notation() : '' });
//   });
//   return out;
// }
// function salvarRangesNomeados_(ss, lista) {
//   var header = ['NOME', 'ABA', 'RANGE']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.nome, it.aba, it.range]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'RANGES_NOMEADOS', rows); return; }
//   var sh = ensureSheet_(ss, 'RANGES_NOMEADOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }
// function coletarGraficos_(sheets) {
//   var out = [];
//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i], nm = sh.getName();
//     if (SAIDAS.indexOf(nm) !== -1) continue;
//     if (sh.isSheetHidden && sh.isSheetHidden()) continue;
//     var charts = sh.getCharts ? (sh.getCharts() || []) : [];
//     for (var c = 0; c < charts.length; c++) {
//       var ch = charts[c];
//       var ranges = (ch.getRanges ? ch.getRanges() : []).map(function (rg) { return rg.getA1Notation(); }).join(' ; ');
//       out.push({ aba: nm, tipo: (ch.getChartType ? String(ch.getChartType()) : ''), ranges: ranges, posicao: (ch.getContainerInfo ? JSON.stringify(ch.getContainerInfo()) : '') });
//     }
//   }
//   return out;
// }
// function salvarGraficos_(ss, lista) {
//   var header = ['ABA', 'TIPO', 'RANGES_FONTE', 'CONTAINER_INFO(JSON)']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.aba, it.tipo, it.ranges, it.posicao]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'GRAFICOS', rows); return; }
//   var sh = ensureSheet_(ss, 'GRAFICOS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }
// function coletarFilterViewsESlicers_(ss, abasSelecionadas) { //
//   if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets || !Sheets.Spreadsheets.get) return [];
//   var id = ss.getId();
//   var resp = Sheets.Spreadsheets.get(id, { fields: 'sheets(properties.title,filterViews,slicers),developerMetadata' });
//   var out = []; var arr = (resp && resp.sheets) || [];

//   // ▼▼▼ BLOCO DE CORREÇÃO ADICIONADO ▼▼▼
//   // Cria um Set (para performance) se a lista de abas selecionadas foi fornecida
//   var filtroAbas = (abasSelecionadas && abasSelecionadas.length > 0) ? new Set(abasSelecionadas) : null;
//   // ▲▲▲ FIM DA CORREÇÃO ▲▲▲

//   for (var i = 0; i < arr.length; i++) {
//     var s = arr[i]; var title = s.properties && s.properties.title;

//     // ▼▼▼ LINHA DE CORREÇÃO ADICIONADA ▼▼▼
//     // Pula esta aba se ela não estiver no filtro de seleção
//     if (filtroAbas && !filtroAbas.has(title)) continue;
//     // ▲▲▲ FIM DA CORREÇÃO ▲▲▲
//     (s.filterViews || []).forEach(function (fv) {
//       out.push({ tipo: 'FILTER_VIEW', aba: title, nome: fv.title || '', range: _gridRangeToA1_(fv.range, title), detalhes: JSON.stringify({ criteria: fv.filterSpecs || [] }) });
//     });
//     (s.slicers || []).forEach(function (sl) {
//       out.push({ tipo: 'SLICER', aba: title, nome: (sl.spec && sl.spec.title) || '', range: _gridRangeToA1_(sl.spec && sl.spec.filterCriteriaRange ? sl.spec.filterCriteriaRange : null, title), detalhes: JSON.stringify(sl.spec || {}) });
//     });
//   }
//   (resp.developerMetadata || []).forEach(function (m) {
//     out.push({ tipo: 'DEVELOPER_METADATA', aba: '', nome: (m.metadataKey || ''), range: '', detalhes: JSON.stringify(m) });
//   });
//   return out;
// }
// function salvarFilterViewsEAfins_(ss, lista) {
//   var header = ['TIPO', 'ABA', 'NOME/TITULO', 'RANGE', 'DETALHES_JSON']; var rows = [header];
//   for (var i = 0; i < lista.length; i++) { var it = lista[i]; rows.push([it.tipo, it.aba, it.nome, it.range, it.detalhes]); }
//   if (MODO_SAIDA_UNICA) { appendSection_(getOutSheet_(), 'FILTER_VIEWS_E_AFINS', rows); return; }
//   var sh = ensureSheet_(ss, 'FILTER_VIEWS_E_AFINS'); sh.clear(); writeTable_(sh, 1, 1, rows, true);
// }

// /* ===================== Helpers de relatório (visual + export) ===================== */
// function beginReport_(ss) {
//   __OUT_SH = ensureSheet_(ss, NOME_SAIDA_UNICA); __OUT_SH.clear();
//   var head = [['RELATORIO DE ANALISE — ' + (ss.getName() || ''), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')]];
//   __OUT_SH.getRange(1, 1, 1, head[0].length).setValues(head).setFontWeight('bold').setFontSize(12);
//   __OUT_SH.getRange(3, 1, 1, 1).setValue('SUMARIO').setFontWeight('bold');
//   __OUT_SH.setFrozenRows(5);
//   __CUR_ROW = 6; __TOC = []; __SECTIONS = []; __SEC_INDEX = 0;
//   if (EXPORT_TO_DRIVE) exportBegin_(ss);
// }
// function finalizeReport_() {
//   var gid = __OUT_SH.getSheetId();
//   var tocRows = [['Seção', 'Ir']];
//   var locale = SpreadsheetApp.getActive().getSpreadsheetLocale() || 'en_US';
//   var sep = (locale.indexOf('en_') === 0) ? ',' : ';';
//   for (var i = 0; i < __TOC.length; i++) {
//     var item = __TOC[i];
//     var formula = '=HYPERLINK("#gid=' + gid + '&range=A' + item.row + '"' + sep + '"' + item.title + '")';
//     tocRows.push([item.title, formula]);
//   }
//   writeTable_(__OUT_SH, 4, 1, tocRows, true);
//   if (EXPORT_TO_DRIVE) exportAllSections_();
// }
// function appendSection_(sheet, titulo, rows) {
//   __SEC_INDEX++; __SECTIONS.push({ idx: __SEC_INDEX, title: titulo, rows: rows });
//   sheet.getRange(__CUR_ROW, 1).setValue('=== BEGIN:' + titulo + ' ===');
//   sheet.getRange(__CUR_ROW + 1, 1, 1, 1).setValue('■ ' + titulo).setFontWeight('bold').setBackground('#efefef');
//   var tableStart = __CUR_ROW + 2;
//   writeTable_(sheet, tableStart, 1, rows, true);
//   sheet.getRange(tableStart + rows.length, 1).setValue('=== END:' + titulo + ' ===');
//   __TOC.push({ title: titulo, row: tableStart });
//   __CUR_ROW = tableStart + rows.length + 2;
// }
// function writeTable_(sh, startRow, startCol, rows, withFormat) {
//   if (!rows || !rows.length) return; var cols = rows[0].length;
//   sh.getRange(startRow, startCol, rows.length, cols).setValues(rows);
//   if (withFormat) {
//     sh.getRange(startRow, startCol, 1, cols).setFontWeight('bold').setBackground('#f5f5f5');
//     sh.getRange(startRow, startCol, rows.length, cols).setWrap(true);
//     try { sh.getRange(startRow, startCol, rows.length, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) { }
//     try { if (!sh.getFilter()) sh.getRange(startRow, startCol, rows.length, cols).createFilter(); } catch (e) { }
//     sh.autoResizeColumns(startCol, cols);
//   }
// }

// /* ===================== Export: YAML + CSV + MD ===================== */
// function exportBegin_(ss) {
//   var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
//   var root = DriveApp.getFoldersByName(EXPORT_FOLDER_BASENAME);
//   var baseFolder = root.hasNext() ? root.next() : DriveApp.createFolder(EXPORT_FOLDER_BASENAME);
//   var folder = baseFolder.createFolder(EXPORT_FOLDER_BASENAME + '_' + stamp);
//   __EXPORT = {
//     folder: folder,
//     folderId: folder.getId(),
//     pathPrefix: EXPORT_FOLDER_BASENAME + '_' + stamp + '/',
//     manifest: {
//       spreadsheet: { id: ss.getId(), name: ss.getName(), url: ss.getUrl() },
//       generatedAt: stamp,
//       sections: []
//     }
//   };
//   var readme = ['# README_AI', '', '- Leia 00_MANIFEST.yml primeiro.', '- Depois, carregue apenas os CSV/MD das secções relevantes.', '- No relatório em planilha, use marcadores === BEGIN:/END: para copiar blocos.', ''].join('\n');
//   folder.createFile('00_README_AI.md', readme, MimeType.PLAIN_TEXT);
// }
// function exportAllSections_() {
//   if (!__EXPORT) return; var folder = __EXPORT.folder; var man = __EXPORT.manifest;
//   for (var i = 0; i < __SECTIONS.length; i++) {
//     var S = __SECTIONS[i];
//     var slug = sanitizeSlug_(S.title);
//     var idx = pad2_(S.idx);
//     var csvName = idx + '_' + slug + '.csv';
//     var csv = csvEncodeRows_(S.rows);
//     folder.createFile(csvName, csv, MimeType.PLAIN_TEXT);
//     var mdBase = idx + '_' + slug;
//     var md = mdFromRows_(S.title, S.rows);
//     var parts = splitText_(md, EXPORT_MD_SPLIT_THRESHOLD);
//     var mdFiles = [];
//     for (var p = 0; p < parts.length; p++) {
//       var nm = mdBase + (parts.length > 1 ? ('.part' + (p + 1)) : '') + '.md';
//       folder.createFile(nm, parts[p], MimeType.PLAIN_TEXT);
//       mdFiles.push(nm);
//     }
//     man.sections.push({ idx: S.idx, title: S.title, csv: csvName, md: mdFiles, rows: S.rows.length, cols: (S.rows[0] || []).length });
//   }
//   var yaml = yamlFromManifest_(man);
//   folder.createFile('00_MANIFEST.yml', yaml, MimeType.PLAIN_TEXT);
// }
// function csvEncodeRows_(rows) {
//   function esc(v) { var s = (v == null ? '' : String(v)); if (/[",\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"'; return s; }
//   var out = []; for (var r = 0; r < rows.length; r++) { out.push(rows[r].map(esc).join(',')); }
//   return out.join('\n');
// }
// function mdFromRows_(title, rows) {
//   var header = '---\nsection: ' + title + '\n---\n\n';
//   var body = '```csv\n' + csvEncodeRows_(rows) + '\n```\n';
//   return header + body;
// }
// function splitText_(txt, max) { if (txt.length <= max) return [txt]; var out = []; for (var i = 0; i < txt.length; i += max) { out.push(txt.substring(i, Math.min(txt.length, i + max))); } return out; }
// function yamlFromManifest_(m) {
//   function esc(s) { return String(s).replace(/\n/g, '\\n'); }
//   var lines = [];
//   lines.push('spreadsheet:');
//   lines.push('  id: ' + esc(m.spreadsheet.id));
//   lines.push('  name: ' + esc(m.spreadsheet.name));
//   lines.push('  url: ' + esc(m.spreadsheet.url));
//   lines.push('generatedAt: ' + esc(m.generatedAt));
//   lines.push('sections:');
//   for (var i = 0; i < m.sections.length; i++) {
//     var s = m.sections[i];
//     lines.push('  - idx: ' + s.idx);
//     lines.push('    title: ' + esc(s.title));
//     lines.push('    rows: ' + s.rows);
//     lines.push('    cols: ' + s.cols);
//     lines.push('    csv: ' + s.csv);
//     if (s.md && s.md.length) {
//       lines.push('    md:');
//       for (var j = 0; j < s.md.length; j++) { lines.push('      - ' + s.md[j]); }
//     }
//   }
//   return lines.join('\n') + '\n';
// }
// function sanitizeSlug_(s) { return String(s).trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, ''); }
// function pad2_(n) { return (n < 10 ? '0' : '') + n; }

// /* ===================== Headers: heurística + leitura da guia ===================== */
// function obterHeaderRow_(sheet) {
//   var g = readGlobalHeaderRowFromSelection_();
//   if (g && g >= 1) return g;
//   var perSheet = readPerSheetHeaderRowFromSelection_(sheet.getName());
//   if (perSheet && perSheet >= 1) return perSheet;
//   var lastRow = sheet.getLastRow();
//   var lastCol = sheet.getLastColumn();
//   if (lastCol < 1) return 1;
//   var maxScan = Math.min(HEADER_SCAN_LIMIT, Math.max(1, lastRow));
//   var best = { row: 1, score: -1 };
//   for (var r = 1; r <= maxScan; r++) {
//     var vals = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
//     var belowVals = (r < lastRow) ? sheet.getRange(r + 1, 1, 1, Math.min(lastCol, 20)).getDisplayValues()[0] : [];
//     var sc = scoreHeaderRow_(vals, belowVals);
//     if (sc > best.score) { best = { row: r, score: sc }; }
//   }
//   if (best.score >= HEADER_SCORE_THRESHOLD) return best.row;
//   return best.row || 1;
// }
// function scoreHeaderRow_(vals, belowVals) {
//   var n = vals.length; var nonEmpty = 0, texty = 0, digits = 0; var set = {};
//   for (var i = 0; i < n; i++) {
//     var v = String(vals[i] || '').trim();
//     if (v !== '') {
//       nonEmpty++; set[v] = 1;
//       if (/^[0-9.,:\-\/]+$/.test(v)) digits++; else texty++;
//     }
//   }
//   var pNon = nonEmpty / n; var pText = texty / Math.max(1, (texty + digits)); var uniq = Object.keys(set).length / Math.max(1, nonEmpty);
//   var keywords = 0; var kwRe = /(^|\b)(ID|DATA|DT|VALOR|PRECO|PREÇO|EMAIL|E-MAIL|STATUS|CPF|CNPJ)($|\b)/i;
//   for (var i2 = 0; i2 < n; i2++) { if (kwRe.test(String(vals[i2] || ''))) keywords++; }
//   var belowNums = 0; for (var j = 0; j < belowVals.length; j++) { if (/^-?\d+[\d.,\/:-]*$/.test(String(belowVals[j] || ''))) belowNums++; }
//   var belowRatio = belowVals.length ? (belowNums / belowVals.length) : 0;
//   return (pNon * 0.6) + (pText * 0.4) + (uniq * 0.3) + (Math.min(1, keywords / 4) * 0.3) + (belowRatio * 0.2);
// }
// function buildHeaderIndex_(headers) {
//   var map = {};
//   for (var c = 0; c < headers.length; c++) {
//     var h = (headers[c] || '').toString().trim(); if (!h) continue;
//     var key = h.toUpperCase(); if (!map[key]) map[key] = columnToLetter_(c + 1);
//   }
//   return map;
// }

// /* ===================== Leitura da guia de seleção (header/abas) ===================== */
// var NOME_SELECAO_ABAS = 'SELECIONAR_ABAS_TEMP';
// function getSelectionSheet_() {
//   return SpreadsheetApp.getActive().getSheetByName(NOME_SELECAO_ABAS) || null;
// }
// function readGlobalHeaderRowFromSelection_() {
//   var sh = getSelectionSheet_(); if (!sh) return null;
//   var v = Number(sh.getRange('E2').getValue()); return isFinite(v) && v >= 1 ? Math.floor(v) : null;
// }
// function readPerSheetHeaderRowFromSelection_(sheetName) {
//   var sh = getSelectionSheet_(); if (!sh) return null;
//   var lastRow = sh.getLastRow(); if (lastRow < 2) return null;
//   var vals = sh.getRange(2, 1, lastRow - 1, 3).getValues();
//   for (var i = 0; i < vals.length; i++) {
//     var nm = String(vals[i][1] || '').trim();
//     if (nm === sheetName) {
//       var num = Number(vals[i][2]); return isFinite(num) && num >= 1 ? Math.floor(num) : null;
//     }
//   }
//   return null;
// }

// /* ====================================================================
//  * UI SIMPLES (PAINEL DE CONTROLO): Selecionar abas e header na própria guia
//  * ==================================================================== */
// function prepararSelecaoAbasTemp() {
//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName(NOME_SELECAO_ABAS) || ss.insertSheet(NOME_SELECAO_ABAS);
//   sh.clear();
//   sh.showSheet();

//   sh.getRange('A1:C1').setValues([['Incluir?', 'Aba', 'Header Row (Opcional)']]).setFontWeight('bold');
//   var sheets = ss.getSheets();
//   var rows = [];
//   for (var i = 0; i < sheets.length; i++) {
//     var nm = sheets[i].getName();
//     if (SAIDAS.indexOf(nm) !== -1 || nm === NOME_SELECAO_ABAS) continue;
//     var sug = _heuristicaHeaderRowNoPrompt_(sheets[i]);
//     rows.push([true, nm, sug]);
//   }
//   if (rows.length) {
//     var rangeAbas = sh.getRange(2, 1, rows.length, 3);
//     rangeAbas.setValues(rows);
//     var dvCheckbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
//     sh.getRange(2, 1, rows.length, 1).setDataValidation(dvCheckbox);
//   }

//   sh.getRange('E1').setValue('Preset de Análise').setFontWeight('bold');
//   const presetCell = sh.getRange('E2');
//   presetCell.setBackground('#fff9c4').setNote('Selecione o tipo de análise a ser executado.');
//   const presetNames = Object.keys(PRESETS);
//   const presetListRange = sh.getRange(2, 6, presetNames.length, 1);
//   presetListRange.setValues(presetNames.map(name => [name]));
//   var dvDropdown = SpreadsheetApp.newDataValidation().requireValueInRange(presetListRange, true).setAllowInvalid(false).build();
//   presetCell.setDataValidation(dvDropdown);
//   presetCell.setValue('FOCO_CODIGO_E_REGRAS');

//   sh.setFrozenRows(1);
//   sh.autoResizeColumns(1, 5);
//   sh.hideColumns(6);
//   SpreadsheetApp.setActiveSheet(sh);
//   SpreadsheetApp.getActive().toast('Painel de controlo pronto. Ajuste as seleções e execute a análise.');
// }

// function executarAnalisePeloPainel() {
//   const shSel = SpreadsheetApp.getActive().getSheetByName(NOME_SELECAO_ABAS);
//   if (!shSel) {
//     SpreadsheetApp.getUi().alert('Painel não encontrado. Por favor, execute "prepararSelecaoAbasTemp" primeiro.');
//     return;
//   }
//   const presetSelecionado = shSel.getRange('E2').getValue();
//   if (!presetSelecionado || !PRESETS[presetSelecionado]) {
//     SpreadsheetApp.getUi().alert(`Preset "${presetSelecionado}" inválido. Selecione uma opção válida no dropdown da célula E2.`);
//     return;
//   }
//   SpreadsheetApp.getActive().toast(`Iniciando análise com o preset: "${presetSelecionado}"...`, 'Progresso', 30);
//   runAnaliseComSelecao(presetSelecionado);
// }

// function runAnaliseComSelecao(presetName) {
//   var ss = SpreadsheetApp.getActive();
//   var shSel = getSelectionSheet_();
//   if (!shSel) {
//     SpreadsheetApp.getUi().alert('Painel não encontrado. Execute "prepararSelecaoAbasTemp".');
//     return;
//   }
//   var selecionadas = _lerSelecaoAbasTemp_();
//   var snap = _snapshotVisibilidade_(ss);
//   try {
//     if (selecionadas.length > 0) {
//       _ocultarExceto_(ss, new Set(selecionadas));
//     } else {
//       SpreadsheetApp.getActive().toast('Nenhuma aba marcada — analisando todas as visíveis.');
//     }
//     documentarEstruturaCompleta(presetName, selecionadas);
//   } finally {
//     _restaurarVisibilidade_(ss, snap);
//   }
// }

// function _lerSelecaoAbasTemp_() {
//   var sh = getSelectionSheet_(); if (!sh) return [];
//   var lastRow = sh.getLastRow(); if (lastRow < 2) return [];
//   var vals = sh.getRange(2, 1, lastRow - 1, 2).getValues();
//   var out = [];
//   for (var i = 0; i < vals.length; i++) {
//     if (vals[i][0] === true) {
//       var nm = String(vals[i][1] || '').trim();
//       if (nm && SAIDAS.indexOf(nm) == -1) out.push(nm);
//     }
//   }
//   return out;
// }
// function _snapshotVisibilidade_(ss) {
//   var snap = {}; var sheets = ss.getSheets();
//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i]; var nm = sh.getName(); snap[nm] = (sh.isSheetHidden && sh.isSheetHidden());
//   }
//   return snap;
// }
// function _restaurarVisibilidade_(ss, snap) {
//   var sheets = ss.getSheets();
//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i]; var nm = sh.getName(); var wasHidden = !!snap[nm];
//     try { if (wasHidden) sh.hideSheet(); else sh.showSheet(); } catch (e) { }
//   }
// }
// function _ocultarExceto_(ss, selectedSet) {
//   var sheets = ss.getSheets();
//   for (var i = 0; i < sheets.length; i++) {
//     var sh = sheets[i]; var nm = sh.getName();
//     if (SAIDAS.indexOf(nm) !== -1 || nm === NOME_SELECAO_ABAS) continue;
//     try { if (!selectedSet.has(nm)) sh.hideSheet(); else sh.showSheet(); } catch (e) { }
//   }
// }

// function _heuristicaHeaderRowNoPrompt_(sheet) {
//   var lastRow = sheet.getLastRow(); var lastCol = sheet.getLastColumn(); if (lastCol < 1) return 1;
//   var maxScan = Math.min(HEADER_SCAN_LIMIT, Math.max(1, lastRow));
//   var best = { row: 1, score: -1 };
//   for (var r = 1; r <= maxScan; r++) {
//     var vals = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
//     var belowVals = (r < lastRow) ? sheet.getRange(r + 1, 1, 1, Math.min(lastCol, 20)).getDisplayValues()[0] : [];
//     var sc = scoreHeaderRow_(vals, belowVals);
//     if (sc > best.score) { best = { row: r, score: sc }; }
//   }
//   return best.row;
// }

// /* ===================== Utilitários ===================== */
// function ensureSheet_(ss, name) { var sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name); return sh; }
// function getOutSheet_() { return __OUT_SH || ensureSheet_(SpreadsheetApp.getActive(), NOME_SAIDA_UNICA); }
// function columnToLetter_(column) { var temp, letter = ''; while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; } return letter; }
// function _pct(n) { return Math.round((n || 0) * 100) + '%'; }
// function _matrizTemAlgum_(mat) { if (!mat || !mat.length) return false; for (var r = 0; r < mat.length; r++) { for (var c = 0; c < mat[r].length; c++) { if (mat[r][c]) return true; } } return false; }
// function _matrizTemAlgumStringNaoVazia_(mat) {
//   if (!mat || !mat.length) return false;
//   for (var r = 0; r < mat.length; r++) {
//     for (var c = 0; c < mat[r].length; c++) {
//       if (String(mat[r][c] || '').trim() !== '') return true;
//     }
//   }
//   return false;
// }
// function calcularDensidade_(sheet, headerRow) {
//   var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
//   if (lastRow <= headerRow || lastCol < 1) return 0;
//   var rows = Math.min(lastRow - headerRow, 200), cols = Math.min(lastCol, 50);
//   var rg = sheet.getRange(headerRow + 1, 1, rows, cols);
//   var vals = rg.getDisplayValues();
//   var total = rows * cols, filled = 0;
//   for (var r = 0; r < vals.length; r++) {
//     for (var c = 0; c < vals[0].length; c++) {
//       if (String(vals[r][c]).trim() !== '') filled++;
//     }
//   }
//   return total ? (filled / total) : 0;
// }
// function inferirTipoPeloHeader_(h) {
//   h = (h || '').toString().trim().toUpperCase(); if (!h) return '';
//   if (/\bDATA\b|\bDT\b/.test(h)) return 'data';
//   if (/VALOR|PRECO|PREÇO|TOTAL|MONTANTE|R\$|\bVLR\b/.test(h)) return 'numero';
//   if (/CPF|CNPJ/.test(h)) return 'documento';
//   if (/EMAIL|E-MAIL/.test(h)) return 'email';
//   if (/ATIVO|APROVADO|SIM|NAO|NÃO|CHECK|STATUS/.test(h)) return 'booleano';
//   if (/\bID\b/.test(h)) return 'identificador';
//   return 'texto';
// }
// function parseColRefToLetter_(sheetName, ref) {
//   if (ref == null) return '';
//   if (typeof ref === 'number') return columnToLetter_(ref);
//   var s = String(ref).trim(); if (!s) return '';
//   if (/^\d+$/.test(s)) return columnToLetter_(parseInt(s, 10));
//   var m = s.match(/^([A-Za-z]+)(:\d.*)?$/); if (m) return m[1].toUpperCase();
//   var idx = __HEADER_INDEX[(sheetName || '').toUpperCase()] || {};
//   var L = idx[s.toUpperCase()] || ''; return L || '';
// }
// function _gridRangeToA1_(gr, title) {
//   if (!gr) return '';
//   function colToA1(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); } return s; }
//   var r1 = (gr.startRowIndex || 0) + 1, c1 = (gr.startColumnIndex || 0) + 1;
//   var r2 = (gr.endRowIndex == null ? r1 : gr.endRowIndex);
//   var c2 = (gr.endColumnIndex == null ? c1 : gr.endColumnIndex);
//   return (title ? "'" + title + "'!" : '') + colToA1(c1) + r1 + ':' + colToA1(c2) + r2;
// }
// function log_(msg) { try { Logger.log(msg); } catch (e) { } }

