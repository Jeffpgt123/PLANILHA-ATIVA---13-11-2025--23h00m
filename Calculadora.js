/**
 * ===================================================================
 * ARQUIVO: CALCULADORA UNIFICADA (VERSÃO 2 - Refatorada)
 * Gerencia as simulações de Prêmio Seguro, PRONAMPE e Simulação de Prêmios de Vida.
 *
 * MUDANÇAS (v2):
 * - Layout redesenhado: Inputs no topo (Linhas 1-15), Outputs na base (Linhas 20+).
 * - Calc 1 (Prestamista): Agora simula 5 valores de capital fixos (A7:A11) e gera tabelas de progressão mensal.
 * - Calc 3 (Seg. Vida): Convertido de Fórmulas para GAS. Inverte a lógica:
 * Usuário define o "Teto do Prêmio" e o script gera uma tabela de simulação de "Capital Individual".
 * - Calc 2 (Pronampe): Lógica mantida, mas referências de células atualizadas para o novo layout.
 * ===================================================================
 */
const NOME_ABA = 'Calculadora🧮';

/**
 * Retorna alíquota aproximada de IOF regressivo (0 a 1) em função do prazo em dias.
 * Simplificação linear oficial: de ~96% (dia 1) até 0% (dia 30).
 */

const BASE_DIAS_ANO_CDI = 252;

function _getAliquotaIofRendaFixa_(dias) {
  if (dias <= 0) return 0;
  if (dias >= 30) return 0;
  return (30 - dias) / 30; // aprox. regressivo
}

/**
 * Tabela regressiva de IR sobre renda fixa (CDB, LCA PJ).
 */
function _getAliquotaIrRendaFixa_(prazoDias) {
  if (prazoDias <= 180) return 0.225;
  if (prazoDias <= 360) return 0.20;
  if (prazoDias <= 720) return 0.175;
  return 0.15;
}


// ===================================================================
// ▶ CONFIGURAÇÕES DE CÉLULAS (v2 - Novo Layout)
// ===================================================================

const CONFIG_CALCULADORA = {
  // --- Bloco 1: Prestamista (Inputs) ---
  CALC1: {
    TAXA: 'B3', // Taxa Mensal (i_m)
    PRAZO_MAXIMO: 'B4', // Prazo Máximo (Meses)
    CAPITAIS_RANGE: 'B6:B10', // MODIFICADO: Era 'A7:A11'.
    // (Output)
    RESULTADO_INICIO_ROW: 13, // MODIFICADO: Era 12. Título fica na 12, cabeçalho/dados começam na 13
    RESULTADO_COL: 1 // Coluna A
  },

  // --- Bloco 3: Seguro de Vida (Inputs) ---
  CALC3: {
    QTD_FUNC: 'E3',
    QTD_SOCIOS: 'E4',
    TETO_PREMIO: 'E6',
    TAXAS_MODULOS: [0.000457, 0.000487, 0.000560, 0.000591], // Taxas como decimais (0.0457% = 0.000457)
    CAPITAL_INCREMENTO: 10000,
    // (Output)
    RESULTADO_INICIO_ROW: 13, // MODIFICADO: Era 12. Título fica na 12, cabeçalho na 13
    RESULTADO_COL: 4 // Coluna D
  },

  // --- Bloco 2: PRONAMPE (Inputs) ---
  CALC2: {
    SELIC: 'K3',
    VALOR: 'K4', // Valor Financiado
    PRAZO: 'K5',
    CARENCIA: 'K6',
    DATA_INICIO: 'K7',
    // (Output Resumo)
    IOF_TOTAL: 'P4', // MODIFICADO: Era P10
    VALOR_LIQUIDO: 'P5', // MODIFICADO: Era P11. (O5 é o Rótulo).
    // (Output Tabela)
    TABELA_INICIO_ROW: 14, // MODIFICADO: Era 13. Título na 12, Cabeçalho na 13, Dados na 14
    TABELA_INICIO_COL: 10 // Coluna J
  },

   // --- Bloco 4: CDB / LCA (CDI Pós) ---
  CALC4: {
    // Inputs (colunas R:U)
    TIPO_TITULO: 'S3',           // CDB / LCA
    TIPO_INVESTIDOR: 'S4',       // PF / PJ
    VALOR_INICIAL: 'S5',         // R$
    PRAZO_DIAS: 'S6',            // dias
    CDI_ANUAL_PERCENT: 'S7',     // % a.a.
    PERCENTUAL_CDI: 'S8',        // % CDI
    OUTROS_CUSTOS_PERCENT: 'S9', // % a.a. (taxas extras)

    // Resultados
    RESULTADO_INICIO_ROW: 12, // primeira linha dos resultados
    LABEL_COL: 18,            // Coluna R
    VALUE_COL: 19             // Coluna S
  }
};


// ===================================================================
// 📞 FUNÇÕES DE MENU E SETUP
// ===================================================================

/**
 * ATENÇÃO: A função onOpen() abaixo foi removida
 * conforme solicitado.
 * * Para ativar os botões, siga as instruções
 * no arquivo 'instrucoes_botoes.md'.
 */
/*
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('⚙️ Calculadora');
  menu.addItem('1. Preparar Layout da Aba', 'prepararCalculadora');
  menu.addSeparator();
  menu.addItem('Calcular 1: Prestamista (Múltiplo)', 'calcularPremioPrestamista');
  menu.addItem('Calcular 2: PRONAMPE', 'calcularPronampe');
  menu.addItem('Calcular 3: Seguro de Vida (por Teto)', 'simularPremioVidaPorTeto');
  menu.addSeparator();
  menu.addItem('Limpar 2: Limpar PRONAMPE', 'limparCalculoPronampe');
  menu.addToUi();
}
*/

/**
 * Prepara a aba Calculadora, limpando e formatando os 3 blocos.
 */
function prepararCalculadora() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOME_ABA);
  if (!sheet) {
    sheet = ss.insertSheet(NOME_ABA);
  }
  
  // Limpa tudo
  sheet.clear();
  sheet.setFrozenRows(0);

  // Define larguras de coluna para espaçamento
  sheet.setColumnWidth(3, 20); // C
  sheet.setColumnWidth(9, 20); // I
  
  // Chama as funções de preparação individuais
  _prepararCalcPrestamista(sheet);
  _prepararCalcPremioVida(sheet);
  _prepararCalcPronampe(sheet);
  _prepararCalcCdbLca(sheet); // NOVO BLOCO 4: CDB / LCA

  
  // Protege os rótulos e títulos (Inputs no Topo)
  sheet.getRange('A1:P15').setValues(sheet.getRange('A1:P15').getValues());
  
  ss.toast(`Aba "${NOME_ABA}" preparada com o novo layout v2.`);
}

/**
 * Prepara o Bloco 1: Seguro Prestamista (Layout)
 */
function _prepararCalcPrestamista(sheet) {
  const cfg = CONFIG_CALCULADORA.CALC1;

  // --- Zona de Inputs (Topo) ---
  // CORREÇÃO: Desmembrar o encadeamento (chaining)
  const rangeA1 = sheet.getRange('A1:B1'); // MODIFICADO: Era 'A1:C1', agora 'A1:B1'
  rangeA1.merge();
  rangeA1.setValue('Calculadora 1: Seg. Prestamista (Simulador)');
  rangeA1.setFontWeight('bold');
  rangeA1.setFontSize(12);
  rangeA1.setBackground('#0c8f0a');
  rangeA1.setHorizontalAlignment('center');
  
  const rangeA3 = sheet.getRange('A3');
  rangeA3.setValue('Taxa Mensal (i_m):');
  rangeA3.setFontWeight('bold');
  
  // MODIFICAÇÃO: Pré-insere a taxa e aplica o formato
  const rangeTaxa = sheet.getRange(cfg.TAXA);
  rangeTaxa.setValue(0.001486); // Pré-insere 0,1486%
  rangeTaxa.setNumberFormat('0.0000%').setBackground('#fff9c4');
  
  const rangeA4 = sheet.getRange('A4');
  rangeA4.setValue('Prazo Máximo (Meses):');
  rangeA4.setFontWeight('bold');
  sheet.getRange(cfg.PRAZO_MAXIMO).setNumberFormat('0').setBackground('#fff9c4');
  
  const rangeA6 = sheet.getRange('A6');
  rangeA6.setValue('Valores de Capital a Simular:');
  rangeA6.setFontWeight('bold');
  sheet.getRange(cfg.CAPITAIS_RANGE).setNumberFormat('R$ #,##0.00').setBackground('#fff9c4'); // Isto agora aplica o formato em B6:B10

  // --- Zona de Resultados (Base) ---
  // CORREÇÃO: Desmembrar o encadeamento (chaining)
  const rangeResultado = sheet.getRange(cfg.RESULTADO_INICIO_ROW - 1, cfg.RESULTADO_COL); // MODIFICADO: -2 -> -1. (Linha 12)
  rangeResultado.setValue('Resultados - Seguro Prestamista');
  rangeResultado.setFontWeight('bold');
  rangeResultado.setFontSize(11);
}

/**
 * Prepara o Bloco 3: Seguro de Vida (Layout)
 */
function _prepararCalcPremioVida(sheet) {
  const cfg = CONFIG_CALCULADORA.CALC3;

  // --- Zona de Inputs (Topo) ---
  const rangeI1 = sheet.getRange('D1:H1');
  rangeI1.merge();
  rangeI1.setValue('Calculadora 3: Seg. de Vida (Simulador)');
  rangeI1.setFontWeight('bold');
  rangeI1.setFontSize(12);
  rangeI1.setBackground('#0c8f0a');
  rangeI1.setHorizontalAlignment('center');

  const rangeI3 = sheet.getRange('D3');
  rangeI3.setValue('Qtd. Funcionários:');
  rangeI3.setFontWeight('bold');
  sheet.getRange(cfg.QTD_FUNC).setBackground('#fff9c4');
  
  const rangeI4 = sheet.getRange('D4');
  rangeI4.setValue('Qtd. Sócios/Adms:');
  rangeI4.setFontWeight('bold');
  sheet.getRange(cfg.QTD_SOCIOS).setBackground('#fff9c4');

  // --- TRAVA VISUAL: TETO CAPITAL INDIVIDUAL (Linha 5) ---
  // Modificação estritamente necessária para indicar o limite
  const rangeI5 = sheet.getRange('D5');
  rangeI5.setValue('Teto Capital Ind. (Fixo):');
  rangeI5.setFontWeight('bold');
  
  sheet.getRange('E5')
       .setValue(100000) // Valor numérico puro para facilitar leitura
       .setNumberFormat('R$ #,##0.00')
       .setBackground('#e0e0e0') // Cinza: indica campo travado/sistema
       .setFontWeight('bold');
  // -------------------------------------------------------
  
  const rangeI6 = sheet.getRange('D6');
  rangeI6.setValue('Teto do Prêmio (R$):');
  rangeI6.setFontWeight('bold');
  
  sheet.getRange(cfg.TETO_PREMIO).setNumberFormat('R$ #,##0.00').setBackground('#fff9c4');

  // --- Zona de Resultados (Base) ---
  const rangeResultado = sheet.getRange(cfg.RESULTADO_INICIO_ROW - 1, cfg.RESULTADO_COL);
  rangeResultado.setValue('Resultados - Simulação Seguro de Vida');
  rangeResultado.setFontWeight('bold');
  rangeResultado.setFontSize(11);
    
  // Cabeçalho da Tabela de Saída
  const headers = [
    'Capital Individual (R$)', 'Prêmio Mód 1', 'Prêmio Mód 2', 'Prêmio Mód 3', 'Prêmio Mód 4'
  ];
  sheet.getRange(cfg.RESULTADO_INICIO_ROW, cfg.RESULTADO_COL, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
}
/**
 * Prepara o Bloco 2: PRONAMPE (Layout)
 */
function _prepararCalcPronampe(sheet) {
  const cfg = CONFIG_CALCULADORA.CALC2;

  // --- Zona de Inputs (Topo) ---
  // CORREÇÃO: Desmembrar o encadeamento (chaining)
  const rangeO1 = sheet.getRange('J1:P1');
  rangeO1.merge();
  rangeO1.setValue('Calculadora 2: Simulação PRONAMPE');
  rangeO1.setFontWeight('bold');
  rangeO1.setFontSize(12);
  rangeO1.setBackground('#0c8f0a');
  rangeO1.setHorizontalAlignment('center');

  const rangeO3 = sheet.getRange('J3');
  rangeO3.setValue('Taxa Selic (% a.a.):');
  rangeO3.setFontWeight('bold');
  sheet.getRange(cfg.SELIC).setNumberFormat('0.00"%"').setBackground('#fff9c4');
  
  const rangeO4 = sheet.getRange('J4');
  rangeO4.setValue('Valor Financiado:');
  rangeO4.setFontWeight('bold');
  sheet.getRange(cfg.VALOR).setNumberFormat('R$ #,##0.00').setBackground('#fff9c4');
  
  const rangeO5 = sheet.getRange('J5');
  rangeO5.setValue('Prazo Pagamento (meses):');
  rangeO5.setFontWeight('bold');
  sheet.getRange(cfg.PRAZO).setBackground('#fff9c4');
  
  const rangeO6 = sheet.getRange('J6');
  rangeO6.setValue('Carência (meses):');
  rangeO6.setFontWeight('bold');
  sheet.getRange(cfg.CARENCIA).setBackground('#fff9c4');
  
  const rangeO7 = sheet.getRange('J7');
  rangeO7.setValue('Data Início Contrato:');
  rangeO7.setFontWeight('bold');
  sheet.getRange(cfg.DATA_INICIO).setNumberFormat('dd/mm/yyyy').setBackground('#fff9c4');

  // CORREÇÃO: Aplicar formatação em linhas separadas. O encadeamento (chaining) não funciona.
  // MODIFICADO: Movido de O9 para O3
  const rangeO9 = sheet.getRange('O3');
  rangeO9.setValue('RESUMO (Saída Imediata):');
  rangeO9.setFontWeight('bold');
  rangeO9.setFontStyle('italic'); // CORREÇÃO: setFontItalic(true) -> setFontStyle('italic')

  sheet.getRange('O4').setValue('IOF Total:'); // MODIFICADO: Era O10
  sheet.getRange(cfg.IOF_TOTAL).setNumberFormat('R$ #,##0.00'); // (cfg.IOF_TOTAL é P4)
  sheet.getRange('O5').setValue('VALOR LÍQUIDO:'); // MODIFICADO: Era O11
  sheet.getRange(cfg.VALOR_LIQUIDO).setNumberFormat('R$ #,##0.00').setBackground('#d9ead3').setFontWeight('bold').setFontSize(13);;; // (cfg.VALOR_LIQUIDO é P5)

  // --- Zona de Resultados (Base) ---
  // CORREÇÃO: Desmembrar o encadeamento (chaining)
  const rangeResultado = sheet.getRange(cfg.TABELA_INICIO_ROW - 2, cfg.TABELA_INICIO_COL); // MODIFICADO: -3 -> -2. (14 - 2 = 12)
  rangeResultado.setValue('Resultados - Amortização PRONAMPE');
  rangeResultado.setFontWeight('bold');
  rangeResultado.setFontSize(11);

  // Cabeçalho da Tabela de Saída
  const headers_pronampe = [
    'Data Vencimento', 'Nº Parcela', 'Total Parcela', 'IOF', 'Juros', 'Amortização', 'Saldo Devedor'
  ];
  sheet.getRange(cfg.TABELA_INICIO_ROW - 1, cfg.TABELA_INICIO_COL, 1, headers_pronampe.length) // (14 - 1 = 13)
    .setValues([headers_pronampe])
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
}

/**
 * Prepara o Bloco 4: CDB / LCA (CDI Pós) nas colunas R:U.
 */
/**
 * Prepara o Bloco 4: CDB / LCA (CDI Pós) nas colunas R:U.
 */
function _prepararCalcCdbLca(sheet) {
  // Garante que "sheet" exista mesmo se a função for executada isoladamente
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getSheetByName(NOME_ABA);
    if (!sheet) {
      sheet = ss.insertSheet(NOME_ABA);
    }
  }

  const cfg = CONFIG_CALCULADORA.CALC4;
  if (!cfg) {
    throw new Error('CONFIG_CALCULADORA.CALC4 não definido em CONFIG_CALCULADORA.');
  }

  // Título do bloco (R1:U1)
  const rangeTitulo = sheet.getRange('R1:T1');
  rangeTitulo.merge();
  rangeTitulo.setValue('Calculadora 4: CDB / LCA (CDI Pós)');
  rangeTitulo.setFontWeight('bold');
  rangeTitulo.setFontSize(12);
  rangeTitulo.setBackground('#0c8f0a');
  rangeTitulo.setHorizontalAlignment('center');

  // Ajusta larguras das colunas do bloco (R:U)
  sheet.setColumnWidth(18, 26); // R
  sheet.setColumnWidth(19, 26); // S
  sheet.setColumnWidth(20, 22); // T
  sheet.setColumnWidth(21, 22); // U

  // Limpa área de inputs/outputs do bloco (somente R:U)
  sheet.getRange('R3:U25').clearContent().clearFormat();

  // Rótulos de inputs (coluna R)
  sheet.getRange('R3').setValue('Tipo do Título (CDB/LCA):').setFontWeight('bold');
  sheet.getRange('R4').setValue('Tipo de Investidor (PF/PJ):').setFontWeight('bold');
  sheet.getRange('R5').setValue('Valor Inicial (R$):').setFontWeight('bold');
  sheet.getRange('R6').setValue('Prazo (dias):').setFontWeight('bold');
  sheet.getRange('R7').setValue('CDI Anual (% a.a.):').setFontWeight('bold');
  sheet.getRange('R8').setValue('% do CDI:').setFontWeight('bold');
  sheet.getRange('R9').setValue('Outros Custos (% a.a.):').setFontWeight('bold');

  // Inputs (coluna S)
  sheet.getRange(cfg.TIPO_TITULO).setBackground('#fff9c4');
  sheet.getRange(cfg.TIPO_INVESTIDOR).setBackground('#fff9c4');

  sheet.getRange(cfg.VALOR_INICIAL)
       .setNumberFormat('R$ #,##0.00')
       .setBackground('#fff9c4');

  sheet.getRange(cfg.PRAZO_DIAS)
       .setNumberFormat('0')
       .setBackground('#fff9c4');

  sheet.getRange(cfg.CDI_ANUAL_PERCENT)
       .setNumberFormat('0.00"%"')
       .setBackground('#fff9c4');

  sheet.getRange(cfg.PERCENTUAL_CDI)
       .setNumberFormat('0.00"%"')
       .setBackground('#fff9c4');

  sheet.getRange(cfg.OUTROS_CUSTOS_PERCENT)
       .setNumberFormat('0.00"%"')
       .setBackground('#fff9c4');

  // Título da seção de resultados
  const rotuloResultados = sheet.getRange(cfg.RESULTADO_INICIO_ROW - 1, cfg.LABEL_COL); // ex.: linha 11, coluna R
  rotuloResultados.setValue('Resultados - CDB / LCA (Líquido de IOF/IR/Taxas)');
  rotuloResultados.setFontWeight('bold');
  rotuloResultados.setFontSize(11);

  // Rótulos dos outputs
  const baseRow = cfg.RESULTADO_INICIO_ROW; // 12
  const labels = [
    'Valor Final Bruto:',
    'Rendimento Bruto:',
    'Rentab. Bruta no Período:',
    'Rentab. Bruta a.a.:',
    'IOF (se houver):',
    'IR:',
    'Outros Custos:',
    'VALOR FINAL LÍQUIDO:',
    'Rendimento Líquido:',
    'Rentab. Líquida no Período:',
    'Rentab. Líquida a.a.:'
  ];

  sheet.getRange(baseRow, cfg.LABEL_COL, labels.length, 1)
    .setValues(labels.map(l => [l]))
    .setFontWeight('bold');

  // Formatação padrão dos outputs (coluna S)
  // 1) Moeda (BRL): S12, S13, S16, S17, S18, S19, S20
  sheet.getRange('S12').setNumberFormat('R$ #,##0.00');
  sheet.getRange('S13').setNumberFormat('R$ #,##0.00');
  sheet.getRange('S16').setNumberFormat('R$ #,##0.00');
  sheet.getRange('S17').setNumberFormat('R$ #,##0.00');
  sheet.getRange('S18').setNumberFormat('R$ #,##0.00');
  sheet.getRange('S19').setNumberFormat('R$ #,##0.00');
  sheet.getRange('S20').setNumberFormat('R$ #,##0.00');

  // 2) Percentual: S14, S15, S21, S22
  sheet.getRange('S14').setNumberFormat('0.00"%"');
  sheet.getRange('S15').setNumberFormat('0.00"%"');
  sheet.getRange('S21').setNumberFormat('0.00"%"');
  sheet.getRange('S22').setNumberFormat('0.00"%"');

  // Destaque no Valor Final Líquido
  const rangeVfLiq = sheet.getRange(baseRow + 7, cfg.VALUE_COL); // linha 19, col S
  rangeVfLiq.setBackground('#d9ead3');
  rangeVfLiq.setFontWeight('bold');
  rangeVfLiq.setFontSize(11);
}


// ===================================================================
// 📞 FUNÇÕES DE EXECUÇÃO (v2)
// ===================================================================

/**
 * AÇÃO 1: (NOVO) Simula o Prêmio Prestamista para 5 valores de capital.
 */
function calcularPremioPrestamista() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA);
  if (!sheet) {
    ui.alert(`Aba "${NOME_ABA}" não encontrada. Execute "prepararCalculadora" primeiro.`);
    return;
  }
  
  try {
    const cfg = CONFIG_CALCULADORA.CALC1;
    
    // 1. Ler Inputs
    const taxa = sheet.getRange(cfg.TAXA).getValue();
    const prazoMaximo = sheet.getRange(cfg.PRAZO_MAXIMO).getValue();
    
    // Lê os 5 valores de capital
    const capitais = sheet.getRange(cfg.CAPITAIS_RANGE).getValues().flat(); 
    
    if (taxa <= 0 || prazoMaximo <= 0) {
      ui.alert('Cálculo 1: Taxa Mensal e Prazo Máximo devem ser maiores que zero.');
      return;
    }

    // 2. Limpar área de resultado
    sheet.getRange(cfg.RESULTADO_INICIO_ROW, cfg.RESULTADO_COL, 
                    sheet.getMaxRows() - cfg.RESULTADO_INICIO_ROW, 2).clear();

    let currentRow = cfg.RESULTADO_INICIO_ROW;
    const allData = []; // Buffer para escrita em lote

    // 3. Loop para cada Capital
    for (const capital of capitais) {
      if (typeof capital !== 'number' || capital <= 0) {
        continue; // Pula se a célula estiver vazia ou inválida
      }

      // Adiciona Título
      // CORREÇÃO: Desmembrar o encadeamento E usar setFontStyle('italic')
      const rangeTitulo = sheet.getRange(currentRow, cfg.RESULTADO_COL);
      rangeTitulo.setValue(`Simulação para ${capital.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}`);
      rangeTitulo.setFontWeight('bold');
      rangeTitulo.setFontStyle('italic');
      currentRow++;

      // Adiciona Cabeçalho
      sheet.getRange(currentRow, cfg.RESULTADO_COL, 1, 2)
        .setValues([['Mês (Prazo)', 'Prêmio Total']])
        .setFontWeight('bold').setBackground('#f3f3f3');
      currentRow++;

      const tabelaResultados = [];
      // 4. Loop de progressão mensal
      for (let n = 1; n <= prazoMaximo; n++) {
        const premio = capital * taxa * n;
        tabelaResultados.push([n, premio]);
      }
      
      // 5. Escreve a tabela de dados
      if (tabelaResultados.length > 0) {
        
        // --- INÍCIO DA CORREÇÃO (BUG Linha 341) ---
        // O erro ocorre porque .setNumberFormats() espera um array de formatos
        // com as *mesmas dimensões* do range (ex: 36 linhas x 2 colunas).
        
        // 1. Criar o array de formatos (ex: 36 linhas x 2 colunas)
        const formatoLinha = ['0', 'R$ #,##0.00'];
        const formatosTabela = [];
        for (let i = 0; i < tabelaResultados.length; i++) {
          formatosTabela.push(formatoLinha);
        }

        // 2. Obter o range e aplicar valores e formatos separadamente
        const rangeSaida = sheet.getRange(currentRow, cfg.RESULTADO_COL, tabelaResultados.length, 2);
        rangeSaida.setValues(tabelaResultados);
        rangeSaida.setNumberFormats(formatosTabela);
        // --- FIM DA CORREÇÃO ---
          
        currentRow += tabelaResultados.length + 2; // +2 para espaçamento
      }
    }
    
    ss.toast(`Cálculo 1 (Prestamista) concluído!`);

  } catch (e) {
    ui.alert(`Ocorreu um erro no Cálculo 1: ${e.message}\n${e.stack}`);
  }
}

/**
 * AÇÃO 3: (MODIFICADA) Simula o Prêmio de Seguro de Vida com TRAVA HARDCODED EM 100k.
 */
function simularPremioVidaPorTeto() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA);
  if (!sheet) {
    ui.alert(`Aba "${NOME_ABA}" não encontrada. Execute "prepararCalculadora" primeiro.`);
    return;
  }

  try {
    const cfg = CONFIG_CALCULADORA.CALC3;

    // 1. Ler Inputs
    const qtdFunc = sheet.getRange(cfg.QTD_FUNC).getValue();
    const qtdSocios = sheet.getRange(cfg.QTD_SOCIOS).getValue();
    const tetoPremio = sheet.getRange(cfg.TETO_PREMIO).getValue();

    if (tetoPremio <= 0 || (qtdFunc + qtdSocios) <= 0) {
      ui.alert('Cálculo 3: Teto do Prêmio, Qtd. Funcionários e Qtd. Sócios devem ser maiores que zero.');
      return;
    }

    // 2. Limpar área de resultado (abaixo do cabeçalho)
    const startRowClear = cfg.RESULTADO_INICIO_ROW + 1; 
    sheet.getRange(startRowClear, cfg.RESULTADO_COL, 
                   sheet.getMaxRows() - startRowClear, cfg.TAXAS_MODULOS.length + 1)
                   .clear();

    let capitalIterativo = cfg.CAPITAL_INCREMENTO;
    let continuarLoop = true;
    const resultados = [];

    // 3. Loop de simulação
    while (continuarLoop) {
      
      // >>>>>>>>>> INÍCIO DA MODIFICAÇÃO (TRAVA DE 100k) <<<<<<<<<<
      // Se o capital atual for maior que 100.000, encerra o loop imediatamente.
      // O valor HARDCODED aqui garante que a tabela pare exatamente onde solicitado.
      if (capitalIterativo > 100000) {
        break; 
      }
      // >>>>>>>>>> FIM DA MODIFICAÇÃO <<<<<<<<<<

      const cgTotal = (qtdFunc * capitalIterativo) + (qtdSocios * capitalIterativo);
      
      const premios = cfg.TAXAS_MODULOS.map(taxa => {
        // Lógica: Capital Global * Taxa Mensal (decimal) * 12 (meses)
        return cgTotal * taxa * 12;
      });

      // Adiciona a linha ao buffer de resultados
      resultados.push([capitalIterativo, ...premios]);

      // Condição de parada original (Teto do Prêmio)
      // Mantida caso o prêmio estoure antes de chegar aos 100k de capital
      if (premios[0] > tetoPremio) {
        continuarLoop = false;
      }
      
      capitalIterativo += cfg.CAPITAL_INCREMENTO;
      
      // Trava de segurança para evitar loops infinitos
      if (resultados.length > 500) {
        continuarLoop = false;
        ui.alert('Cálculo 3: A simulação foi interrompida em 500 linhas para evitar lentidão.');
      }
    }

    // 4. Escrever resultados na planilha
    if (resultados.length > 0) {
      sheet.getRange(startRowClear, cfg.RESULTADO_COL, resultados.length, resultados[0].length)
        .setValues(resultados)
        .setNumberFormat('R$ #,##0.00');
    }

    ss.toast(`Cálculo 3 (Seg. Vida) concluído com teto de R$ 100k!`);

  } catch (e) {
    ui.alert(`Ocorreu um erro no Cálculo 3: ${e.message}\n${e.stack}`);
  }
}


/**
 * AÇÃO 2: (MODIFICADA) Calcula a simulação do PRONAMPE.
 */
function calcularPronampe() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(NOME_ABA);

  if (!aba) {
    ui.alert(`Aba "${NOME_ABA}" não encontrada. Execute "prepararCalculadora" primeiro.`);
    return;
  }
  
  try {
    const cfg = CONFIG_CALCULADORA.CALC2;

    // --- Leituras (Inputs) ---
    const selic = parseFloat(aba.getRange(cfg.SELIC).getValue());
    const valor = parseFloat(aba.getRange(cfg.VALOR).getValue());
    const prazo = parseInt(aba.getRange(cfg.PRAZO).getValue()); 
    const carencia = parseInt(aba.getRange(cfg.CARENCIA).getValue()); 
    const dataInicioRaw = aba.getRange(cfg.DATA_INICIO).getValue();
    const dataInicio = new Date(dataInicioRaw);

    const linhaInicioTabela = cfg.TABELA_INICIO_ROW; // (Linha 14)
    const colInicioTabela = cfg.TABELA_INICIO_COL;

    // Validação
    const periodoAmortizacao = prazo - carencia;
    
    // Prazo total deve ser MAIOR que carência (evita divisão por zero)
    if (periodoAmortizacao <= 0) {
      ui.alert('Cálculo 2: Erro! O Prazo Total deve ser maior que a Carência.');
      return;
    }

    if (isNaN(selic) || isNaN(valor) || isNaN(prazo) || isNaN(carencia) || isNaN(dataInicio.getTime())) {
      ui.alert('Cálculo 2: Verifique se todos os campos de entrada estão preenchidos corretamente.');
      return;
    }
    
    // --- Lógica de Cálculo (PRO RATA DIÁRIO + SAC) ---
    const taxaFixaAA = 6;
    const taxaTotalAA = taxaFixaAA + selic; // SELIC + 6% a.a.
    
    // Taxa diária efetiva (ano civil - 365 dias)
    const BASE_DIAS_ANO = 365;
    const taxaDiaria = Math.pow(1 + taxaTotalAA / 100, 1 / BASE_DIAS_ANO) - 1;
    const MS_POR_DIA = 1000 * 60 * 60 * 24;

    // IOF (mantido como no script original – modelo aproximado)
    const dias = prazo * 30; 
    const iofDiario = Math.min(0.0082 * dias, 3.0);
    const iofAdicional = 0.38;
    const iofTotal = valor * ((iofDiario + iofAdicional) / 100);
    
    // SAÍDA DE RESUMO
    const valorLiquido = valor - iofTotal;
    aba.getRange(cfg.IOF_TOTAL).setValue(iofTotal);
    aba.getRange(cfg.VALOR_LIQUIDO).setValue(valorLiquido);

    let saldoDevedor = valor; 
    const todosOsDados = [];

    // Data de referência para contagem dos períodos (começa na data de início do contrato)
    let dataReferencia = new Date(dataInicio);

    // --- 1. Geração das linhas de CARÊNCIA (juros capitalizados pro rata dia) ---
    for (let i = 0; i < carencia; i++) {
      const mesAtual = i + 1;

      // Próxima data de vencimento: 1 mês após a data de referência
      const dataVencimento = new Date(dataReferencia);
      dataVencimento.setMonth(dataVencimento.getMonth() + 1);

      // Quantidade de dias no período (pro rata dia)
      const diffMs = dataVencimento.getTime() - dataReferencia.getTime();
      const diasPeriodo = Math.round(diffMs / MS_POR_DIA);

      // Juros do período com taxa diária
      const fatorPeriodo = Math.pow(1 + taxaDiaria, diasPeriodo);
      const jurosMes = saldoDevedor * (fatorPeriodo - 1);

      const amortizacao = 0; 
      const parcela = 0; 
      
      // Juros são capitalizados no saldo devedor durante a carência
      saldoDevedor += jurosMes;
      const iofLinha = (i === 0) ? iofTotal : 0; 
      
      todosOsDados.push([
        dataVencimento,   // Data de vencimento
        mesAtual,         // Nº Parcela (carência)
        parcela,          // Parcela (0 na carência)
        iofLinha,         // IOF (somente na 1ª linha)
        jurosMes,         // Juros do período
        amortizacao,      // Amortização (0 na carência)
        saldoDevedor      // Saldo devedor atualizado
      ]);

      // Próximo período inicia na data de vencimento atual
      dataReferencia = new Date(dataVencimento);
    }

    // --- 2. Geração das linhas de AMORTIZAÇÃO (SAC com pro rata dia) ---
    // SAC: amortização constante do saldo após a carência
    const saldoInicialAmortizacao = saldoDevedor; 
    const amortizacaoConstante = saldoInicialAmortizacao / periodoAmortizacao;

    for (let i = 0; i < periodoAmortizacao; i++) {
      const mesAtual = carencia + i + 1;

      // Próxima data de vencimento: 1 mês após a data de referência
      const dataVencimento = new Date(dataReferencia);
      dataVencimento.setMonth(dataVencimento.getMonth() + 1);

      // Dias do período (pro rata dia)
      const diffMs = dataVencimento.getTime() - dataReferencia.getTime();
      const diasPeriodo = Math.round(diffMs / MS_POR_DIA);

      // Juros do período
      const fatorPeriodo = Math.pow(1 + taxaDiaria, diasPeriodo);
      const jurosMes = saldoDevedor * (fatorPeriodo - 1);
      
      // Ajuste final para zerar o saldo (tratamento de dízimas)
      let amortizacao = amortizacaoConstante;
      if (i === periodoAmortizacao - 1) { 
        amortizacao = saldoDevedor; // Quita tudo no último período
      }

      const parcela = amortizacao + jurosMes; // SAC: parcela tende a ser decrescente
      let saldoApos = saldoDevedor - amortizacao;
      
      // Proteção contra -0.00
      if (Math.abs(saldoApos) < 0.01) saldoApos = 0;

      todosOsDados.push([
        dataVencimento,   // Data de vencimento
        mesAtual,         // Nº Parcela
        parcela,          // Parcela total
        0,                // IOF (somente na 1ª linha da carência)
        jurosMes,         // Juros do período
        amortizacao,      // Amortização
        saldoApos         // Saldo devedor após pagamento
      ]);
      
      saldoDevedor = saldoApos;
      dataReferencia = new Date(dataVencimento);
    }
    
    // Limpar saída anterior (Limpa 500 linhas a partir da primeira linha da tabela)
    aba.getRange(linhaInicioTabela, colInicioTabela, 500, 7).clearContent();

    // Inserir todos os dados
    if (todosOsDados.length > 0) {
      const rangeSaida = aba.getRange(linhaInicioTabela, colInicioTabela, todosOsDados.length, 7);
      rangeSaida.setValues(todosOsDados);

      // --- Formatação (Otimizada) ---
      // Coluna Data
      aba.getRange(linhaInicioTabela, colInicioTabela, todosOsDados.length, 1)
         .setNumberFormat('dd/MM/yyyy'); 
      // Coluna Nº Parcela
      aba.getRange(linhaInicioTabela, colInicioTabela + 1, todosOsDados.length, 1)
         .setNumberFormat('0');
      // Colunas de valores monetários
      aba.getRange(linhaInicioTabela, colInicioTabela + 2, todosOsDados.length, 5)
         .setNumberFormat('R$ #,##0.00'); 
    }
    
    ss.toast('Cálculo 2 (PRONAMPE SAC com pro rata diário) concluído!');

  } catch (e) {
    ui.alert(`Ocorreu um erro no Cálculo 2 (PRONAMPE): ${e.message}\n${e.stack}`);
  }
}

/**
 * AÇÃO 4: (MODIFICADA) Limpa os inputs e outputs do PRONAMPE.
 */
function limparCalculoPronampe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA);
  if (!sheet) return;
  
  const cfg = CONFIG_CALCULADORA.CALC2;

  // 1. Limpa os inputs (de VALOR até DATA_INICIO) - mantém a Selic
  sheet.getRange(cfg.VALOR + ':' + cfg.DATA_INICIO).clearContent();
  
  // 2. Limpa os outputs de resumo
  sheet.getRange(cfg.IOF_TOTAL).clearContent();
  sheet.getRange(cfg.VALOR_LIQUIDO).clearContent();

  // 3. Limpa a tabela de amortização a partir da primeira linha de dados
  sheet.getRange(cfg.TABELA_INICIO_ROW, cfg.TABELA_INICIO_COL, 
                 sheet.getMaxRows() - cfg.TABELA_INICIO_ROW, 7).clearContent(); 

  ss.toast('Calculadora PRONAMPE limpa!', 'Limpeza', 5);
}

/**
 * AÇÃO 4: Calcula a rentabilidade de CDB / LCA (CDI Pós) no bloco R:U.
 */
function calcularCdbLca() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA);

  if (!sheet) {
    ui.alert(`Aba "${NOME_ABA}" não encontrada. Execute "prepararCalculadora" primeiro.`);
    return;
  }

  try {
    const cfg = CONFIG_CALCULADORA.CALC4;

    // --- 1. Ler inputs ---
    const tipoTituloRaw = sheet.getRange(cfg.TIPO_TITULO).getValue();
    const tipoInvestidorRaw = sheet.getRange(cfg.TIPO_INVESTIDOR).getValue();
    const valorInicial = parseFloat(sheet.getRange(cfg.VALOR_INICIAL).getValue());
    const prazoDias = parseInt(sheet.getRange(cfg.PRAZO_DIAS).getValue(), 10);
    const cdiAnualPercent = parseFloat(sheet.getRange(cfg.CDI_ANUAL_PERCENT).getValue());
    const percentualCdi = parseFloat(sheet.getRange(cfg.PERCENTUAL_CDI).getValue());
    const outrosCustosPercent = parseFloat(sheet.getRange(cfg.OUTROS_CUSTOS_PERCENT).getValue() || 0);

    const tipoTitulo = String(tipoTituloRaw || '').toUpperCase().trim();
    const tipoInvestidor = String(tipoInvestidorRaw || '').toUpperCase().trim();

    // --- 2. Validações básicas ---
    if (tipoTitulo !== 'CDB' && tipoTitulo !== 'LCA') {
      ui.alert('Bloco CDB/LCA: Informe corretamente o Tipo do Título (CDB ou LCA).');
      return;
    }

    if (tipoInvestidor !== 'PF' && tipoInvestidor !== 'PJ') {
      ui.alert('Bloco CDB/LCA: Informe corretamente o Tipo de Investidor (PF ou PJ).');
      return;
    }

    if (!valorInicial || valorInicial <= 0 || !prazoDias || prazoDias <= 0 ||
        !cdiAnualPercent || cdiAnualPercent <= 0 || !percentualCdi || percentualCdi <= 0) {
      ui.alert('Bloco CDB/LCA: Verifique se Valor Inicial, Prazo, CDI Anual e % do CDI estão preenchidos e maiores que zero.');
      return;
    }

    // Regra LCA: mínimo 90 dias
    if (tipoTitulo === 'LCA' && prazoDias < 90) {
      ui.alert('Bloco CDB/LCA: LCA exige prazo mínimo de 90 dias de aplicação.');
      return;
    }

    // --- 3. Conversões ---
    const cdiAnual = cdiAnualPercent / 100;
    const percCdiDecimal = percentualCdi / 100;
    const taxaOutrosAnual = (isNaN(outrosCustosPercent) ? 0 : outrosCustosPercent) / 100;

    // --- 4. Taxa diária do CDI e do título ---
    const fatorCdiDiario = Math.pow(1 + cdiAnual, 1 / BASE_DIAS_ANO_CDI);
    const cdiDia = fatorCdiDiario - 1;

    const taxaDiaTitulo = cdiDia * percCdiDecimal;
    const fatorRendimentoBruto = Math.pow(1 + taxaDiaTitulo, prazoDias);

    // --- 5. Rendimento bruto ---
    const valorFinalBruto = valorInicial * fatorRendimentoBruto;
    const rendimentoBruto = valorFinalBruto - valorInicial;

    // --- 6. IOF (somente CDB, até 30 dias) ---
    let aliquotaIof = 0;
    if (tipoTitulo === 'CDB' && prazoDias <= 30) {
      aliquotaIof = _getAliquotaIofRendaFixa_(prazoDias); // aproximação regressiva
    }
    const iofValor = rendimentoBruto * aliquotaIof;

    // --- 7. IR (CDB sempre; LCA só PJ) ---
    let aliquotaIr = 0;
    if (tipoTitulo === 'CDB') {
      aliquotaIr = _getAliquotaIrRendaFixa_(prazoDias);
    } else if (tipoTitulo === 'LCA' && tipoInvestidor === 'PJ') {
      aliquotaIr = _getAliquotaIrRendaFixa_(prazoDias);
    } // LCA PF: isento (aliquotaIr = 0)

    let baseIr = rendimentoBruto - iofValor;
    if (baseIr < 0) baseIr = 0;
    const irValor = baseIr * aliquotaIr;

    // --- 8. Outros custos (% a.a.) ---
    let outrosCustosValor = 0;
    if (taxaOutrosAnual > 0) {
      const fatorOutrosDiario = Math.pow(1 + taxaOutrosAnual, 1 / BASE_DIAS_ANO_CDI);
      const taxaOutrosDia = fatorOutrosDiario - 1;
      const fatorCustoOutros = Math.pow(1 + taxaOutrosDia, prazoDias);

      const valorAposCusto = valorFinalBruto / fatorCustoOutros;
      outrosCustosValor = valorFinalBruto - valorAposCusto;
    }

    // --- 9. Valores líquidos ---
    const valorFinalLiquido = valorFinalBruto - iofValor - irValor - outrosCustosValor;
    const rendimentoLiquido = valorFinalLiquido - valorInicial;

    // --- 10. Rentabilidades (período e anualizadas) ---
    const rentabBrutaPeriodo = rendimentoBruto / valorInicial;
    const rentabBrutaAA = Math.pow(valorFinalBruto / valorInicial, BASE_DIAS_ANO_CDI / prazoDias) - 1;

    const rentabLiquidaPeriodo = rendimentoLiquido / valorInicial;
    const rentabLiquidaAA = Math.pow(valorFinalLiquido / valorInicial, BASE_DIAS_ANO_CDI / prazoDias) - 1;

    // --- 11. Escrever resultados no bloco R:U ---
    const baseRow = cfg.RESULTADO_INICIO_ROW;
    const colLabel = cfg.LABEL_COL;
    const colValue = cfg.VALUE_COL;

    // Limpa valores antigos (somente área de resultados)
    sheet.getRange(baseRow, colValue, 11, 1).clearContent();

    const linhas = [
      ['Valor Final Bruto:',          valorFinalBruto],
      ['Rendimento Bruto:',          rendimentoBruto],
      ['Rentab. Bruta no Período:',  rentabBrutaPeriodo],
      ['Rentab. Bruta a.a.:',        rentabBrutaAA],
      ['IOF (se houver):',           iofValor],
      ['IR:',                        irValor],
      ['Outros Custos:',             outrosCustosValor],
      ['VALOR FINAL LÍQUIDO:',       valorFinalLiquido],
      ['Rendimento Líquido:',        rendimentoLiquido],
      ['Rentab. Líquida no Período:',rentabLiquidaPeriodo],
      ['Rentab. Líquida a.a.:',      rentabLiquidaAA]
    ];

    // Rótulos já foram escritos no preparo; aqui só garantimos o valor
    sheet.getRange(baseRow, colValue, linhas.length, 1)
      .setValues(linhas.map(l => [l[1]]));

    // Formatação numérica
    // Linhas em moeda: 1,2,5,6,7,8,9
    const linhasMoeda = [0, 1, 4, 5, 6, 7, 8].map(i => baseRow + i);
    linhasMoeda.forEach(r => {
      sheet.getRange(r, colValue).setNumberFormat('R$ #,##0.00');
    });

    // Linhas em percentual: 3,4,10,11 (índices 2,3,9,10)
    const linhasPercent = [2, 3, 9, 10].map(i => baseRow + i);
    linhasPercent.forEach(r => {
      sheet.getRange(r, colValue).setNumberFormat('0.00%');
    });

    // Reforça destaque no Valor Final Líquido
    const rangeVfLiq = sheet.getRange(baseRow + 7, colValue);
    rangeVfLiq.setBackground('#d9ead3');
    rangeVfLiq.setFontWeight('bold');
    rangeVfLiq.setFontSize(11);

    ss.toast('Cálculo 4 (CDB / LCA) concluído!', 'Calculadora CDB/LCA', 5);

  } catch (e) {
    ui.alert(`Ocorreu um erro no Cálculo 4 (CDB/LCA): ${e.message}\n${e.stack}`);
  }
}

