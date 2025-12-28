/**
 * Módulo de Geração de Documentos Google Docs
 * Gatilho: Comando de texto (ex: "+DOC") na coluna configurada.
 * Integração: Usa BatchOperations e Utils.
 * Visual: Gera link amigável (Hyperlink).
 * Correção: Libera permissão de EDIÇÃO pública (qualquer um com o link edita).
 */
const DocumentGenerator = {

  // CONFIGURAÇÃO DO COMANDO DISPARADOR
  COMANDO_GATILHO: "+DOC", 
  
  // TEXTO QUE APARECERÁ NO LINK
  TEXTO_DO_LINK: "📄 ABRIR DOCUMENTO",

  /**
   * Função principal chamada pelo onEditHandler
   */
  criarDocumento: (e) => {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const aba = sheet.getName();
    const linha = e.range.getRow();
    
    // Validação do comando
    const valorDigitado = String(e.value || "").trim().toUpperCase();
    if (valorDigitado !== DocumentGenerator.COMANDO_GATILHO && valorDigitado !== "DOC") return;

    if (typeof CONFIG === 'undefined') return;
    const configAba = CONFIG.SHEETS[aba];
    if (!configAba || !configAba.GOOGLE_DOC_CONFIG) return;
    
    const docConfig = configAba.GOOGLE_DOC_CONFIG;

    try {
      SpreadsheetApp.getActive().toast("📄 Criando documento editável...", "Aguarde");

      // 1. Definição da Pasta e Nome
      const nomeArquivo = DocumentGenerator._montarNomeArquivo(sheet, linha, docConfig.FILE_NAME);
      const folderId = CONFIG.FOLDER_ID;
      if (!folderId) throw new Error("ID da pasta não configurado.");
      
      const folder = DriveApp.getFolderById(folderId);

      // 2. Criação do Documento
      const doc = DocumentApp.create(nomeArquivo);
      const docId = doc.getId();
      const docFile = DriveApp.getFileById(docId);
      
      folder.addFile(docFile);
      DriveApp.getRootFolder().removeFile(docFile);

      // --- ALTERAÇÃO: LIBERAÇÃO TOTAL DE EDIÇÃO ---
      // Define: Qualquer pessoa com o link (ANYONE_WITH_LINK) pode EDITAR (EDIT)
      try {
        docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      } catch (ePerm) {
        console.warn("Aviso: Não foi possível definir permissão de edição. Verifique as políticas de segurança do domínio.", ePerm);
      }
      // ------------------------------------------------

      // 3. Preenchimento do Conteúdo
      DocumentGenerator._preencherConteudo(doc, sheet, linha, nomeArquivo);
      doc.saveAndClose();

      const docUrl = doc.getUrl();

      // 4. Montagem da Fórmula Visual (Hyperlink)
      const formulaLink = `=HYPERLINK("${docUrl}"; "${DocumentGenerator.TEXTO_DO_LINK}")`;

      // 5. Inserção na Planilha
      const colLinkNum = docConfig.LINK || docConfig.TRIGGER_COL;
      const colLinkIdNum = docConfig.FILE_ID;

      if (typeof BatchOperations !== 'undefined' && BatchOperations.add) {
        const enderecoLink = `'${aba}'!${Utils.colunaParaLetra(colLinkNum)}${linha}`;
        BatchOperations.add('setFormula', enderecoLink, formulaLink);
        
        if (colLinkIdNum) {
          const enderecoId = `'${aba}'!${Utils.colunaParaLetra(colLinkIdNum)}${linha}`;
          BatchOperations.add('setValue', enderecoId, docId);
        }
        BatchOperations.execute('DocumentGenerator');

      } else {
        sheet.getRange(linha, colLinkNum).setFormula(formulaLink);
        if (colLinkIdNum) sheet.getRange(linha, colLinkIdNum).setValue(docId);
      }

      SpreadsheetApp.getActive().toast("✅ Doc editável criado!", "Concluído");

    } catch (erro) {
      console.error(erro);
      SpreadsheetApp.getActive().toast("❌ Erro: " + erro.message);
      e.range.setValue("ERRO: " + erro.message);
    }
  },

  /**
   * Helper: Monta o nome do arquivo dinamicamente
   */
  _montarNomeArquivo: (sheet, linha, configFileName) => {
    let partesNome = [];
    const colunas = Array.isArray(configFileName) ? configFileName : [configFileName];

    colunas.forEach(col => {
      let valor = sheet.getRange(linha, col).getDisplayValue();
      if (valor) partesNome.push(valor.trim());
    });

    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let nomeFinal = partesNome.join(" - ");
    return `${dataHoje} | ${nomeFinal || "Novo Documento"}`;
  },

  /**
   * Helper: Preenche o conteúdo inicial
   */
  _preencherConteudo: (doc, sheet, linha, titulo) => {
    const body = doc.getBody();
    const cabecalho = (CONFIG.MODELO_DOCUMENTO && CONFIG.MODELO_DOCUMENTO.CABECALHO) || "DOCUMENTO AUTOMÁTICO";
    
    body.insertParagraph(0, cabecalho).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Referência: ${titulo}`);
    body.appendHorizontalRule();

    body.appendParagraph("Dados do Registro:");
    
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dados = sheet.getRange(linha, 1, 1, lastCol).getDisplayValues()[0];
    
    const tableData = [];
    for(let i=0; i < headers.length; i++) {
        if(headers[i] && dados[i] && headers[i].toString().trim() !== "") {
            tableData.push([headers[i], dados[i]]);
        }
    }
    
    if (tableData.length > 0) {
        const table = body.appendTable(tableData);
        for (let r = 0; r < table.getNumRows(); r++) {
            table.getRow(r).getCell(0).setBold(true).setWidth(150);
        }
    }
    
    body.appendParagraph("\n\n-- Inserir anotações abaixo --");
  }
};