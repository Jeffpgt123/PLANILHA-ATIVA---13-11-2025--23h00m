function atualizarFormatacaoDiaria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 4 || lastColumn < 1) return;

    var range = sheet.getRange(4, 1, lastRow - 3, lastColumn);

    // Aplica formatação geral
    range.setFontFamily("Arial")
         .setFontSize(13);

    var values = range.getValues();
    var formulas = range.getFormulas();
    var validations = range.getDataValidations();

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        var valor = values[i][j];

        // Ignora fórmulas
        if (formulas[i][j] !== "") continue;

        // Ignora células com validação
        if (validations[i][j] !== null) continue;

        // Verifica se é texto
        if (typeof valor === "string") {
          var maiusculo = valor.toUpperCase();

          // Só altera se for diferente
          if (valor !== maiusculo) {
            sheet.getRange(i + 4, j + 1).setValue(maiusculo);
          }
        }
      }
    }
  });
}
