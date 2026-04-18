function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Essência do Brasil — Sistema de Gestão');
}

function buscarProdutosWeb(termoBusca) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventário");
  var dados = sheet.getDataRange().getValues();

  var HEADER_ROW = 0;
  var START_ROW  = 1;

  var palavras = termoBusca
    .toLowerCase()
    .split("+")
    .map(function(p) { return p.trim(); })
    .filter(function(p) { return p; });

  var resultados = [];

  for (var i = START_ROW; i < dados.length; i++) {
    for (var j = 2; j < dados[0].length; j += 2) {
      var produto = dados[i][j];
      var qtd     = dados[i][j + 1];
      var caixa   = dados[HEADER_ROW][j];

      if (!produto || !qtd || qtd <= 0) continue;

      var nome  = produto.toString().toLowerCase();
      var match = palavras.every(function(p) { return nome.indexOf(p) !== -1; });

      if (match) {
        resultados.push({ produto: produto, caixa: caixa });
      }
    }
  }

  return resultados;
}

function getValidadeData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Validade");
  if (!sheet) return [];
  var rows   = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1]) continue;
    result.push({
      codigoBarras:    String(r[0]  || ""),
      produto:         String(r[1]  || ""),
      dataProducao:    String(r[2]  || ""),
      etiqueta:        String(r[3]  || ""),
      lote:            String(r[4]  || ""),
      validade:        String(r[5]  || ""),
      diasRestantes:   Number(r[6]) || 0,
      status:          String(r[7]  || ""),
      precisaProduzir: String(r[8]  || ""),
      estoque:         String(r[9]  || ""),
      observacao:      String(r[10] || "")
    });
  }
  return result;
}
