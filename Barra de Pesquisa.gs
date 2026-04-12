function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Sistema Essência');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function buscarProdutosWeb(termoBusca) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventário");

  const dados = sheet.getDataRange().getValues();

  const HEADER_ROW = 0;
  const START_ROW = 1;

  // separa por "+"
  const palavras = termoBusca
    .toLowerCase()
    .split("+")
    .map(p => p.trim())
    .filter(p => p);

  const resultados = [];

  for (let i = START_ROW; i < dados.length; i++) {
    for (let j = 2; j < dados[0].length; j += 2) {
      const produto = dados[i][j];
      const qtd = dados[i][j + 1];
      const caixa = dados[HEADER_ROW][j];

      if (!produto || !qtd || qtd <= 0) continue;

      const nome = produto.toString().toLowerCase();

      const match = palavras.every(p => nome.includes(p));

      if (match) {
        resultados.push({
          produto: produto,
          caixa: caixa
        });
      }
    }
  }

  return resultados;
}

function getValidadeData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Validade");
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1]) continue;
    result.push({
      codigoBarras:    String(r[0] || ""),
      produto:         String(r[1] || ""),
      dataProducao:    String(r[2] || ""),
      etiqueta:        String(r[3] || ""),
      lote:            String(r[4] || ""),
      validade:        String(r[5] || ""),
      diasRestantes:   Number(r[6]) || 0,
      status:          String(r[7] || ""),
      precisaProduzir: String(r[8] || ""),
      estoque:         String(r[9] || ""),
      observacao:      String(r[10] || "")
    });
  }
  return result;
}
