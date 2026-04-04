function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
