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

function enviarRelatorioInativo() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Validade");
  if (!sheet) return { ok: false, msg: 'Aba "Validade" não encontrada.' };

  var rows     = sheet.getDataRange().getValues();
  var vencidos = [];
  var criticos = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1]) continue;
    var item = {
      produto:       String(r[1]  || ''),
      validade:      String(r[5]  || ''),
      dias:          Number(r[6]) || 0,
      status:        String(r[7]  || ''),
      estoque:       String(r[9]  || ''),
      observacao:    String(r[10] || '')
    };
    if (item.dias < 0)        vencidos.push(item);
    else if (item.dias <= 30) criticos.push(item);
  }

  if (!vencidos.length && !criticos.length) {
    return { ok: true, msg: 'Nenhum produto crítico ou vencido para reportar.' };
  }

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  function tableRows(list, cor) {
    var out = '';
    for (var k = 0; k < list.length; k++) {
      var p  = list[k];
      var bg = (k % 2 === 0) ? '#fafafa' : '#ffffff';
      out += '<tr style="background:' + bg + '">' +
             '<td style="padding:8px 10px;color:#333">'                      + p.produto   + '</td>' +
             '<td style="padding:8px 10px;text-align:center;color:#555">'    + p.validade  + '</td>' +
             '<td style="padding:8px 10px;text-align:center;color:' + cor + ';font-weight:700">' + p.dias + 'd</td>' +
             '<td style="padding:8px 10px;text-align:center;color:#555">'    + p.estoque   + '</td>' +
             '</tr>';
    }
    return out;
  }

  var html = '' +
    '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;background:#f4f4f4;padding:20px">' +
      '<div style="background:#071a0b;border-radius:12px;padding:24px;margin-bottom:20px">' +
        '<h1 style="margin:0 0 6px;font-size:22px;color:#00e676">Essência do Brasil</h1>' +
        '<p style="margin:0;font-size:11px;color:rgba(0,230,118,0.6);letter-spacing:1px">RELATÓRIO DE VALIDADE — ' + now + '</p>' +
      '</div>';

  if (vencidos.length) {
    html += '' +
      '<div style="background:#fff;border-radius:10px;padding:20px;margin-bottom:16px;border-left:4px solid #c62828">' +
        '<h2 style="color:#c62828;margin:0 0 14px;font-size:16px">&#9888; Vencidos (' + vencidos.length + ')</h2>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
          '<tr style="background:#ffebee">' +
            '<th style="text-align:left;padding:8px 10px;color:#c62828">Produto</th>' +
            '<th style="padding:8px 10px;color:#c62828">Validade</th>' +
            '<th style="padding:8px 10px;color:#c62828">Dias</th>' +
            '<th style="padding:8px 10px;color:#c62828">Estoque</th>' +
          '</tr>' +
          tableRows(vencidos, '#c62828') +
        '</table>' +
      '</div>';
  }

  if (criticos.length) {
    html += '' +
      '<div style="background:#fff;border-radius:10px;padding:20px;margin-bottom:16px;border-left:4px solid #e65100">' +
        '<h2 style="color:#e65100;margin:0 0 14px;font-size:16px">&#9888; Críticos — vence em até 30 dias (' + criticos.length + ')</h2>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
          '<tr style="background:#fff3e0">' +
            '<th style="text-align:left;padding:8px 10px;color:#e65100">Produto</th>' +
            '<th style="padding:8px 10px;color:#e65100">Validade</th>' +
            '<th style="padding:8px 10px;color:#e65100">Dias</th>' +
            '<th style="padding:8px 10px;color:#e65100">Estoque</th>' +
          '</tr>' +
          tableRows(criticos, '#e65100') +
        '</table>' +
      '</div>';
  }

  html += '' +
      '<div style="text-align:center;padding:16px;color:#aaa;font-size:11px">' +
        'Gerado automaticamente pelo Sistema de Gestão Essência do Brasil' +
      '</div>' +
    '</div>';

  var subject = '[Essência do Brasil] Validade — ' +
                vencidos.length + ' vencido(s) · ' +
                criticos.length + ' crítico(s) · ' + now;

  GmailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    subject,
    'Este e-mail requer suporte a HTML para ser exibido corretamente.',
    { htmlBody: html }
  );

  return {
    ok:  true,
    msg: 'Relatório enviado: ' + vencidos.length + ' vencido(s) e ' + criticos.length + ' crítico(s).'
  };
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
