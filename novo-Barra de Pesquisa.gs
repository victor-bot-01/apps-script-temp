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

  var inativos = getInativosSet();
  var resultados = [];

  for (var i = START_ROW; i < dados.length; i++) {
    for (var j = 2; j < dados[0].length; j += 2) {
      var produto = dados[i][j];
      var qtd     = dados[i][j + 1];
      var caixa   = dados[HEADER_ROW][j];

      var qtdNum = Number(qtd);
      if (!produto || isNaN(qtdNum) || qtdNum <= 0) continue;

      var nome  = produto.toString().toLowerCase();
      var match = palavras.every(function(p) { return nome.indexOf(p) !== -1; });

      if (match) {
        var key = produto.toString() + '|' + caixa.toString();
        if (inativos.indexOf(key) !== -1) continue;
        resultados.push({ produto: produto, caixa: caixa, ativo: true });
      }
    }
  }

  return resultados;
}

function getInativosSet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Status_Inativos');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return data
    .filter(function(r) { return r[0] && r[1]; })
    .map(function(r) { return r[0].toString() + '|' + r[1].toString(); });
}

function setProductStatus(produto, caixa, ativo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Status_Inativos');
  if (!sh) {
    sh = ss.insertSheet('Status_Inativos');
    sh.getRange(1, 1, 1, 3).setValues([['Produto', 'Caixa', 'Qtd_Original']]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  var lastRow = sh.getLastRow();
  var existingRow = -1;
  var existingQtdOriginal = 0;

  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString() === produto.toString() &&
          data[i][1].toString() === caixa.toString()) {
        existingRow = i + 2;
        existingQtdOriginal = Number(data[i][2]) || 0;
        break;
      }
    }
  }

  var invSheet = ss.getSheetByName('Inventário');
  var dados = invSheet ? invSheet.getDataRange().getValues() : null;
  var HEADER_ROW = 0;

  if (!ativo) {
    var originalQty = 0;
    if (dados && invSheet) {
      var found = false;
      for (var i = 1; i < dados.length && !found; i++) {
        for (var j = 2; j < dados[0].length && !found; j += 2) {
          if (dados[i][j] && dados[i][j].toString() === produto.toString() &&
              dados[HEADER_ROW][j] && dados[HEADER_ROW][j].toString() === caixa.toString()) {
            originalQty = Number(dados[i][j + 1]) || 0;
            invSheet.getRange(i + 1, j + 2).setValue(0);
            found = true;
          }
        }
      }
    }
    if (existingRow < 0) {
      sh.appendRow([produto, caixa, originalQty]);
    }
  } else {
    if (existingRow > 0) {
      var origQty = existingQtdOriginal || 1;
      if (dados && invSheet) {
        var found2 = false;
        for (var i = 1; i < dados.length && !found2; i++) {
          for (var j = 2; j < dados[0].length && !found2; j += 2) {
            if (dados[i][j] && dados[i][j].toString() === produto.toString() &&
                dados[HEADER_ROW][j] && dados[HEADER_ROW][j].toString() === caixa.toString()) {
              invSheet.getRange(i + 1, j + 2).setValue(origQty);
              found2 = true;
            }
          }
        }
      }
      sh.deleteRow(existingRow);
    }
  }
  return { sucesso: true, produto: produto, ativo: ativo };
}

function getSemEtiquetaData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pedidos_Producao');
  var result = [];

  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    data.filter(function(r) { return r[0]; }).forEach(function(r) {
      result.push({
        produto:     String(r[0]),
        caixa:       String(r[3] || ''),
        qtdOriginal: Number(r[1]) || 0
      });
    });
  }

  return result;
}

function enviarRelatorioInativosEmail(destinatario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pedidos_Producao');
  if (!sh || sh.getLastRow() < 2)
    return { sucesso: true, mensagem: 'Nenhum pedido de produção encontrado.' };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var pedidos = data.filter(function(r) { return r[0]; });
  if (pedidos.length === 0)
    return { sucesso: true, mensagem: 'Nenhum pedido de produção encontrado.' };

  var tz      = Session.getScriptTimeZone();
  var now     = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var linhas  = pedidos.map(function(r) {
    return '<tr>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;">' + r[0] + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#aaa;">' + (r[1] || '') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#777;">' + (r[3] || '') + '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;' +
    'background:#040d06;color:#ccc;padding:32px;border-radius:8px;">' +
    '<h2 style="color:#00ff96;margin:0 0 4px;font-size:20px;">Relatório — Para Produção</h2>' +
    '<p style="color:#666;font-size:13px;margin:0 0 24px;">Essência do Brasil · ' + now + '</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr>' +
      '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">PRODUTO</th>' +
      '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">QTD</th>' +
      '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">LOCAL</th>' +
    '</tr>' +
    linhas +
    '</table><p style="margin-top:20px;font-size:12px;color:#555;">Total: <strong style="color:#00ff96;">' +
    pedidos.length + '</strong> item(s)</p></div>';

  var subject = 'Essencia do Brasil - Para Producao dos Rotulos . ' +
    pedidos.length + ' item(s) . ' + now;

  var recipients = destinatario.split(',').map(function(e) { return e.trim(); }).filter(Boolean);
  recipients.forEach(function(email) {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: html });
  });

  return { sucesso: true, mensagem: 'Relatório enviado para ' + recipients.join(', ') };
}

function gerarPreviewEmailIndisponiveis(destinatarios) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pedidos_Producao');
  if (!sh || sh.getLastRow() < 2)
    return { sucesso: true, html: '', assunto: '', total: 0 };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var pedidos = data.filter(function(r) { return r[0]; });
  if (pedidos.length === 0)
    return { sucesso: true, html: '', assunto: '', total: 0 };

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var linhas = pedidos.map(function(r) {
    return '<tr>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;">' + r[0] + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#aaa;">' + (r[1] || '') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#777;">' + (r[3] || '') + '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;' +
    'background:#040d06;color:#ccc;padding:32px;border-radius:8px;">' +
    '<h2 style="color:#00ff96;margin:0 0 4px;font-size:20px;">Relatório — Para Produção</h2>' +
    '<p style="color:#666;font-size:13px;margin:0 0 24px;">Essência do Brasil · ' + now + '</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr>' +
      '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">PRODUTO</th>' +
      '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">QTD</th>' +
      '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">LOCAL</th>' +
    '</tr>' +
    linhas +
    '</table><p style="margin-top:20px;font-size:12px;color:#555;">Total: <strong style="color:#00ff96;">' +
    pedidos.length + '</strong> item(s)</p></div>';

  var assunto = 'Essencia do Brasil - Para Producao dos Rotulos . ' +
    pedidos.length + ' item(s) . ' + now;
  return { sucesso: true, html: html, assunto: assunto, total: pedidos.length };
}

function gerarPreviewEmailValidade(destinatario, produtos) {
  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  var vencidos = produtos.filter(function(p){ return p.dias !== null && p.dias < 0; });
  var criticos = produtos.filter(function(p){ return p.dias !== null && p.dias >= 0 && p.dias <= 30; });
  var atencao  = produtos.filter(function(p){ return p.dias !== null && p.dias > 30 && p.dias <= 90; });
  var ok       = produtos.filter(function(p){ return p.dias === null || p.dias > 90; });

  var linhas = '';
  function linhasSection(title, list, cor) {
    if (!list.length) return '';
    var out = '<tr><td colspan="4" style="padding:10px 8px 4px;font-weight:700;color:' + cor + ';font-size:12px">' + title + ' (' + list.length + ')</td></tr>';
    for (var k = 0; k < list.length; k++) {
      var p = list[k];
      var diasStr = p.dias === null ? '—' : p.dias + 'd';
      out += '<tr style="background:' + (k%2===0?'#fafafa':'#fff') + '">' +
             '<td style="padding:6px 8px;color:#333">' + (p.produto||'') + '</td>' +
             '<td style="padding:6px 8px;color:#555">' + (p.lote?'Lote '+p.lote:'—') + '</td>' +
             '<td style="padding:6px 8px;color:#555">' + (p.validade||'—') + '</td>' +
             '<td style="padding:6px 8px;color:' + cor + ';font-weight:700">' + diasStr + '</td>' +
             '</tr>';
    }
    return out;
  }

  linhas += linhasSection('Vencidos', vencidos, '#c62828');
  linhas += linhasSection('Críticos', criticos, '#e65100');
  linhas += linhasSection('Atenção', atencao, '#f57f17');
  linhas += linhasSection('OK', ok, '#2e7d32');

  var preview =
    '<div style="font-family:Arial,sans-serif;font-size:13px">' +
    '<div style="background:#071a0b;border-radius:8px;padding:16px;margin-bottom:12px">' +
      '<div style="color:#00e676;font-size:16px;font-weight:700">Essência do Brasil</div>' +
      '<div style="color:rgba(0,230,118,0.6);font-size:10px;letter-spacing:1px">RELATÓRIO DE VALIDADE — ' + now + '</div>' +
    '</div>' +
    '<div style="margin-bottom:8px;color:#555">Para: <strong>' + (destinatario||'—') + '</strong></div>' +
    '<table style="width:100%;border-collapse:collapse">' +
      '<tr style="background:#f0f0f0">' +
        '<th style="text-align:left;padding:6px 8px">Produto</th>' +
        '<th style="padding:6px 8px">Lote</th>' +
        '<th style="padding:6px 8px">Validade</th>' +
        '<th style="padding:6px 8px">Dias</th>' +
      '</tr>' +
      linhas +
    '</table>' +
    '<div style="margin-top:12px;color:#aaa;font-size:11px;text-align:center">' + produtos.length + ' produto(s) selecionado(s)</div>' +
    '</div>';

  return { html: preview };
}

function syncInativosFromInventario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName('Inventário');
  if (!invSheet) return { sucesso: true, adicionados: 0 };

  var sh = ss.getSheetByName('Status_Inativos');
  if (!sh) {
    sh = ss.insertSheet('Status_Inativos');
    sh.getRange(1, 1, 1, 3).setValues([['Produto', 'Caixa', 'Qtd_Original']]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  var existingKeys = {};
  if (sh.getLastRow() >= 2) {
    var existingData = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    existingData.forEach(function(r) {
      if (r[0] && r[1]) existingKeys[r[0].toString() + '|' + r[1].toString()] = true;
    });
  }

  var dados = invSheet.getDataRange().getValues();
  var HEADER_ROW = 0;
  var adicionados = 0;

  for (var i = 1; i < dados.length; i++) {
    for (var j = 2; j < dados[0].length; j += 2) {
      var produto = dados[i][j];
      var qtd     = dados[i][j + 1];
      var caixa   = dados[HEADER_ROW][j];
      if (!produto) continue;
      var qtdNum = Number(qtd);
      if (!isNaN(qtdNum) && qtdNum > 0) continue;
      var key = produto.toString() + '|' + caixa.toString();
      if (!existingKeys[key]) {
        sh.appendRow([produto, caixa, 0]);
        existingKeys[key] = true;
        adicionados++;
      }
    }
  }

  return { sucesso: true, adicionados: adicionados };
}

function enviarRelatorioValidadeEmail(destinatario, produtos) {
  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  var vencidos = produtos.filter(function(p){ return p.dias !== null && p.dias < 0; });
  var criticos = produtos.filter(function(p){ return p.dias !== null && p.dias >= 0 && p.dias <= 30; });
  var atencao  = produtos.filter(function(p){ return p.dias !== null && p.dias > 30 && p.dias <= 90; });
  var ok       = produtos.filter(function(p){ return p.dias === null || p.dias > 90; });

  function tableRows(list, cor) {
    var out = '';
    for (var k = 0; k < list.length; k++) {
      var p  = list[k];
      var bg = (k % 2 === 0) ? '#fafafa' : '#ffffff';
      var diasStr = p.dias === null ? '—' : p.dias + 'd';
      out += '<tr style="background:' + bg + '">' +
             '<td style="padding:8px 10px;color:#333">' + (p.produto||'') + '</td>' +
             '<td style="padding:8px 10px;color:#555">' + (p.lote ? 'Lote '+p.lote : '—') + '</td>' +
             '<td style="padding:8px 10px;text-align:center;color:#555">' + (p.validade||'—') + '</td>' +
             '<td style="padding:8px 10px;text-align:center;color:' + cor + ';font-weight:700">' + diasStr + '</td>' +
             '</tr>';
    }
    return out;
  }

  function section(title, list, borderCor, headerCor, rowCor) {
    if (!list.length) return '';
    return '<div style="background:#fff;border-radius:10px;padding:20px;margin-bottom:16px;border-left:4px solid ' + borderCor + '">' +
      '<h2 style="color:' + headerCor + ';margin:0 0 14px;font-size:16px">' + title + ' (' + list.length + ')</h2>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
        '<tr style="background:#f5f5f5">' +
          '<th style="text-align:left;padding:8px 10px;color:' + headerCor + '">Produto</th>' +
          '<th style="padding:8px 10px;color:' + headerCor + '">Lote</th>' +
          '<th style="padding:8px 10px;color:' + headerCor + '">Validade</th>' +
          '<th style="padding:8px 10px;color:' + headerCor + '">Dias</th>' +
        '</tr>' +
        tableRows(list, rowCor) +
      '</table></div>';
  }

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;background:#f4f4f4;padding:20px">' +
    '<div style="background:#071a0b;border-radius:12px;padding:24px;margin-bottom:20px">' +
      '<h1 style="margin:0 0 6px;font-size:22px;color:#00e676">Essência do Brasil</h1>' +
      '<p style="margin:0;font-size:11px;color:rgba(0,230,118,0.6);letter-spacing:1px">RELATÓRIO DE VALIDADE — ' + now + '</p>' +
    '</div>' +
    section('&#9888; Vencidos', vencidos, '#c62828', '#c62828', '#c62828') +
    section('&#9888; Críticos (≤30 dias)', criticos, '#e65100', '#e65100', '#e65100') +
    section('Atenção (31–90 dias)', atencao, '#f9a825', '#f57f17', '#e65100') +
    section('OK', ok, '#2e7d32', '#2e7d32', '#2e7d32') +
    '<div style="text-align:center;padding:16px;color:#aaa;font-size:11px">Gerado pelo Sistema de Gestão Essência do Brasil</div>' +
    '</div>';

  var subject = '[Essência do Brasil] Validade — ' + vencidos.length + ' vencido(s) · ' +
                criticos.length + ' crítico(s) · ' + now;

  GmailApp.sendEmail(destinatario, subject,
    'Este e-mail requer suporte a HTML.',
    { htmlBody: html }
  );

  return { sucesso: true, mensagem: 'Relatório enviado para ' + destinatario + ' (' + produtos.length + ' produto(s)).' };
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
      produto:    String(r[1]  || ''),
      validade:   String(r[5]  || ''),
      dias:       Number(r[6]) || 0,
      status:     String(r[7]  || ''),
      estoque:    String(r[9]  || ''),
      observacao: String(r[10] || '')
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

function enviarRelatorioParaProducaoEmail(destinatario, produtos) {
  var items = [];
  try { items = JSON.parse(produtos); } catch(e) { items = []; }
  if (!items.length) return { sucesso: true, mensagem: 'Nenhum item selecionado.' };

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  var linhas = items.map(function(r) {
    return '<tr>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;">' + (r.produto||'') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#777;">' + (r.caixa||'') + '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;' +
    'background:#040d06;color:#ccc;padding:32px;border-radius:8px;">' +
    '<h2 style="color:#00ff96;margin:0 0 4px;font-size:20px;">Relatório — Para Produção dos Rótulos</h2>' +
    '<p style="color:#666;font-size:13px;margin:0 0 24px;">Essência do Brasil · ' + now + '</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">PRODUTO</th>' +
    '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">CAIXA</th></tr>' +
    linhas +
    '</table><p style="margin-top:20px;font-size:12px;color:#555;">Total: <strong style="color:#00ff96;">' +
    items.length + '</strong> item(s) selecionado(s)</p></div>';

  var subject = 'Essencia do Brasil - Para Producao dos Rotulos . ' + items.length + ' item(s) selecionado(s) . ' + now;
  var recipients = destinatario.split(',').map(function(e) { return e.trim(); }).filter(Boolean);
  recipients.forEach(function(email) {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: html });
  });

  return { sucesso: true, mensagem: 'Relatório enviado para ' + recipients.join(', ') + ' com ' + items.length + ' item(s).' };
}

function gerarPreviewEmailParaProducao(destinatarios, produtos) {
  var items = [];
  try { items = JSON.parse(produtos); } catch(e) { items = []; }
  if (!items.length) return { sucesso: true, html: '', assunto: '', total: 0 };

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  var linhas = items.map(function(r) {
    return '<tr>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;">' + (r.produto||'') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#777;">' + (r.caixa||'') + '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;' +
    'background:#040d06;color:#ccc;padding:32px;border-radius:8px;">' +
    '<h2 style="color:#00ff96;margin:0 0 4px;font-size:20px;">Relatório — Para Produção dos Rótulos</h2>' +
    '<p style="color:#666;font-size:13px;margin:0 0 24px;">Essência do Brasil · ' + now + '</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">PRODUTO</th>' +
    '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">CAIXA</th></tr>' +
    linhas +
    '</table><p style="margin-top:20px;font-size:12px;color:#555;">Total: <strong style="color:#00ff96;">' +
    items.length + '</strong> item(s) selecionado(s)</p></div>';

  var assunto = 'Essencia do Brasil - Para Producao dos Rotulos . ' + items.length + ' item(s) selecionado(s) . ' + now;
  return { sucesso: true, html: html, assunto: assunto, total: items.length };
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
      validade:        r[4] ? (r[4] instanceof Date ? r[4].toLocaleDateString('pt-BR') : String(r[4])) : "",
      diasRestantes:   (r[5] !== '' && r[5] !== null && r[5] !== undefined && !isNaN(Number(r[5]))) ? Number(r[5]) : null,
      status:          String(r[6]  || ""),
      precisaProduzir: String(r[7]  || ""),
      estoque:         String(r[8]  || ""),
      observacao:      String(r[9]  || "")
    });
  }
  return result;
}

// ─────────────────────────────────────────────
// ETIQUETAS
// ─────────────────────────────────────────────

var LISTA_SHEET_ID = '15ueYlLK9JoLn_lXcP6JiFxs8Ds22KF3Km_FD2cCUMW4';

// Helper: lê Etiquetas_Config como mapa chave→valor
function _getConfigMap(ss) {
  var sh = ss.getSheetByName('Etiquetas_Config');
  if (!sh || sh.getLastRow() < 2) return {};
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var map = {};
  data.forEach(function(r) { if (r[0]) map[String(r[0])] = String(r[1] || ''); });
  return map;
}

// Helper: grava/atualiza um valor em Etiquetas_Config
function _setConfigValue(ss, chave, valor) {
  var sh = ss.getSheetByName('Etiquetas_Config');
  if (!sh) {
    sh = ss.insertSheet('Etiquetas_Config');
    sh.getRange(1, 1, 1, 2).setValues([['Chave', 'Valor']]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  var lastRow = sh.getLastRow();
  var existingRow = -1;
  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === chave) { existingRow = i + 2; break; }
    }
  }
  if (existingRow > 0) {
    sh.getRange(existingRow, 2).setValue(valor);
  } else {
    sh.appendRow([chave, valor]);
  }
}

// Etiquetas_Estado — colunas:
// A: DiscordId | B: Produto | C: ID_Pedido | D: Marketplace | E: Qtd | F: Status_Etiqueta | G: Observacao

function getEtiquetasDoDia() {
  var extSS = SpreadsheetApp.openById(LISTA_SHEET_ID);
  var listaSheet = extSS.getSheetByName('Lista');
  if (!listaSheet || listaSheet.getLastRow() < 2) return { items: [], pedidos: [], priority: [] };

  // Lê Lista col D + E + F (fonte principal — inclui FALTA, TENHO e em branco)
  var listaDEF = listaSheet.getRange(2, 4, listaSheet.getLastRow() - 1, 3).getValues();

  // Planilha local
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Avulsos: produtos adicionados manualmente via Etiquetas_Avulsos
  var avulsosMap = {};
  var avulsoSh = ss.getSheetByName('Etiquetas_Avulsos');
  if (avulsoSh && avulsoSh.getLastRow() >= 2) {
    avulsoSh.getRange(2, 1, avulsoSh.getLastRow() - 1, 2).getValues().forEach(function(r) {
      var prod = String(r[0] || '').trim();
      if (prod) avulsosMap[prod] = Number(r[1]) || 1;
    });
  }

  // Conferencia: msgId → [array de linhas] — um pedido com múltiplos itens gera várias linhas
  var confSheet = extSS.getSheetByName('Conferencia');
  var confMap = {};
  if (confSheet && confSheet.getLastRow() >= 2) {
    var numCols  = Math.min(confSheet.getLastColumn(), 10);
    var confData = confSheet.getRange(2, 1, confSheet.getLastRow() - 1, numCols).getValues();
    confData.forEach(function(r, idx) {
      var msgId = String(r[8] || '').trim(); // col I: DiscordMessageId
      if (!msgId) return;
      if (!confMap[msgId]) confMap[msgId] = [];
      var lineKey = (numCols >= 10 && r[9]) ? String(r[9]).trim() : (msgId + '|' + (idx + 2));
      confMap[msgId].push({
        lineKey:     lineKey,
        idPedido:    String(r[0] || '').trim(),
        marketplace: String(r[1] || '').trim(),
        produtoConf: String(r[3] || '').trim(),
        qtd:         Number(r[4]) || 1,
        status:      String(r[5] || '').trim()
      });
    });
  }

  // Descobre quais msgIds estão presentes na Lista (inclui TENHO, FALTA e em branco)
  var msgIdsNaLista = {};
  listaDEF.forEach(function(r) {
    var idsRaw = String(r[2] || '').trim(); // col F = Discord IDs
    if (!idsRaw) return;
    idsRaw.split(/\s+/).forEach(function(msgId) { if (msgId) msgIdsNaLista[msgId] = true; });
  });

  // Constrói newPedidos: itera confMap (uma entrada por lineKey)
  // Lista col D usada apenas para filtrar quais msgIds existem; nome do produto vem da Conferência col D
  var newPedidos = [];

  Object.keys(confMap).forEach(function(msgId) {
    if (!msgIdsNaLista[msgId]) return; // ignora linhas da Conferência que não estão na Lista
    confMap[msgId].forEach(function(conf) {
      if (!conf.marketplace || !conf.lineKey) return;
      var initialStatus = String(conf.status || '').trim().toUpperCase() === 'TENHO' ? 'Encontrado' : '';
      newPedidos.push({
        discordId:     conf.lineKey,
        produto:       conf.produtoConf, // nome direto da Conferência col D — sem matching, sem ambiguidade
        idPedido:      conf.idPedido,
        marketplace:   conf.marketplace,
        qtd:           conf.qtd,
        initialStatus: initialStatus
      });
    });
  });

  // Avulsos sem correspondência nas ordens normais → ID gerado
  var produtosUsados = {};
  newPedidos.forEach(function(p) { produtosUsados[p.produto] = true; });
  Object.keys(avulsosMap).forEach(function(prod) {
    if (produtosUsados[prod]) return;
    var avulsoId = 'AVULSO_' + prod.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
    newPedidos.push({
      discordId:     avulsoId,
      produto:       prod,
      idPedido:      'AVULSO',
      marketplace:   'AVULSO',
      qtd:           avulsosMap[prod],
      initialStatus: ''
    });
  });

  // Dedup por discordId
  var seenIds = {};
  newPedidos = newPedidos.filter(function(p) {
    if (seenIds[p.discordId]) return false;
    seenIds[p.discordId] = true;
    return true;
  });

  if (!newPedidos.length) return { items: [], pedidos: [], priority: [] };

  // Deriva pendentes dos newPedidos (produtos únicos com qtd total)
  var pendentesMap = {};
  newPedidos.forEach(function(p) {
    pendentesMap[p.produto] = (pendentesMap[p.produto] || 0) + p.qtd;
  });
  var pendentes = Object.keys(pendentesMap).map(function(p) {
    return { produto: p, qtdNecessaria: pendentesMap[p] };
  });
  var pendentesSet = {};
  pendentes.forEach(function(p) { pendentesSet[p.produto] = true; });

  // Lê Etiquetas_Estado atual
  var estadoSh = ss.getSheetByName('Etiquetas_Estado');
  if (!estadoSh) {
    estadoSh = ss.insertSheet('Etiquetas_Estado');
    estadoSh.getRange(1, 1, 1, 7).setValues([['DiscordId', 'Produto', 'ID_Pedido', 'Marketplace', 'Qtd', 'Status_Etiqueta', 'Observacao']]);
    estadoSh.getRange(1, 1, 1, 7).setFontWeight('bold');
  }

  var estadoMap = {}; // discordId → { row, produto, idPedido, marketplace, qtd, status, obs }
  if (estadoSh.getLastRow() >= 2) {
    var estadoData = estadoSh.getRange(2, 1, estadoSh.getLastRow() - 1, 7).getValues();
    estadoData.forEach(function(r, i) {
      var did = String(r[0] || '').trim();
      if (!did) return;
      estadoMap[did] = {
        row:         i + 2,
        produto:     String(r[1] || ''),
        idPedido:    String(r[2] || ''),
        marketplace: String(r[3] || ''),
        qtd:         Number(r[4]) || 1,
        status:      String(r[5] || ''),
        obs:         String(r[6] || '')
      };
    });
  }

  // Remove entradas que saíram da Conferência (pedidos enviados)
  var validIds = {};
  newPedidos.forEach(function(p) { validIds[p.discordId] = true; });

  var rowsToDelete = [];
  Object.keys(estadoMap).forEach(function(did) {
    if (!validIds[did]) {
      rowsToDelete.push(estadoMap[did].row);
      delete estadoMap[did];
    }
  });
  if (rowsToDelete.length) {
    rowsToDelete.sort(function(a, b) { return b - a; }); // de baixo para cima
    rowsToDelete.forEach(function(row) { estadoSh.deleteRow(row); });
  }

  // Adiciona linhas novas para pedidos ainda não registrados
  var rowsToAdd = [];
  newPedidos.forEach(function(p) {
    if (!estadoMap[p.discordId]) {
      rowsToAdd.push([p.discordId, p.produto, p.idPedido, p.marketplace, p.qtd, p.initialStatus || '', '']);
    }
  });
  if (rowsToAdd.length) {
    estadoSh.getRange(estadoSh.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
  }

  // Constrói pedidosFull: linhas do estado + linhas novas (apenas produtos pendentes)
  var pedidosFull = [];
  Object.keys(estadoMap).forEach(function(did) {
    var e = estadoMap[did];
    if (pendentesSet[e.produto]) {
      pedidosFull.push({ discordId: did, produto: e.produto, idPedido: e.idPedido, marketplace: e.marketplace, qtd: e.qtd, status: e.status, obs: e.obs });
    }
  });
  rowsToAdd.forEach(function(r) {
    pedidosFull.push({ discordId: r[0], produto: r[1], idPedido: r[2], marketplace: r[3], qtd: r[4], status: r[5], obs: '' });
  });

  // Prioridade de marketplace
  var configMap   = _getConfigMap(ss);
  var priorityStr = configMap['prioridade_marketplaces'] || '';
  var priority    = priorityStr ? priorityStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];

  // Localização dos produtos no inventário
  var locMap   = {};
  var invSheet = ss.getSheetByName('Inventário');
  if (invSheet && invSheet.getLastRow() >= 2) {
    var dados  = invSheet.getDataRange().getValues();
    var HEADER = 0;
    for (var i = 1; i < dados.length; i++) {
      for (var j = 2; j < dados[0].length; j += 2) {
        var prod  = dados[i][j];
        var qtdI  = Number(dados[i][j + 1]);
        var caixa = dados[HEADER][j];
        if (!prod || isNaN(qtdI) || qtdI <= 0) continue;
        var key = String(prod);
        if (!locMap[key]) locMap[key] = [];
        locMap[key].push({ caixa: String(caixa), qtd: qtdI });
      }
    }
  }

  // Pedidos de produção já solicitados
  var pedidosSet = {};
  var pedidosSheet = ss.getSheetByName('Pedidos_Producao');
  if (pedidosSheet && pedidosSheet.getLastRow() >= 2) {
    var pedidosData = pedidosSheet.getRange(2, 1, pedidosSheet.getLastRow() - 1, 1).getValues();
    pedidosData.forEach(function(r) { if (r[0]) pedidosSet[String(r[0])] = true; });
  }

  // Constrói items agregados por produto
  var items = pendentes.map(function(p) {
    var orders     = pedidosFull.filter(function(o) { return o.produto === p.produto; });
    var qtdEnc     = orders.reduce(function(s, o) { return o.status === 'Encontrado'     ? s + o.qtd : s; }, 0);
    var qtdNaoEnc  = orders.reduce(function(s, o) { return o.status === 'Não Encontrado' ? s + o.qtd : s; }, 0);
    var qtdTotal   = orders.reduce(function(s, o) { return s + o.qtd; }, 0);
    var obs        = orders.length ? orders[0].obs : '';
    // qtdNecessaria = ordens reais presentes; fallback para o configurado se ainda não há ordens
    var qtdNecessaria  = qtdTotal > 0 ? qtdTotal : p.qtdNecessaria;
    var encontrado    = qtdEnc >= qtdNecessaria;
    var naoEncontrado = !encontrado && qtdTotal > 0 && (qtdEnc + qtdNaoEnc) >= qtdTotal;
    return {
      produto:          p.produto,
      qtdNecessaria:    qtdNecessaria,
      qtdEncontrada:    Math.min(qtdEnc, p.qtdNecessaria),
      qtdNaoEncontrada: qtdNaoEnc,
      encontrado:       encontrado,
      naoEncontrado:    naoEncontrado,
      observacao:       obs,
      locais:           locMap[p.produto] || [],
      pedidoMais:       !!pedidosSet[p.produto]
    };
  });

  return { items: items, pedidos: pedidosFull, priority: priority };
}

// Marca status de pedidos específicos por discordId
function salvarStatusPedidos(discordIds, status) {
  if (!discordIds || !discordIds.length) return { sucesso: true };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Etiquetas_Estado');
  if (!sh || sh.getLastRow() < 2) return { sucesso: true };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    var did = String(data[i][0] || '').trim();
    if (discordIds.indexOf(did) !== -1) {
      sh.getRange(i + 2, 6).setValue(status);
    }
  }
  return { sucesso: true };
}

// Salva observação para todas as linhas de um produto
function salvarObservacao(produto, obs) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Etiquetas_Estado');
  if (!sh || sh.getLastRow() < 2) return { sucesso: true };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1] || '') === produto) {
      sh.getRange(i + 2, 7).setValue(obs);
    }
  }
  return { sucesso: true };
}

// Salva ordem de prioridade de marketplaces
function salvarMarketplacePriority(marketplaces) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _setConfigValue(ss, 'prioridade_marketplaces', (marketplaces || []).join(','));
  return { sucesso: true };
}

// Limpa Etiquetas_Estado, Etiquetas_Avulsos e Etiquetas_Historico (mantém cabeçalhos)
function limparEtiquetasEstado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Etiquetas_Estado');
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  }
  var avulsoSh = ss.getSheetByName('Etiquetas_Avulsos');
  if (avulsoSh && avulsoSh.getLastRow() >= 2) {
    avulsoSh.getRange(2, 1, avulsoSh.getLastRow() - 1, 2).clearContent();
  }
  var historicSh = ss.getSheetByName('Etiquetas_Historico');
  if (historicSh && historicSh.getLastRow() >= 2) {
    historicSh.getRange(2, 1, historicSh.getLastRow() - 1, 4).clearContent();
  }
  return { sucesso: true };
}

// Salva snapshot, computa diff filtrado pela Conferência e retorna o texto pronto
function salvarECopiarMarketplace(currentMktJson) {
  var atual = JSON.parse(currentMktJson);
  var anterior = getUltimoHistoricoMarketplace();

  // prodAtivo: produtos ainda ativos na Conferência (Confirmado ≠ SIM) — protegidos de "Removidos"
  // prodConfirmado: produtos com Confirmado = SIM — removidos do snapshot atual antes de comparar
  var prodAtivo = {};
  var prodConfirmado = {};
  try {
    var extSS = SpreadsheetApp.openById(LISTA_SHEET_ID);

    // dadosMap: col D Conferência → [col A Lista]
    var dadosMap = {};
    var dadosSheet = extSS.getSheetByName('Dados');
    if (dadosSheet && dadosSheet.getLastRow() >= 1) {
      dadosSheet.getDataRange().getValues().forEach(function(row) {
        for (var c = 0; c < row.length - 1; c++) {
          var nomeAd = String(row[c] || '').trim();
          var desMemRaw = String(row[c + 1] || '').trim();
          if (nomeAd && desMemRaw) {
            dadosMap[nomeAd] = desMemRaw.split(';').map(function(s) { return s.trim(); }).filter(Boolean);
            c++;
          }
        }
      });
    }

    // msgId → [Lista col A produtos]
    var msgProdMap = {};
    var listaSheet = extSS.getSheetByName('Lista');
    if (listaSheet && listaSheet.getLastRow() >= 2) {
      listaSheet.getRange(2, 4, listaSheet.getLastRow() - 1, 3).getValues().forEach(function(r) {
        var prod = String(r[0] || '').trim();
        var idsRaw = String(r[2] || '').trim();
        if (!prod || !idsRaw) return;
        idsRaw.split(/\s+/).forEach(function(id) {
          if (!id) return;
          if (!msgProdMap[id]) msgProdMap[id] = [];
          if (msgProdMap[id].indexOf(prod) === -1) msgProdMap[id].push(prod);
        });
      });
    }

    // Varre Conferência: separa produtos confirmados (SIM) dos ainda ativos (FALTA/TENHO)
    var confSheet = extSS.getSheetByName('Conferencia');
    if (confSheet && confSheet.getLastRow() >= 2) {
      var numCols = Math.max(confSheet.getLastColumn(), 12);
      confSheet.getRange(2, 1, confSheet.getLastRow() - 1, numCols).getValues().forEach(function(r) {
        var confirmado = String(r[11] || '').trim().toUpperCase(); // col L = Confirmado
        var produtoConf = String(r[3] || '').trim();
        var msgId = String(r[8] || '').trim();
        if (confirmado === 'SIM') {
          // Pedido confirmado: excluir do snapshot atual
          (dadosMap[produtoConf] || []).forEach(function(p) { prodConfirmado[p] = true; });
          (msgProdMap[msgId] || []).forEach(function(p) { prodConfirmado[p] = true; });
          return;
        }
        // Pedido ainda ativo: proteger de aparecer como "Removido"
        (dadosMap[produtoConf] || []).forEach(function(p) { prodAtivo[p] = true; });
        (msgProdMap[msgId] || []).forEach(function(p) { prodAtivo[p] = true; });
      });
    }
  } catch(e) {}

  // Remove do snapshot atual os produtos confirmados (SIM) — para que apareçam como "Removido" no diff
  Object.keys(atual).forEach(function(mkt) {
    Object.keys(atual[mkt]).forEach(function(prod) {
      if (prodConfirmado[prod]) delete atual[mkt][prod];
    });
    if (!Object.keys(atual[mkt]).length) delete atual[mkt];
  });

  var isFirst = !anterior || !Object.keys(anterior).length;
  var diffText = null;

  if (!isFirst) {
    var adicionados = [], removidos = [], semAlteracao = [];
    var allMkts = {};
    Object.keys(anterior).forEach(function(m) { allMkts[m] = true; });
    Object.keys(atual).forEach(function(m) { allMkts[m] = true; });

    Object.keys(allMkts).sort().forEach(function(mkt) {
      var ant = anterior[mkt] || {};
      var atu = atual[mkt]    || {};
      var allProds = {};
      Object.keys(ant).forEach(function(p) { allProds[p] = true; });
      Object.keys(atu).forEach(function(p) { allProds[p] = true; });

      var mktAdd = [], mktRem = [];
      Object.keys(allProds).sort().forEach(function(prod) {
        var qAnt = ant[prod] || 0;
        var qAtu = atu[prod] || 0;
        if (qAtu > qAnt) {
          mktAdd.push('📦 ' + mkt + ' → ' + prod +
            (qAnt === 0
              ? ' \xD7 ' + qAtu + ' un. (novo)'
              : ' foi de \xD7 ' + qAnt + ' para \xD7 ' + qAtu));
        } else if (qAtu < qAnt) {
          if (prodAtivo[prod]) return; // ainda na Conferência com status ≠ SIM — não mostrar
          mktRem.push('📦 ' + mkt + ' → ' + prod +
            (qAtu === 0 ? ' (removido)' : ' foi de \xD7 ' + qAnt + ' para \xD7 ' + qAtu));
        }
      });

      if (!mktAdd.length && !mktRem.length) {
        semAlteracao.push(mkt);
      } else {
        adicionados = adicionados.concat(mktAdd);
        removidos   = removidos.concat(mktRem);
      }
    });

    var lines = [];
    if (adicionados.length) {
      lines.push('Adicionados:');
      adicionados.forEach(function(l) { lines.push('• ' + l); });
    }
    if (removidos.length) {
      if (lines.length) lines.push('');
      lines.push('Removidos/Reduzidos:');
      removidos.forEach(function(l) { lines.push('• ' + l); });
    }
    if (semAlteracao.length) {
      if (lines.length) lines.push('');
      lines.push('Sem alteração: ' + semAlteracao.join(', ') + '.');
    }
    if (!lines.length) lines.push('Sem alterações em relação ao último envio.');
    diffText = lines.join('\n');
  }

  salvarHistoricoMarketplace(JSON.stringify(atual));
  return { isFirst: isFirst, diffText: diffText };
}

// Retorna o último snapshot de Etiquetas_Historico como objeto {mkt: {produto: qtd}}
function getUltimoHistoricoMarketplace() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Etiquetas_Historico');
  if (!sh || sh.getLastRow() < 2) return null;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();

  var lastId = '';
  data.forEach(function(r) {
    var sid = String(r[0] || '').trim();
    if (sid > lastId) lastId = sid;
  });
  if (!lastId) return null;

  var map = {};
  data.forEach(function(r) {
    if (String(r[0] || '').trim() !== lastId) return;
    var mkt  = String(r[1] || '').trim();
    var prod = String(r[2] || '').trim();
    var qtd  = Number(r[3]) || 0;
    if (!mkt || !prod) return;
    if (!map[mkt]) map[mkt] = {};
    map[mkt][prod] = qtd;
  });
  return map;
}

// Salva um novo snapshot em Etiquetas_Historico
function salvarHistoricoMarketplace(snapshotJson) {
  var snapshot = JSON.parse(snapshotJson);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Etiquetas_Historico');
  if (!sh) {
    sh = ss.insertSheet('Etiquetas_Historico');
    sh.getRange(1, 1, 1, 4).setValues([['Snapshot_ID', 'Marketplace', 'Produto', 'Qtd']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  var tz  = Session.getScriptTimeZone();
  var sid = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  var rows = [];
  Object.keys(snapshot).sort().forEach(function(mkt) {
    Object.keys(snapshot[mkt]).sort().forEach(function(prod) {
      rows.push([sid, mkt, prod, snapshot[mkt][prod]]);
    });
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  return { sucesso: true };
}

function pedirMaisEtiqueta(produto, qtdNecessaria, local) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pedidos_Producao');
  if (!sh) {
    sh = ss.insertSheet('Pedidos_Producao');
    sh.getRange(1, 1, 1, 4).setValues([['Produto', 'Qtd Necessária', 'Data do Pedido', 'Local']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(produto)) return { sucesso: true, jaExiste: true };
    }
  }

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  sh.appendRow([produto, qtdNecessaria || 1, now, local || '']);
  return { sucesso: true, jaExiste: false };
}

function produtoFaltaNoPedido(status, nomeProduto) {
  if (!status || !status.toUpperCase().startsWith('FALTA')) return false;
  if (status.trim().toUpperCase() === 'FALTA') return true;
  var faltaRaw = status.replace(/^FALTA\s*-\s*/i, '').trim();
  var nome     = nomeProduto.toLowerCase();
  // Status pode listar múltiplos produtos separados por ";" — basta um coincidir
  var itens = faltaRaw.split(/\s*;\s*/);
  return itens.some(function(item) {
    var keywords = item.toLowerCase().split(/\s+/).filter(function(w) { return w.length > 1; });
    return keywords.every(function(kw) { return nome.indexOf(kw) !== -1; });
  });
}

// Busca produtos na Lista (col D) por termo — para modal de avulsos
function buscarProdutosAvulsos(termo) {
  if (!termo || termo.length < 2) return [];
  var extSS = SpreadsheetApp.openById(LISTA_SHEET_ID);
  var listaSheet = extSS.getSheetByName('Lista');
  if (!listaSheet || listaSheet.getLastRow() < 2) return [];

  // Col D = produto, Col E = qtd (colunas 4 e 5, índice 0-based: 3 e 4)
  var data = listaSheet.getRange(2, 4, listaSheet.getLastRow() - 1, 2).getValues();
  var keywords = termo.toLowerCase().split(/\s+/).filter(function(w) { return w.length > 1; });
  if (!keywords.length) return [];

  var seen = {};
  var results = [];
  data.forEach(function(r) {
    var produto = String(r[0] || '').trim();
    var qtd     = Number(r[1]) || 1;
    if (!produto || seen[produto]) return;
    var nome  = produto.toLowerCase();
    var match = keywords.every(function(kw) { return nome.indexOf(kw) !== -1; });
    if (match) { seen[produto] = true; results.push({ produto: produto, qtd: qtd }); }
  });

  return results;
}

// Adiciona produto avulso em Etiquetas_Avulsos (auto-criada se necessário)
function adicionarAvulso(produto, qtd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Etiquetas_Avulsos');
  if (!sh) {
    sh = ss.insertSheet('Etiquetas_Avulsos');
    sh.getRange(1, 1, 1, 2).setValues([['Produto', 'Qtd']]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  if (sh.getLastRow() >= 2) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(produto).trim()) {
        return { sucesso: true, jaExiste: true };
      }
    }
  }

  sh.appendRow([produto, qtd || 1]);
  return { sucesso: true, jaExiste: false };
}
