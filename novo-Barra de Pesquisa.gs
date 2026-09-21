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

  // Produtos pendentes: Lista col A (produto) + col B (qtd)
  var listaAB = listaSheet.getRange(2, 1, listaSheet.getLastRow() - 1, 2).getValues();
  var pendentes = listaAB
    .filter(function(r) { return r[0]; })
    .map(function(r) { return { produto: String(r[0]).trim(), qtdNecessaria: Number(r[1]) || 1 }; });

  if (!pendentes.length) return { items: [], pedidos: [], priority: [] };

  var pendentesSet = {};
  pendentes.forEach(function(p) { pendentesSet[p.produto] = true; });

  // Mapeamento msgId → [produtos FALTA]: Lista col D + col F (match exato col D = col A)
  var listaDF = listaSheet.getRange(2, 4, listaSheet.getLastRow() - 1, 3).getValues();

  var msgProdutosMap = {};
  listaDF.forEach(function(r) {
    var produto = String(r[0] || '').trim();
    var idsRaw  = String(r[2] || '').trim();
    if (!produto || !idsRaw || !pendentesSet[produto]) return;
    idsRaw.split(/\s+/).forEach(function(id) {
      if (!id) return;
      if (!msgProdutosMap[id]) msgProdutosMap[id] = [];
      if (msgProdutosMap[id].indexOf(produto) === -1) msgProdutosMap[id].push(produto);
    });
  });

  // Planilha local (abre aqui para leitura dos avulsos)
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Avulsos: produtos adicionados manualmente via Etiquetas_Avulsos
  var avulsosMap = {};
  var avulsoSh = ss.getSheetByName('Etiquetas_Avulsos');
  if (avulsoSh && avulsoSh.getLastRow() >= 2) {
    var avulsoData = avulsoSh.getRange(2, 1, avulsoSh.getLastRow() - 1, 2).getValues();
    avulsoData.forEach(function(r) {
      var prod = String(r[0] || '').trim();
      // Ignora avulsos que já são FALTA (já estão em pendentesSet)
      if (prod && !pendentesSet[prod]) avulsosMap[prod] = Number(r[1]) || 1;
    });
  }

  // Mapeamento avulso: msgId → [avulso produtos] (busca via col D, sem filtro FALTA)
  var avulsoMsgMap = {};
  if (Object.keys(avulsosMap).length) {
    listaDF.forEach(function(r) {
      var produto = String(r[0] || '').trim();
      var idsRaw  = String(r[2] || '').trim();
      if (!produto || !idsRaw || !avulsosMap[produto]) return;
      idsRaw.split(/\s+/).forEach(function(id) {
        if (!id) return;
        if (!avulsoMsgMap[id]) avulsoMsgMap[id] = [];
        if (avulsoMsgMap[id].indexOf(produto) === -1) avulsoMsgMap[id].push(produto);
      });
    });
  }

  // Adiciona avulsos a pendentes (somente os que não são FALTA)
  Object.keys(avulsosMap).forEach(function(prod) {
    pendentes.push({ produto: prod, qtdNecessaria: avulsosMap[prod] });
    pendentesSet[prod] = true;
  });

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
      // Usa col J (ex: "7413#01") como chave única da linha; fallback para msgId|rowNum
      var lineKey = (numCols >= 10 && r[9]) ? String(r[9]).trim() : (msgId + '|' + (idx + 2));
      confMap[msgId].push({
        lineKey:     lineKey,
        idPedido:    String(r[0] || '').trim(),
        marketplace: String(r[1] || '').trim(),
        produtoConf: String(r[3] || '').trim(), // col D: nome do produto no pedido
        qtd:         Number(r[4]) || 1,
        status:      String(r[5] || '').trim()
      });
    });
  }

  // Pedidos novos: FALTA
  var newPedidos = [];
  Object.keys(msgProdutosMap).forEach(function(msgId) {
    var produtos = msgProdutosMap[msgId];
    var confRows = confMap[msgId];
    if (!confRows || !confRows.length) return;

    confRows.forEach(function(conf) {
      if (!conf.marketplace) return;
      if (!conf.status || !conf.status.toUpperCase().startsWith('FALTA')) return;

      var faltaRaw    = conf.status.replace(/^FALTA\s*-\s*/i, '').trim();
      var statusItens = faltaRaw ? faltaRaw.split(/\s*;\s*/).filter(Boolean) : [];

      if (statusItens.length > 0) {
        // Status lista produtos explicitamente (ex: "FALTA - A;B;C").
        // Cada item do status é casado contra TODOS os pendentes, independente do msgProdutosMap.
        // Isso garante que todos os produtos entrem mesmo que apenas 1 tenha o msgId na col F.
        var seenProds = {};
        statusItens.forEach(function(statusItem, idx) {
          var itemLow = statusItem.toLowerCase();
          var itemKws = itemLow.split(/\s+/).filter(function(w) { return w.length > 1; });
          if (!itemKws.length) return;

          var bestMatch = null, bestFwd = 0, bestRev = 0;
          pendentes.forEach(function(p) {
            var nome = p.produto.toLowerCase();
            var fwd  = itemKws.filter(function(kw) { return nome.indexOf(kw) !== -1; }).length;
            if (fwd < itemKws.length) return;
            var prodKws = nome.split(/\s+/).filter(function(w) { return w.length > 1; });
            var rev = prodKws.filter(function(kw) { return itemLow.indexOf(kw) !== -1; }).length;
            if (fwd > bestFwd || (fwd === bestFwd && rev > bestRev)) {
              bestFwd = fwd; bestRev = rev; bestMatch = p.produto;
            }
          });

          if (!bestMatch || seenProds[bestMatch]) return;
          seenProds[bestMatch] = true;
          newPedidos.push({
            discordId:   conf.lineKey + (idx > 0 ? '_' + idx : ''),
            produto:     bestMatch,
            idPedido:    conf.idPedido,
            marketplace: conf.marketplace,
            qtd:         conf.qtd
          });
        });
      } else {
        // "FALTA" puro → usa candidatos do msgProdutosMap.
        // Se há múltiplos candidatos (ex: Kit → 3 perfumes individuais), cria uma entrada para cada.
        var candidates = produtos.filter(function(p) {
          return produtoFaltaNoPedido(conf.status, p);
        });
        if (!candidates.length) return;

        if (candidates.length === 1) {
          newPedidos.push({
            discordId:   conf.lineKey,
            produto:     candidates[0],
            idPedido:    conf.idPedido,
            marketplace: conf.marketplace,
            qtd:         conf.qtd
          });
        } else {
          candidates.forEach(function(p, idx) {
            newPedidos.push({
              discordId:   conf.lineKey + (idx > 0 ? '_' + idx : ''),
              produto:     p,
              idPedido:    conf.idPedido,
              marketplace: conf.marketplace,
              qtd:         conf.qtd
            });
          });
        }
      }
    });
  });

  // Pedidos novos: Avulsos — name-based matching em col D da Conferencia (sem filtro de status)
  Object.keys(avulsoMsgMap).forEach(function(msgId) {
    var avulsoProds = avulsoMsgMap[msgId];
    var confRows = confMap[msgId];
    if (!confRows || !confRows.length) return;

    confRows.forEach(function(conf) {
      if (!conf.marketplace || !conf.produtoConf) return;
      var confNome = conf.produtoConf.toLowerCase();
      var candidates = avulsoProds.filter(function(p) {
        var kws = p.toLowerCase().split(/\s+/).filter(function(w) { return w.length > 2; });
        return kws.length > 0 && kws.every(function(kw) { return confNome.indexOf(kw) !== -1; });
      });
      if (!candidates.length) return;

      var matched = candidates[0];
      if (candidates.length > 1) {
        var bestScore = -1;
        candidates.forEach(function(p) {
          var kws = p.toLowerCase().split(/\s+/).filter(function(w) { return w.length > 2; });
          var score = kws.filter(function(kw) { return confNome.indexOf(kw) !== -1; }).length;
          if (score > bestScore) { bestScore = score; matched = p; }
        });
      }
      newPedidos.push({
        discordId:   conf.lineKey,
        produto:     matched,
        idPedido:    conf.idPedido,
        marketplace: conf.marketplace,
        qtd:         conf.qtd
      });
    });
  });

  // Dedup newPedidos por discordId (FALTA tem prioridade por vir primeiro)
  var seenIds = {};
  newPedidos = newPedidos.filter(function(p) {
    if (seenIds[p.discordId]) return false;
    seenIds[p.discordId] = true;
    return true;
  });

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

  // Adiciona linhas novas para pedidos ainda não registrados
  var rowsToAdd = [];
  newPedidos.forEach(function(p) {
    if (!estadoMap[p.discordId]) {
      rowsToAdd.push([p.discordId, p.produto, p.idPedido, p.marketplace, p.qtd, '', '']);
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
    pedidosFull.push({ discordId: r[0], produto: r[1], idPedido: r[2], marketplace: r[3], qtd: r[4], status: '', obs: '' });
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
    var encontrado    = qtdEnc >= p.qtdNecessaria;
    var naoEncontrado = !encontrado && qtdTotal > 0 && (qtdEnc + qtdNaoEnc) >= qtdTotal;
    return {
      produto:          p.produto,
      qtdNecessaria:    p.qtdNecessaria,
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

// Limpa Etiquetas_Estado e Etiquetas_Avulsos (mantém cabeçalhos)
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
