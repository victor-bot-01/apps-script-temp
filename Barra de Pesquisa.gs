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
  var sh = ss.getSheetByName('Status_Inativos');
  var result = [];
  var keysInResult = {};

  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    data.filter(function(r) { return r[0] && r[1]; }).forEach(function(r) {
      var key = r[0].toString() + '|' + r[1].toString();
      keysInResult[key] = true;
      result.push({ produto: String(r[0]), caixa: String(r[1]), qtdOriginal: r[2] || 0 });
    });
  }

  var invSheet = ss.getSheetByName('Inventário');
  if (invSheet && invSheet.getLastRow() >= 2) {
    var dados = invSheet.getDataRange().getValues();
    var HEADER_ROW = 0;
    for (var i = 1; i < dados.length; i++) {
      for (var j = 2; j < dados[0].length; j += 2) {
        var produto = dados[i][j];
        var qtd     = dados[i][j + 1];
        var caixa   = dados[HEADER_ROW][j];
        if (!produto) continue;
        var qtdNum = Number(qtd);
        if (isNaN(qtdNum) || qtdNum > 0) continue;
        var key = produto.toString() + '|' + caixa.toString();
        if (!keysInResult[key]) {
          keysInResult[key] = true;
          result.push({ produto: String(produto), caixa: String(caixa), qtdOriginal: 0 });
        }
      }
    }
  }

  return result;
}

function enviarRelatorioInativosEmail(destinatario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Status_Inativos');
  if (!sh || sh.getLastRow() < 2)
    return { sucesso: true, mensagem: 'Nenhum produto indisponível encontrado.' };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var inativos = data.filter(function(r) { return r[0] && r[1]; });
  if (inativos.length === 0)
    return { sucesso: true, mensagem: 'Nenhum produto indisponível encontrado.' };

  var tz      = Session.getScriptTimeZone();
  var now     = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var linhas  = inativos.map(function(r) {
    return '<tr>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;">' + r[0] + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#777;">' + r[1] + '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;' +
    'background:#040d06;color:#ccc;padding:32px;border-radius:8px;">' +
    '<h2 style="color:#00ff96;margin:0 0 4px;font-size:20px;">Relatório — Produtos Indisponíveis</h2>' +
    '<p style="color:#666;font-size:13px;margin:0 0 24px;">Essência do Brasil · ' + now + '</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">PRODUTO</th>' +
    '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">CAIXA</th></tr>' +
    linhas +
    '</table><p style="margin-top:20px;font-size:12px;color:#555;">Total: <strong style="color:#00ff96;">' +
    inativos.length + '</strong> item(s) indisponível(s)</p></div>';

  var subject = 'Essencia do Brasil - Para Producao dos Rotulos . ' +
    inativos.length + ' item(s) indisponivel(s) . ' + now;

  var recipients = destinatario.split(',').map(function(e) { return e.trim(); }).filter(Boolean);
  recipients.forEach(function(email) {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: html });
  });

  return { sucesso: true, mensagem: 'Relatório enviado para ' + recipients.join(', ') };
}

function gerarPreviewEmailIndisponiveis(destinatarios) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Status_Inativos');
  if (!sh || sh.getLastRow() < 2)
    return { sucesso: true, html: '', assunto: '', total: 0 };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var inativos = data.filter(function(r) { return r[0] && r[1]; });
  if (inativos.length === 0)
    return { sucesso: true, html: '', assunto: '', total: 0 };

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var linhas = inativos.map(function(r) {
    return '<tr>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;">' + r[0] + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #1e3a22;color:#777;">' + r[1] + '</td></tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;' +
    'background:#040d06;color:#ccc;padding:32px;border-radius:8px;">' +
    '<h2 style="color:#00ff96;margin:0 0 4px;font-size:20px;">Relatório — Produtos Indisponíveis</h2>' +
    '<p style="color:#666;font-size:13px;margin:0 0 24px;">Essência do Brasil · ' + now + '</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">PRODUTO</th>' +
    '<th style="text-align:left;padding:10px 14px;color:#00ff96;border-bottom:2px solid #1e4a28;font-size:12px;letter-spacing:1px;">CAIXA</th></tr>' +
    linhas +
    '</table><p style="margin-top:20px;font-size:12px;color:#555;">Total: <strong style="color:#00ff96;">' +
    inativos.length + '</strong> item(s) indisponível(s)</p></div>';

  var assunto = 'Essencia do Brasil - Para Producao dos Rotulos . ' +
    inativos.length + ' item(s) indisponivel(s) . ' + now;

  return { sucesso: true, html: html, assunto: assunto, total: inativos.length };
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
