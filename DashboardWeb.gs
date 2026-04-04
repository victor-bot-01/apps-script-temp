function doGet() {
  return HtmlService
    .createTemplateFromFile("DashboardPage")
    .evaluate()
    .setTitle("ERP Dashboard");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cacheKey = "dashboard_data_v3";

  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CFG.SHEET_NAME);

  if (!sh) {
    throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    const emptyData = {
      generatedAt: formatDateTime_(new Date()),
      kpis: {
        total: 0,
        pending: 0,
        inProgress: 0,
        done: 0,
        blocked: 0,
        overdue: 0
      },
      statusCounts: {
        pending: 0,
        inProgress: 0,
        blocked: 0,
        done: 0
      },
      priorityCounts: {
        urgente: 0,
        alta: 0,
        media: 0,
        baixa: 0
      },
      responsibleCounts: [],
      productivity: {
        last7Days: []
      },
      metrics: {
        avgCompletionHours: ""
      },
      recentHistory: [],
      tasks: []
    };

    try {
      cache.put(cacheKey, JSON.stringify(emptyData), 20);
    } catch (e) {}

    return emptyData;
  }

  const headers = values[0];
  const col = mapHeaders_(headers);

  const recentHistory = getRecentHistory_(ss, 80);
  const historyByTaskId = groupHistoryByTaskId_(recentHistory, 20);
  const commentsByTaskId = getCommentsByTaskId_(ss, 20);

  const tasks = [];
  let total = 0;
  let pending = 0;
  let inProgress = 0;
  let blocked = 0;
  let done = 0;
  let overdue = 0;

  let priorityUrgente = 0;
  let priorityAlta = 0;
  let priorityMedia = 0;
  let priorityBaixa = 0;

  const responsibleMap = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const id = getCellRaw_(row, col, "ID");
    const responsible = getCellRaw_(row, col, "Responsável");
    const task = getCellRaw_(row, col, "Tarefa");
    const dueDate = getDateCell_(row, col, "Prazo");
    const hour = getHourCell_(row, col, "Hora");
    const priority = formatPriorityDashboard_(getCellRaw_(row, col, "Prioridade"));
    const statusRaw = getCellRaw_(row, col, "Status");
    const status = normalizeStatusDashboard_(statusRaw);
    const link = getCellRaw_(row, col, "Link");
    const updatedBy = getCellRaw_(row, col, "AtualizadoPor");
    const updatedAt = getDateTimeCell_(row, col, "AtualizadoEm");

    if (!id && !task && !responsible) {
      continue;
    }

    total++;

    if (status === "PENDENTE") pending++;
    if (status === "ANDAMENTO") inProgress++;
    if (status === "BLOQUEADA") blocked++;
    if (status === "CONCLUÍDA") done++;

    if (isOverdueDashboard_(dueDate, status)) overdue++;

    const priorityUpper = normalizePriorityDashboard_(priority);
    if (priorityUpper === "URGENTE") priorityUrgente++;
    if (priorityUpper === "ALTA") priorityAlta++;
    if (priorityUpper === "MÉDIA") priorityMedia++;
    if (priorityUpper === "BAIXA") priorityBaixa++;

    const responsibleKey = String(responsible || "").trim() || "Sem responsável";
    responsibleMap[responsibleKey] = (responsibleMap[responsibleKey] || 0) + 1;

    const taskIdKey = String(id || "").trim();

    tasks.push({
      id: id || "",
      task: task || "",
      responsible: responsible || "",
      dueDate: dueDate || "",
      hour: hour || "",
      priority: priority || "",
      status: status || "",
      statusLabel: statusLabelDashboard_(status),
      link: link || "",
      updatedBy: updatedBy || "",
      updatedAt: updatedAt || "",
      history: historyByTaskId[taskIdKey] || [],
      comments: commentsByTaskId[taskIdKey] || []
    });
  }

  const productivity = getProductivityData_(ss, 7);
  const avgCompletionHours = getAverageCompletionHours_(ss);

  const response = {
    generatedAt: formatDateTime_(new Date()),
    kpis: {
      total: total,
      pending: pending,
      inProgress: inProgress,
      blocked: blocked,
      done: done,
      overdue: overdue
    },
    statusCounts: {
      pending: pending,
      inProgress: inProgress,
      blocked: blocked,
      done: done
    },
    priorityCounts: {
      urgente: priorityUrgente,
      alta: priorityAlta,
      media: priorityMedia,
      baixa: priorityBaixa
    },
    responsibleCounts: Object.keys(responsibleMap)
      .sort(function(a, b) {
        return responsibleMap[b] - responsibleMap[a];
      })
      .map(function(name) {
        return {
          name: name,
          count: responsibleMap[name]
        };
      }),
    productivity: productivity,
    metrics: {
      avgCompletionHours: avgCompletionHours
    },
    recentHistory: recentHistory.slice(0, 20),
    tasks: tasks
  };

  try {
    cache.put(cacheKey, JSON.stringify(response), 20);
  } catch (e) {}

  return response;
}

function updateTaskStatus(taskId, newStatus) {
  const lock = LockService.getDocumentLock();
  let hasLock = false;

  try {
    lock.waitLock(30000);
    hasLock = true;

    const normalizedTaskId = String(taskId || "").trim();
    const status = normalizeStatusDashboard_(newStatus);

    if (!normalizedTaskId) {
      throw new Error("ID da tarefa não informado.");
    }

    if (!status || ["PENDENTE", "ANDAMENTO", "BLOQUEADA", "CONCLUÍDA"].indexOf(status) === -1) {
      throw new Error("Status inválido.");
    }

    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CFG.SHEET_NAME);

    if (!sh) {
      throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      throw new Error("Nenhuma tarefa encontrada.");
    }

    const headers = values[0];
    const col = mapHeaders_(headers);

    if (col["ID"] === undefined) {
      throw new Error('Cabeçalho "ID" não encontrado.');
    }

    if (col["Status"] === undefined) {
      throw new Error('Cabeçalho "Status" não encontrado.');
    }

    const rowIndex = findTaskRowById_(values, col, normalizedTaskId);
    if (rowIndex === null) {
      throw new Error("Tarefa não encontrada.");
    }

    const rowData = values[rowIndex - 1];
    const oldStatus = normalizeStatusDashboard_(rowData[col["Status"]]);
    const taskName = getCellRaw_(rowData, col, "Tarefa");
    const responsible = getCellRaw_(rowData, col, "Responsável");

    const updatedBy = getCurrentUserLabel_();
    const now = new Date();

    sh.getRange(rowIndex, col["Status"] + 1).setValue(status);

    if (col["AtualizadoPor"] !== undefined) {
      sh.getRange(rowIndex, col["AtualizadoPor"] + 1).setValue(updatedBy);
    }

    if (col["AtualizadoEm"] !== undefined) {
      sh.getRange(rowIndex, col["AtualizadoEm"] + 1).setValue(now);
    }

    SpreadsheetApp.flush();

    if (oldStatus !== status) {
      appendHistoryLog_(ss, {
        changedAt: now,
        taskId: normalizedTaskId,
        taskName: taskName,
        responsible: responsible,
        oldStatus: oldStatus,
        newStatus: status,
        updatedBy: updatedBy
      });
    }

    clearDashboardCache_();

    return {
      success: true,
      taskId: normalizedTaskId,
      status: status,
      statusLabel: statusLabelDashboard_(status),
      updatedBy: updatedBy,
      updatedAt: formatDateTime_(now),
      previousStatus: oldStatus
    };

  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function saveTaskComment(taskId, commentText) {
  const lock = LockService.getDocumentLock();
  let hasLock = false;

  try {
    lock.waitLock(30000);
    hasLock = true;

    const normalizedTaskId = String(taskId || "").trim();
    const text = String(commentText || "").trim();

    if (!normalizedTaskId) {
      throw new Error("ID da tarefa não informado.");
    }

    if (!text) {
      throw new Error("Comentário vazio.");
    }

    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    const mainSheet = ss.getSheetByName(CFG.SHEET_NAME);

    if (!mainSheet) {
      throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');
    }

    const values = mainSheet.getDataRange().getValues();
    if (values.length < 2) {
      throw new Error("Nenhuma tarefa encontrada.");
    }

    const headers = values[0];
    const col = mapHeaders_(headers);

    if (col["ID"] === undefined) {
      throw new Error('Cabeçalho "ID" não encontrado.');
    }

    const rowIndex = findTaskRowById_(values, col, normalizedTaskId);
    if (rowIndex === null) {
      throw new Error("Tarefa não encontrada.");
    }

    const rowData = values[rowIndex - 1];
    const taskName = getCellRaw_(rowData, col, "Tarefa");
    const responsible = getCellRaw_(rowData, col, "Responsável");

    const author = getCurrentUserLabel_();
    const now = new Date();

    appendCommentLog_(ss, {
      createdAt: now,
      taskId: normalizedTaskId,
      taskName: taskName,
      responsible: responsible,
      commentText: text,
      author: author
    });

    if (col["AtualizadoPor"] !== undefined) {
      mainSheet.getRange(rowIndex, col["AtualizadoPor"] + 1).setValue(author);
    }

    if (col["AtualizadoEm"] !== undefined) {
      mainSheet.getRange(rowIndex, col["AtualizadoEm"] + 1).setValue(now);
    }

    SpreadsheetApp.flush();
    clearDashboardCache_();

    return {
      success: true,
      taskId: normalizedTaskId,
      author: author,
      commentText: text,
      createdAt: formatDateTime_(now)
    };

  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function appendHistoryLog_(ss, data) {
  const sh = getOrCreateHistorySheet_(ss);

  sh.appendRow([
    data.changedAt || new Date(),
    data.taskId || "",
    data.taskName || "",
    data.responsible || "",
    data.oldStatus || "",
    data.newStatus || "",
    data.updatedBy || ""
  ]);
}

function appendCommentLog_(ss, data) {
  const sh = getOrCreateCommentsSheet_(ss);

  sh.appendRow([
    data.createdAt || new Date(),
    data.taskId || "",
    data.taskName || "",
    data.responsible || "",
    data.commentText || "",
    data.author || ""
  ]);
}

function getOrCreateHistorySheet_(ss) {
  const sheetName = getHistorySheetName_();
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, 7).setValues([[
      "DataHora",
      "ID",
      "Tarefa",
      "Responsável",
      "StatusAnterior",
      "NovoStatus",
      "AtualizadoPor"
    ]]);
    sh.setFrozenRows(1);
  }

  return sh;
}

function getOrCreateCommentsSheet_(ss) {
  const sheetName = getCommentsSheetName_();
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, 6).setValues([[
      "DataHora",
      "ID",
      "Tarefa",
      "Responsável",
      "Comentário",
      "Autor"
    ]]);
    sh.setFrozenRows(1);
  }

  return sh;
}

function getRecentHistory_(ss, limit) {
  const sh = ss.getSheetByName(getHistorySheetName_());
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const rows = values.slice(1).filter(function(row) {
    return row.join("").trim() !== "";
  });

  rows.sort(function(a, b) {
    const dateA = a[0] instanceof Date ? a[0].getTime() : 0;
    const dateB = b[0] instanceof Date ? b[0].getTime() : 0;
    return dateB - dateA;
  });

  return rows.slice(0, limit || 20).map(function(row) {
    const oldStatus = normalizeStatusDashboard_(row[4]);
    const newStatus = normalizeStatusDashboard_(row[5]);

    return {
      dateTime: row[0] instanceof Date ? formatDateTime_(row[0]) : String(row[0] || ""),
      id: String(row[1] || "").trim(),
      task: String(row[2] || "").trim(),
      responsible: String(row[3] || "").trim(),
      oldStatus: oldStatus,
      oldStatusLabel: statusLabelDashboard_(oldStatus),
      newStatus: newStatus,
      newStatusLabel: statusLabelDashboard_(newStatus),
      updatedBy: String(row[6] || "").trim()
    };
  });
}

function groupHistoryByTaskId_(historyItems, perTaskLimit) {
  const map = {};
  const maxPerTask = perTaskLimit || 10;

  historyItems.forEach(function(item) {
    const key = String(item.id || "").trim();
    if (!key) return;

    if (!map[key]) {
      map[key] = [];
    }

    if (map[key].length < maxPerTask) {
      map[key].push({
        dateTime: item.dateTime || "",
        updatedBy: item.updatedBy || "",
        oldStatus: item.oldStatus || "",
        oldStatusLabel: item.oldStatusLabel || "",
        newStatus: item.newStatus || "",
        newStatusLabel: item.newStatusLabel || ""
      });
    }
  });

  return map;
}

function getCommentsByTaskId_(ss, perTaskLimit) {
  const sh = ss.getSheetByName(getCommentsSheetName_());
  const map = {};
  const maxPerTask = perTaskLimit || 20;

  if (!sh) return map;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return map;

  const rows = values.slice(1).filter(function(row) {
    return row.join("").trim() !== "";
  });

  rows.sort(function(a, b) {
    const dateA = a[0] instanceof Date ? a[0].getTime() : 0;
    const dateB = b[0] instanceof Date ? b[0].getTime() : 0;
    return dateB - dateA;
  });

  rows.forEach(function(row) {
    const taskId = String(row[1] || "").trim();
    if (!taskId) return;

    if (!map[taskId]) {
      map[taskId] = [];
    }

    if (map[taskId].length < maxPerTask) {
      map[taskId].push({
        dateTime: row[0] instanceof Date ? formatDateTime_(row[0]) : String(row[0] || ""),
        task: String(row[2] || "").trim(),
        responsible: String(row[3] || "").trim(),
        text: String(row[4] || "").trim(),
        author: String(row[5] || "").trim()
      });
    }
  });

  return map;
}

function getProductivityData_(ss, days) {
  const sh = ss.getSheetByName(getHistorySheetName_());
  const totalDays = days || 7;
  const timeZone = Session.getScriptTimeZone();

  const result = [];
  const baseMap = {};

  for (let i = totalDays - 1; i >= 0; i--) {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - i);

    const key = Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
    baseMap[key] = {
      label: Utilities.formatDate(d, timeZone, "dd/MM"),
      date: Utilities.formatDate(d, timeZone, "dd/MM"),
      key: key,
      count: 0,
      totalChanges: 0,
      movedToDone: 0,
      movedToBlocked: 0
    };
  }

  if (!sh) {
    return {
      last7Days: Object.keys(baseMap).map(function(key) {
        return baseMap[key];
      })
    };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return {
      last7Days: Object.keys(baseMap).map(function(key) {
        return baseMap[key];
      })
    };
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const changedAt = row[0];

    if (!(changedAt instanceof Date)) continue;

    const key = Utilities.formatDate(changedAt, timeZone, "yyyy-MM-dd");
    if (!baseMap[key]) continue;

    const newStatus = normalizeStatusDashboard_(row[5]);

    baseMap[key].count++;
    baseMap[key].totalChanges++;

    if (newStatus === "CONCLUÍDA") {
      baseMap[key].movedToDone++;
    }

    if (newStatus === "BLOQUEADA") {
      baseMap[key].movedToBlocked++;
    }
  }

  Object.keys(baseMap).forEach(function(key) {
    result.push(baseMap[key]);
  });

  return {
    last7Days: result
  };
}

function getAverageCompletionHours_(ss) {
  const sh = ss.getSheetByName(getHistorySheetName_());
  if (!sh) return "";

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return "";

  const firstSeenByTask = {};
  const doneSeenByTask = {};
  const durationsHours = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const changedAt = row[0];
    const taskId = String(row[1] || "").trim();
    const newStatus = normalizeStatusDashboard_(row[5]);

    if (!(changedAt instanceof Date) || !taskId) continue;

    if (!firstSeenByTask[taskId]) {
      firstSeenByTask[taskId] = changedAt;
    }

    if (newStatus === "CONCLUÍDA" && !doneSeenByTask[taskId]) {
      doneSeenByTask[taskId] = changedAt;
    }
  }

  Object.keys(doneSeenByTask).forEach(function(taskId) {
    const start = firstSeenByTask[taskId];
    const end = doneSeenByTask[taskId];
    if (start && end && end.getTime() >= start.getTime()) {
      const hours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
      durationsHours.push(hours);
    }
  });

  if (!durationsHours.length) return "";

  const total = durationsHours.reduce(function(sum, value) {
    return sum + value;
  }, 0);

  return Math.round((total / durationsHours.length) * 10) / 10;
}

function getHistorySheetName_() {
  if (typeof CFG !== "undefined" && CFG.HISTORY_SHEET_NAME) {
    return String(CFG.HISTORY_SHEET_NAME).trim();
  }
  return "HistoricoDashboard";
}

function getCommentsSheetName_() {
  if (typeof CFG !== "undefined" && CFG.COMMENTS_SHEET_NAME) {
    return String(CFG.COMMENTS_SHEET_NAME).trim();
  }
  return "ComentariosDashboard";
}

function clearDashboardCache_() {
  try {
    CacheService.getScriptCache().remove("dashboard_data_v3");
  } catch (e) {}
}

function isOverdueDashboard_(dueDateStr, status) {
  if (!dueDateStr) return false;
  if (status === "CONCLUÍDA") return false;

  const due = parseDateDashboard_(dueDateStr);
  if (!due) return false;

  const today = new Date();
  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  return due.getTime() < todayOnly.getTime();
}

function parseDateDashboard_(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const s = String(value).trim();

  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }

  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  return null;
}

function findTaskRowById_(values, col, taskId) {
  const targetId = String(taskId || "").trim();

  for (let i = 1; i < values.length; i++) {
    const currentId = String(values[i][col["ID"]] || "").trim();
    if (currentId === targetId) {
      return i + 1;
    }
  }

  return null;
}

function getCurrentUserLabel_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) return email;
  } catch (e) {}

  return "Dashboard";
}

function mapHeaders_(headers) {
  const map = {};
  headers.forEach(function(h, i) {
    const key = String(h || "").trim();
    if (key) map[key] = i;
  });
  return map;
}

function getCellRaw_(row, col, key) {
  if (col[key] === undefined) return "";

  const v = row[col[key]];
  if (v === null || v === undefined) return "";

  return String(v).trim();
}

function getDateCell_(row, col, key) {
  if (col[key] === undefined) return "";

  const v = row[col[key]];
  if (!v) return "";

  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  return String(v).trim();
}

function getHourCell_(row, col, key) {
  if (col[key] === undefined) return "";

  const v = row[col[key]];
  if (!v) return "";

  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
  }

  return String(v).trim();
}

function getDateTimeCell_(row, col, key) {
  if (col[key] === undefined) return "";

  const v = row[col[key]];
  if (!v) return "";

  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  }

  return String(v).trim();
}

function normalizeStatusDashboard_(value) {
  const s = String(value || "").trim().toUpperCase();

  if (s === "PENDENTE") return "PENDENTE";
  if (s === "ANDAMENTO") return "ANDAMENTO";
  if (s === "BLOQUEADA") return "BLOQUEADA";
  if (s === "CONCLUÍDA" || s === "CONCLUIDA") return "CONCLUÍDA";

  return s;
}

function statusLabelDashboard_(status) {
  if (status === "PENDENTE") return "Pendente";
  if (status === "ANDAMENTO") return "Em andamento";
  if (status === "BLOQUEADA") return "Bloqueada";
  if (status === "CONCLUÍDA") return "Concluída";
  return status || "Sem status";
}

function normalizePriorityDashboard_(value) {
  const s = String(value || "").trim().toUpperCase();

  if (s === "MEDIA") return "MÉDIA";
  if (s === "MÉDIA") return "MÉDIA";
  if (s === "BAIXA") return "BAIXA";
  if (s === "ALTA") return "ALTA";
  if (s === "URGENTE") return "URGENTE";

  return s;
}

function formatPriorityDashboard_(value) {
  const upper = normalizePriorityDashboard_(value);

  if (upper === "ALTA") return "Alta";
  if (upper === "MÉDIA") return "Média";
  if (upper === "BAIXA") return "Baixa";
  if (upper === "URGENTE") return "Urgente";

  return String(value || "").trim();
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}
