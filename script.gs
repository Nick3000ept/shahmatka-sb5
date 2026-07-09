// ═══════════════════════════════════════════════════════════════════
// СБ5 Adminка — Google Apps Script
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Факт';
const SPREADSHEET_ID = '1Vm0W09F1QNvBi3RK0pWjBa2KB19pXXex9kkd9vHq8zk';
const ADMIN_PASSWORD = 'adminACCB3';
const SK_PASSWORD    = 'priemkaCB3';

const C = {
  ROW_ID    : 1,
  CORPUS    : 2,
  FLOOR     : 3,
  EXTRA1    : 4,
  VOLUME    : 5,
  WORK      : 6,
  ORG       : 7,
  STATUS    : 8,
  DATE_END  : 9,
  DATE_RECV : 10,
  PCT       : 11,
  COMMENT   : 12,
  DATE_CHG  : 13,
  AUTHOR    : 14,
  // O+ не трогаем никогда
};

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var action = p.action || '';

    if (action === 'getRows') return jsonOut(getRows(p.corpus || ''));
    if (action === 'ping')    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);

    if (action === 'debugRows') {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheets = ss.getSheets().map(function(s){ return s.getName(); });
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) return jsonOut({error: 'Лист не найден', sheets: sheets});
      var lastRow = sheet.getLastRow();
      var sample = lastRow > 1 ? sheet.getRange(2, 1, Math.min(3, lastRow - 1), 7).getValues() : [];
      var passed = 0, total = 0;
      if (lastRow > 1) {
        var vals = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
        for (var i = 0; i < vals.length; i++) {
          total++;
          var r = vals[i];
          var corpus = String(r[1]).trim();
          var floor  = String(r[2]).trim();
          var work   = String(r[5]).trim();
          if (corpus && floor && work) passed++;
        }
      }
      return jsonOut({sheetName: SHEET_NAME, lastRow: lastRow, totalRows: total, passedFilter: passed, sheets: sheets, sample: sample});
    }

    if (action === 'getContractors') {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName('Подрядчики');
      if (!sheet) return jsonOut({contractors: []});
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return jsonOut({contractors: []});
      var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var list = vals.map(function(r){ return String(r[0]).trim(); }).filter(Boolean);
      var seen = {};
      list = list.filter(function(c){ if (seen[c]) return false; seen[c] = true; return true; });
      return jsonOut({contractors: list});
    }

    if (action === 'clearCache') { clearCache(); return jsonOut({ok: true}); }
    if (action === 'debugDict') {
      clearCache();
      var dict = getWorkDict();
      var keys = Object.keys(dict);
      var q = p.q || '';
      var match = q ? (dict[q] || null) : null;
      var sample = keys.slice(0, 20);
      return jsonOut({total: keys.length, query: q, match: match, sample: sample});
    }
    if (action === 'getCheckLists') return jsonOut(getCheckLists());
    if (action === 'getStaffing')   return jsonOut(getStaffing());
    if (action === 'getProtocol')   return jsonOut(getProtocol());


    if (action === 'checkPassword') {
      var pwd = p.pwd || '';
      if (pwd === ADMIN_PASSWORD) return jsonOut({ok: true, role: 'admin'});
      if (pwd === SK_PASSWORD)    return jsonOut({ok: true, role: 'sk'});
      return jsonOut({ok: false});
    }

    try {
      return HtmlService.createHtmlOutputFromFile('index.html')
        .setTitle('СБ5 · Админ Панель')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (e2) {
      return HtmlService.createHtmlOutput('<body style="font-family:sans-serif;padding:40px"><h2>Добавьте index.html</h2></body>');
    }
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'saveRow') {
      saveOneRow(body);
      clearCache();
      return jsonOut({ok: true});
    }

    if (body.action === 'saveAll') {
      var rows = body.rows || [];
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = findSheet(ss);
      if (!sheet) return jsonOut({error: 'Лист не найден'});

      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return jsonOut({ok: true, saved: 0});

        var now = new Date();
        var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy');

        // Читаем rowId (столбец A) отдельно — для построения карты
        var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

        // Читаем только G-N (7 столбцов: org/status/dateEnd/dateRecv/pct/comment/dateChg/author)
        // A-F не трогаем — там dropdown-валидация, запись вызовет ошибку
        var EDIT_START = 7; // столбец G
        var EDIT_COLS  = 8; // G..N
        var allValues = sheet.getRange(2, EDIT_START, lastRow - 1, EDIT_COLS).getValues();

        // Строим карту rowId → индекс в массиве
        var rowMap = {};
        idValues.forEach(function(row, i) {
          var id = String(row[0]).trim();
          if (id) rowMap[id] = i;
        });

        // Смещение: индекс в allValues = C.X - EDIT_START
        var off = EDIT_START; // 7
        var changedIndices = [];
        rows.forEach(function(r) {
          var idx = rowMap[String(r.rowId)];
          if (idx === undefined) return;

          // Меняем ТОЛЬКО пришедшие поля — остальные остаются как были
          var anyChange = false;
          if (r.status  !== undefined) { allValues[idx][C.STATUS   - off] = r.status  || ''; anyChange = true; }
          if (r.pct     !== undefined) { allValues[idx][C.PCT      - off] = r.pct     || ''; anyChange = true; }
          if (r.comment !== undefined) { allValues[idx][C.COMMENT  - off] = r.comment || ''; anyChange = true; }
          if (r.org     !== undefined) { allValues[idx][C.ORG      - off] = r.org     || ''; anyChange = true; }
          if (r.dateEnd !== undefined) { allValues[idx][C.DATE_END - off] = r.dateEnd || ''; anyChange = true; }
          if (r.author)                  allValues[idx][C.AUTHOR   - off] = r.author;
          // DATE_CHG только если действительно что-то изменилось
          if (anyChange) allValues[idx][C.DATE_CHG - off] = nowStr;
          if (anyChange) changedIndices.push(idx);
        });

        // Пишем ТОЛЬКО изменённые строки — не трогаем остальные (избегаем ошибок валидации)
        // G может иметь dropdown-валидацию: при ошибке fallback на H-N
        if (changedIndices.length > 0) {
          changedIndices.forEach(function(idx) {
            try {
              sheet.getRange(idx + 2, EDIT_START, 1, EDIT_COLS).setValues([allValues[idx]]);
            } catch(e) {
              // G (org) вызвал ошибку валидации — пишем только H-N
              sheet.getRange(idx + 2, 8, 1, 7).setValues([allValues[idx].slice(1)]);
            }
          });
          SpreadsheetApp.flush();
        }

        clearCache();
      } finally {
        lock.releaseLock();
      }
      return jsonOut({ok: true, saved: changedIndices.length, requested: rows.length});
    }

    if (body.action === 'addProtocol') {
      addProtocolEntry(body);
      return jsonOut({ok: true});
    }

    return jsonOut({error: 'Unknown action: ' + body.action});
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

function getWorkDict() {
  // CacheService: справочник работ кешируется на 2 часа (~20-36 КБ, хорошо укладывается в лимит 100 КБ/ключ)
  // clearCache() сбрасывает ключ sb5_work_dict при любом сохранении данных
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sb5_work_dict');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Факт_работы');
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var dict = {};
  values.forEach(function(row) {
    var name = String(row[2]).trim();
    if (!name) return;
    var place   = String(row[3]).trim();
    var lvl1    = String(row[4]).trim();
    var lvl2    = String(row[5]).trim();
    var kp      = String(row[9]).trim();
    var factNum = String(row[10]).trim();
    // Не перезаписываем уже заполненную запись пустой (защита от дублей без данных)
    var existing = dict[name];
    if (existing && (existing.place || existing.lvl1 || existing.lvl2) && !place && !lvl1 && !lvl2) return;
    dict[name] = {place: place, lvl1: lvl1, lvl2: lvl2, kp: kp, factNum: factNum};
  });

  try { cache.put('sb5_work_dict', JSON.stringify(dict), 7200); } catch(e) {}
  return dict;
}

function getRows(filterCorpus) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || findSheet(ss);
  if (!sheet) return {error: 'Лист не найден'};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {rows: []};
  var dict = getWorkDict();
  var values = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var corpus = String(row[1]).trim();
    var floor  = String(row[2]).trim();
    var work   = String(row[5]).trim();
    if (!corpus || !floor || !work) continue;
    if (filterCorpus && corpus !== filterCorpus) continue;
    var rowId = String(row[0]).trim() || ('row_' + (i + 2));
    var attrs = dict[work] || {place: '', lvl1: '', lvl2: '', kp: '', factNum: ''};
    rows.push({
      rowId      : rowId,
      corpus     : corpus,
      floor      : parseFloor(floor),
      work       : work,
      extra1     : String(row[3]).trim(),
      org        : String(row[6]).trim(),
      status     : String(row[7]).trim(),
      dateEnd    : formatDateOut(row[8]),
      pct        : String(row[10]).trim(),
      comment    : String(row[11]).trim(),
      dateChg    : formatDateOut(row[12]),
      author     : String(row[13]).trim(),
      place      : attrs.place,
      lvl1       : attrs.lvl1,
      lvl2       : attrs.lvl2,
      kp         : attrs.kp,
      factNum    : attrs.factNum,
      baseDate   : formatDateOut(row[14]),
      currentDate: formatDateOut(row[15]),
      volume     : String(row[16]).trim(),
      unit       : String(row[17]).trim(),
      idFact     : String(row[19]).trim()
    });
  }
  rows.sort(function(a, b) {
    var cc = a.corpus < b.corpus ? -1 : a.corpus > b.corpus ? 1 : 0;
    return cc !== 0 ? cc : (b.floor || 0) - (a.floor || 0);
  });
  return {rows: rows};
}

// Построчное сохранение — для малого числа строк
function saveOneRow(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss);
  if (!sheet) throw new Error('Лист не найден');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Таблица пуста');

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var targetRowSheet = -1;
  for (var i = 0; i < ids.length; i++) {
    var rowId = String(ids[i][0]).trim();
    if (rowId && rowId === String(data.rowId)) { targetRowSheet = i + 2; break; }
  }
  // Запасной вариант для синтетических ID (row_N)
  if (targetRowSheet < 0 && String(data.rowId).startsWith('row_')) {
    var synRow = parseInt(String(data.rowId).replace('row_', ''));
    if (!isNaN(synRow)) targetRowSheet = synRow;
  }
  if (targetRowSheet < 0) throw new Error('Строка не найдена: ' + data.rowId);

  // Читаем только G-N (org/status/dateEnd/dateRecv/pct/comment/dateChg/author)
  // A-F не трогаем — там dropdown-валидация, запись вызовет ошибку
  var EDIT_START = 7; // столбец G
  var EDIT_COLS  = 8; // G..N
  var rowValues = sheet.getRange(targetRowSheet, EDIT_START, 1, EDIT_COLS).getValues()[0];
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

  var off = EDIT_START; // 7 — смещение: индекс в rowValues = C.X - off
  var anyChange = false;
  if (data.status  !== undefined) { rowValues[C.STATUS   - off] = data.status  || ''; anyChange = true; }
  if (data.dateEnd !== undefined) { rowValues[C.DATE_END - off] = data.dateEnd || ''; anyChange = true; }
  if (data.pct     !== undefined) { rowValues[C.PCT      - off] = data.pct     || ''; anyChange = true; }
  if (data.org     !== undefined) { rowValues[C.ORG      - off] = data.org     || ''; anyChange = true; }
  if (data.comment !== undefined) { rowValues[C.COMMENT  - off] = data.comment || ''; anyChange = true; }
  if (data.author)                  rowValues[C.AUTHOR   - off] = data.author;
  // DATE_CHG только если действительно что-то изменилось
  if (anyChange) rowValues[C.DATE_CHG - off] = nowStr;
  if (!anyChange) return; // нечего писать — выходим без записи в таблицу

  // G (org) может иметь dropdown-валидацию: при ошибке fallback на H-N
  try {
    sheet.getRange(targetRowSheet, EDIT_START, 1, EDIT_COLS).setValues([rowValues]);
  } catch(e) {
    sheet.getRange(targetRowSheet, 8, 1, 7).setValues([rowValues.slice(1)]);
  }
  SpreadsheetApp.flush();
}

function findSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) return sheet;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().indexOf('общ') >= 0) return sheets[i];
  }
  return null;
}

function parseFloor(v) {
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? v : n;
}

function formatDateOut(v) {
  if (!v) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return pad(v.getDate()) + '.' + pad(v.getMonth() + 1) + '.' + v.getFullYear();
  }
  return String(v).trim();
}

function pad(n) { return n < 10 ? '0' + n : String(n); }

function getStaffing() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Численность_монтажников');
  if (!sheet) return {items: []};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {items: []};
  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var items = [];
  values.forEach(function(row) {
    var date = formatDateOut(row[0]);
    var contractor = String(row[1]).trim();
    var count = parseFloat(String(row[2]).replace(',', '.'));
    if (!date || !contractor || isNaN(count) || count <= 0) return;
    items.push({date: date, contractor: contractor, count: count});
  });
  return {items: items};
}

function getCheckLists() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sb5_checklists');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Чек_листы');
  if (!sheet) return {items: []};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {items: []};
  // A:AB = 28 столбцов
  var values = sheet.getRange(2, 1, lastRow - 1, 28).getValues();
  var items = [];
  values.forEach(function(row) {
    var idFact = String(row[27]).trim(); // AB
    if (!idFact) return;
    var corpus = String(row[2]).trim();  // C
    var floor  = String(row[3]).trim();  // D
    if (!corpus || !floor) return;
    items.push({
      corpus       : corpus,
      floor        : floor,
      razdel       : String(row[4]).trim(),   // E
      podrazdel    : String(row[5]).trim(),   // F
      work         : String(row[6]).trim(),   // G
      actNum       : String(row[7]).trim(),   // H
      date         : formatDateOut(row[8]),   // I
      status       : String(row[9]).trim(),   // J
      contractor   : String(row[10]).trim(),  // K
      remarksTotal : String(row[12]).trim(),  // M
      remarksOpen  : String(row[13]).trim(),  // N
      linkS3       : String(row[14]).trim(),  // O
      linkDrive    : String(row[15]).trim(),  // P
      comment      : String(row[16]).trim(),  // Q
      remarksText  : String(row[18]).trim(),  // S - текст замечаний
      idCl         : String(row[26]).trim(),  // AA
      idFact       : idFact
    });
  });
  var result = {items: items};
  try { cache.put('sb5_checklists', JSON.stringify(result), 7200); } catch(e) {}
  return result;
}

// ─── ПРОТОКОЛ ──────────────────────────────────────────────────────
// Лист «Протокол»: A=дата/время, B=автор, C=текст комментария.
// Создаётся автоматически при первом обращении.
function getProtocolSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Протокол');
  if (!sheet) {
    sheet = ss.insertSheet('Протокол');
    sheet.getRange(1, 1, 1, 3).setValues([['Дата', 'Автор', 'Комментарий']]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 500);
  }
  return sheet;
}

function getProtocol() {
  var sheet = getProtocolSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {items: []};
  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var items = [];
  values.forEach(function(row) {
    var text = String(row[2]).trim();
    if (!text) return;
    items.push({
      date  : formatDateTimeOut(row[0]),
      author: String(row[1]).trim(),
      text  : text
    });
  });
  return {items: items};
}

function addProtocolEntry(data) {
  var text = String(data.text || '').trim();
  if (!text) throw new Error('Пустой комментарий');
  var author = String(data.author || '').trim() || 'Аноним';
  var sheet = getProtocolSheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    sheet.appendRow([nowStr, author, text]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function formatDateTimeOut(v) {
  if (!v) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return pad(v.getDate()) + '.' + pad(v.getMonth() + 1) + '.' + v.getFullYear() + ' ' + pad(v.getHours()) + ':' + pad(v.getMinutes());
  }
  return String(v).trim();
}

function clearCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('sb5_rows_all');
    cache.remove('sb5_work_dict');
    cache.remove('sb5_checklists');
    cache.remove('sb5_staffing');
    ['К1','К2','К3','К4','К5','К6','К7','К8','К9','К10','К11','К12'].forEach(function(c) {
      cache.remove('sb5_rows_' + c);
    });
  } catch(e) {}
}
