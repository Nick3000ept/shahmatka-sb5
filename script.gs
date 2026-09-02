// ═══════════════════════════════════════════════════════════════════
// СБ5 Adminка — Google Apps Script
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Факт';
const SPREADSHEET_ID = '1Vm0W09F1QNvBi3RK0pWjBa2KB19pXXex9kkd9vHq8zk';
const ADMIN_PASSWORD = 'adminACCB3';
const SK_PASSWORD    = 'priemkaCB3';
const TASKS_SHEET    = 'Поручения';
const VOL_SHEET      = 'Объемы';
const VOL_HIST_SHEET = 'Объемы_история';
const PRESETS_SHEET  = 'Пресеты';

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

// Экранирование пользовательского ввода перед записью в таблицу:
// значения на =, +, -, @ Google Sheets трактует как формулу — ставим апостроф впереди.
function safeCell_(v) {
  var s = (v === undefined || v === null) ? '' : String(v);
  if (s.length > 5000) s = s.substring(0, 5000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
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

    // ВРЕМЕННО для диагностики раздутого листа (удалить после)
    if (action === 'debugFar') {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
      var probe = [6085, 6092, 6100, 8000, 15000, 25000, lastRow].filter(function(r){ return r >= 2 && r <= lastRow; });
      var out = probe.map(function(r){
        var vals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
        var filled = [];
        vals.forEach(function(v, i){ if (String(v).trim() !== '') filled.push((i + 1) + ':' + String(v).slice(0, 40)); });
        return {row: r, filled: filled};
      });
      return jsonOut({lastRow: lastRow, lastCol: lastCol, probe: out});
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
    if (action === 'getTasks')      return jsonOut(getTasks(p.all === '1'));
    if (action === 'getSysPresets') return jsonOut(getSysPresets());

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
      if (body.fact !== undefined) saveVolFacts_([{rowId: body.rowId, fact: body.fact}], body.author);
      clearCache();
      return jsonOut({ok: true});
    }

    if (body.action === 'saveAll') {
      var rows = body.rows || [];
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = findSheet(ss);
      if (!sheet) return jsonOut({error: 'Лист не найден'});

      var savedVol = 0;
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

        // Факт по объёмам — в отдельный лист «Объемы» (лист «Факт» не затрагивается)
        var volItems = rows.filter(function(r) { return r.fact !== undefined; })
                           .map(function(r) { return {rowId: r.rowId, fact: r.fact}; });
        if (volItems.length) {
          var vAuthor = '';
          for (var vi = 0; vi < rows.length; vi++) { if (rows[vi].author) { vAuthor = rows[vi].author; break; } }
          savedVol = saveVolFacts_(volItems, vAuthor);
        }

        clearCache();
      } finally {
        lock.releaseLock();
      }
      return jsonOut({ok: true, saved: changedIndices.length, savedVol: savedVol, requested: rows.length});
    }

    if (body.action === 'addProtocol') {
      addProtocolEntry(body);
      return jsonOut({ok: true});
    }

    // Общие пресеты фильтров — менять может только администратор
    if (body.action === 'saveSysPreset') {
      if (body.pwd !== ADMIN_PASSWORD) return jsonOut({error: 'Нет прав'});
      return jsonOut(saveSysPreset(body));
    }

    if (body.action === 'deleteSysPreset') {
      if (body.pwd !== ADMIN_PASSWORD) return jsonOut({error: 'Нет прав'});
      return jsonOut(deleteSysPreset(body));
    }

    // Поручения — создание и правка только под паролем администратора
    if (body.action === 'addTasks') {
      if (body.pwd !== ADMIN_PASSWORD) return jsonOut({error: 'Нет прав'});
      return jsonOut(addTasks(body));
    }

    if (body.action === 'updateTask') {
      if (body.pwd !== ADMIN_PASSWORD) return jsonOut({error: 'Нет прав'});
      return jsonOut(updateTask(body));
    }

    return jsonOut({error: 'Unknown action: ' + body.action});
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

// ═══════════════════════════════════════════════════════════════════
// СИСТЕМНЫЕ ПРЕСЕТЫ (лист «Пресеты», создаётся автоматически)
// Общие пресеты фильтров: видны всем пользователям, менять и удалять может
// только администратор (проверка пароля в doPost). Столбцы:
// A=Название | B=Настройки (JSON: corpus/org/place/lvl1/lvl2/works/kp) | C=Обновлён
// «Удаление» — очистка ячеек строки, не deleteRow (строки листов скриптом не удаляем);
// пустые строки при чтении пропускаются.
// ═══════════════════════════════════════════════════════════════════
function ensurePresetsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PRESETS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRESETS_SHEET);
    sheet.appendRow(['Название', 'Настройки (JSON)', 'Обновлён']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 420);
  }
  return sheet;
}

function findPresetRow_(sheet, name) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === name) return i + 2;
  }
  return 0;
}

function getSysPresets() {
  var sheet = ensurePresetsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {presets: {}};
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var presets = {};
  values.forEach(function(row) {
    var name = String(row[0]).trim();
    if (!name) return;
    try { presets[name] = JSON.parse(String(row[1])); } catch (e) {}
  });
  return {presets: presets};
}

function saveSysPreset(body) {
  var name = String(body.name || '').trim().slice(0, 40);
  if (!name) return {error: 'Укажите название'};
  if (!body.cfg || typeof body.cfg !== 'object') return {error: 'Нет настроек пресета'};
  var json = JSON.stringify(body.cfg);
  if (json.length > 5000) return {error: 'Слишком большой пресет'};
  var sheet = ensurePresetsSheet_();
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  var rowIdx = findPresetRow_(sheet, name);
  if (rowIdx) sheet.getRange(rowIdx, 1, 1, 3).setValues([[safeCell_(name), json, nowStr]]);
  else sheet.appendRow([safeCell_(name), json, nowStr]);
  SpreadsheetApp.flush();
  return {ok: true};
}

function deleteSysPreset(body) {
  var name = String(body.name || '').trim();
  if (!name) return {error: 'Укажите название'};
  var sheet = ensurePresetsSheet_();
  var rowIdx = findPresetRow_(sheet, name);
  if (!rowIdx) return {error: 'Пресет не найден'};
  sheet.getRange(rowIdx, 1, 1, 3).setValues([['', '', '']]);
  SpreadsheetApp.flush();
  return {ok: true};
}

// ─── ЖУРНАЛ ОБЪЁМОВ ────────────────────────────────────────────────
// Лист «Объемы_история»: A=Дата | B=rowId | C=Корпус | D=Этаж | E=Работа |
// F=Было | G=Стало | H=Прирост | I=Автор
// Только добавление строк, ничего не перезаписывается. Отсюда считается «за 7/30 дней».
const VHIST_READ_LIMIT = 5000; // читаем только хвост журнала

function ensureVolHistSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(VOL_HIST_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VOL_HIST_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['Дата', 'rowId', 'Корпус', 'Этаж', 'Работа', 'Было', 'Стало', 'Прирост', 'Автор']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 260);
  }
  return sh;
}

// Приросты факта по строкам за последние 60 дней: {rowId: [{ms, delta}]}
function getVolHist() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sb5_vol_hist');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(VOL_HIST_SHEET);
  var out = {};
  if (sh) {
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var startRow = Math.max(2, lastRow - VHIST_READ_LIMIT + 1);
      var vals = sh.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
      var cutoff = Date.now() - 60 * 24 * 3600 * 1000;
      vals.forEach(function(r) {
        var id = String(r[1]).trim();
        if (!id) return;
        var ms = 0;
        var d = r[0];
        if (d instanceof Date) ms = d.getTime();
        else {
          var m = String(d).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
          if (m) ms = new Date(+m[3], +m[2] - 1, +m[1]).getTime();
        }
        if (!ms || ms < cutoff) return;
        var delta = parseFloat(String(r[7]).replace(',', '.'));
        if (isNaN(delta) || delta === 0) return;
        if (!out[id]) out[id] = [];
        out[id].push({ms: ms, d: delta});
      });
    }
  }
  try { cache.put('sb5_vol_hist', JSON.stringify(out), 7200); } catch(e) {}
  return out;
}

// Сумма приростов по строке за последние N дней
function volSumDays_(list, days) {
  if (!list || !list.length) return 0;
  var from = Date.now() - days * 24 * 3600 * 1000;
  var sum = 0;
  list.forEach(function(x) { if (x.ms >= from) sum += x.d; });
  return Math.round(sum * 100) / 100;
}

// ─── ОБЪЁМЫ ────────────────────────────────────────────────────────
// Лист «Объемы»: A=rowId | B=Корпус | C=Этаж | D=Работа | E=Ед.изм. | F=Объём итого |
// G=Факт | H=Изменено | I=Автор
// Сайт пишет ТОЛЬКО G-I. B-F заполняются вручную; A (rowId) считает ARRAYFORMULA
// в ячейке A2 листа — НИКОГДА не писать в столбец A скриптом, это сломает формулу.
// Лист «Факт» при вводе факта не трогаем нигде, кроме столбца K (процент) — как и раньше.
const VOL = {ROW_ID: 1, CORPUS: 2, FLOOR: 3, WORK: 4, UNIT: 5, TOTAL: 6, FACT: 7, DATE_CHG: 8, AUTHOR: 9};

function ensureVolSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(VOL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VOL_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['rowId', 'Корпус', 'Этаж', 'Работа', 'Ед.изм.', 'Объём итого', 'Факт', 'Изменено', 'Автор']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 90); sh.setColumnWidth(4, 260);
  }
  return sh;
}

// Число из пользовательского ввода: "25,5" → 25.5; мусор → null
function volNum_(v) {
  if (v === undefined || v === null || String(v).trim() === '') return null;
  var n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? null : n;
}

function getVolMap() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sb5_vol_map');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(VOL_SHEET);
  var map = {};
  if (sh) {
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var vals = sh.getRange(2, 1, lastRow - 1, 9).getValues();
      vals.forEach(function(r) {
        var id = String(r[VOL.ROW_ID - 1]).trim();
        if (!id) return;
        map[id] = {
          unit : String(r[VOL.UNIT  - 1]).trim(),
          total: String(r[VOL.TOTAL - 1]).trim(),
          fact : String(r[VOL.FACT  - 1]).trim()
        };
      });
    }
  }
  try { cache.put('sb5_vol_map', JSON.stringify(map), 7200); } catch(e) {}
  return map;
}

// Запись факта: [{rowId, fact}], возвращает число записанных строк
function saveVolFacts_(items, author) {
  if (!items || !items.length) return 0;
  var sh = ensureVolSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  // Читаем весь лист сразу: нужны и rowId, и старое значение факта, и описание работы
  var all = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  var map = {};
  all.forEach(function(r, i) { var id = String(r[VOL.ROW_ID - 1]).trim(); if (id) map[id] = i; });

  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var n = 0, hist = [];
  items.forEach(function(it) {
    var i = map[String(it.rowId)];
    if (i === undefined) return; // объёма для этой работы нет — факт писать некуда
    var was = volNum_(all[i][VOL.FACT - 1]);
    var now = volNum_(it.fact);
    sh.getRange(i + 2, VOL.FACT, 1, 3).setValues([[now === null ? '' : now, nowStr, safeCell_(author || '')]]);
    n++;
    // В журнал — только если значение реально изменилось
    var delta = (now === null ? 0 : now) - (was === null ? 0 : was);
    if (delta !== 0) {
      hist.push([nowStr, String(it.rowId), all[i][VOL.CORPUS - 1], all[i][VOL.FLOOR - 1], all[i][VOL.WORK - 1],
                 was === null ? '' : was, now === null ? '' : now, Math.round(delta * 100) / 100, safeCell_(author || '')]);
    }
  });

  if (hist.length) {
    var hs = ensureVolHistSheet_();
    hs.getRange(hs.getLastRow() + 1, 1, hist.length, 9).setValues(hist);
  }
  if (n) SpreadsheetApp.flush();
  return n;
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
  var volMap = getVolMap();
  var volHist = getVolHist();
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
    var vol = volMap[rowId] || null;
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
      volume     : vol ? vol.total : String(row[16]).trim(),
      unit       : vol ? vol.unit  : String(row[17]).trim(),
      fact       : vol ? vol.fact  : '',
      volWeek    : vol ? volSumDays_(volHist[rowId], 7)  : 0,
      volMonth   : vol ? volSumDays_(volHist[rowId], 30) : 0,
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

// ─── ПОРУЧЕНИЯ ─────────────────────────────────────────────────────
// Лист «Поручения»: A=id | B=Подрядчик | C=Текст | D=Срок | E=Статус |
//                   F=Создано | G=Автор | H=Комментарий | I=Приоритет
// Создаётся автоматически при первом обращении.
function ensureTasksSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASKS_SHEET);
    sheet.appendRow(['id', 'Подрядчик', 'Текст', 'Срок', 'Статус', 'Создано', 'Автор', 'Комментарий', 'Приоритет']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 420);
    sheet.setColumnWidth(8, 420);
  }
  return sheet;
}

function getTasks(includeAll) {
  var sheet = ensureTasksSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {tasks: []};
  var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var tasks = [];
  values.forEach(function(row) {
    var id = String(row[0]).trim();
    if (!id) return;
    var status = String(row[4]).trim() || 'открыто';
    if (!includeAll && status !== 'открыто') return;
    tasks.push({
      id      : id,
      org     : String(row[1]).trim(),
      text    : String(row[2]).trim(),
      due     : formatDateOut(row[3]),
      status  : status,
      created : formatDateOut(row[5]),
      author  : String(row[6]).trim(),
      comment : String(row[7]).trim(),
      priority: String(row[8]).trim()
    });
  });
  return {tasks: tasks};
}

// Пакетное добавление: все поручения одним запросом, одна запись в лист.
// Либо записывается весь пакет, либо (при ошибке валидации) ничего.
function addTasks(body) {
  var list = body.tasks;
  if (!list || !list.length) return {error: 'Нет поручений'};
  var sheet = ensureTasksSheet_();
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var rowsOut = [];
  for (var i = 0; i < list.length; i++) {
    var t = list[i] || {};
    var text = String(t.text || '').trim();
    if (!text) return {error: 'У каждого поручения нужен текст (строка ' + (i + 1) + ')'};
    var id = 't' + Date.now() + i + Math.floor(Math.random() * 1000);
    rowsOut.push([id, safeCell_(t.org), safeCell_(text), safeCell_(t.due), 'открыто',
                  nowStr, safeCell_(body.author), safeCell_(t.comment), '']);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rowsOut.length, 9).setValues(rowsOut);
  SpreadsheetApp.flush();
  return {ok: true, added: rowsOut.length};
}

function updateTask(body) {
  var id = String(body.id || '').trim();
  if (!id) return {error: 'Нет id поручения'};
  var sheet = ensureTasksSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {error: 'Поручение не найдено'};
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      var r = i + 2;
      if (body.status   !== undefined) sheet.getRange(r, 5).setValue(safeCell_(body.status));
      if (body.org      !== undefined) sheet.getRange(r, 2).setValue(safeCell_(body.org));
      if (body.text     !== undefined) sheet.getRange(r, 3).setValue(safeCell_(body.text));
      if (body.due      !== undefined) sheet.getRange(r, 4).setValue(safeCell_(body.due));
      if (body.comment  !== undefined) sheet.getRange(r, 8).setValue(safeCell_(body.comment));
      if (body.priority !== undefined) sheet.getRange(r, 9).setValue(safeCell_(body.priority));
      SpreadsheetApp.flush();
      return {ok: true};
    }
  }
  return {error: 'Поручение не найдено'};
}

// ─── ПРОТОКОЛ ──────────────────────────────────────────────────────
// Лист «Протокол»: A=дата/время, B=автор, C=текст комментария, D=привязанная работа.
// Создаётся автоматически при первом обращении.
function getProtocolSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Протокол');
  if (!sheet) {
    sheet = ss.insertSheet('Протокол');
    sheet.getRange(1, 1, 1, 4).setValues([['Дата', 'Автор', 'Комментарий', 'Работа']]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 500);
    sheet.setColumnWidth(4, 220);
  } else if (!String(sheet.getRange(1, 4).getValue()).trim()) {
    // Лист создан старой версией без столбца «Работа» — дописываем заголовок
    sheet.getRange(1, 4).setValue('Работа');
  }
  return sheet;
}

function getProtocol() {
  var sheet = getProtocolSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {items: []};
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var items = [];
  values.forEach(function(row) {
    var text = String(row[2]).trim();
    if (!text) return;
    items.push({
      date  : formatDateTimeOut(row[0]),
      author: String(row[1]).trim(),
      text  : text,
      work  : String(row[3]).trim()
    });
  });
  return {items: items};
}

function addProtocolEntry(data) {
  var text = String(data.text || '').trim();
  if (!text) throw new Error('Пустой комментарий');
  var author = String(data.author || '').trim() || 'Аноним';
  var work   = String(data.work || '').trim();
  var sheet = getProtocolSheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    sheet.appendRow([nowStr, author, text, work]);
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
    cache.remove('sb5_vol_map');
    cache.remove('sb5_vol_hist');
    cache.remove('sb5_checklists');
    cache.remove('sb5_staffing');
    ['К1','К2','К3','К4','К5','К6','К7','К8','К9','К10','К11','К12'].forEach(function(c) {
      cache.remove('sb5_rows_' + c);
    });
  } catch(e) {}
}
