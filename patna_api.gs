// ═══════════════════════════════════════════════════════════
//  ПЪТНА КНИЖКА — Google Apps Script API
//
//  Инсталация:
//  1. Създай нов Google Sheets файл → наименувай "Пътна Книжка"
//  2. Разширения → Apps Script → постави кода
//  3. Script Properties → добави: ANTHROPIC_KEY = sk-ant-...
//     (Settings → Script properties → Add property)
//  4. Деплой → New deployment → Web App
//     Execute as: Me | Who: Anyone
//  5. Копирай URL → постави в приложението
// ═══════════════════════════════════════════════════════════

var SS = SpreadsheetApp.getActiveSpreadsheet();
var MONTHS_BG = ['','Януари','Февруари','Март','Април','Май','Юни','Юли','Август','Септември','Октомври','Ноември','Декември'];

function doGet(e) {
  try {
    var p = e.parameter;
    var action = p.action;
    if (!action) return err('Missing action');
    if (action === 'get_entries')       return ok(getEntries(p));
    if (action === 'add_entry')         return ok(addEntry(p));
    if (action === 'update_entry')      return ok(updateEntry(p));
    if (action === 'delete_entry')      return ok(deleteEntry(p));
    if (action === 'get_settings')      return ok(getSettings());
    if (action === 'save_settings')     return ok(saveSettings(p));
    if (action === 'ocr_image')         return ok(ocrImage(p.image, p.mime || 'image/jpeg'));
    if (action === 'get_last_km')       return ok(getLastKm());
    return err('Unknown action: ' + action);
  } catch(e) {
    return err(e.message);
  }
}

// ─── SETTINGS ────────────────────────────────────────────────
function getSettings() {
  ensureSheet('settings', ['car_reg','car_make','default_route']);
  var sheet = SS.getSheetByName('settings');
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2 || !data[1][0]) {
    return { car_reg: 'СВ0773МХ', car_make: 'Дачия', default_route: 'Димитровград - Кърджали - Димитровград' };
  }
  return { car_reg: data[1][0], car_make: data[1][1], default_route: data[1][2] || '' };
}

function saveSettings(p) {
  ensureSheet('settings', ['car_reg','car_make','default_route']);
  var sheet = SS.getSheetByName('settings');
  if (sheet.getLastRow() < 2) sheet.appendRow(['','','']);
  sheet.getRange(2, 1, 1, 3).setValues([[p.car_reg || '', p.car_make || '', p.default_route || '']]);
  return { saved: true };
}

// ─── ENTRIES ─────────────────────────────────────────────────
function getEntries(p) {
  var ym = p.month_year || getCurrentMonthYear();
  ensureDataSheet(ym);
  var sheet = SS.getSheetByName('data_' + ym);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { entries: [], total_km: 0, month_year: ym };
  var heads   = data[0];
  var entries = [];
  var total   = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = {};
    for (var j = 0; j < heads.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) val = Utilities.formatDate(val, 'Europe/Sofia', 'yyyy-MM-dd');
      row[heads[j]] = val;
    }
    var km = Number(row.km_driven) || 0;
    total += km;
    entries.push(row);
  }
  entries.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  return { entries: entries, total_km: total, month_year: ym };
}

function addEntry(p) {
  var date     = p.date || todayStr();
  var ym       = date.substring(0, 7);
  var startKm  = parseInt(p.start_km) || 0;
  var endKm    = parseInt(p.end_km) || 0;
  var kmDriven = endKm > startKm ? endKm - startKm : 0;
  var route    = p.route || '';

  ensureDataSheet(ym);
  var sheet = SS.getSheetByName('data_' + ym);
  var data  = sheet.getDataRange().getValues();
  var heads = data[0];

  // Check if entry for this date already exists → update instead
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], 'Europe/Sofia', 'yyyy-MM-dd')
      : String(data[i][0]);
    if (rowDate === date) {
      var updateP = { id: String(i + 1), date: date, start_km: p.start_km, end_km: p.end_km, route: route, month_year: ym };
      return updateEntry(updateP);
    }
  }

  // New row
  var newId = (data.length > 1 ? data.length : 1);
  var row = heads.map(function(h) {
    if (h === 'id')        return newId;
    if (h === 'date')      return date;
    if (h === 'start_km')  return startKm;
    if (h === 'end_km')    return endKm;
    if (h === 'km_driven') return kmDriven;
    if (h === 'route')     return route;
    if (h === 'created_at') return new Date().toISOString();
    return '';
  });
  sheet.appendRow(row);

  // Also update previous open entry end_km if this is a new morning reading
  if (p.update_prev === 'true' && startKm > 0) {
    updatePrevEntryEndKm(ym, date, startKm);
  }

  return { saved: true, id: newId, date: date, start_km: startKm, end_km: endKm, km_driven: kmDriven };
}

function updatePrevEntryEndKm(ym, todayDate, todayStartKm) {
  // Find the most recent entry before today with no end_km set
  var sheet = SS.getSheetByName('data_' + ym);
  if (!sheet) {
    // Check previous month
    var d = new Date(todayDate);
    d.setDate(1); d.setMonth(d.getMonth() - 1);
    var prevYm = Utilities.formatDate(d, 'Europe/Sofia', 'yyyy-MM');
    sheet = SS.getSheetByName('data_' + prevYm);
  }
  if (!sheet) return;
  var data  = sheet.getDataRange().getValues();
  var heads = data[0];
  var endIdx = heads.indexOf('end_km');
  var kmIdx  = heads.indexOf('km_driven');
  var stIdx  = heads.indexOf('start_km');
  // Find last row with empty end_km
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][0]) continue;
    if (!data[i][endIdx] || data[i][endIdx] === 0) {
      sheet.getRange(i + 1, endIdx + 1).setValue(todayStartKm);
      var startKm = Number(data[i][stIdx]) || 0;
      if (todayStartKm > startKm) {
        sheet.getRange(i + 1, kmIdx + 1).setValue(todayStartKm - startKm);
      }
      return;
    }
  }
}

function updateEntry(p) {
  var ym    = p.month_year || p.date.substring(0, 7);
  var sheet = SS.getSheetByName('data_' + ym);
  if (!sheet) return { error: 'Sheet not found' };
  var data  = sheet.getDataRange().getValues();
  var heads = data[0];
  var rowNum = parseInt(p.id);
  if (isNaN(rowNum) || rowNum > data.length) return { error: 'Row not found' };

  var updates = {};
  if (p.date)     updates.date     = p.date;
  if (p.start_km) updates.start_km = parseInt(p.start_km);
  if (p.end_km)   updates.end_km   = parseInt(p.end_km);
  if (p.route !== undefined) updates.route = p.route;

  var startKm = parseInt(p.start_km) || Number(data[rowNum - 1][heads.indexOf('start_km')]) || 0;
  var endKm   = parseInt(p.end_km)   || Number(data[rowNum - 1][heads.indexOf('end_km')])   || 0;
  if (endKm > startKm) updates.km_driven = endKm - startKm;

  heads.forEach(function(h, j) {
    if (updates[h] !== undefined) {
      sheet.getRange(rowNum, j + 1).setValue(updates[h]);
    }
  });
  return { saved: true };
}

function deleteEntry(p) {
  var ym    = p.month_year;
  var sheet = SS.getSheetByName('data_' + ym);
  if (!sheet) return { error: 'Sheet not found' };
  var rowNum = parseInt(p.id);
  sheet.deleteRow(rowNum);
  return { deleted: true };
}

function getLastKm() {
  // Find the most recent km reading across all data sheets
  var ym    = getCurrentMonthYear();
  var sheet = SS.getSheetByName('data_' + ym);
  var last  = 0;
  if (sheet) {
    var data  = sheet.getDataRange().getValues();
    var heads = data[0];
    var stIdx = heads.indexOf('start_km');
    var enIdx = heads.indexOf('end_km');
    for (var i = data.length - 1; i >= 1; i--) {
      if (!data[i][0]) continue;
      var en = Number(data[i][enIdx]);
      var st = Number(data[i][stIdx]);
      last = en || st || last;
      if (last) break;
    }
  }
  return { last_km: last };
}

// ─── OCR via Claude API ───────────────────────────────────────
function ocrImage(base64Image, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!apiKey) return { error: 'ANTHROPIC_KEY не е зададен в Script Properties' };
  if (!base64Image) return { error: 'Няма изображение' };

  var payload = {
    model: 'claude-opus-4-5',
    max_tokens: 50,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Image } },
        { type: 'text', text: 'Какъв е показанията на километража на това табло? Отговори САМО с числото, без никакъв текст, без точки, без интервали. Само цифрите.' }
      ]
    }]
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  if (result.error) return { error: result.error.message };

  var text = result.content[0].text.trim().replace(/[^\d]/g, '');
  var km   = parseInt(text);
  if (isNaN(km) || km < 1000 || km > 9999999) return { error: 'Не успях да прочета километража. Опитай отново.', raw: text };
  return { km: km };
}

// ─── Helpers ─────────────────────────────────────────────────
function ensureSheet(name, headers) {
  var sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function ensureDataSheet(ym) {
  var name = 'data_' + ym;
  var sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(['id','date','start_km','end_km','km_driven','route','created_at']);
  }
  return sheet;
}

function getCurrentMonthYear() {
  return Utilities.formatDate(new Date(), 'Europe/Sofia', 'yyyy-MM');
}

function todayStr() {
  return Utilities.formatDate(new Date(), 'Europe/Sofia', 'yyyy-MM-dd');
}

function ok(data)  { return res({ ok: true, data: data }); }
function err(msg)  { return res({ ok: false, error: msg }); }
function res(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── ТЕСТ ───────────────────────────────────────────────────
function testSetup() {
  Logger.log('Settings: ' + JSON.stringify(getSettings()));
  Logger.log('Last KM: ' + JSON.stringify(getLastKm()));
  Logger.log('API Key set: ' + !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY'));
}
