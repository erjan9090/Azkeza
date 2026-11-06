/**
 * Приём записей с калькулятора аскез (ходьба, статика, вода) в Google-таблицу.
 * Развернуть как веб-приложение (Выполнять от имени: Я, Доступ: Все).
 * После замены кода: Развернуть → Управление развёртываниями → ✏ → Версия: «Новая версия».
 * Не создавайте «Новое развёртывание» — у него будет другая ссылка.
 */
// ID таблицы (часть адреса таблицы между /d/ и /edit)
const SPREADSHEET_ID = '148TOaD7odSF3Qh27jNZ093CRCjp6DPbCcofpQ1mBxRI';
const SHEET_NAME = 'Журнал';
const PLANS_SHEET = 'Графики';
const HEADERS = ['Дата', 'Имя', 'Упражнение', 'Время', 'Норма, мин', 'Метров', 'Температура, °C',
                 'Цель, слоёв', 'Засчитано слоёв', 'Слоёв точно', 'Цель выполнена', 'ID', 'График'];
const ID_COL = 12;   // столбец ID (L); «График» добавлен после него, чтобы не сдвигать старые строки
const TYPES = { walk: 'Ходьба', static: 'Статика', water: 'Вода' };
const PLAN_HEADERS = ['График', 'Вода, раз', 'Вода, слоёв', 'Статика, раз', 'Статика, слоёв',
                      'Ходьба, раз', 'Ходьба, слоёв', 'Дни (подсказка)'];
const PLAN_DEFAULTS = [
  ['1', 2, 5, 3, 5, 2, 1, 'Вода пн,пт · Статика пн,ср,сб · Ходьба вт,чт'],
  ['2', 2, 5, 3, 5, 2, 1, 'Вода пн,пт · Статика вт,чт,сб · Ходьба ср,вс'],
  ['3', 1, 5, 3, 5, 2, 1, 'Вода пн · Статика пн,ср,сб · Ходьба вт,чт'],
  ['4', 1, 5, 3, 5, 2, 1, 'Вода пн · Статика вт,чт,сб · Ходьба ср,вс']
];

// GET ?name=Имя — последние записи этого человека и список графиков.
// GET без параметров — проверка: должно быть {"ok":true,...}
function doGet(e) {
  const cb = e && e.parameter && e.parameter.callback;
  const out = doGet_(e);
  // JSONP: страница может загрузить ответ через <script>, если обычный запрос блокируется
  if (cb && /^[A-Za-z_$][\w$]{0,60}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function doGet_(e) {
  const json_ = function (o) { return JSON.stringify(o); };
  try {
    const sheet = getSheet_();
    const plans = getPlans_();
    if (e && e.parameter && e.parameter.group) {
      return json_(group_(sheet, plans, Number(e.parameter.from) || 0, Number(e.parameter.to) || Date.now()));
    }
    const name = e && e.parameter && e.parameter.name ? String(e.parameter.name).trim().toLowerCase() : '';
    if (!name) {
      return json_({ ok: true, table: sheet.getParent().getName(), sheet: sheet.getName(),
                     rows: sheet.getLastRow() - 1, plans: plans });
    }
    const last = sheet.getLastRow();
    const rows = last > 1 ? sheet.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
    const typeKey = {};
    Object.keys(TYPES).forEach(function (k) { typeKey[TYPES[k]] = k; });
    const records = [];
    for (let i = rows.length - 1; i >= 0 && records.length < 60; i--) {
      const r = rows[i];
      if (String(r[1]).trim().toLowerCase() !== name) continue;
      const time = r[3] instanceof Date
        ? Utilities.formatDate(r[3], 'GMT', 'H:mm')   // на случай, если ячейка стала временем
        : String(r[3]);
      records.push({
        ts: r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || 0,
        name: String(r[1]),
        type: typeKey[r[2]] || 'walk',
        time: time,
        dist: r[5],
        temp: r[6],
        goal: r[7],
        layers: r[8],
        id: String(r[11]),
        plan: String(r[12] || '')
      });
    }
    return json_({ ok: true, records: records, plans: plans });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    // Запрос журнала: {action:'list', name:'Имя'}
    if (d.action === 'group') {
      return ContentService.createTextOutput(doGet_({ parameter: { group: '1', from: d.from, to: d.to } }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (d.action === 'list') {
      return ContentService.createTextOutput(doGet_({ parameter: { name: String(d.name || ' ') } }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (!d.id || !d.type) return json_({ ok: false, error: 'пустая запись' });
    const sheet = getSheet_();

    // Защита от повторной отправки одной и той же записи
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, ID_COL, last - 1, 1).getValues().flat();
      if (ids.indexOf(String(d.id)) !== -1) return json_({ ok: true, duplicate: true });
    }

    sheet.appendRow([
      new Date(Number(d.ts) || Date.now()),
      safe_(d.name),
      TYPES[d.type] || TYPES.walk,
      safe_(d.time),
      Number(d.minutes) || '',
      d.dist === '' || d.dist == null ? '' : Number(d.dist) || 0,
      d.temp === '' || d.temp == null ? '' : Number(d.temp),
      Number(d.goal) || '',
      Number(d.layers) || 0,
      Number(d.exact) || 0,
      d.done ? 'да' : 'нет',
      String(d.id),
      safe_(d.plan || '')
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Активность всех участников за период [from, to) и последний выбранный график каждого
function group_(sheet, plans, from, to) {
  const last = sheet.getLastRow();
  const rows = last > 1 ? sheet.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  const typeKey = {};
  Object.keys(TYPES).forEach(function (k) { typeKey[TYPES[k]] = k; });
  const records = [], people = {};
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    const name = String(r[1]).trim();
    if (!name || !r[3]) continue;
    const key = name.toLowerCase();
    if (!(key in people)) people[key] = { name: name, plan: '' };
    if (!people[key].plan && r[12]) people[key].plan = String(r[12]);
    const ts = r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || 0;
    if (ts >= from && ts < to) {
      records.push({ ts: ts, name: name, type: typeKey[r[2]] || 'walk', goal: r[7], layers: r[8] });
    }
  }
  return { ok: true, records: records, people: people, plans: plans };
}

function ss_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Таблица не найдена: впишите SPREADSHEET_ID');
  return ss;
}

function getSheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd.MM.yyyy HH:mm');
    sheet.getRange('D:D').setNumberFormat('@');
  } else if (sheet.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
    sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1]);   // столбец «График»
  }
  return sheet;
}

// Лист «Графики»: одна строка — один вариант недельного графика
function getPlans_() {
  const ss = ss_();
  let sh = ss.getSheetByName(PLANS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PLANS_SHEET);
    sh.getRange(1, 1, 1, PLAN_HEADERS.length).setValues([PLAN_HEADERS]);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(2, 1, PLAN_DEFAULTS.length, PLAN_HEADERS.length).setValues(PLAN_DEFAULTS);
    sh.setFrozenRows(1);
  }
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, PLAN_HEADERS.length).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        water: { times: Number(r[1]) || 0, layers: Number(r[2]) || 0 },
        static: { times: Number(r[3]) || 0, layers: Number(r[4]) || 0 },
        walk: { times: Number(r[5]) || 0, layers: Number(r[6]) || 0 },
        days: String(r[7] || '')
      };
    });
}

// Текст из формы не должен превращаться в формулу
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 60);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
