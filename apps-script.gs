/**
 * Приём записей с калькулятора аскез (ходьба, статика, вода) в Google-таблицу.
 * Вставить в Расширения → Apps Script таблицы и развернуть как веб-приложение
 * (Выполнять от имени: Я, Доступ: Все).
 * После замены кода: Развернуть → Управление развёртываниями → ✏ → Версия: новая.
 */
// Если скрипт создан отдельно на script.google.com, впишите сюда ID таблицы
// (часть адреса таблицы между /d/ и /edit). Если скрипт открыт из самой таблицы — оставьте пусто.
const SPREADSHEET_ID = '';
const SHEET_NAME = 'Журнал';
const HEADERS = ['Дата', 'Имя', 'Упражнение', 'Время', 'Норма, мин', 'Метров', 'Температура, °C',
                 'Цель, слоёв', 'Засчитано слоёв', 'Слоёв точно', 'Цель выполнена', 'ID'];
const TYPES = { walk: 'Ходьба', static: 'Статика', water: 'Вода' };

// GET ?name=Имя — последние записи этого человека (для журнала на странице).
// GET без параметров — проверка: должно быть {"ok":true,...}
function doGet(e) {
  try {
    const sheet = getSheet_();
    const name = e && e.parameter && e.parameter.name ? String(e.parameter.name).trim().toLowerCase() : '';
    if (!name) {
      return json_({ ok: true, table: sheet.getParent().getName(), sheet: sheet.getName(), rows: sheet.getLastRow() - 1 });
    }
    const last = sheet.getLastRow();
    const rows = last > 1 ? sheet.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
    const typeKey = {};
    Object.keys(TYPES).forEach(function (k) { typeKey[TYPES[k]] = k; });
    const records = [];
    for (let i = rows.length - 1; i >= 0 && records.length < 20; i--) {
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
        id: String(r[11])
      });
    }
    return json_({ ok: true, records: records });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    // Защита от повторной отправки одной и той же записи
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, HEADERS.length, last - 1, 1).getValues().flat();
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
      String(d.id)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Таблица не найдена: скрипт создан отдельно — впишите SPREADSHEET_ID');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd.MM.yyyy HH:mm');
    sheet.getRange('D:D').setNumberFormat('@');
  }
  return sheet;
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
