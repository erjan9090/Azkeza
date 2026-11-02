/**
 * Приём записей с калькулятора аскез (ходьба, статика, вода) в Google-таблицу.
 * Вставить в Расширения → Apps Script таблицы и развернуть как веб-приложение
 * (Выполнять от имени: Я, Доступ: Все).
 * После замены кода: Развернуть → Управление развёртываниями → ✏ → Версия: новая.
 */
const SHEET_NAME = 'Журнал';
const HEADERS = ['Дата', 'Имя', 'Упражнение', 'Время', 'Норма, мин', 'Метров', 'Температура, °C',
                 'Цель, слоёв', 'Засчитано слоёв', 'Слоёв точно', 'Цель выполнена', 'ID'];
const TYPES = { walk: 'Ходьба', static: 'Статика', water: 'Вода' };

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
