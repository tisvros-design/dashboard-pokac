/**
 * Дежурства РК — v40.0 debug
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'jsonp') {
    return jsonp(e);
  }
  try {
    const data = parseDutySheet();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({error:error.toString(),success:false}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function jsonp(e) {
  try {
    const data = parseDutySheet();
    return ContentService.createTextOutput('window._dutyData='+JSON.stringify(data)+';')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (error) {
    return ContentService.createTextOutput('window._dutyData={error:"'+error.toString()+'"};')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function parseDutySheet() {
  const ss = SpreadsheetApp.openById('16J0EEb9yYndMdUZVSX9pP1WfcXM8-hjs5vHxY9Cnfxc');
  const sheet = ss.getSheetByName('Дежурства РК');
  if (!sheet) throw new Error('Лист не найден');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const now = new Date();
  const mNow = new Date(now.toLocaleString('en-US', {timeZone: 'Europe/Moscow'}));

  // Ищем ВСЕ строки "ФИО"
  const fioRows = [];
  for (let i = 0; i < data.length; i++) {
    const v = String(data[i][0] || '').trim().toLowerCase();
    if (v === 'фио') {
      fioRows.push(i);
    }
  }

  Logger.log('Найдено строк ФИО: ' + fioRows.length + ' → ' + JSON.stringify(fioRows));
  if (fioRows.length === 0) throw new Error('Нет заголовков ФИО');

  // Берём последнюю секцию
  const fioRow = fioRows[fioRows.length - 1];
  const peopleStart = fioRow + 1;

  Logger.log('Используем секцию: ФИО=' + fioRow + ', люди=' + peopleStart);

  // Даты в ТОЙ ЖЕ строке что и ФИО (col 1..N)
  const fioRow_data = data[fioRow];
  Logger.log('Строка ФИО (длина=' + fioRow_data.length + '): ' + JSON.stringify(fioRow_data).substring(0, 500));

  const dates = [];
  for (let c = 1; c < fioRow_data.length; c++) {
    const cell = fioRow_data[c];

    if (cell === '' || cell === null || cell === undefined) continue;

    let day = null;

    if (cell instanceof Date) {
      day = cell.getDate();
    } else if (typeof cell === 'number') {
      day = Math.round(cell);
    } else {
      const s = String(cell).trim();
      if (!s) continue;
      // "1.7" → 1, "13.7" → 13, "1" → 1, "01.07" → 1
      const m1 = s.match(/^(\d{1,2})\./);
      const m2 = s.match(/^(\d{1,2})$/);
      if (m1) day = parseInt(m1[1]);
      else if (m2) day = parseInt(m2[1]);
    }

    if (day && day >= 1 && day <= 31) {
      dates.push({col: c, day: day});
    }
  }

  Logger.log('Распознано дат: ' + dates.length + ' → ' + JSON.stringify(dates));

  // Сотрудники — начинаются СРАЗУ после ФИО
  const employees = [];

  for (let r = peopleStart; r < data.length; r++) {
    const name = String(data[r][0] || '').trim();
    if (!name || name.length < 2) continue;
    // Стоп: следующая секция (новое "ФИО" или заголовок месяца)
    if (/^фио$/i.test(name)) break;
    if (/^\d{4}/.test(name)) break;
    if (/^(январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь)/i.test(name)) break;

    const duties = {};
    dates.forEach(function(d) {
      const v = String(data[r][d.col] || '').trim().toLowerCase();
      if (v === '1') duties[d.day] = 'day';
      else if (v === '2') duties[d.day] = 'evening';
      else if (v === 'в' || v === 'б') duties[d.day] = 'weekend';
      else if (v === 'о') duties[d.day] = 'vacation';
      else if (v.indexOf('1-2') >= 0 || v.indexOf('1 2') >= 0) duties[d.day] = 'day';
    });

    employees.push({name: name, duties: duties});
    Logger.log('Сотрудник: ' + name + ' → ' + Object.keys(duties).length + ' дежурств');
  }

  const today = mNow.getDate();
  const todayD = {day: [], evening: [], weekend: [], vacation: []};
  employees.forEach(function(e) {
    const d = e.duties[String(today)];
    if (d === 'day') todayD.day.push(e.name);
    else if (d === 'evening') todayD.evening.push(e.name);
    else if (d === 'weekend') todayD.weekend.push(e.name);
  });

  Logger.log('Сегодня ' + today + ': дневные=' + JSON.stringify(todayD.day) + ', вечерние=' + JSON.stringify(todayD.evening));

  const dn = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
  const mn = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const cm = mNow.getMonth() + 1;

  return {
    success: true,
    month: mn[cm - 1] + ' ' + mNow.getFullYear(),
    date: {
      full: mNow.toISOString().split('T')[0],
      day: mNow.getDate(),
      month: cm,
      year: mNow.getFullYear(),
      dayOfWeek: dn[mNow.getDay()],
      dateRu: mNow.getDate() + ' ' + mn[cm - 1]
    },
    today: todayD,
    employees: employees
  };
}

// Тестовый запуск — вызови из редактора
function testParse() {
  const result = parseDutySheet();
  Logger.log('=== РЕЗУЛЬТАТ ===');
  Logger.log(JSON.stringify(result, null, 2));
}