/**
 * Приём показаний ХВС/ГВС с сайта на Tilda -> Google Таблица.
 * Установка: см. form/README.md
 *
 * Листы (создаются автоматически):
 *   «Настройки» — окно приёма и параметры (меняйте значения в столбце B)
 *   «Показания» — по одной строке на квартиру за месяц (повторная отправка обновляет строку).
 *                 Столбцы: Период | Квартира | ХВС | ГВС | Принято | Примечание — копируется в 1С
 *   «Журнал»    — все отправки подряд, без перезаписи
 */
var SETTINGS = [
  ['День начала приёма', 20, 'число месяца, включительно'],
  ['День окончания приёма', 25, 'число месяца, включительно'],
  ['Приём включён', 'ДА', 'ДА / НЕТ — аварийный выключатель, перекрывает даты'],
  ['Всего квартир', 100, 'допустимые номера: 1..N; 0 — не проверять'],
  ['Текст при закрытом окне', 'Показания принимаются с {start} по {end} число каждого месяца.', '{start} и {end} подставляются сами']
];
var HEAD = ['Период', 'Квартира', 'ХВС', 'ГВС', 'Принято', 'Примечание'];

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) { sh.appendRow(header); sh.setFrozenRows(1); sh.getRange(1, 1, 1, header.length).setFontWeight('bold'); }
  }
  return sh;
}
function settings_() {
  var sh = sheet_('Настройки');
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, SETTINGS.length, 3).setValues(SETTINGS); sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 260); }
  var v = sh.getRange(1, 2, SETTINGS.length, 1).getValues();
  return { start: +v[0][0], end: +v[1][0], on: String(v[2][0]).toUpperCase() !== 'НЕТ', flats: +v[3][0], closedText: String(v[4][0]) };
}
function status_() {
  var s = settings_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var day = +Utilities.formatDate(new Date(), tz, 'd');
  var open = s.on && day >= s.start && day <= s.end;
  return { open: open, start: s.start, end: s.end, flats: s.flats,
           message: open ? '' : s.closedText.replace('{start}', s.start).replace('{end}', s.end) };
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Форма при загрузке спрашивает, открыто ли окно приёма.
function doGet() { return json_(status_()); }

function num_(x) { var n = parseFloat(String(x).replace(/\s/g, '').replace(',', '.')); return isFinite(n) ? n : NaN; }

function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents), st = status_();
    if (d.website) return json_({ ok: true });                       // ловушка для ботов
    if (!st.open) return json_({ ok: false, error: st.message || 'Приём показаний закрыт.' });
    var flat = Math.round(num_(d.flat)), cold = num_(d.cold), hot = num_(d.hot);
    if (!(flat >= 1) || (st.flats > 0 && flat > st.flats)) return json_({ ok: false, error: 'Неверный номер квартиры.' });
    if (!(cold >= 0) || !(hot >= 0)) return json_({ ok: false, error: 'Показания должны быть неотрицательными числами.' });

    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(), now = new Date();
    var period = Utilities.formatDate(now, tz, 'yyyy-MM'), stamp = Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm');
    sheet_('Журнал', ['Принято', 'Квартира', 'ХВС', 'ГВС']).appendRow([stamp, flat, cold, hot]);

    var sh = sheet_('Показания', HEAD), data = sh.getDataRange().getValues(), row = -1, prev = null;
    for (var i = 1; i < data.length; i++) {
      if (+data[i][1] !== flat) continue;
      if (String(data[i][0]) === period) row = i + 1;
      else if (!prev || String(data[i][0]) > String(prev[0])) prev = data[i];
    }
    var notes = [];
    if (prev && cold < +prev[2]) notes.push('ХВС меньше прошлого (' + prev[2] + ')');
    if (prev && hot < +prev[3]) notes.push('ГВС меньше прошлого (' + prev[3] + ')');
    var line = [period, flat, cold, hot, stamp, notes.join('; ')];
    if (row > 0) sh.getRange(row, 1, 1, 6).setValues([line]); else sh.appendRow(line);
    sh.getRange('A:A').setNumberFormat('@');
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'Ошибка сервера.' });
  } finally { lock.releaseLock(); }
}
