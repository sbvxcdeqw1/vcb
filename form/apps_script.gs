/**
 * Приём показаний ХВС/ГВС с сайта на Tilda -> Google Таблица.  Установка: см. README.md
 *
 * Листы (создаются автоматически):
 *   «Настройки»      — окно приёма и параметры (меняйте значения в столбце B)
 *   «Лицевые счета»  — справочник: Лицевой счёт | Помещение (загружается из accounts.csv)
 *   «Показания»      — Период | Лицевой счёт | Квартира | Место | ХВС | ГВС | Принято | Примечание
 *                      (копируется в 1С; повторная отправка за месяц заменяет строки счёта)
 *   «Журнал»         — все отправки подряд
 */
var SETTINGS = [
  ['День начала приёма', 20, 'число месяца, включительно'],
  ['День окончания приёма', 25, 'включительно; если меньше дня начала — окно переходит на следующий месяц (28 → 3)'],
  ['Приём включён', 'ДА', 'ДА / НЕТ — аварийный выключатель, перекрывает даты'],
  ['Текст при закрытом окне', 'Показания принимаются с {start} по {end} число каждого месяца.', '{start} и {end} подставляются сами']
];
var HEAD = ['Период', 'Лицевой счёт', 'Квартира', 'Место', 'ХВС', 'ГВС', 'Принято', 'Примечание'];
var PLACES = ['ванна', 'кухня'];
var VALUE_RE = /^\d{1,5},\d{3}$/;      // 27,124  144,000  (форма приводит ввод к этому виду)

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) { sh.appendRow(header); sh.setFrozenRows(1); sh.getRange(1, 1, 1, header.length).setFontWeight('bold'); }
    sh.getRange('A:H').setNumberFormat('@');                            // всё как текст: без потери запятых и нулей
  }
  return sh;
}
function settings_() {
  var sh = sheet_('Настройки');
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, SETTINGS.length, 3).setValues(SETTINGS); sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 260); }
  var v = sh.getRange(1, 2, SETTINGS.length, 1).getValues();
  return { start: +v[0][0], end: +v[1][0], on: String(v[2][0]).toUpperCase() !== 'НЕТ', closedText: String(v[3][0]) };
}
function status_() {
  var s = settings_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var now = new Date(), day = +Utilities.formatDate(now, tz, 'd');
  var wrap = s.start > s.end;
  var inWin = wrap ? (day >= s.start || day <= s.end) : (day >= s.start && day <= s.end);
  var open = s.on && inWin;
  var ref = (wrap && day <= s.end) ? new Date(now.getFullYear(), now.getMonth() - 1, 15) : now;   // период = месяц начала окна
  return { open: open, start: s.start, end: s.end, period: Utilities.formatDate(ref, tz, 'yyyy-MM'),
           message: open ? '' : s.closedText.replace('{start}', s.start).replace('{end}', s.end) };
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Квартира по лицевому счёту — только при полном совпадении (12–13 цифр). Список счетов наружу не отдаётся.
function lookup_(account) {
  account = String(account || '').trim();
  if (!/^\d{12,13}$/.test(account)) return null;
  var d = sheet_('Лицевые счета', ['Лицевой счёт', 'Помещение']).getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (String(d[i][0]).trim() === account) return String(d[i][1]).trim();
  return null;
}

// GET без параметров — состояние окна приёма; GET ?account=… — проверка лицевого счёта.
function doGet(e) {
  var st = status_();
  if (e && e.parameter && e.parameter.account !== undefined) {
    var place = lookup_(e.parameter.account);
    st.found = place !== null; st.place = place || '';
  }
  return json_(st);
}

function toNum_(s) { return parseFloat(String(s).replace(',', '.')); }

function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents), st = status_();
    if (d.website) return json_({ ok: true });                                  // ловушка для ботов
    if (!st.open) return json_({ ok: false, error: st.message || 'Приём показаний закрыт.' });

    var account = String(d.account || '').trim(), flat = lookup_(account);
    if (flat === null) return json_({ ok: false, error: 'Лицевой счёт не найден. Проверьте номер.' });

    var rows = d.rows;
    if (!Array.isArray(rows) || rows.length < 1 || rows.length > 2) return json_({ ok: false, error: 'Укажите один или два прибора учёта.' });
    var places = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!VALUE_RE.test(String(r.cold)) || !VALUE_RE.test(String(r.hot))) return json_({ ok: false, error: 'Проверьте подлинность значений ХВС и ГВС.' });
      if (rows.length > 1) {
        if (PLACES.indexOf(r.place) < 0) return json_({ ok: false, error: 'Укажите, где установлен прибор: ванна или кухня.' });
        if (places[r.place]) return json_({ ok: false, error: 'Места приборов учёта должны различаться.' });
        places[r.place] = 1;
      } else r.place = '';
    }

    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(), stamp = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
    var log = sheet_('Журнал', ['Принято', 'Лицевой счёт', 'Помещение', 'Место', 'ХВС', 'ГВС']);
    var sh = sheet_('Показания', HEAD), data = sh.getDataRange().getValues();

    // прошлые показания этого счёта (последний период до текущего) — для пометки «меньше прошлого»
    var prev = {}, prevPeriod = '';
    for (var k = 1; k < data.length; k++)
      if (String(data[k][1]) === account && String(data[k][0]) < st.period && String(data[k][0]) > prevPeriod) prevPeriod = String(data[k][0]);
    for (k = 1; k < data.length; k++)
      if (String(data[k][1]) === account && String(data[k][0]) === prevPeriod) prev[String(data[k][3])] = data[k];

    // убрать прежние строки этого счёта за этот период (повторная отправка заменяет их)
    for (k = data.length - 1; k >= 1; k--)
      if (String(data[k][0]) === st.period && String(data[k][1]) === account) sh.deleteRow(k + 1);

    rows.forEach(function (r) {
      var notes = [], p = prev[r.place];
      if (p && toNum_(r.cold) < toNum_(p[4])) notes.push('ХВС меньше прошлого (' + p[4] + ')');
      if (p && toNum_(r.hot) < toNum_(p[5])) notes.push('ГВС меньше прошлого (' + p[5] + ')');
      sh.appendRow([st.period, account, flat, r.place, r.cold, r.hot, stamp, notes.join('; ')]);
      log.appendRow([stamp, account, flat, r.place, r.cold, r.hot]);
    });
    return json_({ ok: true, flat: flat });
  } catch (err) {
    return json_({ ok: false, error: 'Ошибка сервера.' });
  } finally { lock.releaseLock(); }
}
