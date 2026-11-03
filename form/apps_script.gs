/**
 * Приём показаний ХВС/ГВС с сайта на Tilda -> Google Таблица -> 1С.  Установка и обновление: README.md
 *
 * Листы:
 *   «Настройки»     — параметры по названию в столбце A, значение в столбце B (недостающие строки добавляются сами)
 *   «Лицевые счета» — справочник из accounts.csv: Лицевой счёт | Помещение | Объект учета | Площадь
 *   «Приборы»       — справочник из meters.csv:   Порядок | Объект учета | Прибор учета | Заводской номер | Тип | Место | Пред. показание | Дата пред.
 *   «Показания»     — всё принятое (повторная отправка за период заменяет строки счёта)
 *   «Журнал»        — все отправки подряд
 *   «Для 1С», «Не подали» — формируются из меню «Показания» в таблице
 */
var SETTINGS = [
  ['День начала приёма', 20, 'число месяца, включительно'],
  ['День окончания приёма', 25, 'включительно; если меньше дня начала — окно переходит на следующий месяц (28 → 3)'],
  ['Приём включён', 'ДА', 'ДА / НЕТ — аварийный выключатель, перекрывает даты'],
  ['Текст при закрытом окне', 'Показания принимаются с {start} по {end} число каждого месяца.', '{start} и {end} подставляются сами'],
  ['Проверять площадь', 'ДА', 'ДА / НЕТ — жилец вводит общую площадь из квитанции, она сверяется со справочником'],
  ['Макс. расход в месяц, м³', 15, 'больше — форма просит подтвердить, в «Примечании» отметка'],
  ['Подтверждение на почту', 'НЕТ', 'ДА / НЕТ — поле e-mail в форме и письмо после приёма (лимит Gmail ~100 писем в сутки)'],
  ['Имя отправителя письма', 'УК «Водники-1»', 'так письмо подписано во «Входящих» жильца'],
  ['Адрес для ответов', '', 'куда уйдёт ответ жильца на письмо; пусто — на адрес аккаунта Google']
];
var SHOW_HEAD = ['Период', 'Лицевой счёт', 'Помещение', 'Объект учета', 'Прибор учета', 'Заводской номер',
                 'Предыдущее показание', 'Дата пред.', 'Текущее показание', 'Объем', 'Принято', 'Примечание'];
var VALUE_RE = /^\d{1,5},\d{3}$/;
var MAX_AREA_FAILS = 5;                          // неверных площадей подряд -> блокировка счёта на час

// ---------- общее ----------
function ss_() { return SpreadsheetApp.getActive(); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function sheet_(name, header, textCols) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    if (header) { sh.appendRow(header); sh.setFrozenRows(1); sh.getRange(1, 1, 1, header.length).setFontWeight('bold'); }
    if (textCols) sh.getRange(textCols).setNumberFormat('@');
  }
  return sh;
}
// Лист с другими столбцами (от прежней версии) переименовывается, чтобы данные не смешались.
function ensureHeader_(name, header) {
  var sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() === 0) return;
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  if (h.slice(0, header.length).join('|') !== header.join('|')) sh.setName(name + ' (старая версия ' + Utilities.formatDate(new Date(), tz_(), 'dd.MM.yyyy HH.mm') + ')');
}
var LOG_HEAD = ['Принято', 'Лицевой счёт', 'Прибор учета', 'Заводской номер', 'Показание', 'Согласие на обработку ПДн', 'E-mail'];
function rows_(name, header) { var v = sheet_(name, header).getDataRange().getValues(); return v.slice(1); }
function num_(x) { return typeof x === 'number' ? x : parseFloat(String(x).replace(/\s/g, '').replace(',', '.')); }
function f3_(n) { return n.toFixed(3).replace('.', ','); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function yes_(v) { return String(v).trim().toUpperCase() === 'ДА'; }
function dateKey_(v) {                            // Date | 'dd.MM.yyyy[ HH:mm]' -> 'yyyyMMdd'
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyyMMdd');
  var m = String(v || '').match(/(\d\d)\.(\d\d)\.(\d{4})/);
  return m ? m[3] + m[2] + m[1] : '';
}
function keyToDate_(k) { return new Date(+k.slice(0, 4), +k.slice(4, 6) - 1, +k.slice(6, 8)); }
function keyToText_(k) { return k ? k.slice(6, 8) + '.' + k.slice(4, 6) + '.' + k.slice(0, 4) : ''; }

// ---------- настройки (по названию, не по номеру строки) ----------
function settings_() {
  var sh = sheet_('Настройки'), v = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 2).getValues() : [], map = {};
  v.forEach(function (r) { map[String(r[0]).trim()] = r[1]; });
  SETTINGS.forEach(function (s) {
    if (!(s[0] in map)) { sh.appendRow(s); map[s[0]] = s[1]; }
  });
  return { start: +map[SETTINGS[0][0]], end: +map[SETTINGS[1][0]], on: yes_(map[SETTINGS[2][0]]),
           closedText: String(map[SETTINGS[3][0]]), area: yes_(map[SETTINGS[4][0]]),
           maxMonth: num_(map[SETTINGS[5][0]]) || 0, email: yes_(map[SETTINGS[6][0]]),
           mailName: String(map[SETTINGS[7][0]] || '').trim(), mailReply: String(map[SETTINGS[8][0]] || '').trim() };
}
function status_() {
  var s = settings_(), now = new Date(), day = +Utilities.formatDate(now, tz_(), 'd');
  var wrap = s.start > s.end;
  var open = s.on && (wrap ? (day >= s.start || day <= s.end) : (day >= s.start && day <= s.end));
  var y = +Utilities.formatDate(now, tz_(), 'yyyy'), mo = +Utilities.formatDate(now, tz_(), 'M');
  if (wrap && day <= s.end) { mo -= 1; if (mo === 0) { mo = 12; y -= 1; } }     // период = месяц начала окна
  return { open: open, period: y + '-' + ('0' + mo).slice(-2), needArea: s.area, email: s.email, maxMonth: s.maxMonth,
           mailName: s.mailName, mailReply: s.mailReply,
           message: open ? '' : s.closedText.replace('{start}', s.start).replace('{end}', s.end) };
}

// ---------- справочники ----------
function account_(acc) {
  acc = String(acc || '').trim();
  if (!/^\d{12,13}$/.test(acc)) return null;
  var d = rows_('Лицевые счета', ['Лицевой счёт', 'Помещение', 'Объект учета', 'Площадь']);
  for (var i = 0; i < d.length; i++)
    if (String(d[i][0]).trim() === acc) return { acc: acc, place: String(d[i][1]).trim(), obj: String(d[i][2]).trim(), area: num_(d[i][3]) };
  return null;
}
var KIND = { 'ХВ': 'ХВС', 'ГВ': 'ГВС' };
function meters_(obj) {
  if (!obj) return [];
  return rows_('Приборы', ['Порядок', 'Объект учета', 'Прибор учета', 'Заводской номер', 'Тип', 'Место', 'Пред. показание', 'Дата пред.'])
    .filter(function (r) { return String(r[1]).trim() === obj; })
    .map(function (r) {
      var serial = String(r[3]).trim(), kind = String(r[4]).trim(), place = String(r[5]).trim();
      var label = KIND[kind] ? KIND[kind] + (place ? ', ' + place : '') : String(r[2]).trim();
      return { id: 'r' + r[0], order: +r[0], name: String(r[2]), serial: serial, kind: kind,
               label: label + (serial ? ' · №…' + serial.slice(-4) : ''),
               prev: r[6] === '' ? null : num_(r[6]), prevKey: dateKey_(r[7]) };
    });
}
function manualMeters_(rows) {                   // счёт без приборов в базе: строки «ванна/кухня» из формы
  var out = [];
  rows.forEach(function (r, i) {
    var suf = r.place === 'ванна' ? ' ван' : r.place === 'кухня' ? ' кух' : '';
    [['ХВ', 'cold'], ['ГВ', 'hot']].forEach(function (k) {
      out.push({ id: 'm' + i + k[1][0], order: 1e9, name: k[0] + suf, serial: '', kind: k[0], value: r[k[1]],
                 label: KIND[k[0]] + (r.place ? ', ' + r.place : ''), prev: null, prevKey: '' });
    });
  });
  return out;
}

// ---------- GET: состояние окна / проверка счёта ----------
function doGet(e) {
  var st = status_(), p = (e && e.parameter) || {};
  if (p.account !== undefined) {
    var a = account_(p.account);
    st.found = !!a;
    if (a) {
      st.place = a.place;
      st.meters = meters_(a.obj).map(function (m) { return { id: m.id, label: m.label, kind: m.kind }; });  // без показаний
    }
  }
  delete st.maxMonth; delete st.mailName; delete st.mailReply;
  return json_(st);
}

// ---------- POST: приём ----------
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return json_(accept_(JSON.parse(e.postData.contents))); }
  catch (err) { console.error(err); return json_({ ok: false, error: 'Ошибка сервера.' }); }
  finally { lock.releaseLock(); }
}

function accept_(d) {
  var st = status_();
  if (d.website) return { ok: true };                                            // ловушка для ботов
  if (!st.open) return { ok: false, error: st.message || 'Приём показаний закрыт.' };
  if (d.consent !== true) return { ok: false, error: 'Нужно согласие на обработку персональных данных.' };
  var a = account_(d.account);
  if (!a) return { ok: false, error: 'Лицевой счёт не найден. Проверьте номер.' };

  if (st.needArea) {
    var cache = CacheService.getScriptCache(), ck = 'areafail_' + a.acc, fails = +(cache.get(ck) || 0);
    if (fails >= MAX_AREA_FAILS) return { ok: false, error: 'Слишком много неверных попыток. Попробуйте через час или обратитесь в управляющую компанию.' };
    if (!(Math.abs(num_(d.area) - a.area) < 0.051)) {
      cache.put(ck, String(fails + 1), 3600);
      return { ok: false, error: 'Площадь не совпадает с данными лицевого счёта. Проверьте квитанцию.', fields: { area: 'Не совпадает' } };
    }
    cache.remove(ck);
  }

  // приборы: из базы 1С или ручной ввод
  var list = meters_(a.obj), manual = !list.length, fields = {};
  if (manual) {
    var rows = d.rows;
    if (!Array.isArray(rows) || rows.length < 1 || rows.length > 2) return { ok: false, error: 'Укажите один или два набора приборов.' };
    if (rows.length > 1) {
      var pl = rows.map(function (r) { return r.place; });
      if (pl.some(function (p) { return p !== 'ванна' && p !== 'кухня'; }) || pl[0] === pl[1])
        return { ok: false, error: 'Укажите для каждой строки разное место: ванна или кухня.' };
    } else rows[0].place = '';
    list = manualMeters_(rows);
  } else {
    var vals = d.readings || {};
    list.forEach(function (m) { m.value = vals[m.id]; });
  }
  list.forEach(function (m) { if (!VALUE_RE.test(String(m.value))) fields[m.id] = 'Проверьте подлинность значения'; });
  if (Object.keys(fields).length) return { ok: false, error: 'Проверьте подлинность значений.', fields: fields };

  // предыдущее показание: из базы 1С или принятое формой раньше (что свежее)
  ensureHeader_('Показания', SHOW_HEAD); ensureHeader_('Журнал', LOG_HEAD);
  var shown = rows_('Показания', SHOW_HEAD), todayKey = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd');
  list.forEach(function (m) {
    shown.forEach(function (r) {
      if (String(r[1]) !== a.acc || String(r[0]) >= st.period) return;
      if (String(r[4]).trim() !== m.name.trim() || String(r[5]).trim() !== m.serial) return;
      var k = dateKey_(r[10]);
      if (k > m.prevKey) { m.prev = num_(r[8]); m.prevKey = k; }
    });
  });

  var warnings = [], errors = [];
  list.forEach(function (m) {
    m.num = num_(m.value); m.note = manual ? 'прибор не найден в базе 1С' : '';
    if (m.prev === null || !isFinite(m.prev)) return;
    if (m.num < m.prev - 1e-9) { fields[m.id] = 'Меньше предыдущего (' + f3_(m.prev) + ')'; errors.push(m.label + ': меньше предыдущего ' + f3_(m.prev)); return; }
    var months = m.prevKey ? Math.max(1, Math.round((keyToDate_(todayKey) - keyToDate_(m.prevKey)) / 864e5 / 30.4)) : 1;
    if (st.maxMonth > 0 && m.num - m.prev > st.maxMonth * months) {
      warnings.push(m.label + ': расход ' + f3_(m.num - m.prev) + ' м³');
      m.note = (m.note ? m.note + '; ' : '') + 'большой расход, подтверждён жильцом';
    }
  });
  if (errors.length) return { ok: false, error: 'Показание не может быть меньше предыдущего. Если прибор заменён — сообщите в управляющую компанию.', fields: fields };
  if (warnings.length && d.confirmBig !== true) return { ok: false, confirm: true, warnings: warnings };

  // запись
  var stamp = Utilities.formatDate(new Date(), tz_(), 'dd.MM.yyyy HH:mm');
  var sh = sheet_('Показания', SHOW_HEAD, 'A:F'), data = sh.getDataRange().getValues();
  for (var k = data.length - 1; k >= 1; k--) if (String(data[k][0]) === st.period && String(data[k][1]) === a.acc) sh.deleteRow(k + 1);
  var out = list.map(function (m) {
    var has = m.prev !== null && isFinite(m.prev);
    return [st.period, a.acc, a.place, a.obj, m.name, m.serial, has ? m.prev : '', keyToText_(m.prevKey), m.num,
            has ? Math.round((m.num - m.prev) * 1000) / 1000 : '', stamp, m.note];
  });
  var first = sh.getLastRow() + 1;
  sh.getRange(first, 1, out.length, SHOW_HEAD.length).setValues(out);
  sh.getRange(first, 7, out.length, 1).setNumberFormat('0.000');
  sh.getRange(first, 9, out.length, 2).setNumberFormat('0.000');

  var email = String(d.email || '').trim(), mailOk = st.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email);
  var log = sheet_('Журнал', LOG_HEAD, 'A:G');
  log.getRange(log.getLastRow() + 1, 1, out.length, 7).setValues(out.map(function (r) { return [stamp, a.acc, r[4], r[5], f3_(r[8]), 'да', mailOk ? email : '']; }));

  var mailed = false;
  if (mailOk) {
    try {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var opts = {};
        if (st.mailName) opts.name = st.mailName;
        if (st.mailReply) opts.replyTo = st.mailReply;
        MailApp.sendEmail(email, 'Показания приняты — ' + (/^\d+$/.test(a.place) ? 'кв. ' : '') + a.place,
          'Здравствуйте!\n\nПоказания по лицевому счёту ' + a.acc + ' (' + (/^\d+$/.test(a.place) ? 'квартира ' : '') + a.place +
          ') приняты ' + stamp + ', расчётный период ' + st.period + ':\n\n' +
          list.map(function (m) { return m.label + ': ' + f3_(m.num) + ' м³'; }).join('\n') +
          '\n\nЕсли вы не передавали эти показания или заметили ошибку, сообщите в управляющую компанию' +
          (st.mailReply ? ' — ответьте на это письмо.' : '.') +
          '\nПовторная отправка формы в этом периоде заменяет показания.\n\n' + (st.mailName || ''), opts);
        mailed = true;
      } else console.error('почта: исчерпан дневной лимит');
    } catch (err) { console.error('почта', err); }
  }
  return { ok: true, mailed: mailed };
}

// ---------- меню в таблице ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Показания')
    .addItem('Сформировать лист «Для 1С»', 'buildFor1C')
    .addItem('Сформировать лист «Не подали»', 'buildNotSubmitted')
    .addToUi();
}
function askPeriod_() {
  var def = status_().period, ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Период', 'Период в формате ГГГГ-ММ (по умолчанию ' + def + '):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var p = r.getResponseText().trim() || def;
  if (!/^\d{4}-\d\d$/.test(p)) { ui.alert('Неверный период: ' + p); return null; }
  return p;
}
function fresh_(name, header) {
  var sh = ss_().getSheetByName(name);
  if (sh) sh.clear(); else sh = ss_().insertSheet(name);
  sh.appendRow(header); sh.setFrozenRows(1); sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  return sh;
}
// Тот же порядок и столбцы, что в документе 1С «Ввод показаний индивидуальных приборов учета».
function buildFor1C(period) {
  period = period || askPeriod_(); if (!period) return;
  var order = {};
  rows_('Приборы').forEach(function (r) { order[String(r[1]).trim() + '|' + String(r[2]).trim() + '|' + String(r[3]).trim()] = +r[0]; });
  var data = rows_('Показания', SHOW_HEAD).filter(function (r) { return String(r[0]) === period; });
  data.sort(function (x, y) {
    var ox = order[String(x[3]).trim() + '|' + String(x[4]).trim() + '|' + String(x[5]).trim()] || 1e9;
    var oy = order[String(y[3]).trim() + '|' + String(y[4]).trim() + '|' + String(y[5]).trim()] || 1e9;
    return ox - oy || String(x[3]).localeCompare(String(y[3]));
  });
  var head = ['N', 'Объект учета', 'Прибор учета', 'Заводской номер', 'Предыдущее показание', 'Текущее показание', 'Объем',
              'Дата пред. показания', 'Дата показания', 'Комментарий'];
  var sh = fresh_('Для 1С', head);
  if (!data.length) { SpreadsheetApp.getActive().toast('За ' + period + ' показаний нет'); return; }
  var out = data.map(function (r, i) {
    return [i + 1, r[3], r[4], r[5], r[6], r[8], r[9], r[7], String(r[10]).slice(0, 10), r[11]];
  });
  sh.getRange(2, 1, out.length, head.length).setValues(out);
  sh.getRange(2, 4, out.length, 1).setNumberFormat('@');
  sh.getRange(2, 5, out.length, 3).setNumberFormat('0.000');
  sh.autoResizeColumns(1, head.length);
  SpreadsheetApp.getActive().toast('Для 1С: ' + out.length + ' приборов за ' + period);
}
function buildNotSubmitted(period) {
  period = period || askPeriod_(); if (!period) return;
  var done = {};
  rows_('Показания', SHOW_HEAD).forEach(function (r) { if (String(r[0]) === period) done[String(r[1])] = 1; });
  var withMeters = {};
  rows_('Приборы').forEach(function (r) { withMeters[String(r[1]).trim()] = 1; });
  var out = rows_('Лицевые счета').filter(function (r) { return r[0] !== '' && !done[String(r[0]).trim()]; })
    .map(function (r) { return [String(r[0]).trim(), r[1], r[2], withMeters[String(r[2]).trim()] ? 'да' : 'нет']; });
  var sh = fresh_('Не подали', ['Лицевой счёт', 'Помещение', 'Объект учета', 'Приборы в базе 1С']);
  if (out.length) { sh.getRange(2, 1, out.length, 1).setNumberFormat('@'); sh.getRange(2, 1, out.length, 4).setValues(out); }
  SpreadsheetApp.getActive().toast('Не подали за ' + period + ': ' + out.length);
}

// Запустить один раз после установки/обновления: создаёт листы, дописывает настройки, запрашивает разрешения.
function setup() {
  settings_();
  ensureHeader_('Показания', SHOW_HEAD); ensureHeader_('Журнал', LOG_HEAD);
  sheet_('Журнал', LOG_HEAD, 'A:G');
  sheet_('Лицевые счета', ['Лицевой счёт', 'Помещение', 'Объект учета', 'Площадь'], 'A:C');
  sheet_('Приборы', ['Порядок', 'Объект учета', 'Прибор учета', 'Заводской номер', 'Тип', 'Место', 'Пред. показание', 'Дата пред.'], 'B:F');
  sheet_('Показания', SHOW_HEAD, 'A:F');
  MailApp.getRemainingDailyQuota();
  CacheService.getScriptCache();
}
