// Тесты формы показаний: node form/tests/run.js   (браузерная часть — если установлен playwright)
// Данные в fixtures/ вымышленные: реальные справочники в репозиторий не кладутся.
const assert = require('assert'), fs = require('fs'), path = require('path');
const { load } = require('./gas_mock');
const DATA = path.join(__dirname, 'fixtures');
let passed = 0;
const t = (name, f) => Promise.resolve().then(f).then(() => { passed++; console.log('  ok ', name); },
  e => { console.log('  FAIL', name, '\n', e); process.exitCode = 1; });

function fresh(now) {
  const G = load({ now });
  G.setup();
  G.loadCsv('Лицевые счета', path.join(DATA, 'accounts.csv'));
  G.loadCsv('Приборы', path.join(DATA, 'meters.csv'));
  return G;
}
const KV1 = '990000000122', OK = { consent: true, area: '76,5' };
const kv1 = (vals) => ({ ...OK, account: KV1, readings: { r1: vals[0], r2: vals[1], r3: vals[2], r4: vals[3] } });

(async () => {
  console.log('Сервер (apps_script.gs):');
  await t('настройки по названию; старая строка «Всего квартир» не ломает текст', () => {
    const G = load({ now: new Date(2026, 10, 10) });
    G.setup();
    const s = G.ss.sheets['Настройки'];
    s.rows = [['День начала приёма', 20], ['День окончания приёма', 25], ['Приём включён', 'ДА'], ['Всего квартир', 100],
              ['Текст при закрытом окне', 'Закрыто с {start} по {end}']];
    const st = G.get();
    assert.strictEqual(st.open, false); assert.strictEqual(st.message, 'Закрыто с 20 по 25');
    assert(s.rows.some(r => r[0] === 'Проверять площадь'), 'недостающие настройки дописаны');
  });
  await t('окно через границу месяцев: 2 ноября -> период 2026-10', () => {
    const G = fresh(new Date(2026, 10, 2));
    G.setSetting('День начала приёма', 28); G.setSetting('День окончания приёма', 3);
    const st = G.get(); assert.strictEqual(st.open, true); assert.strictEqual(st.period, '2026-10');
  });
  await t('январь, окно 28→3: период — декабрь прошлого года', () => {
    const G = fresh(new Date(2027, 0, 2));
    G.setSetting('День начала приёма', 28); G.setSetting('День окончания приёма', 3);
    assert.strictEqual(G.get().period, '2026-12');
  });
  await t('поиск счёта: приборы без показаний, 13 цифр, неверный', () => {
    const G = fresh();
    const r = G.get({ account: KV1 });
    assert.strictEqual(r.place, '1'); assert.strictEqual(r.meters.length, 4);
    assert(!JSON.stringify(r).includes('27,759'), 'предыдущие показания не раскрываются');
    assert.match(r.meters.find(m => m.id === 'r3').label, /^ХВС, ванная · №…5228$/);
    assert.strictEqual(G.get({ account: '9900000009022' }).place, '90');
    assert.strictEqual(G.get({ account: '990000000999' }).found, false);
    assert.strictEqual(G.get({ account: '99000000012' }).found, false);
  });
  await t('отказы: без согласия, неверная площадь, формат, меньше предыдущего', () => {
    const G = fresh();
    assert.match(G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), consent: false }).error, /согласие/);
    assert.match(G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), area: '70' }).error, /Площадь/);
    const f = G.post(kv1(['28', '29,000', '35,000', '13,000'])); assert.ok(f.fields.r1);
    const less = G.post(kv1(['27,000', '29,000', '35,000', '13,000']));
    assert.match(less.error, /меньше предыдущего/); assert.match(less.fields.r1, /27,759/);
    const miss = G.post({ ...OK, account: KV1, readings: { r1: '28,000' } }); assert.ok(miss.fields.r2);
  });
  await t('блокировка после 5 неверных площадей', () => {
    const G = fresh();
    for (let i = 0; i < 5; i++) G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), area: '1' });
    assert.match(G.post(kv1(['28,000', '29,000', '35,000', '13,000'])).error, /Слишком много/);
  });
  await t('большой расход: сначала подтверждение, потом приём с пометкой', () => {
    const G = fresh();
    const r1 = G.post(kv1(['28,000', '29,000', '80,000', '13,000']));
    assert.strictEqual(r1.confirm, true); assert.match(r1.warnings[0], /ХВС, ванная.*45,709/);
    assert.strictEqual(G.post({ ...kv1(['28,000', '29,000', '80,000', '13,000']), confirmBig: true }).ok, true);
    const row = G.ss.sheets['Показания'].rows.find(r => r[4] === 'ХВ ван');
    assert.match(row[11], /большой расход/);
  });
  await t('порог учитывает месяцы с прошлой даты (24.08 -> ~3 мес)', () => {
    const G = fresh();
    const r = G.post({ ...OK, area: '35,9', account: '990000000222', readings: { r5: '81,000', r6: '190,000' } });
    assert.strictEqual(r.ok, true, JSON.stringify(r));                // +34.5 м³ за 3 месяца при пороге 15/мес — без вопроса
  });
  await t('запись, повторная отправка заменяет, «Для 1С» в порядке 1С, «Не подали»', () => {
    const G = fresh();
    assert.strictEqual(G.post(kv1(['28,000', '29,000', '35,000', '13,000'])).ok, true);
    assert.strictEqual(G.post(kv1(['28,100', '29,000', '35,000', '13,000'])).ok, true);
    const rows = G.ss.sheets['Показания'].rows;
    assert.strictEqual(rows.length, 5);
    const gv = rows.find(r => r[4] === 'ГВ ван');
    assert.deepStrictEqual([gv[6], gv[7], gv[8], gv[9]], [27.759, '23.10.2026', 28.1, 0.341]);
    G.buildFor1C('2026-11');
    const out = G.ss.sheets['Для 1С'].rows;
    assert.deepStrictEqual(out.slice(1).map(r => r[2]), ['ГВ ван', 'ГВ кух', 'ХВ ван', 'ХВ кух']);
    assert.deepStrictEqual(out[0].slice(0, 7), ['N', 'Объект учета', 'Прибор учета', 'Заводской номер', 'Предыдущее показание', 'Текущее показание', 'Объем']);
    G.buildNotSubmitted('2026-11');
    const ns = G.ss.sheets['Не подали'].rows.slice(1).map(r => r[0]);
    assert(!ns.includes(KV1) && ns.includes('990000000222'));
  });
  await t('следующий месяц: предыдущим становится принятое формой', () => {
    const G = fresh(new Date(2026, 10, 22));
    G.post(kv1(['28,000', '29,000', '35,000', '13,000']));
    G.g.now = new Date(2026, 11, 22);
    const r = G.post(kv1(['27,900', '29,500', '35,500', '13,500']));
    assert.match(r.fields.r1, /28,000/);
  });
  await t('ручной режим (нет приборов в базе): ванна/кухня обязательны', () => {
    const G = fresh();
    const base = { ...OK, area: '41,7', account: '990000026222' };
    assert.strictEqual(G.get({ account: '990000026222' }).meters.length, 0);
    assert.match(G.post({ ...base, rows: [{ cold: '1,000', hot: '1,000' }, { cold: '2,000', hot: '2,000' }] }).error, /ванна или кухня/);
    assert.match(G.post({ ...base, rows: [{ place: 'ванна', cold: '1,000', hot: '1,000' }, { place: 'ванна', cold: '2,000', hot: '2,000' }] }).error, /разное место/);
    assert.strictEqual(G.post({ ...base, rows: [{ place: 'ванна', cold: '1,000', hot: '1,000' }, { place: 'кухня', cold: '2,000', hot: '2,000' }] }).ok, true);
    const names = G.ss.sheets['Показания'].rows.slice(1).map(r => r[4] + ':' + r[11]);
    assert.deepStrictEqual(names, ['ХВ ван:прибор не найден в базе 1С', 'ГВ ван:прибор не найден в базе 1С', 'ХВ кух:прибор не найден в базе 1С', 'ГВ кух:прибор не найден в базе 1С']);
  });
  await t('письмо-подтверждение только при включённой настройке', () => {
    const G = fresh();
    G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), email: 'a@b.ru' }); assert.strictEqual(G.g.mail.length, 0);
    G.setSetting('Подтверждение на почту', 'ДА');
    const r = G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), email: 'a@b.ru' }); assert.strictEqual(G.g.mail.length, 1);
    assert.strictEqual(r.mailed, true);
    const m = G.g.mail[0];
    assert.strictEqual(m.o.name, 'УК «Водники-1»'); assert.strictEqual(m.o.replyTo, undefined);
    assert.match(m.s, /кв\. 1$/); assert.match(m.b, /ХВС, ванная · №…5228: 35,000 м³/);
    G.setSetting('Адрес для ответов', 'uk@example.ru');
    G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), email: 'a@b.ru' }); assert.strictEqual(G.g.mail[1].o.replyTo, 'uk@example.ru');
    G.g.quota = 0;
    assert.strictEqual(G.post({ ...kv1(['28,000', '29,000', '35,000', '13,000']), email: 'a@b.ru' }).mailed, false);
    assert.strictEqual(G.get().mailReply, undefined, 'настройки почты не уходят на сайт');
  });
  await t('форма: поле e-mail появляется при включённой настройке', () => {
    const G = fresh(); assert.strictEqual(G.get().email, false);
    G.setSetting('Подтверждение на почту', 'ДА'); assert.strictEqual(G.get().email, true);
  });
  await t('старый лист «Показания» с другими столбцами переименовывается', () => {
    const H = fresh();
    H.ss.sheets['Показания'].rows = [['Период', 'Лицевой счёт', 'Квартира', 'Место', 'ХВС', 'ГВС', 'Принято', 'Примечание'], ['2026-10', KV1, '1', '', '1,000', '1,000', '', '']];
    assert.strictEqual(H.post(kv1(['28,000', '29,000', '35,000', '13,000'])).ok, true);
    assert(Object.keys(H.ss.sheets).some(n => n.startsWith('Показания (старая версия')));
    assert.strictEqual(H.ss.sheets['Показания'].rows[0][3], 'Объект учета');
  });

  let chromium; try { ({ chromium } = require('playwright')); } catch (e) { console.log('playwright не найден — браузерные тесты пропущены'); }
  if (chromium) {
    console.log('Форма (tilda_block.html) + сервер в браузере:');
    const browser = await chromium.launch();
    const open = async (G, width = 400) => {
      const p = await browser.newPage({ viewport: { width, height: 900 } });
      p.on('pageerror', e => { throw e; });
      await p.route('https://script.google.com/**', r => {
        const u = new URL(r.request().url()), params = Object.fromEntries(u.searchParams);
        const body = r.request().method() === 'POST' ? G.doPost({ postData: { contents: r.request().postData() } }).t : G.doGet({ parameter: params }).t;
        r.fulfill({ contentType: 'application/json', headers: { 'access-control-allow-origin': '*' }, body });
      });
      p.dialogs = []; p.on('dialog', d => { p.dialogs.push(d.message()); d.accept(); });
      await p.setContent(fs.readFileSync(path.join(__dirname, '..', 'tilda_block.html'), 'utf8').replace('ВАШ_ID', 'X'));
      return p;
    };
    const typeAcc = async (p, a) => { await p.fill('#vcb-acc', a); await p.waitForFunction(() => !/Проверяем/.test(document.getElementById('vcb-flat').textContent)); };
    await t('известные приборы: ХВС первыми, согласие и площадь обязательны, приём', async () => {
      const G = fresh(), p = await open(G);
      await typeAcc(p, KV1);
      assert.strictEqual((await p.textContent('#vcb-flat')).trim(), 'Квартира 1');
      const labels = await p.$$eval('.known label', ls => ls.map(l => l.textContent));
      assert.deepStrictEqual(labels.map(l => l.split(' · ')[0]), ['ХВС, ванная', 'ХВС, кухня', 'ГВС, ванная', 'ГВС, кухня']);
      assert.strictEqual(await p.isDisabled('#vcb-send'), true, 'без согласия кнопка неактивна');
      await p.check('#vcb-consent'); assert.strictEqual(await p.isDisabled('#vcb-send'), false);
      const ins = p.locator('.known input');
      for (const [i, v] of [[0, '35'], [1, '13'], [2, '28,0'], [3, '29']]) await ins.nth(i).fill(v);
      await p.click('#vcb-send');
      assert.match(await p.textContent('#vcb-msg'), /Проверьте выделенные/); // нет площади
      await p.fill('#vcb-area', '76,5'); await p.click('#vcb-send');
      await p.waitForFunction(() => /приняты/.test(document.getElementById('vcb-msg').textContent));
      assert.deepStrictEqual(await ins.count(), 0);
      assert.strictEqual(G.ss.sheets['Показания'].rows.length, 5);
      assert.match(p.dialogs[0], /ХВС, ванная · №…5228: 35,000/);
    });
    await t('меньше предыдущего подсвечивается у поля', async () => {
      const G = fresh(), p = await open(G);
      await typeAcc(p, KV1); await p.fill('#vcb-area', '76,5'); await p.check('#vcb-consent');
      const ins = p.locator('.known input');
      await ins.nth(0).fill('30'); await ins.nth(1).fill('13'); await ins.nth(2).fill('28'); await ins.nth(3).fill('29');
      await p.click('#vcb-send'); await p.waitForFunction(() => /меньше предыдущего/.test(document.getElementById('vcb-msg').textContent));
      assert.match(await p.locator('.known .fe').first().textContent(), /34,291/);
    });
    await t('большой расход: подтверждение в окне браузера', async () => {
      const G = fresh(), p = await open(G);
      await typeAcc(p, KV1); await p.fill('#vcb-area', '76,5'); await p.check('#vcb-consent');
      const ins = p.locator('.known input');
      await ins.nth(0).fill('90'); await ins.nth(1).fill('13'); await ins.nth(2).fill('28'); await ins.nth(3).fill('29');
      await p.click('#vcb-send'); await p.waitForFunction(() => /приняты/.test(document.getElementById('vcb-msg').textContent));
      assert(p.dialogs.some(d => /расход больше обычного/.test(d)));
    });
    await t('ручной режим и «Помещение ТАС»; ширина на телефоне', async () => {
      const G = fresh(), p = await open(G, 360);
      await typeAcc(p, '990000026222');
      assert.strictEqual((await p.textContent('#vcb-flat')).trim(), 'Помещение ТАС');
      await p.click('#vcb-add'); await p.fill('#vcb-area', '41,7'); await p.check('#vcb-consent');
      await p.locator('.m-cold').nth(0).fill('1'); await p.locator('.m-hot').nth(0).fill('2');
      await p.locator('.m-cold').nth(1).fill('3'); await p.locator('.m-hot').nth(1).fill('4');
      await p.click('#vcb-send'); assert.match(await p.textContent('#vcb-msg'), /Проверьте выделенные/);
      await p.selectOption('.m-place >> nth=0', 'ванна'); await p.selectOption('.m-place >> nth=1', 'кухня');
      await p.click('#vcb-send'); await p.waitForFunction(() => /приняты/.test(document.getElementById('vcb-msg').textContent));
      const box = await p.evaluate(() => [document.getElementById('vcb-acc').getBoundingClientRect(), document.getElementById('vcb-send').getBoundingClientRect(), document.documentElement.scrollWidth]);
      assert.strictEqual(Math.round(box[0].left), Math.round(box[1].left)); assert.strictEqual(Math.round(box[0].right), Math.round(box[1].right));
      assert(box[0].left >= 16, 'отступ от края'); assert.strictEqual(box[2], 360, 'нет горизонтальной прокрутки');
    });
    await t('e-mail: поле видно, после приёма — сообщение о письме', async () => {
      const G = fresh(); G.setSetting('Подтверждение на почту', 'ДА');
      const p = await open(G);
      await typeAcc(p, KV1);
      assert.strictEqual(await p.isVisible('#vcb-email'), true);
      await p.fill('#vcb-area', '76,5'); await p.fill('#vcb-email', 'a@b.ru'); await p.check('#vcb-consent');
      const ins = p.locator('.known input');
      await ins.nth(0).fill('35'); await ins.nth(1).fill('13'); await ins.nth(2).fill('28'); await ins.nth(3).fill('29');
      await p.click('#vcb-send'); await p.waitForFunction(() => /приняты/.test(document.getElementById('vcb-msg').textContent));
      assert.match(await p.textContent('#vcb-msg'), /Подтверждение отправлено на a@b\.ru/);
      assert.strictEqual(G.g.mail.length, 1);
    });
    await t('окно закрыто — кнопка неактивна, текст из настроек', async () => {
      const G = fresh(new Date(2026, 10, 10)), p = await open(G);
      await p.waitForFunction(() => /Показания принимаются/.test(document.getElementById('vcb-msg').textContent));
      await typeAcc(p, KV1); await p.check('#vcb-consent');
      assert.strictEqual(await p.isDisabled('#vcb-send'), true);
    });
    await browser.close();
  }
  console.log(`\nПройдено: ${passed}` + (process.exitCode ? ' — ЕСТЬ ОШИБКИ' : ''));
})();
