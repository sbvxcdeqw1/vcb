const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1920, height: 1080 } });
  await p.goto('file:///home/user/vcb/bpm_docs/deck.html');
  await p.pdf({ path: '/home/user/vcb/bpm_docs/Документация_бизнес-процессов.pdf', width: '1920px', height: '1080px', printBackground: true, preferCSSPageSize: true });
  const n = await p.evaluate(() => document.querySelectorAll('.slide').length);
  // screenshots for QA
  const fs = require('fs'); fs.mkdirSync('/tmp/qa', { recursive: true });
  for (let i = 0; i < n; i++) {
    const el = (await p.$$('.slide'))[i];
    await el.screenshot({ path: `/tmp/qa/s${String(i+1).padStart(2,'0')}.png` });
  }
  await b.close(); console.log('rendered', n);
})();
