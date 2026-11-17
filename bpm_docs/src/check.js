const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1920, height: 1080 } });
  await p.goto('file:///home/user/vcb/bpm_docs/deck.html');
  const r = await p.evaluate(() => {
    const out = [];
    document.querySelectorAll('.slide').forEach((sl, i) => {
      sl.querySelectorAll('.box').forEach(bx => {
        const t = bx.classList.contains('dec') ? bx.querySelector('.in') : bx;
        if (t.scrollHeight > t.clientHeight + 1) out.push(`slide ${i+1}: ${t.textContent.slice(0,60)} (${t.scrollHeight}>${t.clientHeight})`);
      });
      // content overflowing bottom
      const foot = sl.querySelector('.foot').getBoundingClientRect().top;
      sl.querySelectorAll('table.t, .card').forEach(el => { const r = el.getBoundingClientRect(); if (r.bottom > foot + 2) out.push(`slide ${i+1}: content overflow ${Math.round(r.bottom)}>${Math.round(foot)}`); });
    });
    return out;
  });
  console.log(r.join('\n') || 'no overflow');
  await b.close();
})();
