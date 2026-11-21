// Reads the state of the add-in pages inside PowerPoint's WebView2 over the DevTools protocol (test only).
//   node cdp.js <part of the add-in url> <label>
// Needs the HKLM WebView2 policy --remote-debugging-port=9222 and `npm i playwright-core` next to this file.
const { chromium } = require('playwright-core');
(async () => {
  const [part, label] = process.argv.slice(2);
  const browser = await chromium.connectOverCDP('http://127.0.0.1:9222', { timeout: 15000 });
  const out = [];
  for (const ctx of browser.contexts()) {
    for (const page of ctx.pages()) {
      if (!page.url().includes(part)) continue;
      const info = {};
      try {
        info.host = await page.evaluate(() => ({
          view: document.body.className,
          note: document.getElementById('note').textContent.slice(0, 40),
          panel: getComputedStyle(document.getElementById('panel')).display,
          shown: document.getElementById('view').style.visibility,
          size: innerWidth + 'x' + innerHeight,
          edg: /Edg\/\d+/.exec(navigator.userAgent) && /Edg\/\d+/.exec(navigator.userAgent)[0],
        }));
        // Office can add frames of its own to the page: the deck is the frame that carries the slide marker.
        let deck = null;
        let other = null;
        for (const f of page.frames()) {
          if (f === page.mainFrame()) continue;
          let marked = false;
          try { marked = await f.evaluate(() => typeof window.__LIVE_SLIDE__ !== 'undefined'); } catch (e) { /* frame gone */ }
          if (marked) { deck = f; break; }
          other = other || f;
        }
        deck = deck || other;
        info.host.frames = page.frames().length;
        if (deck) {
          info.deck = await deck.evaluate(() => {
            const q = (s) => document.querySelectorAll(s).length;
            const on = document.querySelector('.slide.on');
            const el = document.activeElement;
            return {
              slide: window.__LIVE_SLIDE__, on: on && on.id, liveSlide: !!window.liveSlide,
              flipsOpen: q('.flip.open'), ok: q('.opt.ok'), no: q('.opt.no'), checks: q('.check.on'),
              frac: (document.getElementById('frac') || {}).textContent || null,
              name: (document.getElementById('name') || {}).value || null,
              clock: (document.getElementById('clock') || {}).textContent || null,
              pick: (document.getElementById('pick') || {}).textContent || null,
              focus: el ? el.tagName + (el.id ? '#' + el.id : '') : null,
            };
          });
        } else { info.deck = null; }
      } catch (e) { info.err = String(e).slice(0, 160); }
      out.push(info);
    }
  }
  console.log('DOM ' + (label || '') + ' ' + JSON.stringify(out));
  await browser.close();
})().catch((e) => { console.log('DOM ' + (process.argv[3] || '') + ' failed: ' + String(e).slice(0, 200)); });
