// 슬라이드별 미리보기 그림을 찍는다.  node snap.js <deck.html> <outDir> <slides> <width> <height>
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

(async () => {
  const [file, outDir, n, w, h] = process.argv.slice(2);
  const html = fs.readFileSync(file, 'utf8');
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: +w, height: +h } });
  for (let i = 1; i <= +n; i++) {
    const boot = `<script>window.__LIVE_SLIDE__=${i};</script>`;
    const m = /<head[^>]*>/i.exec(html);
    const doc = m ? html.slice(0, m.index + m[0].length) + boot + html.slice(m.index + m[0].length) : boot + html;
    await page.setContent(doc, { waitUntil: 'networkidle' }).catch(() => page.setContent(doc));
    await page.waitForTimeout(700);
    await page.screenshot({ path: path.join(outDir, `slide${i}.png`) });
  }
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
