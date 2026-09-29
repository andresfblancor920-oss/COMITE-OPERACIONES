// Genera Politica_Devolucion_KAIR.pdf a partir de politica_devolucion_kair.html
// Uso: NODE_PATH=$(npm root -g) node generar_pdf.js
const path = require('path');
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + path.join(__dirname, 'politica_devolucion_kair.html'), { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);
  await page.pdf({ path: path.join(__dirname, 'Politica_Devolucion_KAIR.pdf'), format: 'Letter', printBackground: true, preferCSSPageSize: true });
  await browser.close();
  console.log('PDF generado: Politica_Devolucion_KAIR.pdf');
})();
