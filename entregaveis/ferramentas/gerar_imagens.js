// Gera, a partir do protótipo, as imagens e as posições usadas na montagem da tela no Aspen Process Graphics Editor.
// Uso: node gerar_imagens.js   (requer Playwright com Chromium). Depois rode montar_pacote_aspen.py.
// Saídas em ../03_aspen/:
//   fundo_1920x1080.png        imagem de fundo (somente a parte fixa)
//   mapa_posicionamento.png    fundo com a posição numerada de cada objeto dinâmico
//   tela_exemplo_*.png         tela completa com dados simulados, por cenário
//   posicoes_px.json           posição e tamanho de cada objeto dinâmico, em pixels
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const html = 'file://' + path.resolve(__dirname, '../02_prototipo/index.html');
const saida = path.resolve(__dirname, '../03_aspen');
fs.mkdirSync(saida, { recursive: true });

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
  const abrir = async (q) => { await page.goto(`${html}?captura=1&${q}`); await page.evaluate(() => document.fonts.ready); };

  await abrir('modo=fundo');
  await page.screenshot({ path: path.join(saida, 'fundo_1920x1080.png') });

  await abrir('modo=mapa');
  await page.screenshot({ path: path.join(saida, 'mapa_posicionamento.png') });
  const slots = await page.$$eval('[data-slot]', els => els.map(e => {
    const r = e.getBoundingClientRect();
    return { n: +e.dataset.slot, x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) };
  }));
  slots.sort((a, b) => a.n - b.n);
  fs.writeFileSync(path.join(saida, 'posicoes_px.json'), JSON.stringify(slots, null, 1));

  for (const c of ['normal', 'f200', 'bad', 'manual']) {
    await abrir(`modo=demo&cenario=${c}`);
    await page.screenshot({ path: path.join(saida, `tela_exemplo_${c}.png`) });
  }
  await browser.close();
  console.log(`${slots.length} objetos dinâmicos mapeados em ${saida}`);
})();
