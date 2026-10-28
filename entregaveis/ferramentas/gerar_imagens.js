// Gera, a partir do protótipo, as imagens e as posições usadas na montagem da tela no Aspen Process Graphics Editor.
// Uso: node gerar_imagens.js   (requer Playwright com Chromium). Depois rode montar_pacote_aspen.py.
// A tela do Aspen mede 50 x 35 unidades (proporção 10:7, a mesma da tela PQ US04). O protótipo é desenhado
// em 1600 x 1120 (32 px por unidade) e as imagens saem em 1400 x 980 (28 px por unidade), perto do tamanho
// em que a tela aparece no monitor, para o Aspen quase não precisar redimensionar a imagem.
// Saídas em ../03_aspen/:
//   fundo_1400x980.png         imagem de fundo (somente a parte fixa)
//   mapa_posicionamento.png    fundo com a posição numerada de cada objeto dinâmico
//   tela_exemplo_*.png         tela completa com dados simulados, por cenário
//   posicoes.json              posição e tamanho de cada objeto dinâmico, em unidades do Aspen
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const html = 'file://' + path.resolve(__dirname, '../02_prototipo/index.html');
const saida = path.resolve(__dirname, '../03_aspen');
fs.mkdirSync(saida, { recursive: true });

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 1120 }, deviceScaleFactor: 1400 / 1600 });
  const abrir = async (q) => { await page.goto(`${html}?captura=1&${q}`); await page.evaluate(() => document.fonts.ready); };

  await abrir('modo=fundo');
  await page.screenshot({ path: path.join(saida, 'fundo_1400x980.png') });

  await abrir('modo=mapa');
  await page.screenshot({ path: path.join(saida, 'mapa_posicionamento.png') });
  const slots = await page.$$eval('[data-slot]', els => els.map(e => {
    const r = e.getBoundingClientRect();
    const u = v => Math.round(v / 32 * 100) / 100;  // 32 px do protótipo = 1 unidade do Aspen
    return { n: +e.dataset.slot, x: u(r.left), y: u(r.top), w: u(r.width), h: u(r.height) };
  }));
  slots.sort((a, b) => a.n - b.n);
  fs.writeFileSync(path.join(saida, 'posicoes.json'), JSON.stringify(slots, null, 1));

  for (const c of ['normal', 'f200', 'bad', 'manual']) {
    await abrir(`modo=demo&cenario=${c}`);
    await page.screenshot({ path: path.join(saida, `tela_exemplo_${c}.png`) });
  }
  await browser.close();
  console.log(`${slots.length} objetos dinâmicos mapeados em ${saida}`);
})();
