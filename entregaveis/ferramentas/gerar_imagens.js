// Gera, a partir do protótipo, as imagens e as posições usadas na montagem da tela no Aspen Process Graphics Editor.
// Uso: node gerar_imagens.js   (requer Playwright com Chromium). Depois rode montar_pacote_aspen.py.
// A tela do Aspen mede 50 x 35 unidades (proporção 10:7, a mesma da tela PQ US04). O protótipo é desenhado
// em 1600 x 1120 (32 px por unidade) e as imagens saem em 1400 x 980 (28 px por unidade), perto do tamanho
// em que a tela aparece no monitor, para o Aspen quase não precisar redimensionar a imagem.
// Saídas em ../03_aspen/:
//   objetos_tela.json          painéis, linhas, textos, símbolos e campos, para gerar o arquivo .atgraphic
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

  // Parte fixa da tela: painéis, linhas, textos e símbolos, para virar objetos nativos do Aspen.
  await abrir('modo=fundo');
  const estatico = await page.evaluate(() => {
    const palco = document.getElementById('palco');
    const P = palco.getBoundingClientRect();
    const box = r => ({ x: r.left - P.left, y: r.top - P.top, w: r.width, h: r.height });
    const visivel = c => c && c !== 'transparent' && !/rgba\(.*,\s*0\)$/.test(c);
    const fora = el => el.closest('.slot') || el.closest('svg') || el.closest('#cenarios');
    const fundoDe = el => { for (let e = el; e; e = e.parentElement) { const c = getComputedStyle(e).backgroundColor; if (visivel(c)) return c; } return getComputedStyle(palco).backgroundColor; };
    const formas = [], textos = [], simbolos = [];
    for (const el of palco.querySelectorAll('*')) {
      if (fora(el)) continue;
      const cs = getComputedStyle(el), r = box(el.getBoundingClientRect());
      if (el.classList.contains('simb')) { simbolos.push({ ...r, forma: [...el.classList].find(c => c !== 'simb'), cor: cs.color, fundo: fundoDe(el.parentElement) }); continue; }
      const bw = ['Top', 'Right', 'Bottom', 'Left'].map(l => parseFloat(cs['border' + l + 'Width']) || 0);
      const raio = Math.min(parseFloat(cs.borderTopLeftRadius) || 0, r.h / 2, r.w / 2);
      const temFundo = visivel(cs.backgroundColor);
      if (bw.every(v => v > 0) || temFundo) {
        formas.push({ ...r, raio, preenchimento: temFundo ? cs.backgroundColor : null, linha: bw.every(v => v > 0) ? cs.borderTopColor : null, fundo: fundoDe(el.parentElement), sobre: !!el.closest('.sobre') });
      }
      if (bw[0] > 0 && !bw.every(v => v > 0)) formas.push({ x: r.x, y: r.y, w: r.w, h: bw[0], raio: 0, preenchimento: cs.borderTopColor, linha: null });
      if (bw[2] > 0 && !bw.every(v => v > 0)) formas.push({ x: r.x, y: r.y + r.h - bw[2], w: r.w, h: bw[2], raio: 0, preenchimento: cs.borderBottomColor, linha: null });
    }
    const it = document.createTreeWalker(palco, NodeFilter.SHOW_TEXT);
    for (let n = it.nextNode(); n; n = it.nextNode()) {
      const el = n.parentElement;
      if (fora(el) || el.classList.contains('logo')) continue;
      const t = n.textContent.replace(/\u00a0/g, ' ');
      const i = t.search(/\S/); if (i < 0) continue;
      const j = t.length - t.split('').reverse().join('').search(/\S/);
      const rg = document.createRange(); rg.setStart(n, i); rg.setEnd(n, j);
      const linhas = rg.getClientRects().length;
      const cs = getComputedStyle(el);
      textos.push({ ...box(rg.getBoundingClientRect()), texto: t.slice(i, j).replace(/\s+/g, ' '), linhas, tamanho: parseFloat(cs.fontSize), peso: parseInt(cs.fontWeight), cor: cs.color, fundo: fundoDe(el) });
    }
    const campos = [...palco.querySelectorAll('[data-slot]')].map(el => {
      const cs = getComputedStyle(el);
      return { n: +el.dataset.slot, ...box(el.getBoundingClientRect()), tamanho: parseFloat(cs.fontSize), peso: parseInt(cs.fontWeight), alinhamento: cs.textAlign, fundo: fundoDe(el.parentElement) };
    });
    return { largura: P.width, altura: P.height, fundoTela: getComputedStyle(palco).backgroundColor, formas, textos, simbolos, campos };
  });
  fs.writeFileSync(path.join(saida, 'objetos_tela.json'), JSON.stringify(estatico, null, 1));
  const quebrados = estatico.textos.filter(t => t.linhas > 1);
  if (quebrados.length) console.warn('Textos com quebra de linha:', quebrados.map(t => t.texto));

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
