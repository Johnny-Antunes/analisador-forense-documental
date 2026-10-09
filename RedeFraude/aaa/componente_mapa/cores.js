/*
 * Paletas do mapa em OKLab: interpolar em OKLab mantém a luminosidade
 * crescendo de forma regular sobre o fundo escuro (em sRGB os tons do meio
 * ficam acinzentados). Conversões adaptadas de open-apuracao-brazil
 * (src/data/mocks.js) — Copyright (c) 2026 Bruno Pinheiro, Licença MIT.
 */

function paraLab(hex) {
  const [r, g, b] = hex.match(/[a-f\d]{2}/gi).map(n => parseInt(n, 16) / 255)
    .map(v => v <= .04045 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4);
  const l = Math.cbrt(.4122214708 * r + .5363325363 * g + .0514459929 * b);
  const m = Math.cbrt(.2119034982 * r + .6806995451 * g + .1073969566 * b);
  const s = Math.cbrt(.0883024619 * r + .2817188376 * g + .6299787005 * b);
  return [.2104542553 * l + .793617785 * m - .0040720468 * s, 1.9779984951 * l - 2.428592205 * m + .4505937099 * s, .0259040371 * l + .7827717662 * m - .808675766 * s];
}
function deLab([L, a, b]) {
  const l = (L + .3963377774 * a + .2158037573 * b) ** 3;
  const m = (L - .1055613458 * a - .0638541728 * b) ** 3;
  const s = (L - .0894841775 * a - 1.291485548 * b) ** 3;
  return '#' + [4.0767416621 * l - 3.3077115913 * m + .2309699292 * s, -1.2684380046 * l + 2.6097574011 * m - .3413193965 * s, -.0041960863 * l - .7034186147 * m + 1.707614701 * s]
    .map(v => Math.round(Math.max(0, Math.min(1, v <= .0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - .055)) * 255).toString(16).padStart(2, '0')).join('');
}
const rampa = (de, ate, passos) => {
  const a = paraLab(de), b = paraLab(ate);
  return passos.map(t => deLab(a.map((v, i) => v + (b[i] - v) * t)));
};

export const FUNDO = '#121214';
export const BASE = '#1d1d21';        // município com dados, sem destaque
export const VAZIO = '#17171a';       // sem acionamentos no mês
export const TINTA_2 = '#A1A1AA';
export const COR = { ambar: '#C9A66B', laranja: '#FF7300', carmim: '#C0625F', foco: '#FFFFFF' };

// Severidade do alerta: 1 = um alerta; 2 = explosão ou dois alertas; 3 = explosão + outro (ou 3+).
export const SEVERIDADE = [BASE, COR.ambar, COR.laranja, COR.carmim];
export function severidade(bits) {
  if (!bits) return 0;
  let n = 0;
  for (let b = bits; b; b >>= 1) n += b & 1;
  const pontos = n + (bits & 1 ? 1 : 0);   // explosão de volume pesa dobrado
  return Math.min(3, pontos);
}

export const RAMPA_VOLUME = rampa(BASE, COR.carmim, [.3, .5, .7, .85, 1]);
export const RAMPA_DESVIO = rampa(BASE, COR.laranja, [.3, .55, .8, 1]);
export const RAMPA_TAXA = rampa(BASE, COR.ambar, [.3, .5, .7, .85, 1]);
export const LIMITES_DESVIO = [1.5, 3, 6];   // × média histórica

export const passo = (valor, limites) => {
  const i = limites.findIndex(l => valor < l);
  return i < 0 ? limites.length : i;
};

/** Faixa i = valores <= limite i (rampas por quantis: "até X"). */
export const passoAte = (valor, limites) => {
  const i = limites.findIndex(l => valor <= l);
  return i < 0 ? limites.length : i;
};

/** Cortes por quantis (sem repetir valores) para rampas de 5 faixas. */
export function quantis(valores, n = 5) {
  const v = valores.filter(x => x > 0).sort((a, b) => a - b);
  if (!v.length) return [];
  const cortes = [];
  for (let i = 1; i < n; i++) {
    const c = v[Math.min(v.length - 1, Math.floor(v.length * i / n) - 1)];
    if ((!cortes.length || c > cortes[cortes.length - 1]) && c < v[v.length - 1]) cortes.push(c);
  }
  return cortes;
}

/** Cor de texto legível sobre um preenchimento #rrggbb. */
export function tintaSobre(hex) {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map(v => v <= .04045 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4);
  return .2126 * r + .7152 * g + .0722 * b > .3 ? '#141821' : '#ffffff';
}
