/*
 * Mapa territorial do LCFO — Canvas 2D puro, sem dependências externas.
 *
 * Modos (payload.modo):
 *   "radar": Centro de Comando. Linha do tempo mensal, métricas Alerta/Volume/Desvio/Taxa,
 *            unidades Estados/Municípios/Bolhas, drill Brasil -> UF -> município.
 *   "caso":  mesa de um caso. Volume do caso por município, câmera enquadrando o caso.
 *
 * Ponte com o Python (mesmo protocolo do componente do grafo):
 *   iframe -> pai: streamlit:componentReady; streamlit:setComponentValue {tipo: "mapa_estado"|"mapa_pronto", ...}
 *   pai -> iframe: streamlit:render_url {url} (payload por arquivo); lcfo:mapa_estado {uf, municipio, mes};
 *                  lcfo:layout {topo, base, esq, dir} (área coberta por HUD/painéis do app: o canvas
 *                  ocupa a tela toda, mas a câmera e os controles respeitam essa área)
 * Mudanças vindas do Python não geram mensagem de volta (sem eco).
 *
 * Desempenho (medido em Chromium sem GPU, 5.570 municípios): o custo estava em rasterizar ~5.570
 * preenchimentos por quadro, e o quadro inteiro era refeito a cada movimento do mouse. Agora:
 *   - os municípios são agrupados por cor em poucos Path2D (um preenchimento por faixa);
 *   - o desenho é agendado por requestAnimationFrame (no máx. um por quadro);
 *   - hover/foco/contorno de UF vivem numa segunda camada (canvas #sobre), que é barata;
 *   - a dica só é refeita quando o município sob o cursor muda.
 * `?perf=1` mostra o tempo de desenho no canto e expõe window.__lcfoMapa para medição.
 *
 * Interação de câmera/desenho adaptada de open-apuracao-brazil (src/map/ElectionMap.js)
 * Copyright (c) 2026 Bruno Pinheiro — Licença MIT.
 */

import { criarGeografia, cameraPara } from './geografia.js';
import * as C from './cores.js';

const MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const NOME_UF = {
  AC: 'Acre', AL: 'Alagoas', AP: 'Amapá', AM: 'Amazonas', BA: 'Bahia', CE: 'Ceará', DF: 'Distrito Federal',
  ES: 'Espírito Santo', GO: 'Goiás', MA: 'Maranhão', MT: 'Mato Grosso', MS: 'Mato Grosso do Sul', MG: 'Minas Gerais',
  PA: 'Pará', PB: 'Paraíba', PR: 'Paraná', PE: 'Pernambuco', PI: 'Piauí', RJ: 'Rio de Janeiro', RN: 'Rio Grande do Norte',
  RS: 'Rio Grande do Sul', RO: 'Rondônia', RR: 'Roraima', SC: 'Santa Catarina', SP: 'São Paulo', SE: 'Sergipe', TO: 'Tocantins',
};
const UNIDADES = { estados: 'Estados', municipios: 'Municípios', bolhas: 'Bolhas' };
const MARCADORES = { circulos: 'Círculos', area: 'Só área' };   // mapa do caso
const METRICAS = { alerta: 'Alerta', volume: 'Volume', desvio: 'Desvio', taxa: 'Taxa /10 mil' };
const CAMERA_MS = 480, ARRASTE_MIN = 5, PASSO_ZOOM = 1.55, PLAY_MS = 1100;
const BOLHA_MAIOR = .045;      // maior bolha = 4,5% da largura do Brasil
const MAX_ROTULOS_CASO = 14, MAX_ROTULOS_UF = 10;
const LARGURA_LEGENDA_COMPACTA = 720;   // abaixo disso a legenda nasce recolhida
const PERF = new URLSearchParams(location.search).get('perf') === '1';

const fmtMes = m => (!m || m === 'caso') ? 'Total' : `${MESES[+m.slice(5, 7) - 1]}/${m.slice(0, 4)}`;
const fmtInt = n => n == null ? '—' : Math.round(n).toLocaleString('pt-BR');
const fmtPct = n => `${(n * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%`;
const fmtX = n => n == null ? '—' : `${n.toLocaleString('pt-BR', { maximumFractionDigits: 1 })}×`;
const fmtDesvio = (r, media) => !media ? 'sem histórico' : fmtX(r);
const semAcento = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
const $ = id => document.getElementById(id);

const S = {
  geo: null, payload: null, urlPendente: null,
  mes: 0, uf: null, municipio: null, unidade: 'municipios', metrica: 'alerta',
  camera: null, alvo: null, quadro: 0, hover: null, tam: { w: 600, h: 600 },
  toques: new Map(), inter: {}, play: null, serieUF: {}, cache: null,
  layout: { topo: 0, base: 0, esq: 0, dir: 0 }, alturaTempo: 0, camManual: false,
  legendaRecolhida: false, legendaManual: false, dicaId: null,
  rafPendente: false, sujo: false, sujoSobre: false, sujoUI: false,
  perf: { base: 0, sobre: 0, quadros: 0 },
};

/* ------------------------------------------------------------------ ponte */
const enviar = valor => parent.postMessage({ isStreamlitMessage: true, type: 'streamlit:setComponentValue', value: valor, dataType: 'json' }, '*');
const emitirEstado = () => enviar({ tipo: 'mapa_estado', uf: S.uf, municipio: S.municipio?.id || null, mes: S.mes, metrica: S.metrica, unidade: S.unidade });

window.addEventListener('message', ev => {
  const d = ev.data || {};
  if (d.type === 'streamlit:render_url') carregarPayload(d.url);
  else if (d.type === 'lcfo:mapa_estado') aplicarEstado(d, false);
  else if (d.type === 'lcfo:layout') aplicarLayout(d);
});

async function carregarPayload(url) {
  if (!S.geo) { S.urlPendente = url; return; }
  const r = await fetch(url);
  if (!r.ok) { mostrarErro(`Não foi possível carregar os dados do mapa (HTTP ${r.status}).`); return; }
  definirPayload(await r.json());
}

function definirPayload(p) {
  const mantemMes = S.payload && S.payload.meses.length === p.meses.length && S.payload.modo === p.modo;
  S.payload = p;
  if (!mantemMes) S.mes = p.meses.length - 1;
  if (p.modo === 'caso') { S.unidade = 'municipios'; if (!['volume', 'taxa'].includes(S.metrica)) S.metrica = 'volume'; }
  if (p.layout) definirLayout(p.layout);
  // Série mensal por UF (para o desvio do estado e os totais).
  S.serieUF = {};
  for (const [cod, d] of Object.entries(p.municipios)) {
    const m = S.geo.porId.get(cod); if (!m) continue;
    const s = S.serieUF[m.uf] ||= Array(p.meses.length).fill(0);
    d.v.forEach((v, i) => { s[i] += v || 0; });
  }
  S.cache = null;
  montarControles();
  irPara(true);
  atualizarUI(true);
  S.sujo = true; desenharTudo();
  enviar({ tipo: 'mapa_pronto' });
}

/** Estado vindo do Python ou de um clique. */
function aplicarEstado({ uf = S.uf, municipio = S.municipio?.id || null, mes = S.mes, metrica, unidade }, emitir) {
  const novoMun = municipio ? S.geo.porId.get(String(municipio)) || null : null;
  const novaUf = novoMun ? novoMun.uf : (uf || null);
  const mudouLugar = novaUf !== S.uf || (novoMun?.id || null) !== (S.municipio?.id || null);
  const mudouMes = mes != null && S.payload && mes !== S.mes && mes >= 0 && mes < S.payload.meses.length;
  S.uf = novaUf; S.municipio = novoMun;
  if (mudouMes) { S.mes = mes; S.cache = null; }
  if (metrica && METRICAS[metrica] && metrica !== S.metrica) { S.metrica = metrica; S.cache = null; }
  if (unidade && UNIDADES[unidade] && unidade !== S.unidade) { S.unidade = unidade; S.cache = null; }
  if (mudouLugar) { S.camManual = false; irPara(false); S.cache = null; }
  agendar(true); atualizarUI();
  if (emitir) emitirEstado();
}

/* ------------------------------------------------------------------ layout (área coberta pelo app) */
function definirLayout(l) {
  S.layout = { topo: +l.topo || 0, base: +l.base || 0, esq: +l.esq || 0, dir: +l.dir || 0 };
  const q = $('quadro').style;
  q.setProperty('--seguro-topo', S.layout.topo + 'px');
  q.setProperty('--seguro-base', S.layout.base + 'px');
  q.setProperty('--seguro-esq', S.layout.esq + 'px');
  q.setProperty('--seguro-dir', S.layout.dir + 'px');
  agendarAjusteControles();
}

/** Linha do tempo estreita (tela dividida): os rótulos dos meses se atropelavam ("setoutnovdez...").
 *  Mantém um a cada N (mais o último) com ~30 px por rótulo; visibility preserva o espaçamento. */
function rarearMarcasDoTempo() {
  const marcas = $('marcas'), n = marcas?.children.length || 0;
  if (!n) return;
  const passo = Math.max(1, Math.ceil(n * 30 / Math.max(1, marcas.clientWidth)));
  [...marcas.children].forEach((s, i) => { s.style.visibility = (i % passo === 0 || i === n - 1) ? '' : 'hidden'; });
}

/** No máx. uma medição por quadro, depois de tudo o que mexe no topo (layout, migalhas, controles). */
function agendarAjusteControles() {
  if (S.ajustePendente) return;
  S.ajustePendente = true;
  requestAnimationFrame(() => { S.ajustePendente = false; ajustarControles(); });
}

/** Faixa livre estreita (painéis do app abertos dos dois lados): os botões de modo, que ficam no topo
 *  à direita, colidiam com as migalhas/busca à esquerda. Nesse caso descem para uma segunda linha. */
function ajustarControles() {
  rarearMarcasDoTempo();
  const quadroEl = $('quadro'), topo = document.querySelector('.topo'), modos = document.querySelector('.modos');
  if (!topo || !modos) return;
  quadroEl.classList.remove('modos-abaixo');
  const a = topo.getBoundingClientRect(), b = modos.getBoundingClientRect();
  if (a.right + 8 > b.left) quadroEl.classList.add('modos-abaixo');
}

function aplicarLayout(l) {
  definirLayout(l);
  medirLinhaTempo();
  if (S.geo && S.payload) {
    if (!S.camManual) irPara(false);
    agendar(true);
  }
}

/** Região livre do canvas: abaixo do HUD, ao lado dos painéis e acima da linha do tempo. */
function areaUtil() {
  const L = S.layout, W = S.tam.w, H = S.tam.h;
  const x = L.esq, y = L.topo;
  const tempo = S.alturaTempo ? S.alturaTempo + 16 : 0;      // linha do tempo + a margem dela até a borda
  return { x, y, w: Math.max(80, W - L.esq - L.dir), h: Math.max(80, H - L.topo - L.base - tempo) };
}

function medirLinhaTempo() {
  const lt = $('linha-tempo');
  S.alturaTempo = lt.hidden ? 0 : lt.offsetHeight + 8;
  $('quadro').style.setProperty('--altura-tempo', S.alturaTempo + 'px');
}

/* ------------------------------------------------------------------ dados do mês */
const enfase = m => S.uf && m.uf !== S.uf ? .16 : S.municipio && m.id !== S.municipio.id ? .55 : 1;

function dadosDoMes() {
  if (S.cache) return S.cache;
  const p = S.payload, i = S.mes, geo = S.geo;
  const mun = new Map(), uf = {};
  for (const [cod, d] of Object.entries(p.municipios)) {
    const g = geo.porId.get(cod); if (!g) continue;
    const reg = { id: cod, nome: g.nome, uf: g.uf, pop: g.populacao, v: d.v[i] || 0, n: d.n ? (d.n[i] || 0) : 0, bits: d.a[i] || 0, media: d.m[i], razao: d.r[i] };
    reg.taxa = reg.pop ? reg.v / reg.pop * 1e4 : 0;
    mun.set(cod, reg);
    const e = uf[g.uf] ||= { id: g.uf, nome: NOME_UF[g.uf] || g.uf, uf: g.uf, v: 0, nAlerta: 0, bits: 0, pop: geo.ufs[g.uf]?.populacao || 0 };
    e.v += reg.v;
    if (reg.bits) { e.nAlerta++; e.bits |= reg.bits; }
  }
  for (const e of Object.values(uf)) {
    const s = S.serieUF[e.uf] || [];
    const hist = s.slice(0, i).filter(v => v > 0);
    e.media = hist.length ? hist.reduce((a, b) => a + b, 0) / hist.length : null;
    e.razao = e.media ? e.v / e.media : null;
    e.taxa = e.pop ? e.v / e.pop * 1e4 : 0;
  }
  const porMun = S.uf || S.unidade !== 'estados';
  const lista = porMun ? [...mun.values()].filter(m => !S.uf || m.uf === S.uf) : Object.values(uf);
  const cortes = S.metrica === 'volume' ? C.quantis(lista.map(x => x.v)) : S.metrica === 'taxa' ? C.quantis(lista.map(x => x.taxa)) : null;
  S.cache = { mun, uf, porMun, lista, cortes, baldes: null, legenda: null, emAlerta: null, bolhas: null };
  return S.cache;
}

/** Cor e índice de faixa (para a legenda) de um município ou UF no mês. */
function faixa(reg, ehUf) {
  if (!reg || (!reg.v && !reg.bits)) return { cor: C.VAZIO, f: -1 };
  const { cortes } = dadosDoMes();
  switch (S.metrica) {
    case 'alerta': {
      if (ehUf) { const f = C.passo(reg.nAlerta, [1, 3, 6]); return { cor: C.SEVERIDADE[f], f }; }
      const f = C.severidade(reg.bits); return { cor: C.SEVERIDADE[f], f };
    }
    // Sem cortes (um único valor, ex.: caso num município só): cor plena, faixa única.
    case 'volume': { if (!cortes.length) return { cor: C.RAMPA_VOLUME.at(-1), f: 0 }; const f = C.passoAte(reg.v, cortes); return { cor: C.RAMPA_VOLUME[f], f }; }
    case 'taxa': { if (!cortes.length) return { cor: C.RAMPA_TAXA.at(-1), f: 0 }; const f = C.passoAte(reg.taxa, cortes); return { cor: C.RAMPA_TAXA[f], f }; }
    case 'desvio': {
      if (reg.razao == null || !reg.media) return { cor: C.BASE, f: -2 };
      const f = C.passo(reg.razao, C.LIMITES_DESVIO); return { cor: C.RAMPA_DESVIO[f], f };
    }
  }
  return { cor: C.BASE, f: 0 };
}

/**
 * Municípios agrupados por UF e, dentro dela, por (cor, ênfase): poucos Path2D por UF = poucos
 * preenchimentos. O agrupamento por UF permite pular UFs fora da tela (um Path2D único do Brasil
 * inteiro obrigava o navegador a recortar toda a geometria a cada quadro de pan/zoom).
 */
function baldesDe(dm) {
  if (dm.baldes) return dm.baldes;
  const porUf = new Map();
  for (const m of S.geo.municipios) {
    let u = porUf.get(m.uf);
    if (!u) porUf.set(m.uf, u = { box: [Infinity, Infinity, -Infinity, -Infinity], grupos: new Map() });
    u.box[0] = Math.min(u.box[0], m.box[0]); u.box[1] = Math.min(u.box[1], m.box[1]);
    u.box[2] = Math.max(u.box[2], m.box[2]); u.box[3] = Math.max(u.box[3], m.box[3]);   // inclui Noronha (a caixa de UF da câmera não)
    const cor = faixa(dm.mun.get(m.id)).cor, alpha = enfase(m), chave = cor + alpha;
    let g = u.grupos.get(chave);
    if (!g) u.grupos.set(chave, g = { cor, alpha, path: new Path2D() });
    g.path.addPath(m.path);
  }
  return (dm.baldes = [...porUf.values()].map(u => ({ box: u.box, grupos: [...u.grupos.values()] })));
}

function emAlertaDe(dm) {
  return dm.emAlerta ||= [...dm.mun.values()].filter(r => r.bits).sort((a, b) => C.severidade(a.bits) - C.severidade(b.bits));
}

function faixasDaLegenda() {
  const dm = dadosDoMes();
  if (dm.legenda) return dm.legenda;
  const { lista, cortes, porMun } = dm;
  const fmt = S.metrica === 'taxa' ? (v => v.toLocaleString('pt-BR', { maximumFractionDigits: 1 })) : fmtInt;
  let def;
  if (S.metrica === 'alerta') def = porMun
    ? [['sem alerta', C.BASE], ['1 alerta', C.COR.ambar], ['explosão ou 2 alertas', C.COR.laranja], ['explosão + outro', C.COR.carmim]]
    : [['nenhum município em alerta', C.BASE], ['1–2 em alerta', C.COR.ambar], ['3–5 em alerta', C.COR.laranja], ['6+ em alerta', C.COR.carmim]];
  else if (S.metrica === 'desvio') def = [['< 1,5× a média', C.RAMPA_DESVIO[0]], ['1,5–3×', C.RAMPA_DESVIO[1]], ['3–6×', C.RAMPA_DESVIO[2]], ['6× ou mais', C.RAMPA_DESVIO[3]]];
  else {
    const rampa = S.metrica === 'volume' ? C.RAMPA_VOLUME : C.RAMPA_TAXA;
    const cs = cortes || [];
    def = cs.length
      ? [...cs.map(c => `até ${fmt(c)}`), `mais de ${fmt(cs[cs.length - 1])}`].map((rotulo, i) => [rotulo, rampa[i]])
      : [['com acionamentos', rampa.at(-1)]];
  }
  const tot = lista.reduce((a, x) => a + x.v, 0) || 1;
  const out = def.map(([rotulo, cor]) => ({ rotulo, cor, n: 0, v: 0 }));
  for (const x of lista) {
    const { f } = faixa(x, !porMun);
    if (f >= 0 && out[f]) { out[f].n++; out[f].v += x.v; }
  }
  return (dm.legenda = out.map(b => ({ ...b, parcela: b.v / tot })));
}

/* ------------------------------------------------------------------ câmera */
function irPara(instantaneo) {
  S.camManual = false;
  const caixa = S.payload?.modo === 'caso' && !S.uf && !S.municipio ? caixaDoCaso() : null;
  S.alvo = cameraPara(S.geo, { uf: S.uf, municipio: S.municipio, caixa }, S.tam.w, S.tam.h, areaUtil());
  if (instantaneo || !S.camera) { S.camera = { ...S.alvo }; agendar(true); return; }
  animar();
}

function caixaDoCaso() {
  const box = [Infinity, Infinity, -Infinity, -Infinity];
  for (const cod of Object.keys(S.payload.municipios)) {
    const m = S.geo.porId.get(cod); if (!m) continue;
    box[0] = Math.min(box[0], m.box[0]); box[1] = Math.min(box[1], m.box[1]);
    box[2] = Math.max(box[2], m.box[2]); box[3] = Math.max(box[3], m.box[3]);
  }
  return isFinite(box[0]) ? box : null;
}

function animar() {
  cancelAnimationFrame(S.quadro);
  limparHover();
  const de = { ...S.camera }, para = { ...S.alvo }, t0 = performance.now();
  const reduzido = matchMedia('(prefers-reduced-motion: reduce)').matches;
  // Interpola o centro (coordenadas do mapa) e o zoom em escala geométrica: num voo
  // Brasil -> município (zoom ×100) a interpolação linear de k/x/y "despenca" no fim.
  const { w, h } = S.tam;
  const c0 = [(w / 2 - de.x) / de.k, (h / 2 - de.y) / de.k], c1 = [(w / 2 - para.x) / para.k, (h / 2 - para.y) / para.k];
  const tick = agora => {
    // max(0, ...): o 1º quadro pode chegar com carimbo de tempo anterior a t0; t < 0
    // fazia a curva "voltar" e o zoom ficar negativo (arc() com raio negativo).
    const t = reduzido ? 1 : Math.max(0, Math.min(1, (agora - t0) / CAMERA_MS)), e = 1 - (1 - t) ** 4;
    const k = de.k * (para.k / de.k) ** e, cx = c0[0] + (c1[0] - c0[0]) * e, cy = c0[1] + (c1[1] - c0[1]) * e;
    S.camera = t >= 1 ? { ...para } : { k, x: w / 2 - cx * k, y: h / 2 - cy * k };
    S.sujo = true; desenharTudo();
    if (t < 1) S.quadro = requestAnimationFrame(tick);
  };
  S.quadro = requestAnimationFrame(tick);
}

function zoom(fator, ancora = [S.tam.w / 2, S.tam.h / 2], suave = true) {
  if (!pronto()) return;
  S.camManual = true;
  limparHover();
  const v = S.camera, base = cameraPara(S.geo, { uf: S.uf, municipio: S.municipio }, S.tam.w, S.tam.h, areaUtil());
  const k = Math.max(base.k * .5, Math.min(base.k * 40, v.k * fator));
  S.alvo = { k, x: ancora[0] - (ancora[0] - v.x) * (k / v.k), y: ancora[1] - (ancora[1] - v.y) * (k / v.k) };
  if (suave) animar(); else moverCamera({ ...S.alvo });
}

/* ------------------------------------------------------------------ desenho */
const canvas = $('mapa'), sobre = $('sobre');

/**
 * Pan/zoom por gesto: em vez de rasterizar o Brasil inteiro a cada quadro, a imagem já desenhada é
 * deslocada/escalada por CSS (só composição) e o desenho nítido é refeito quando o gesto para.
 * Se a câmera se afastar demais da imagem pronta (zoom > ~2×), redesenha de verdade.
 */
function moverCamera(nova) {
  // Gesto manual (roda, arraste, pinça) manda: cancela um voo animado em andamento, que senão
  // continuava sobrescrevendo a câmera a cada quadro e o mapa "pulava".
  cancelAnimationFrame(S.quadro);
  S.camera = nova; S.camManual = true; limparHover();
  const b = S.camBase;
  if (!b) { agendar(true); return; }
  const f = nova.k / b.k;
  if (f > 2.2 || f < .45) { agendar(true); return; }
  canvas.style.transformOrigin = '0 0';
  canvas.style.transform = `translate(${nova.x - b.x * f}px, ${nova.y - b.y * f}px) scale(${f})`;
  clearTimeout(S.tReraster);
  S.tReraster = setTimeout(rasterizarAoParar, 130);
}
/** Redesenho nítido depois do gesto — nunca com o botão ainda pressionado: uma pausa curta no meio do
 *  arraste disparava o desenho completo do Brasil e o quadro pesado caía bem quando o mouse voltava a andar. */
function rasterizarAoParar() {
  if (S.toques.size) { S.tReraster = setTimeout(rasterizarAoParar, 130); return; }
  agendar(true);
}
const ctxTeste = document.createElement('canvas').getContext('2d');

/** Ponto para o marcador: o centroide, ou — em municípios côncavos, onde ele cai
 *  fora — o ponto de uma grade interna mais próximo dele. */
function ancora(m) {
  if (m.ancora) return m.ancora;
  if (ctxTeste.isPointInPath(m.path, m.centro[0], m.centro[1])) return (m.ancora = m.centro);
  let melhor = m.centro, dist = Infinity;
  const [x0, y0, x1, y1] = m.box, N = 14;
  for (let i = 1; i < N; i++) for (let j = 1; j < N; j++) {
    const x = x0 + (x1 - x0) * i / N, y = y0 + (y1 - y0) * j / N;
    const d = Math.hypot(x - m.centro[0], y - m.centro[1]);
    if (d < dist && ctxTeste.isPointInPath(m.path, x, y)) { dist = d; melhor = [x, y]; }
  }
  return (m.ancora = melhor);
}

/** Agenda o redesenho para o próximo quadro (no máx. um por quadro, qualquer que seja o nº de eventos). */
function agendar(base = false) {
  if (base) S.sujo = true; else S.sujoSobre = true;
  if (S.rafPendente) return;
  S.rafPendente = true;
  requestAnimationFrame(() => {
    S.rafPendente = false;
    if (S.sujoUI) { S.sujoUI = false; atualizarUI(true); }
    if (S.sujo || S.sujoSobre) desenharTudo();
  });
}

function prepararCanvas() {
  const dpr = Math.min(devicePixelRatio || 1, 2), { w, h } = S.tam;
  if (canvas.width !== Math.round(w * dpr) || canvas.height !== Math.round(h * dpr)) { canvas.width = Math.round(w * dpr); canvas.height = Math.round(h * dpr); }
  return dpr;
}

function desenharTudo() {
  if (!S.geo || !S.camera || !S.payload) return;
  const t0 = performance.now();
  if (S.sujo || !S.baseDesenhada) { desenharBase(); S.sujo = false; S.baseDesenhada = true; if (S.hover) S.sujoSobre = true; }
  const t1 = performance.now();
  if (S.sujoSobre) { desenharSobre(); S.sujoSobre = false; }
  if (PERF) {
    S.perf.base = S.perf.base * .8 + (t1 - t0) * .2; S.perf.sobre = S.perf.sobre * .8 + (performance.now() - t1) * .2; S.perf.quadros++;
    $('perf').textContent = `base ${S.perf.base.toFixed(1)} ms · sobre ${S.perf.sobre.toFixed(1)} ms · ${S.perf.quadros} quadros`;
  }
}

/** Some com o hover (e a dica) enquanto a câmera se mexe: assim a camada de cima não precisa acompanhar cada quadro. */
function limparHover() {
  if (S.hover) { S.hover = null; S.sujoSobre = true; }
  esconderDica();
}

function desenharBase() {
  const ctx = canvas.getContext('2d'), dpr = prepararCanvas(), { w, h } = S.tam;
  const { k, x, y } = S.camera, geo = S.geo, dm = dadosDoMes();
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);
  ctx.translate(x, y); ctx.scale(k, k);
  ctx.lineJoin = 'round';
  const visivel = m => !(m.box[2] * k + x < 0 || m.box[0] * k + x > w || m.box[3] * k + y < 0 || m.box[1] * k + y > h);

  if (!dm.porMun) {
    for (const e of Object.values(geo.ufs)) { ctx.fillStyle = faixa(dm.uf[e.uf], true).cor; ctx.fill(e.fill); }
  } else if (S.unidade === 'bolhas' && S.payload.modo !== 'caso') {
    ctx.fillStyle = C.VAZIO;
    for (const e of Object.values(geo.ufs)) { ctx.globalAlpha = S.uf && e.uf !== S.uf ? .4 : 1; ctx.fill(e.fill); }
    const bolhas = dm.bolhas ||= [...dm.mun.values()].filter(r => r.v).sort((a, b) => b.v - a.v);
    const maxV = Math.max(1, bolhas[0]?.v || 1);
    const escala = BOLHA_MAIOR * (geo.box[2] - geo.box[0]) / Math.sqrt(maxV);
    const amortecer = Math.sqrt(k / cameraPara(geo, {}, w, h).k);
    ctx.strokeStyle = C.FUNDO; ctx.lineWidth = .6 / k;
    for (const r of bolhas) {
      const m = geo.porId.get(r.id);
      if (!visivel(m)) continue;
      ctx.globalAlpha = enfase(m) * .9;
      ctx.fillStyle = faixa(r).cor;
      ctx.beginPath();
      ctx.arc(...ancora(m), Math.max(1.5 / k, Math.sqrt(r.v) * escala / amortecer), 0, Math.PI * 2);
      ctx.fill(); ctx.stroke();
    }
  } else {
    for (const u of baldesDe(dm)) {
      if (u.box[2] * k + x < 0 || u.box[0] * k + x > w || u.box[3] * k + y < 0 || u.box[1] * k + y > h) continue;   // UF fora da tela
      for (const g of u.grupos) { ctx.globalAlpha = g.alpha; ctx.fillStyle = g.cor; ctx.fill(g.path); }
    }
    if (S.uf || S.payload.modo === 'caso') {
      ctx.globalAlpha = 1; ctx.strokeStyle = 'rgba(18,18,20,.55)'; ctx.lineWidth = .5 / k;
      ctx.stroke(geo.bordas.municipio);
    }
  }

  ctx.globalAlpha = 1;
  ctx.strokeStyle = C.FUNDO; ctx.lineWidth = 1.3 / k; ctx.stroke(geo.bordas.uf);
  ctx.strokeStyle = 'rgba(255,255,255,.10)'; ctx.lineWidth = 1 / k; ctx.stroke(geo.bordas.costa);

  // Watchlist: contorno âmbar nos municípios em watchlist com alerta no mês.
  if (dm.porMun && S.payload.modo === 'radar') {
    ctx.strokeStyle = C.COR.ambar; ctx.lineWidth = 1.2 / k;
    for (const r of emAlertaDe(dm)) if (r.bits & 8) { const m = geo.porId.get(r.id); if (visivel(m)) ctx.stroke(m.path); }
  }
  // Área engana: um município pequeno em alerta some no mapa nacional. Todo município
  // em alerta ganha um marcador no centroide, do mesmo tamanho na tela.
  if (dm.porMun && S.metrica === 'alerta' && S.unidade !== 'bolhas' && S.payload.modo === 'radar') {
    const raio = (S.uf ? 4 : 5) / k;
    for (const r of emAlertaDe(dm)) {
      const m = geo.porId.get(r.id);
      if (!visivel(m)) continue;
      ctx.globalAlpha = enfase(m);
      const [ax, ay] = ancora(m); ctx.beginPath(); ctx.arc(ax, ay, raio, 0, Math.PI * 2);
      ctx.fillStyle = C.SEVERIDADE[C.severidade(r.bits)]; ctx.fill();
      ctx.strokeStyle = 'rgba(255,255,255,.85)'; ctx.lineWidth = 1.2 / k; ctx.stroke();
    }
    ctx.globalAlpha = 1;
  }
  ctx.globalAlpha = 1;
  if (S.payload.modo === 'caso') desenharCamadaCaso(ctx, k, dm, visivel);
  if (S.uf) { ctx.strokeStyle = 'rgba(255,255,255,.6)'; ctx.lineWidth = 1.4 / k; ctx.stroke(geo.ufs[S.uf].contorno); }
  if (S.municipio) {
    ctx.save();
    ctx.shadowColor = 'rgba(255,255,255,.45)'; ctx.shadowBlur = 8 * dpr;
    ctx.strokeStyle = C.COR.foco; ctx.lineWidth = 2.2 / k; ctx.stroke(S.municipio.path);
    ctx.restore();
  }
  S.camBase = { ...S.camera };
  canvas.style.transform = '';
  // Rótulos em coordenadas de tela, no próprio canvas: <span> sobre o canvas custava ~40 ms/quadro
  // em pintura/composição (uma camada por rótulo, com text-shadow).
  ctx.setTransform(1, 0, 0, 1, 0, 0);       // sprites já vêm em pixels do dispositivo
  desenharRotulos(ctx, dpr);
}

/**
 * Mapa do caso: arcos de deslocamento impossível (mesma placa/telefone em municípios distantes em
 * pouco tempo) e círculos proporcionais ao volume acumulado até o mês. O círculo tem tamanho de TELA
 * (não encolhe no zoom) e quem esteve ativo no mês ganha um anel laranja — ao reproduzir a linha do
 * tempo o perito vê onde a célula começou e para onde migrou.
 */
const BOLHA_CASO_MAX_PX = 28, BOLHA_CASO_MIN_PX = 4;
function desenharCamadaCaso(ctx, k, dm, visivel) {
  const geo = S.geo, mes = S.mes, foco = S.municipio?.id, comTempo = S.payload.meses.length > 1;
  ctx.save();
  for (const d of deslocamentosAteMes()) {
    const a = geo.porId.get(d.de), b = geo.porId.get(d.para);
    if (!a || !b) continue;
    const [x1, y1] = ancora(a), [x2, y2] = ancora(b);
    const mx = (x1 + x2) / 2, my = (y1 + y2) / 2, dx = x2 - x1, dy = y2 - y1;
    const cx = mx - dy * .22, cy = my + dx * .22;          // curva para o lado: arcos paralelos não se sobrepõem
    const destaque = foco && (d.de === foco || d.para === foco);
    ctx.globalAlpha = foco && !destaque ? .25 : .9;
    ctx.strokeStyle = destaque ? '#FFFFFF' : C.COR.carmim;
    ctx.lineWidth = (1.4 + Math.log2(1 + d.n)) / k;
    ctx.setLineDash([6 / k, 4 / k]);
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.quadraticCurveTo(cx, cy, x2, y2); ctx.stroke();
  }
  ctx.setLineDash([]);
  const regs = [...dm.mun.values()].filter(r => r.v).sort((a, b) => b.v - a.v);
  ctx.globalAlpha = 1;
  if ((S.marcadores || 'circulos') === 'area') {
    // Só área: sem marcadores; quem esteve ativo no mês ganha o contorno laranja do próprio município.
    if (comTempo) {
      ctx.strokeStyle = C.COR.laranja; ctx.lineWidth = 2 / k;
      for (const r of regs) { const m = geo.porId.get(r.id); if (r.n && visivel(m)) ctx.stroke(m.path); }
    }
    ctx.restore();
    return;
  }
  // Círculos acompanham o zoom: tamanho cheio no enquadramento do caso, até ~1/3 afastando
  // (antes ficavam com 28px fixos e, com o mapa afastado, viravam bolhas sobrepostas).
  const escala = Math.min(1, Math.max(.35, Math.sqrt(k / kReferenciaCaso())));
  const maxV = Math.max(1, regs[0]?.v || 1);
  for (const r of regs) {
    const m = geo.porId.get(r.id);
    if (!visivel(m)) continue;
    const px = Math.max(BOLHA_CASO_MIN_PX * escala, Math.sqrt(r.v / maxV) * BOLHA_CASO_MAX_PX * escala);
    const [ax, ay] = ancora(m);
    ctx.beginPath(); ctx.arc(ax, ay, px / k, 0, Math.PI * 2);
    ctx.fillStyle = 'rgba(192,98,95,.45)'; ctx.fill();
    ctx.strokeStyle = r.id === foco ? '#FFFFFF' : C.COR.carmim; ctx.lineWidth = 1.1 / k; ctx.stroke();
    if (comTempo && r.n) {                                   // ativo neste mês
      if (px >= 7) {                                         // anel só quando o círculo é legível
        ctx.beginPath(); ctx.arc(ax, ay, (px + 2.5) / k, 0, Math.PI * 2);
        ctx.strokeStyle = C.COR.laranja; ctx.lineWidth = 1.6 / k; ctx.stroke();
      } else {                                               // pequeno demais: preenche de laranja
        ctx.beginPath(); ctx.arc(ax, ay, px / k, 0, Math.PI * 2);
        ctx.fillStyle = C.COR.laranja; ctx.fill();
      }
    }
  }
  ctx.restore();
}

/** Zoom do enquadramento inicial do caso (referência para o tamanho dos círculos). */
function kReferenciaCaso() {
  const caixa = caixaDoCaso();
  return caixa ? cameraPara(S.geo, { caixa }, S.tam.w, S.tam.h, areaUtil()).k : (S.camera?.k || 1);
}

/** Deslocamentos impossíveis já ocorridos até o mês da linha do tempo. */
function deslocamentosAteMes() {
  return (S.payload.deslocamentos || []).filter(d => (d.mes ?? 0) <= S.mes);
}

const SPRITES = new Map();
/** Cada rótulo (texto + estilo) é desenhado uma vez num canvas pequeno; depois é só drawImage. */
function spriteDe(texto, tipo, dpr) {
  const chave = `${tipo}|${dpr}|${texto}`;
  let sp = SPRITES.get(chave);
  if (sp) return sp;
  const fonte = tipo === 'forte' ? '600 12px "Segoe UI", sans-serif' : tipo === 'uf' ? '600 10px "Segoe UI", sans-serif' : '600 11px "Segoe UI", sans-serif';
  const m = ctxTeste; m.font = fonte;
  const w = Math.ceil(m.measureText(texto).width) + 10, h = 20;
  const c = document.createElement('canvas'); c.width = Math.ceil(w * dpr); c.height = h * dpr;
  const g = c.getContext('2d'); g.scale(dpr, dpr);
  g.font = fonte; g.textAlign = 'center'; g.textBaseline = 'middle'; g.lineJoin = 'round';
  g.lineWidth = 3.2; g.strokeStyle = 'rgba(18,18,20,.92)'; g.strokeText(texto, w / 2, h / 2);
  g.fillStyle = tipo === 'forte' ? '#FFFFFF' : tipo === 'uf' ? C.TINTA_2 : '#E4E4E7'; g.fillText(texto, w / 2, h / 2);
  if (SPRITES.size > 600) SPRITES.clear();
  SPRITES.set(chave, sp = { c, w, h });
  return sp;
}

function desenharRotulos(ctx, dpr) {
  for (const o of rotulosDaVista()) {
    const sp = spriteDe(o.texto, o.forte ? 'forte' : o.uf ? 'uf' : 'normal', dpr);
    ctx.drawImage(sp.c, Math.round((o.x - sp.w / 2) * dpr), Math.round((o.y - sp.h / 2) * dpr), sp.c.width, sp.c.height);
  }
}

/** Camada de cima: só o hover. É a única parte que muda a cada movimento do mouse. */
function desenharSobre() {
  const hv = S.hover;
  if (!hv || S.inter.arrastando) { sobre.hidden = true; return; }
  const { k, x, y } = S.camera, geo = S.geo, dm = dadosDoMes(), { w, h } = S.tam, dpr = Math.min(devicePixelRatio || 1, 2);
  let alvo = null, box = null;
  if (hv.mun && (S.uf || dm.porMun)) { alvo = hv.mun.path; box = hv.mun.box; }
  else if (hv.uf) { alvo = geo.ufs[hv.uf].contorno; box = geo.ufs[hv.uf].box; }
  if (!alvo) { sobre.hidden = true; return; }
  // O canvas só cobre a caixa do alvo (+ margem): cobrir a janela inteira obrigava o navegador a
  // recompor uma camada do tamanho da tela a cada quadro (~10 ms em renderização por software).
  const pad = 4;
  const x0 = Math.max(0, Math.floor(box[0] * k + x - pad)), y0 = Math.max(0, Math.floor(box[1] * k + y - pad));
  const x1 = Math.min(w, Math.ceil(box[2] * k + x + pad)), y1 = Math.min(h, Math.ceil(box[3] * k + y + pad));
  if (x1 <= x0 || y1 <= y0) { sobre.hidden = true; return; }
  sobre.hidden = false;
  sobre.style.left = x0 + 'px'; sobre.style.top = y0 + 'px'; sobre.style.width = (x1 - x0) + 'px'; sobre.style.height = (y1 - y0) + 'px';
  sobre.width = Math.round((x1 - x0) * dpr); sobre.height = Math.round((y1 - y0) * dpr);
  const ctx = sobre.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, -x0 * dpr, -y0 * dpr);
  ctx.translate(x, y); ctx.scale(k, k);
  ctx.lineJoin = 'round';
  ctx.strokeStyle = 'rgba(255,115,0,.95)'; ctx.lineWidth = 1.5 / k;
  ctx.stroke(alvo);
}

/* ------------------------------------------------------------------ teste de clique */
// Posição do cursor relativa ao #quadro, que NUNCA se move. Medir contra o próprio canvas era errado
// durante o gesto: o canvas está deslocado por CSS transform (moverCamera), o retângulo dele anda junto
// e cada pointermove "devolvia" parte do deslocamento — o mapa andava metade do mouse (medido: 150 de 300 px)
// e a âncora do zoom pela roda escorregava entre eventos seguidos.
const quadro = $('quadro');
/** Mapa pronto para interação: malha, dados e câmera carregados. O iframe é recriado a cada troca de
 *  painel (~1 s de carga); mexer o mouse ou a roda nesse intervalo lia S.camera.x com a câmera ainda
 *  null — "Cannot read properties of null (reading 'x')" — e deixava o gesto em estado inconsistente. */
const pronto = () => !!(S.geo && S.payload && S.camera);

function posicao(ev) {
  const b = S.retQuadro || (S.retQuadro = quadro.getBoundingClientRect());
  return [ev.clientX - b.left, ev.clientY - b.top];
}
window.addEventListener('resize', () => { S.retQuadro = null; });

function acertar(ev) {
  const [sx, sy] = posicao(ev), c = S.camera, x = (sx - c.x) / c.k, y = (sy - c.y) / c.k;
  const ctx = ctxTeste;
  const candidatos = S.uf ? S.geo.ufs[S.uf].municipios : S.geo.municipios;
  for (const m of candidatos) {
    if (x >= m.box[0] && x <= m.box[2] && y >= m.box[1] && y <= m.box[3] && ctx.isPointInPath(m.path, x, y)) return { uf: m.uf, mun: m, sx, sy };
  }
  if (S.uf) for (const m of S.geo.municipios) {     // clique fora da UF aberta: troca de UF
    if (m.uf !== S.uf && x >= m.box[0] && x <= m.box[2] && y >= m.box[1] && y <= m.box[3] && ctx.isPointInPath(m.path, x, y)) return { uf: m.uf, mun: m, sx, sy, outraUf: true };
  }
  return null;
}

canvas.addEventListener('pointerdown', ev => {
  if (ev.button !== 0 || !pronto()) return;
  S.retQuadro = null;                       // retângulo novo a cada gesto (o iframe pode ter mudado de lugar)
  const p = posicao(ev);
  S.toques.set(ev.pointerId, p);
  S.inter = { inicio: p, anterior: p, arrastando: false };
  if (S.toques.size === 2) S.inter.pinca = Math.hypot(...(() => { const [a, b] = [...S.toques.values()]; return [a[0] - b[0], a[1] - b[1]]; })());
  try { canvas.setPointerCapture(ev.pointerId); } catch { /* ponteiro sintético ou já liberado */ }
});
canvas.addEventListener('pointermove', ev => {
  if (!pronto()) return;
  const p = posicao(ev), it = S.inter;
  if (S.toques.has(ev.pointerId)) {
    S.toques.set(ev.pointerId, p);
    if (S.toques.size === 2) {
      const [a, b] = [...S.toques.values()], d = Math.hypot(a[0] - b[0], a[1] - b[1]);
      if (it.pinca) zoom(d / it.pinca, [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2], false);
      it.pinca = d; it.arrastando = true; return;
    }
    if (it.inicio && Math.hypot(p[0] - it.inicio[0], p[1] - it.inicio[1]) > ARRASTE_MIN) it.arrastando = true;
    if (it.arrastando) {
      cancelAnimationFrame(S.quadro);
      moverCamera({ ...S.camera, x: S.camera.x + p[0] - it.anterior[0], y: S.camera.y + p[1] - it.anterior[1] });
      it.anterior = p; canvas.classList.add('arrastando');
      return;
    }
  }
  if (ev.pointerType === 'mouse') atualizarHover(acertar(ev));
});
canvas.addEventListener('pointerup', ev => {
  if (!pronto()) { S.toques.clear(); S.inter = {}; return; }
  const arrastou = S.inter.arrastando;
  S.toques.delete(ev.pointerId);
  canvas.classList.remove('arrastando');
  if (!arrastou) {
    const a = acertar(ev);
    if (a) {
      if (!S.uf || a.outraUf) aplicarEstado({ uf: a.uf, municipio: null }, true);
      else aplicarEstado({ uf: a.uf, municipio: a.mun.id }, true);
    }
  }
  S.inter = S.toques.size ? { ...S.inter, arrastando: true, anterior: [...S.toques.values()][0] } : {};
  if (arrastou && !S.toques.size) { clearTimeout(S.tReraster); agendar(true); }   // soltou: nítido já no próximo quadro
});
canvas.addEventListener('pointercancel', () => { S.toques.clear(); S.inter = {}; });
canvas.addEventListener('pointerleave', () => atualizarHover(null));
canvas.addEventListener('wheel', ev => { ev.preventDefault(); if (!pronto()) return; zoom(Math.exp(-ev.deltaY * .0015), posicao(ev), false); }, { passive: false });

/** Hover: a camada de cima só é refeita (e a dica só reescrita) quando o alvo sob o cursor muda. */
function atualizarHover(h) {
  const id = h ? (h.mun?.id || h.uf) : null, antes = S.hover ? (S.hover.mun?.id || S.hover.uf) : null;
  S.hover = h;
  if (id !== antes) { agendar(false); mostrarDica(true); }
  else if (h) posicionarDica(h.sx, h.sy);
}

function subir() {
  if (S.municipio) aplicarEstado({ uf: S.uf, municipio: null }, true);
  else if (S.uf) aplicarEstado({ uf: null, municipio: null }, true);
}

/* ------------------------------------------------------------------ HTML: controles, legenda, dica, rótulos */
function segmentado(id, opcoes, valor, aoMudar) {
  const caixa = $(id); caixa.innerHTML = '';
  for (const [chave, rotulo] of Object.entries(opcoes)) {
    const b = document.createElement('button');
    b.textContent = rotulo; b.dataset.chave = chave; b.setAttribute('aria-pressed', String(chave === valor));
    b.onclick = () => aoMudar(chave);
    caixa.appendChild(b);
  }
}

function marcarSegmentado(id, valor) {
  for (const b of $(id).children) b.setAttribute('aria-pressed', String(b.dataset.chave === valor));
}

/** Constrói os controles uma vez por payload (antes eram refeitos a cada tick do slider). */
function montarControles() {
  const caso = S.payload.modo === 'caso';
  agendarAjusteControles();
  $('unidades').hidden = caso;
  $('marcadores').hidden = !caso;
  if (caso) segmentado('marcadores', MARCADORES, S.marcadores || 'circulos', v => { S.marcadores = v; S.sujo = true; montarControles(); montarLegenda(true); agendar(true); });
  segmentado('unidades', UNIDADES, S.unidade, u => aplicarEstado({ unidade: u }, true));
  const metricas = caso ? { volume: METRICAS.volume, taxa: METRICAS.taxa } : METRICAS;
  segmentado('metricas', metricas, S.metrica, m => aplicarEstado({ metrica: m }, true));
  const linha = $('linha-tempo'), p = S.payload;
  linha.hidden = p.meses.length < 2;   // caso com datas: linha do tempo da célula (rota de expansão)
  const r = $('mes');
  r.min = Math.min(p.mes_inicial ?? 0, p.meses.length - 1); r.max = p.meses.length - 1; r.value = S.mes;
  const marcas = $('marcas'); marcas.innerHTML = '';
  for (let i = +r.min; i <= +r.max; i++) {
    const s = document.createElement('span'); s.textContent = MESES[+p.meses[i].slice(5, 7) - 1] || ''; marcas.appendChild(s);
  }
  medirLinhaTempo();
}

/** Atualiza só o que mudou: estados dos botões, rótulo do mês, migalhas e legenda. */
function atualizarUI(agora = false) {
  agendarAjusteControles();            // as migalhas mudam de largura (Brasil › UF › município)
  if (!S.payload) return;
  if (!agora) { S.sujoUI = true; agendar(false); return; }
  const caso = S.payload.modo === 'caso';
  marcarSegmentado('unidades', S.unidade); marcarSegmentado('metricas', S.metrica);
  $('mes').value = S.mes;
  $('mes-rotulo').textContent = fmtMes(S.payload.meses[S.mes]);

  const migalhas = $('migalhas'); migalhas.innerHTML = '';
  const passos = [['Brasil', () => aplicarEstado({ uf: null, municipio: null }, true)]];
  if (S.uf) passos.push([NOME_UF[S.uf] || S.uf, () => aplicarEstado({ uf: S.uf, municipio: null }, true)]);
  if (S.municipio) passos.push([S.municipio.nome, null]);
  passos.forEach(([rot, acao], i) => {
    if (i) { const sep = document.createElement('i'); sep.textContent = '›'; migalhas.appendChild(sep); }
    const b = document.createElement(acao && i < passos.length - 1 ? 'button' : 'b');
    b.textContent = rot; if (acao && i < passos.length - 1) b.onclick = acao;
    migalhas.appendChild(b);
  });
  $('voltar').hidden = !S.uf;
  montarLegenda(caso);
  if (S.hover) mostrarDica(true);
}

function montarLegenda(caso) {
  const leg = $('legenda'), dm = dadosDoMes();
  const recolhida = S.legendaManual ? S.legendaRecolhida : S.tam.w < LARGURA_LEGENDA_COMPACTA;
  leg.classList.toggle('recolhida', recolhida);
  leg.innerHTML = '';
  const cab = document.createElement('div'); cab.className = 'legenda-cab';
  const titulo = document.createElement('div'); titulo.className = 'legenda-titulo';
  const comMes = !caso || S.payload.meses.length > 1;
  titulo.textContent = `${METRICAS[S.metrica]} · ${dm.porMun ? 'municípios' : 'estados'}${comMes ? ' · ' + fmtMes(S.payload.meses[S.mes]) : ''}`;
  const alternar = document.createElement('button');
  alternar.className = 'legenda-alternar'; alternar.textContent = recolhida ? '▸' : '▾';
  alternar.title = recolhida ? 'Mostrar a legenda' : 'Recolher a legenda'; alternar.setAttribute('aria-expanded', String(!recolhida));
  alternar.onclick = () => { S.legendaManual = true; S.legendaRecolhida = !recolhida; montarLegenda(caso); };
  cab.append(titulo, alternar); leg.appendChild(cab);
  if (recolhida) return;
  for (const f of faixasDaLegenda()) {
    const l = document.createElement('div'); l.className = 'legenda-linha';
    l.innerHTML = `<i style="background:${f.cor}"></i><span>${f.rotulo}</span><b>${fmtInt(f.n)}</b><em>${fmtPct(f.parcela)}</em>`;
    leg.appendChild(l);
  }
  if (caso) {
    const extra = document.createElement('div'); extra.className = 'legenda-caso';
    const n = deslocamentosAteMes().length, crit = S.payload.criterio_deslocamento || {};
    const soArea = (S.marcadores || 'circulos') === 'area';
    extra.innerHTML = (soArea ? '' : `<span><i class="leg-circulo"></i>círculo = acionamentos acumulados</span>`)
      + (S.payload.meses.length > 1 ? `<span><i class="leg-anel"></i>${soArea ? 'contorno laranja' : 'anel'} = ativo no mês</span>` : '')
      + `<span title="Mesma placa ou telefone em municípios a ≥ ${crit.km_mesmo_dia || 300} km no mesmo dia, ou ≥ ${crit.km_janela || 600} km em até ${crit.janela_dias ?? 1} dia(s). A base só tem datas, sem hora.">`
      + `<i class="leg-arco"></i>${n ? `${n} deslocamento${n > 1 ? 's' : ''} impossíve${n > 1 ? 'is' : 'l'}` : 'nenhum deslocamento impossível'}</span>`;
    leg.appendChild(extra);
  }
  const rod = document.createElement('div'); rod.className = 'legenda-rodape';
  rod.textContent = `${fmtPct(S.payload.pct_mapeado ?? 1)} dos acionamentos mapeados · ${S.payload.fonte || 'IBGE'}`;
  leg.appendChild(rod);
}

/** Rótulos de texto sobre o mapa: UFs no Brasil; municípios do caso; municípios em alerta numa UF. */
function rotulosDaVista() {
  const p = S.payload, c = S.camera, dm = dadosDoMes(), out = [];
  const aMapa = (m, texto, forte) => { const [ax, ay] = ancora(m); out.push({ x: ax * c.k + c.x, y: ay * c.k + c.y, texto, forte }); };
  if (p.modo === 'caso') {
    const regs = [...dm.mun.values()].filter(r => r.v).sort((a, b) => b.v - a.v);
    const foco = S.municipio?.id;
    for (const r of regs.slice(0, MAX_ROTULOS_CASO)) aMapa(S.geo.porId.get(r.id), `${r.nome} · ${fmtInt(r.v)}`, r.id === foco);
    if (foco && !out.some(o => o.forte)) { const r = dm.mun.get(foco); if (r) aMapa(S.municipio, `${r.nome} · ${fmtInt(r.v)}`, true); }
  } else if (!S.uf) {
    for (const e of Object.values(S.geo.ufs)) out.push({ x: e.centro[0] * c.k + c.x, y: e.centro[1] * c.k + c.y, texto: e.uf, forte: false, uf: true });
  } else {
    for (const r of emAlertaDe(dm).filter(r => r.uf === S.uf).slice(-MAX_ROTULOS_UF).reverse()) aMapa(S.geo.porId.get(r.id), r.nome, r.id === S.municipio?.id);
  }
  // Sem sobreposição: o que tem prioridade (vem primeiro) fica; o que colidir some.
  const { w, h } = S.tam, L = S.layout, caixas = [], vis = [];
  for (const o of out) {
    if (o.x < L.esq || o.x > w - L.dir || o.y < L.topo || o.y > h - L.base) continue;
    const lw = o.texto.length * 6 + 10, lh = 14, r = [o.x - lw / 2, o.y - lh / 2, o.x + lw / 2, o.y + lh / 2];
    if (!o.forte && caixas.some(b => r[0] < b[2] && r[2] > b[0] && r[1] < b[3] && r[3] > b[1])) continue;
    caixas.push(r); vis.push(o);
  }
  return vis;
}

function mostrarDica(reescrever) {
  const dica = $('dica'), hv = S.hover;
  if (!hv || S.inter.arrastando) { esconderDica(); return; }
  const dm = dadosDoMes(), caso = S.payload.modo === 'caso';
  if (reescrever || S.dicaId !== (hv.mun?.id || hv.uf)) {
    let html;
    if (dm.porMun || S.uf) {
      const r = dm.mun.get(hv.mun.id);
      const alertas = r && r.bits ? S.payload.alertas.filter(a => r.bits & a.bit).map(a => `<li>${a.rotulo}</li>`).join('') : '';
      html = `<strong>${hv.mun.nome} / ${hv.mun.uf}</strong><small>IBGE ${hv.mun.id} · ${fmtInt(hv.mun.populacao)} hab.</small>`
        + (caso && S.payload.meses.length > 1
        ? `<span>Acumulado até ${fmtMes(S.payload.meses[S.mes])}: <b>${fmtInt(r?.v || 0)}</b> · no mês: <b>${fmtInt(r?.n || 0)}</b></span>`
        : `<span>${caso ? 'Acionamentos do caso' : 'Acionamentos no mês'}: <b>${fmtInt(r?.v || 0)}</b></span>`)
        + (!caso && r?.media ? `<span>Média histórica: ${fmtInt(r.media)} · desvio ${fmtDesvio(r.razao, r.media)}</span>` : '')
        + (S.metrica === 'taxa' && r ? `<span>Taxa: ${r.taxa.toLocaleString('pt-BR', { maximumFractionDigits: 2 })} /10 mil hab.</span>` : '')
        + (alertas ? `<ul>${alertas}</ul>` : '');
    } else {
      const e = dm.uf[hv.uf];
      html = `<strong>${NOME_UF[hv.uf] || hv.uf}</strong>`
        + `<span>Acionamentos no mês: <b>${fmtInt(e?.v || 0)}</b></span>`
        + (e?.media ? `<span>Média histórica: ${fmtInt(e.media)} · desvio ${fmtX(e.razao)}</span>` : '')
        + `<span>Municípios em alerta: <b>${fmtInt(e?.nAlerta || 0)}</b></span>`;
    }
    dica.innerHTML = html; S.dicaId = hv.mun?.id || hv.uf;
    S.dicaAltura = null;
  }
  dica.hidden = false;
  posicionarDica(hv.sx, hv.sy);
}

function posicionarDica(sx, sy) {
  const dica = $('dica');
  if (dica.hidden) return;
  S.dicaAltura ??= dica.offsetHeight;      // mede uma vez por conteúdo, não a cada movimento do mouse
  const x = Math.max(8, Math.min(S.tam.w - 250, sx + 14)), y = Math.max(8, Math.min(S.tam.h - S.dicaAltura - 8, sy - 20));
  dica.style.transform = `translate(${x}px, ${y}px)`;
}

function esconderDica() { $('dica').hidden = true; S.dicaId = null; }

/* ------------------------------------------------------------------ linha do tempo */
$('mes').addEventListener('input', ev => { S.mes = +ev.target.value; S.cache = null; agendar(true); atualizarUI(); });
$('mes').addEventListener('change', () => emitirEstado());
$('play').addEventListener('click', () => {
  if (S.play) { pararPlay(); return; }
  const r = $('mes');
  if (S.mes >= +r.max) { S.mes = +r.min; }
  $('play').setAttribute('aria-pressed', 'true'); $('play').textContent = '❚❚';
  S.play = setInterval(() => {
    if (S.mes >= +r.max) { pararPlay(); return; }
    S.mes++; S.cache = null; agendar(true); atualizarUI();
  }, PLAY_MS);
  S.cache = null; agendar(true); atualizarUI();
});
function pararPlay() {
  clearInterval(S.play); S.play = null;
  $('play').setAttribute('aria-pressed', 'false'); $('play').textContent = '▶';
  emitirEstado();
}

/* ------------------------------------------------------------------ busca "/" */
// Combobox no padrão do projeto de referência: ↑/↓ movem a opção ativa (com volta), Enter abre,
// o mouse também move a opção ativa, Esc fecha. Sem termo, sugere os municípios em alerta do mês.
const MAX_RESULTADOS = 25;
const busca = { itens: [], ativo: 0 };

function abrirBusca() {
  $('busca').hidden = false; $('busca-campo').value = ''; preencherBusca(''); $('busca-campo').focus();
}
function fecharBusca() { $('busca').hidden = true; }

function lugaresPara(termo) {
  const t = semAcento(termo.trim()), noMes = id => S.payload?.municipios[id]?.v[S.mes] || 0;
  if (!t) {
    if (!S.payload || S.payload.modo !== 'radar') return [];
    return [...dadosDoMes().mun.values()].filter(r => r.bits).sort((a, b) => C.severidade(b.bits) - C.severidade(a.bits) || b.v - a.v)
      .slice(0, 8).map(r => ({ uf: r.uf, mun: r.id, nome: r.nome, detalhe: `Em alerta · ${fmtInt(r.v)} no mês`, sugestao: true }));
  }
  const estados = Object.entries(NOME_UF).filter(([uf, nome]) => semAcento(uf) === t || semAcento(nome).includes(t))
    .map(([uf, nome]) => ({ uf, nome, detalhe: `Estado · ${S.geo.ufs[uf]?.municipios.length || 0} municípios` }));
  const muns = [];
  for (const m of S.geo.municipios) {
    const n = semAcento(m.nome), pos = n.indexOf(t);
    if (pos >= 0) muns.push({ m, pos });
  }
  // Começo do nome primeiro ("SANTO" acha Santo André antes de Espírito Santo), depois população.
  muns.sort((a, b) => (a.pos !== 0) - (b.pos !== 0) || (b.m.populacao || 0) - (a.m.populacao || 0));
  return [...estados, ...muns.slice(0, MAX_RESULTADOS).map(({ m }) => {
    const v = noMes(m.id);
    return { uf: m.uf, mun: m.id, nome: m.nome, detalhe: `Município · ${NOME_UF[m.uf] || m.uf}${v ? ` · ${fmtInt(v)} no mês` : ''}` };
  })];
}

function preencherBusca(termo) {
  busca.itens = lugaresPara(termo); busca.ativo = 0;
  const lista = $('busca-lista'); lista.innerHTML = '';
  if (!termo.trim() && busca.itens.length) {
    const cap = document.createElement('p'); cap.className = 'busca-legenda'; cap.textContent = 'Em alerta neste mês'; lista.appendChild(cap);
  }
  if (!busca.itens.length && termo.trim()) {
    const vazio = document.createElement('p'); vazio.className = 'busca-legenda';
    vazio.textContent = `Nenhum lugar com “${termo.trim()}”. Confira a grafia ou tente só o começo do nome.`;
    lista.appendChild(vazio);
  }
  busca.itens.forEach((r, i) => {
    const b = document.createElement('button');
    b.id = `busca-opcao-${i}`; b.setAttribute('role', 'option'); b.tabIndex = -1;
    b.innerHTML = `<span class="busca-uf">${r.uf}</span><span class="busca-nome"><strong></strong><small></small></span>`;
    b.querySelector('strong').textContent = r.nome; b.querySelector('small').textContent = r.detalhe;
    b.onclick = () => escolher(i);
    b.onpointermove = () => { if (busca.ativo !== i) marcarAtivo(i); };
    lista.appendChild(b);
  });
  marcarAtivo(0);
}

function marcarAtivo(i) {
  busca.ativo = i;
  const opcoes = $('busca-lista').querySelectorAll('[role="option"]');
  opcoes.forEach((o, j) => o.setAttribute('aria-selected', String(j === i)));
  const campo = $('busca-campo');
  if (opcoes[i]) { campo.setAttribute('aria-activedescendant', opcoes[i].id); opcoes[i].scrollIntoView({ block: 'nearest' }); }
  else campo.removeAttribute('aria-activedescendant');
}

function escolher(i) {
  const r = busca.itens[i];
  if (!r) return;
  fecharBusca();
  aplicarEstado({ uf: r.uf, municipio: r.mun || null }, true);
}

$('busca-campo').addEventListener('input', ev => preencherBusca(ev.target.value));
$('busca-campo').addEventListener('keydown', ev => {
  const n = busca.itens.length;
  if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
    ev.preventDefault();
    if (n) marcarAtivo((busca.ativo + (ev.key === 'ArrowDown' ? 1 : -1) + n) % n);
  } else if (ev.key === 'Enter') { ev.preventDefault(); escolher(busca.ativo); }
  else if (ev.key === 'Escape') { ev.stopPropagation(); fecharBusca(); }
});
$('busca').addEventListener('click', ev => { if (ev.target.id === 'busca') fecharBusca(); });

document.addEventListener('keydown', ev => {
  if (ev.target.tagName === 'INPUT' && ev.target.type !== 'range') return;
  if (ev.key === '/') { ev.preventDefault(); abrirBusca(); }
  else if (ev.key === 'Escape') { if (!$('busca').hidden) fecharBusca(); else subir(); }
  else if (ev.key === '+' || ev.key === '=') zoom(PASSO_ZOOM);
  else if (ev.key === '-') zoom(1 / PASSO_ZOOM);
});

/* ------------------------------------------------------------------ botões */
$('zoom-mais').onclick = () => zoom(PASSO_ZOOM);
$('zoom-menos').onclick = () => zoom(1 / PASSO_ZOOM);
$('centralizar').onclick = () => { irPara(false); };
$('voltar').onclick = subir;
$('abrir-busca').onclick = abrirBusca;
$('exportar').onclick = exportarPNG;

/** PNG com o mapa, título, rótulos e legenda (tudo desenhado num canvas novo). */
function exportarPNG() {
  S.sujo = true; desenharTudo();            // garante a imagem nítida (o gesto pode ter deixado só um transform CSS)
  const dpr = canvas.width / S.tam.w, faixas = faixasDaLegenda();
  const out = document.createElement('canvas');
  out.width = canvas.width; out.height = canvas.height;
  const ctx = out.getContext('2d');
  ctx.fillStyle = C.FUNDO; ctx.fillRect(0, 0, out.width, out.height);
  ctx.drawImage(canvas, 0, 0);          // a base já inclui contornos e rótulos
  ctx.scale(dpr, dpr);
  const lugar = [S.municipio?.nome, S.uf && (NOME_UF[S.uf] || S.uf), S.payload.modo === 'caso' ? 'mapa do caso' : 'Brasil'].filter(Boolean)[0];
  ctx.fillStyle = '#EDEDED'; ctx.font = '600 15px Segoe UI, sans-serif';
  ctx.fillText(`LCFO · Radar territorial — ${lugar}`, 16, 26);
  ctx.fillStyle = '#A1A1AA'; ctx.font = '12px Segoe UI, sans-serif';
  ctx.fillText(`${METRICAS[S.metrica]} · ${fmtMes(S.payload.meses[S.mes])}`, 16, 44);
  let y = S.tam.h - 18 - faixas.length * 18;
  for (const f of faixas) {
    ctx.fillStyle = f.cor; ctx.fillRect(16, y - 10, 12, 12);
    ctx.fillStyle = '#D4D4D8'; ctx.fillText(`${f.rotulo} — ${fmtInt(f.n)} (${fmtPct(f.parcela)})`, 34, y);
    y += 18;
  }
  ctx.fillStyle = '#71717A'; ctx.textAlign = 'right';
  ctx.fillText(S.payload.fonte || 'Fonte: IBGE', S.tam.w - 16, S.tam.h - 14);
  out.toBlob(blob => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `radar-${S.municipio?.id || S.uf || 'brasil'}-${S.payload.meses[S.mes]}.png`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });
}

function mostrarErro(msg) { $('aviso').textContent = msg; $('aviso').hidden = false; }
window.addEventListener('error', ev => mostrarErro(`Erro no mapa: ${ev.message}`));
window.addEventListener('unhandledrejection', ev => mostrarErro(`Erro no mapa: ${ev.reason?.message || ev.reason}`));

/* ------------------------------------------------------------------ início */
new ResizeObserver(entradas => {
  agendarAjusteControles();
  S.retQuadro = null;
  const { width, height } = entradas[0].contentRect;
  if (!width || !height) return;
  S.tam = { w: width, h: height };
  medirLinhaTempo();
  if (S.geo && S.payload) {
    if (S.camManual) { S.sujo = true; } else irPara(true);
    if (!S.legendaManual) atualizarUI(true);
    agendar(true);
  }
}).observe($('quadro'));

(async () => {
  if (PERF) { $('perf').hidden = false; window.__lcfoMapa = { S, desenharTudo, dadosDoMes }; }
  try {
    const r = await fetch('data/brasil.topo.json');
    if (!r.ok) throw new Error(`malha não encontrada (HTTP ${r.status})`);
    S.geo = criarGeografia(await r.json());
  } catch (e) {
    mostrarErro(`O mapa não carregou: ${e.message}. Rode tools/gerar_malha.py.`);
    return;
  }
  parent.postMessage({ isStreamlitMessage: true, type: 'streamlit:componentReady', apiVersion: 1 }, '*');
  if (S.urlPendente) carregarPayload(S.urlPendente);
  // Página aberta direto (sem o app): ?payload=URL para testar o componente isolado.
  // Também aceita ?uf=SP&mun=3547809&metrica=volume&unidade=bolhas para abrir num estado específico,
  // e ?topo=108&dir=296 (etc.) para simular a área coberta por HUD/painéis do app.
  const q = new URLSearchParams(location.search);
  if (q.get('payload')) {
    await carregarPayload(q.get('payload'));
    if (['topo', 'base', 'esq', 'dir'].some(k => q.get(k))) aplicarLayout(Object.fromEntries(['topo', 'base', 'esq', 'dir'].map(k => [k, q.get(k)])));
    aplicarEstado({ uf: q.get('uf'), municipio: q.get('mun'), metrica: q.get('metrica'), unidade: q.get('unidade') }, false);
    if (S.camera && S.alvo) { cancelAnimationFrame(S.quadro); S.camera = { ...S.alvo }; S.sujo = true; desenharTudo(); }
  }
})();
