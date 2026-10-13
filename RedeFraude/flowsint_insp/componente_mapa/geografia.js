/*
 * TopoJSON (pré-projetado, metros) -> Path2D do Canvas, contornos de UF e câmeras.
 *
 * Adaptado de open-apuracao-brazil (src/map/geography.js)
 * Copyright (c) 2026 Bruno Pinheiro — Licença MIT (https://opensource.org/licenses/MIT).
 * Mudanças: sem zonas eleitorais; câmera com área útil configurável (parâmetro `area`); malha
 * própria gerada por tools/gerar_malha.py a partir do IBGE.
 */

const NORONHA = '2605459';   // desenhado, mas não amplia a câmera do Brasil

const caixaVazia = () => [Infinity, Infinity, -Infinity, -Infinity];
export const estender = (box, x, y) => {
  box[0] = Math.min(box[0], x); box[1] = Math.min(box[1], y);
  box[2] = Math.max(box[2], x); box[3] = Math.max(box[3], y);
};

export function criarGeografia(topologia) {
  const { scale, translate } = topologia.transform;
  const box = caixaVazia();
  const arcos = topologia.arcs.map(arco => {
    let x = 0, y = 0;
    return arco.map(([dx, dy]) => {
      x += dx; y += dy;
      const p = [(x * scale[0] + translate[0]) / 1000, -(y * scale[1] + translate[1]) / 1000];  // km, y para baixo
      estender(box, ...p);
      return p;
    });
  });
  const origem = [box[0], box[1]];
  for (const arco of arcos) for (const p of arco) { p[0] -= origem[0]; p[1] -= origem[1]; }

  const donosDoArco = arcos.map(() => []);
  const caixaBrasil = caixaVazia();
  const municipios = topologia.objects.municipios.geometries.map((g, indice) => {
    const path = new Path2D(), limites = caixaVazia();
    const poligonos = g.type === 'Polygon' ? [g.arcs] : g.arcs;
    let area = 0, centro = [0, 0], maiorArea = 0;
    for (const poligono of poligonos) for (const [iAnel, anel] of poligono.entries()) {
      const pontos = [];
      for (const idArco of anel) {
        const id = idArco < 0 ? ~idArco : idArco;
        donosDoArco[id].push(indice);
        const p = idArco < 0 ? [...arcos[id]].reverse() : arcos[id];
        pontos.push(...p.slice(pontos.length ? 1 : 0));
      }
      path.moveTo(...pontos[0]);
      for (const ponto of pontos.slice(1)) path.lineTo(...ponto);
      path.closePath();
      for (const p of pontos) estender(limites, ...p);
      if (iAnel === 0) {
        let a2 = 0, cx = 0, cy = 0;
        for (let i = 0, j = pontos.length - 1; i < pontos.length; j = i++) {
          const f = pontos[j][0] * pontos[i][1] - pontos[i][0] * pontos[j][1];
          a2 += f; cx += (pontos[j][0] + pontos[i][0]) * f; cy += (pontos[j][1] + pontos[i][1]) * f;
        }
        const a = Math.abs(a2 / 2); area += a;
        if (a > maiorArea) { maiorArea = a; centro = a2 ? [cx / (3 * a2), cy / (3 * a2)] : pontos[0]; }
      }
    }
    const p = g.properties;
    if (p.id !== NORONHA) { estender(caixaBrasil, limites[0], limites[1]); estender(caixaBrasil, limites[2], limites[3]); }
    return { indice, id: String(p.id), nome: p.n, uf: p.uf, populacao: p.p, path, box: limites, centro, area };
  });

  const ufs = {};
  for (const m of municipios) {
    const e = ufs[m.uf] ||= { uf: m.uf, municipios: [], box: caixaVazia(), centro: [0, 0], area: 0, contorno: new Path2D(), fill: new Path2D(), populacao: 0 };
    e.municipios.push(m); e.area += m.area; e.populacao += m.populacao || 0;
    e.fill.addPath(m.path);
    e.centro[0] += m.centro[0] * m.area; e.centro[1] += m.centro[1] * m.area;
    if (m.id !== NORONHA) { estender(e.box, m.box[0], m.box[1]); estender(e.box, m.box[2], m.box[3]); }
  }
  for (const e of Object.values(ufs)) e.centro = e.centro.map(n => n / e.area);

  const bordas = { uf: new Path2D(), municipio: new Path2D(), costa: new Path2D() };
  for (let i = 0; i < arcos.length; i++) {
    const donos = [...new Set(donosDoArco[i])];
    const tipo = donos.length < 2 ? 'costa' : municipios[donos[0]].uf !== municipios[donos[1]].uf ? 'uf' : 'municipio';
    const tracar = path => { path.moveTo(...arcos[i][0]); for (const p of arcos[i].slice(1)) path.lineTo(...p); };
    tracar(bordas[tipo]);
    if (tipo !== 'municipio') for (const d of donos) tracar(ufs[municipios[d].uf].contorno);
  }
  return { municipios, ufs, bordas, box: caixaBrasil, porId: new Map(municipios.map(m => [m.id, m])) };
}

/**
 * Câmera {k, x, y} que enquadra uma caixa (ou o Brasil / UF / município).
 * `area` {x, y, w, h} é a região livre do canvas (abaixo do HUD, ao lado dos painéis
 * flutuantes): o mapa ocupa a tela toda, mas o enquadramento é centrado nessa região.
 */
export function cameraPara(geo, { uf, municipio, caixa }, largura, altura, area) {
  const A = area || { x: 0, y: 0, w: largura, h: altura };
  let box = caixa || municipio?.box || (uf ? geo.ufs[uf].box : geo.box);
  // Município: deixa ~3× o seu tamanho de contexto em volta (vizinhos visíveis).
  const margem = municipio ? 1.1 : caixa ? .6 : uf ? .08 : .03;   // caixa do caso: com vizinhos em volta
  const bx = Math.max(box[2] - box[0], 1), by = Math.max(box[3] - box[1], 1);
  box = [box[0] - bx * margem, box[1] - by * margem, box[2] + bx * margem, box[3] + by * margem];
  const w = box[2] - box[0], h = box[3] - box[1];
  const k = Math.min(Math.max(A.w, 40) / w, Math.max(A.h, 40) / h);
  return { k, x: A.x + A.w / 2 - (box[0] + box[2]) / 2 * k, y: A.y + A.h / 2 - (box[1] + box[3]) / 2 * k };
}
