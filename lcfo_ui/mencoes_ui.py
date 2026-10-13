"""
Menções nos pareceres (item 11): autocomplete com @ no editor e clique numa menção já escrita.

JS -> Python por um elemento-ponte (#lcfo-mencoes-bridge, como a ponte do grafo):
  {acao: 'buscar', req, termo}  -> resposta window.lcfoMencoes.resultado(req, grupos)
  {acao: 'abrir', ref}          -> foca o nó / "Abrir em" / município no Centro de Comando
As regras (formato, conversão do legado, busca, retrolinks) ficam em lcfo_dominio/mencoes.py.
"""

from __future__ import annotations

import json
import time
from typing import Any, Dict, List, Tuple

from nicegui import ui

from lcfo_dominio import mencoes, sketch_repo, tipos
from lcfo_ui import overlay
from lcfo_ui.dados import DADOS
from lcfo_ui.servicos import listar_sketches_do_caso

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

_CACHE_ENT: Dict[str, Tuple[float, List[Tuple[str, str]]]] = {}
VALIDADE_INDICE = 60.0

JS = r"""
<style>
  .lcfo-mencao { background:#1E293B; color:#38BDF8; padding:1px 6px; border-radius:4px; font-weight:600;
                 border:1px solid rgba(56,189,248,.3); cursor:pointer; white-space:nowrap; }
  .lcfo-mencao[data-tipo="municipio"] { color:#6FA98A; border-color:rgba(111,169,138,.35); background:#14251B; }
  .lcfo-mencao[data-tipo="sketch"] { color:#C9A66B; border-color:rgba(201,166,107,.35); background:#2A2216; }
  #lcfo-mencoes-menu { position:fixed; z-index:9000; width:320px; max-height:340px; overflow-y:auto; display:none;
                       background:rgba(24,24,26,.97); border:1px solid #2B2B2F; border-radius:10px; padding:4px;
                       box-shadow:0 12px 32px rgba(0,0,0,.55); font:12px "Segoe UI",sans-serif; color:#EDEDED; }
  #lcfo-mencoes-menu .g { font-size:9px; letter-spacing:.08em; text-transform:uppercase; color:#71717A; padding:6px 8px 2px; }
  #lcfo-mencoes-menu .i { padding:5px 8px; border-radius:6px; cursor:pointer; display:flex; gap:6px; align-items:baseline; }
  #lcfo-mencoes-menu .i.ativo, #lcfo-mencoes-menu .i:hover { background:rgba(255,115,0,.14); }
  #lcfo-mencoes-menu .i small { color:#71717A; margin-left:auto; font-size:10px; }
</style>
<script>
window.lcfoMencoes = window.lcfoMencoes || (function() {
  let req = 0, estado = null, itens = [], idx = 0, timer = null;
  const ponte = () => document.getElementById('lcfo-mencoes-bridge');
  const enviar = d => { const p = ponte(); if (p) p.dispatchEvent(new CustomEvent('mencao', { detail: JSON.stringify(d) })); };
  function menu() {
    let m = document.getElementById('lcfo-mencoes-menu');
    if (!m) { m = document.createElement('div'); m.id = 'lcfo-mencoes-menu'; document.body.appendChild(m); }
    return m;
  }
  function fechar() { estado = null; itens = []; menu().style.display = 'none'; }
  function verificar(conteudo) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return fechar();
    const r = sel.getRangeAt(0), no = r.startContainer;
    if (no.nodeType !== 3) return fechar();
    const antes = no.textContent.slice(0, r.startOffset);
    const m = antes.match(/(^|\s)@([^\s@][^@\n]{0,39})$/);      // aceita nomes com espaço ("Sao Bernardo")
    if (!m) return fechar();
    estado = { conteudo, no, ini: r.startOffset - m[2].length - 1, fim: r.startOffset, rect: r.getBoundingClientRect() };
    clearTimeout(timer);
    timer = setTimeout(() => { req += 1; enviar({ acao: 'buscar', req, termo: m[2] }); }, 120);
  }
  function desenhar(grupos) {
    const m = menu(); m.innerHTML = ''; itens = []; idx = 0;
    Object.entries(grupos).forEach(([g, lista]) => {
      if (!lista.length) return;
      const t = document.createElement('div'); t.className = 'g'; t.textContent = g; m.appendChild(t);
      lista.forEach(it => {
        const el = document.createElement('div'); el.className = 'i';
        const a = document.createElement('span'); a.textContent = it.rotulo; el.appendChild(a);
        if (it.detalhe) { const s = document.createElement('small'); s.textContent = it.detalhe; el.appendChild(s); }
        const k = itens.length; el.onmousedown = ev => { ev.preventDefault(); idx = k; inserir(); };
        itens.push({ ...it, el }); m.appendChild(el);
      });
    });
    if (!itens.length) { m.style.display = 'none'; return; }
    marcar();
    const r = estado.rect;
    m.style.left = Math.min(r.left, innerWidth - 330) + 'px';
    m.style.top = (r.bottom + 6 + 340 > innerHeight ? Math.max(8, r.top - 346) : r.bottom + 6) + 'px';
    m.style.display = 'block';
  }
  function marcar() { itens.forEach((it, i) => it.el.classList.toggle('ativo', i === idx)); if (itens[idx]) itens[idx].el.scrollIntoView({ block: 'nearest' }); }
  function inserir() {
    const it = itens[idx]; if (!it || !estado) return;
    const { conteudo, no, ini, fim } = estado;
    const r = document.createRange(); r.setStart(no, Math.max(0, ini)); r.setEnd(no, fim); r.deleteContents();
    const span = document.createElement('span');
    span.className = 'lcfo-mencao'; span.dataset.ref = it.ref; span.dataset.tipo = it.tipo;
    span.setAttribute('contenteditable', 'false'); span.title = it.rotulo; span.textContent = '@' + it.rotulo;
    const espaco = document.createTextNode(' ');
    r.insertNode(espaco); r.insertNode(span);
    const sel = window.getSelection(), fimR = document.createRange();
    fimR.setStartAfter(espaco); fimR.collapse(true); sel.removeAllRanges(); sel.addRange(fimR);
    fechar();
    conteudo.dispatchEvent(new Event('input', { bubbles: true }));         // o q-editor sincroniza o valor
  }
  document.addEventListener('input', ev => {
    const c = ev.target && ev.target.closest && ev.target.closest('.q-editor__content');
    if (c) verificar(c);
  }, true);
  document.addEventListener('keydown', ev => {
    if (!estado || !itens.length || menu().style.display === 'none') return;
    if (ev.key === 'ArrowDown') { idx = (idx + 1) % itens.length; marcar(); ev.preventDefault(); ev.stopPropagation(); }
    else if (ev.key === 'ArrowUp') { idx = (idx - 1 + itens.length) % itens.length; marcar(); ev.preventDefault(); ev.stopPropagation(); }
    else if (ev.key === 'Enter' || ev.key === 'Tab') { inserir(); ev.preventDefault(); ev.stopPropagation(); }
    else if (ev.key === 'Escape') { fechar(); ev.preventDefault(); ev.stopPropagation(); }
  }, true);
  document.addEventListener('click', ev => {
    const s = ev.target && ev.target.closest && ev.target.closest('.lcfo-mencao');
    if (!s || s.closest('.q-editor__content')) return;                    // editando: não navega
    enviar({ acao: 'abrir', ref: s.dataset.ref });
  });
  return { resultado(r, grupos) { if (r === req && estado) desenhar(grupos); }, fechar };
})();
</script>
"""


def instalar(ctx: "Pagina") -> None:
    ui.add_body_html(JS)
    ponte = ui.element('div').props('id=lcfo-mencoes-bridge').style('display:none')

    def _receber(e):
        try:
            d = json.loads(e.args.get("detail") if isinstance(e.args, dict) else e.args)
        except Exception:
            return
        if d.get("acao") == "buscar":
            grupos = mencoes.buscar(str(d.get("termo") or ""), _entidades_do_dossie(ctx), _sketches_do_dossie(ctx))
            ui.run_javascript(f"window.lcfoMencoes.resultado({int(d.get('req') or 0)}, {json.dumps(grupos, ensure_ascii=False)})")
        elif d.get("acao") == "abrir" and d.get("ref"):
            abrir(ctx, str(d["ref"]))
    ponte.on('mencao', _receber, args=['detail'])


# ------------------------------------------------------------------ fontes do autocomplete
def _sketches_do_dossie(ctx: "Pagina") -> List[Tuple[str, str]]:
    caso = ctx.st.caso_ativo_id
    return [(s["id_sketch"], s["nome_sketch"]) for s in listar_sketches_do_caso(caso)] if caso else []


def _rotulo_no(nid: str) -> str:
    if DADOS.G is not None and nid in DADOS.G:
        a = DADOS.G.nodes[nid]
        nome = a.get("nome_titular")
        return nome if nome and nome != "N/D" else (a.get("valor") or tipos.formatar_id(nid))
    return tipos.formatar_id(nid)


def _ids_do_sketch(sk: Dict[str, Any]) -> List[str]:
    c = next((c for c in DADOS.cluster_info if c["id"] == sk.get("cluster_origem_id")), None) if sk.get("cluster_origem_id") else None
    return list(c["nodes"]) if c else []


def _entidades_do_dossie(ctx: "Pagina") -> List[Tuple[str, str]]:
    """Índice em memória (cache curto) das entidades do dossiê: clusters dos sketches + overlay."""
    caso = ctx.st.caso_ativo_id or ""
    agora = time.monotonic()
    if caso in _CACHE_ENT and agora - _CACHE_ENT[caso][0] < VALIDADE_INDICE:
        return _CACHE_ENT[caso][1]
    vistos: Dict[str, str] = {}
    for sk in (listar_sketches_do_caso(caso) if caso else []):
        for nid in _ids_do_sketch(sk):
            vistos.setdefault(nid, _rotulo_no(nid))
        for nid, no in sketch_repo.carregar_overlay(sk["id_sketch"])["nodes"].items():
            vistos.setdefault(nid, no.get("label") or tipos.formatar_id(nid))
    lista = list(vistos.items())
    _CACHE_ENT[caso] = (agora, lista)
    return lista


# ------------------------------------------------------------------ clique numa menção
def _focar_no_grafo(ctx: "Pagina", ref: str) -> bool:
    for cg in ctx.st.contexto_grafo_ativo.values():
        presentes = cg.get("ids_oficiais", set()) | set(overlay.carregar(ctx.st, cg["identificador_caso"])["nodes"])
        if ref in presentes:
            ctx.st.entidade_foco = ref
            cg["grafo"].enviar_mensagem({"type": "lcfo:focar", "id": ref})        # leve: sem reenviar o payload
            return True
    return False


def abrir(ctx: "Pagina", ref: str) -> None:
    ref = mencoes.resolver(ref)                                                  # menção a nó fundido segue o alias
    st = ctx.st
    if ref.startswith("MUN_"):
        import territorio_engine as te
        m = te.carregar_municipios_ibge().get(ref[4:])
        st.radar_uf, st.radar_mun = (m["uf"] if m else None), ref[4:]
        st.tela_ativa = "Centro de Comando"
        ctx.voltar_ao_dashboard()
        return
    if ref.startswith("SKT_"):
        ctx.abrir_mesa_sketch(ref[4:])
        return
    if _focar_no_grafo(ctx, ref):
        return
    caso = st.caso_ativo_id
    onde = []
    for sk in (listar_sketches_do_caso(caso) if caso else []):
        if ref in _ids_do_sketch(sk) or ref in sketch_repo.carregar_overlay(sk["id_sketch"])["nodes"]:
            onde.append(sk)
    with ui.dialog() as dlg, ui.card().classes('bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 w-[420px] gap-2'):
        ui.label(tipos.formatar_id(ref)).classes('text-sm font-bold text-gray-100')
        if onde:
            ui.label('Abrir em:').classes('text-xs text-gray-300')
            for sk in onde:
                ui.button(sk["nome_sketch"], icon='hub', color=None,
                          on_click=lambda s=sk: (dlg.close(), setattr(st, 'entidade_foco', ref), ctx.abrir_mesa_sketch(s["id_sketch"]))).props(
                    'flat dense no-caps').classes('w-full justify-start text-xs text-[#FF7300]')
        else:
            ui.label('Fora do acervo do dossiê: nenhum sketch dele tem esta entidade.').classes('text-xs text-[#C9A66B]')
            if st.sketch_ativo_id:
                def _criar():
                    t = tipos.tipo_do_id(ref)
                    try:
                        sketch_repo.criar_entidade(st.sketch_ativo_id, t.chave, tipos.valor_do_id(ref))
                    except Exception as e:
                        ui.notify(f"Não foi possível criar: {e}", type="warning")
                        return
                    dlg.close()
                    ov = overlay.carregar(st, st.sketch_ativo_id, forcar=True)
                    for cg in st.contexto_grafo_ativo.values():
                        if cg.get("identificador_caso") == st.sketch_ativo_id and ref in ov["nodes"]:
                            cg["grafo"].enviar_patch({"add": {"nodes": [ov["nodes"][ref]], "edges": []}, "focus": ref})
                    ui.notify("Criada no sketch aberto.", type="positive")
                if tipos.tipo_do_id(ref):
                    ui.button('Criar no sketch aberto', color=None, on_click=_criar).props('dense no-caps').classes(
                        'text-xs bg-[#FF7300] text-white rounded-md px-3')
        ui.button('Fechar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('self-end text-xs text-gray-400')
    dlg.open()
