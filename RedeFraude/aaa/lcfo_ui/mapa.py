"""
Componente do mapa territorial (componente_mapa/, Canvas 2D) dentro do NiceGUI.

Mesmo transporte do grafo: o payload é gravado em componente_mapa/_payloads/
e só a URL trafega pelo WebSocket; a malha (brasil.topo.json) é um arquivo
estático que o navegador guarda em cache. Mensagens do iframe chegam pela
ponte global (lcfo-graph-bridge) e são entregues ao receptor registrado em
ctx.receptores_iframe[chave].
"""

from __future__ import annotations

import json
import re
import uuid
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable, Dict, Optional

from nicegui import app, ui

if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

_DIR_MAPA = Path(__file__).resolve().parent.parent / "componente_mapa"
_DIR_PAYLOADS_MAPA = _DIR_MAPA / "_payloads"
if _DIR_MAPA.exists():
    app.add_static_files('/lcfo_mapa_assets', str(_DIR_MAPA))
    _DIR_PAYLOADS_MAPA.mkdir(exist_ok=True)


class MapaTerritorialNiceGUI:
    def __init__(self, ctx: "Pagina", chave: str, ao_mudar: Callable[[Dict[str, Any]], None]):
        self.ctx = ctx
        self.chave = chave
        self.sufixo = re.sub(r'[^a-zA-Z0-9_-]', '_', chave)
        self.iframe_id = f"mapa-iframe-{self.sufixo}-{uuid.uuid4().hex[:6]}"
        self.ao_mudar = ao_mudar

    def _gravar_payload(self, payload: Dict[str, Any]) -> str:
        arq = _DIR_PAYLOADS_MAPA / f"{self.sufixo}.json"
        arq.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        return f"/lcfo_mapa_assets/_payloads/{self.sufixo}.json?v={uuid.uuid4().hex[:8]}"

    def render(self, payload: Dict[str, Any]) -> "MapaTerritorialNiceGUI":
        ui.html(
            f'<iframe id="{self.iframe_id}" src="/lcfo_mapa_assets/index.html?v={uuid.uuid4().hex[:8]}" '
            f'title="Mapa territorial" style="position:absolute; inset:0; width:100%; height:100%; border:none; display:block;"></iframe>',
            sanitize=False,
        ).classes('absolute inset-0 w-full h-full')
        url = self._gravar_payload(payload)
        self.ctx.receptores_iframe[self.chave] = self.ao_mudar
        ui.run_javascript(f"window.lcfoRegistrarGrafoIframe('{self.chave}', '{self.iframe_id}', {json.dumps(url)});")
        return self

    def enviar_layout(self, topo: int = 0, base: int = 0, esq: int = 0, dir: int = 0) -> None:
        """Informa ao mapa quanto de cada borda está coberto por HUD/painéis do app (px).
        O canvas continua ocupando tudo; câmera e controles respeitam a área livre."""
        msg = json.dumps({"type": "lcfo:layout", "topo": topo, "base": base, "esq": esq, "dir": dir})
        ui.run_javascript(f"""
            (function() {{
                const el = document.getElementById('{self.iframe_id}');
                if (el && el.contentWindow) el.contentWindow.postMessage({msg}, '*');
            }})();
        """)

    def enviar_estado(self, uf: Optional[str], municipio: Optional[str], mes: Optional[int],
                      metrica: Optional[str] = None, unidade: Optional[str] = None) -> None:
        """Sincroniza o mapa com o Python (lista clicada, mês do gráfico...). Não gera eco."""
        msg = json.dumps({"type": "lcfo:mapa_estado", "uf": uf, "municipio": municipio, "mes": mes,
                          "metrica": metrica, "unidade": unidade})
        ui.run_javascript(f"""
            (function() {{
                const el = document.getElementById('{self.iframe_id}');
                if (el && el.contentWindow) el.contentWindow.postMessage({msg}, '*');
            }})();
        """)
