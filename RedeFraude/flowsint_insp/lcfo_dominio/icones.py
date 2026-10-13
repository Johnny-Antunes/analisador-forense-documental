"""
Pictogramas dos nós (SVG 24×24, traço branco) e o círculo colorido que os envolve.

Antes viviam num if/elif dentro de graph_engine._criar_svg_nodo_flowsint; tipo novo
exigia mexer lá. Agora o registro de tipos (tipos.py) aponta para uma chave daqui.
"""

from __future__ import annotations

import urllib.parse
from functools import lru_cache

_B = 'fill="none" stroke="#FFFFFF"'

PICTOGRAMAS = {
    "pessoa": f'<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" {_B} stroke-width="2.2"/><circle cx="12" cy="7" r="4" {_B} stroke-width="2.2"/>',
    "telefone": f'<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z" {_B} stroke-width="2.2"/>',
    "carro": f'<path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.4-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.5 2.8C2 10.9 2 11.2 2 11.5V16c0 .6.4 1 1 1h2" {_B} stroke-width="2"/><circle cx="7" cy="17" r="2" {_B} stroke-width="2"/><circle cx="17" cy="17" r="2" {_B} stroke-width="2"/>',
    "predio": f'<path d="M3 21h18M5 21V7l7-4 7 4v14M9 9v.01M15 9v.01M9 13v.01M15 13v.01M9 17v.01M15 17v.01" {_B} stroke-width="2"/>',
    "empresa": f'<rect x="4" y="2" width="16" height="20" rx="2" {_B} stroke-width="2"/><path d="M9 6h.01M15 6h.01M9 10h.01M15 10h.01M9 14h.01M15 14h.01" {_B} stroke-width="2.5"/>',
    "endereco": f'<path d="M12 22s7-6.2 7-12a7 7 0 0 0-14 0c0 5.8 7 12 7 12z" {_B} stroke-width="2"/><circle cx="12" cy="10" r="2.5" {_B} stroke-width="2"/>',
    "clinica": f'<rect x="3" y="3" width="18" height="18" rx="3" {_B} stroke-width="2"/><path d="M12 7v10M7 12h10" {_B} stroke-width="2.4"/>',
    "pet": f'<circle cx="7" cy="9" r="1.8" {_B} stroke-width="1.8"/><circle cx="12" cy="6" r="1.8" {_B} stroke-width="1.8"/><circle cx="17" cy="9" r="1.8" {_B} stroke-width="1.8"/><path d="M12 12c-3 0-5.5 3.2-5.5 5.4 0 1.7 1.4 2.6 3 2.6 1 0 1.6-.5 2.5-.5s1.5.5 2.5.5c1.6 0 3-.9 3-2.6 0-2.2-2.5-5.4-5.5-5.4z" {_B} stroke-width="1.8"/>',
    "documento": f'<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" {_B} stroke-width="2"/><path d="M14 2v6h6M8 13h8M8 17h5" {_B} stroke-width="2"/>',
}


# Contorno do nó (aparência escolhida pelo analista, item 9); o pictograma do tipo fica dentro.
FORMAS = {
    "circulo": '<circle cx="24" cy="24" r="21" {a}/>',
    "quadrado": '<rect x="4" y="4" width="40" height="40" rx="7" {a}/>',
    "losango": '<polygon points="24,2 46,24 24,46 2,24" {a}/>',
    "hexagono": '<polygon points="24,2 43,13 43,35 24,46 5,35 5,13" {a}/>',
    "triangulo": '<polygon points="24,3 46,43 2,43" {a}/>',
}


@lru_cache(maxsize=512)
def svg_no(icone: str, cor_bg: str, cor_borda: str, is_hub: bool = False, forma: str = "circulo", realce: str = "") -> str:
    """Data URI do nó: contorno na cor do tipo com o pictograma branco (hub: borda laranja grossa).
    `realce` = cor da flag: vai na borda do próprio desenho (num nó-imagem o vis ignora a cor de borda)."""
    stroke_w = "3.5" if (is_hub or realce) else "1.8"
    borda_final = realce or ("#FF7300" if is_hub else cor_borda)
    path_icon = PICTOGRAMAS.get(icone, PICTOGRAMAS["empresa"])
    contorno = FORMAS.get(forma, FORMAS["circulo"]).format(a=f'fill="{cor_bg}" stroke="{borda_final}" stroke-width="{stroke_w}"')
    desloc = "translate(12, 15)" if forma == "triangulo" else "translate(12, 12)"
    svg_raw = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48">'
        f'{contorno}'
        f'<g transform="{desloc}">{path_icon}</g>'
        f'</svg>'
    )
    return f"data:image/svg+xml;utf8,{urllib.parse.quote(svg_raw)}"
