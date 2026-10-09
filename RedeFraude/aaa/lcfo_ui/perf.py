"""
Instrumentação de desempenho: mede o tempo das funções pesadas e imprime
linhas "[PERF] nome: 1.23s" no terminal.

Ligada por padrão; desligue com a variável de ambiente LCFO_PERF=0.
Não altera nenhuma lógica — só envolve a função e cronometra.
"""

from __future__ import annotations

import functools
import os
import time

PERF_ATIVO = os.environ.get("LCFO_PERF", "1") != "0"


def medir(fn):
    if not PERF_ATIVO:
        return fn

    @functools.wraps(fn)
    def _w(*a, **k):
        t0 = time.perf_counter()
        try:
            return fn(*a, **k)
        finally:
            print(f"[PERF] {fn.__name__}: {time.perf_counter() - t0:.2f}s", flush=True)
    return _w
