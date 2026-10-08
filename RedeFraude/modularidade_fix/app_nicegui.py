"""
Módulo: app_nicegui.py
Objetivo: Ponto de entrada da Mesa de Inteligência Forense de Fraudes em
          Assistências (LCFO) — NiceGUI.

Toda a interface vive no pacote lcfo_ui/ (ver lcfo_ui/README_ARQUITETURA.md).
Este arquivo apenas registra o tema, a página e sobe o servidor:

    python app_nicegui.py

Variáveis de ambiente úteis:
    LCFO_PERF=0   desliga as linhas [PERF] de medição de tempo no terminal.
"""

from nicegui import ui

from lcfo_ui.tema import instalar_tema
from lcfo_ui import pagina  # noqa: F401  (registra a rota '/')

instalar_tema()

# reconnect_timeout: o padrão do NiceGUI é 3s — qualquer cálculo síncrono
# mais longo que isso derrubava a conexão ("Connection lost") e a página
# recarregava do zero. 30s é um paliativo: a tela ainda congela durante o
# cálculo; a correção real é tirar o cálculo pesado do event loop.
ui.run(title="LCFO Forensic Intelligence", dark=True, port=8501, reload=False, reconnect_timeout=30)
