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

from lcfo_dominio import sketch_repo, tipos_custom
from lcfo_ui.tema import instalar_tema
from lcfo_ui import pagina  # noqa: F401  (registra a rota '/')

# Tipos de entidade criados pelo analista (tela Configurações) entram no registro único.
try:
    tipos_custom.registrar_todos()
except Exception as _e:
    print(f"[TIPOS] tipos customizados não registrados: {_e}")

# Sketches: o blob nos_extras_json vira linhas nas tabelas do sketch (uma vez por sketch,
# com backup em backups/ antes; sem backup, não migra). A coluna antiga fica só para leitura.
try:
    _rel = sketch_repo.migrar_nos_extras()
    if _rel["sketches"] or _rel["ignorados"]:
        print(f"[MIGRAÇÃO] sketches: {_rel['sketches']} · nós: {_rel['nos']} · vínculos: {_rel['vinculos']}"
              f" · backup: {_rel['backup']} · itens ignorados: {len(_rel['ignorados'])}")
        for _i in _rel["ignorados"][:50]:
            print("   ignorado:", _i)
except Exception as _e:  # o app sobe mesmo assim; o JSON antigo continua intacto
    print(f"[MIGRAÇÃO] não executada: {_e}")

# Menções: indexa uma vez os pareceres antigos (o texto deles não é reescrito).
try:
    from lcfo_dominio import mencoes as _mencoes
    import database as _db
    _c = _db.get_db_connection()
    try:
        _vazio = _c.execute("SELECT COUNT(*) FROM parecer_mencoes").fetchone()[0] == 0
    finally:
        _c.close()
    if _vazio:
        _mencoes.reindexar_todos()
except Exception as _e:
    print(f"[MENÇÕES] índice não reconstruído: {_e}")

instalar_tema()

# reconnect_timeout: o padrão do NiceGUI é 3s — qualquer cálculo síncrono
# mais longo que isso derrubava a conexão ("Connection lost") e a página
# recarregava do zero. 30s é um paliativo: a tela ainda congela durante o
# cálculo; a correção real é tirar o cálculo pesado do event loop.
ui.run(title="LCFO Forensic Intelligence", dark=True, port=8501, reload=False, reconnect_timeout=30)
