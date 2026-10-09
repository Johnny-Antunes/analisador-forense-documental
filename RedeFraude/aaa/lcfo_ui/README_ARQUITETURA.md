# Arquitetura da interface (lcfo_ui)

`app_nicegui.py` é só o ponto de entrada (`python app_nicegui.py`). A interface
inteira vive neste pacote, um assunto por arquivo.

## Onde fica cada coisa

| Arquivo | Responsabilidade |
|---|---|
| `pagina.py` | Rota `/`: monta a página na ordem certa chamando as fábricas abaixo |
| `contexto.py` | `Pagina` (o `ctx`): estado da sessão + refreshables + navegação, compartilhado por todos os módulos |
| `estado.py` | `SessionState`: o que o analista tem aberto (caso, sketch, filtros, painéis) |
| `servicos.py` | **Fachada única** para database / graph_engine / correlation / anomalias. Telas importam daqui, nunca direto dos motores |
| `perf.py` | Medição `[PERF]` das funções pesadas (desliga com `LCFO_PERF=0`) |
| `dados.py` | `DADOS` (grafo + clusters + casos em memória) e `recarregar()` |
| `config.py` | Constantes, paletas, opções da toolbar, módulos opcionais (enrich / EXIF) |
| `tema.py` | CSS global |
| `js_globais.py` | JS global: ponte postMessage do grafo, resize de painéis, bloqueio de atalhos |
| `grafo.py` | Componente do grafo (iframe vis.js), transporte do payload por arquivo, `montar_payload_grafo` |
| `mapa.py` | Componente do mapa territorial (iframe Canvas em `componente_mapa/`), mesmo transporte do grafo |
| `ponte_grafo.py` | Mensagens vindas do grafo (clique, layout, enrich, seleção) |
| `navegacao.py` | `abrir_caso_overview`, `abrir_mesa_sketch`, `voltar_ao_dashboard`, ... |
| `navbar.py`, `rail.py`, `sidebar.py` | Moldura da página |
| `dialogos.py` | Command palette, ingestão, novo sketch, remover sketch |
| `analises.py` | Pareceres técnicos (criar / editar) |
| `workspace.py` | Decide qual tela desenhar |
| `telas/*.py` | Uma tela por arquivo: `casos`, `overview`, `mesa`, `base_mestra`, `watchlist`, `centro_comando` |
| `telas/paineis/*.py` | Um painel da mesa por arquivo + registro em `paineis/__init__.py` |

## Fora do pacote

| Arquivo | Responsabilidade |
|---|---|
| `territorio_engine.py` | Junção município → código IBGE, alertas do radar mês a mês (mesmas regras do radar), payloads do mapa |
| `componente_mapa/` | Mapa em Canvas 2D sem dependências: `geografia.js` (malha → Path2D, câmera), `cores.js` (OKLab), `mapa.js` (desenho, drill, linha do tempo, ponte). Adaptado de open-apuracao-brazil (MIT) |
| `componente_mapa/data/brasil.topo.json` | Malha própria (IBGE, Censo 2022), gerada por `tools/gerar_malha.py` |
| `tools/gerar_malha.py` | Regera a malha a partir das APIs do IBGE (precisa de internet; rodar só se o IBGE atualizar a malha) |
| `tools/gerar_criacoes_teste.py` | Gera Criações Diárias sintéticas com cenários de alerta plantados; `--remover` apaga do banco |

Testar o mapa isolado (sem o app): `python -m http.server` dentro de `componente_mapa/` e abrir
`index.html?payload=_payloads/arquivo.json&uf=SP&metrica=volume`.

## Como evoluir sem perder recursos

1. **Antes de mexer:** `python -m pytest -q` precisa passar (teste de fumaça percorre todas as telas e painéis contra o banco real).
2. **Tela nova:** crie `telas/minha_tela.py` com `def render(ctx)`, e adicione um `elif` em `workspace.py` e um botão em `rail.py`.
3. **Painel novo na mesa:** crie `telas/paineis/meu_painel.py` com `def render(m, lado)`, registre em `PAINEIS` e em `config.OPCOES_TOOLBAR` / `ICONE_TOOLBAR`. Os dados já calculados chegam em `m` (`ContextoMesa`).
4. **Substituir um painel (ex.: o mapa da Fase 2):** troque só o arquivo do painel; o resto não muda.
5. **Função nova de banco/motor:** escreva no motor (`database.py` etc.), exponha em `servicos.py` (com `medir(...)` se for pesada) e importe de lá.
6. **Depois de mexer:** rode os testes de novo e faça um commit pequeno (`git add -A && git commit -m "..."`). Se algo quebrar, `git diff` mostra exatamente o que mudou.

## Regras

- Telas não falam com SQLite nem com NetworkX direto: sempre via `servicos.py`.
- Nada de estado global por usuário: o que é da sessão vai em `ctx.st`.
- `DADOS` é compartilhado entre todas as abas; após ingerir ou promover, chame `DADOS.recarregar()`.
