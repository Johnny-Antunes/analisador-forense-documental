"""
Diálogos globais (command palette, ingestão) e diálogos de sketch.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict

from nicegui import ui

from lcfo_ui.config import CLASSE_MENU_PADRAO
from lcfo_ui.servicos import (
    CAMINHO_REDE_OFICIAL, PASTA_LOCAL_BLACKLIST, CAMINHO_REDE_CRIACAO, PASTA_LOCAL_CRIACAO, carregar_arquivos_para_sqlite, carregar_criacoes_diarias_para_sqlite, reparar_mojibake_historico, criar_sketch, listar_sketches_do_caso, remover_sketch, vincular_cluster_como_sketch,
)
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_dialogos_globais(ctx: "Pagina") -> None:
    """Command palette (Ctrl+J) e gestão de bases & ingestão."""

    # -------------------------------------------------------------------
    # AÇÕES DE INGESTÃO
    # -------------------------------------------------------------------
    def executar_ingestao_bl(caminho: str):
        _, qtd, msg, _ = carregar_arquivos_para_sqlite(caminho, forcar_releitura=False)
        DADOS.recarregar()
        ui.notify(f"{qtd:,} assistências inseridas." if qtd > 0 else (msg or "Base já atualizada."))
        ctx.dlg_ingestao.close()
        ctx.navbar_breadcrumbs.refresh()
        ctx.sidebar_container.refresh()
        ctx.workspace.refresh()

    def executar_ingestao_cr(caminho: str):
        _, qtd, msg, _ = carregar_criacoes_diarias_para_sqlite(caminho, forcar_releitura=False)
        DADOS.recarregar()
        ui.notify(f"{qtd:,} registros de criação ingeridos." if qtd > 0 else (msg or "Nenhum arquivo novo."))
        ctx.dlg_ingestao.close()
        ctx.navbar_breadcrumbs.refresh()
        ctx.sidebar_container.refresh()
        ctx.workspace.refresh()

    def executar_reparo_mojibake():
        res = reparar_mojibake_historico()
        DADOS.recarregar()
        ui.notify(f"{sum(res.values())} registros corrigidos.")
        ctx.dlg_ingestao.close()
        ctx.workspace.refresh()

    # -------------------------------------------------------------------
    # DIALOG: COMMAND PALETTE (Ctrl+J)
    # -------------------------------------------------------------------
    with ui.dialog() as dlg_palette, ui.card().classes('w-[520px] bg-[#18181A] border border-[#2B2B2F] p-0 rounded-xl shadow-2xl overflow-hidden'):
        with ui.row().classes('h-11 border-b border-[#2B2B2F] px-3 items-center gap-2 w-full'):
            ui.icon('search', size='15px').classes('text-[#71717A]')
            ipt_pal = ui.input(placeholder='Type a command or search investigations...').props('dense borderless autofocus').classes('flex-1 text-white text-xs bg-transparent')

        box_pal_res = ui.column().classes('w-full gap-0.5 p-1.5 max-h-72 overflow-y-auto custom-scroll')

        def filtrar_palette(termo_bruto: str):
            termo = (termo_bruto or "").strip().upper()
            box_pal_res.clear()
            if not termo:
                return
            with box_pal_res:
                encontrados = 0
                ui.label('INVESTIGAÇÕES').classes('text-[10px] font-semibold text-[#71717A] tracking-wider px-2.5 py-1 uppercase')
                for c in DADOS.cluster_info[:100]:
                    nome_c = (DADOS.casos_cadastrados.get(c['id'], {}).get('nome') or c['hub_label']).upper()
                    if termo in nome_c or termo in c['id']:
                        with ui.row().classes('w-full items-center justify-between p-2 rounded-md hover:bg-[#222225] cursor-pointer transition-colors').on('click', lambda cid=c['id']: ctx.abrir_caso_overview(cid)):
                            with ui.row().classes('items-center gap-2'):
                                ui.icon('folder', size='14px').classes('text-[#71717A]')
                                ui.label(f"{nome_c[:28]} ({c['id']})").classes('text-xs text-gray-200 font-medium')
                            ui.label('Abrir').classes('text-[10px] text-[#71717A] tracking-widest uppercase')
                        encontrados += 1
                        if encontrados >= 8:
                            break
        ipt_pal.on_value_change(lambda e: filtrar_palette(e.value))

    # -------------------------------------------------------------------
    # DIALOG: GESTÃO DE BASES & INGESTÃO
    # -------------------------------------------------------------------
    with ui.dialog() as dlg_ingestao, ui.card().classes('w-[560px] bg-[#18181A] border border-[#2B2B2F] p-5 rounded-xl shadow-2xl'):
        ui.label('Gestão e Ingestão de Dados').classes('text-sm font-bold text-[#FF7300] uppercase tracking-wider mb-3')
        with ui.expansion('Base Blacklist (.xlsx / .csv)', icon='folder').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg mb-2'):
            pasta_padrao_bl = str(CAMINHO_REDE_OFICIAL) if Path(CAMINHO_REDE_OFICIAL).exists() else str(PASTA_LOCAL_BLACKLIST)
            caminho_bl = ui.input(value=pasta_padrao_bl, label='Caminho:').props('dense outlined').classes('w-full text-xs my-2')
            ui.button('Ingerir Arquivos Blacklist', on_click=lambda: executar_ingestao_bl(caminho_bl.value)).classes('w-full bg-[#FF7300] text-white text-xs font-bold')

        with ui.expansion('Criações Diárias (.xlsx / .csv)', icon='update').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg mb-2'):
            pasta_padrao_cr = str(CAMINHO_REDE_CRIACAO) if Path(CAMINHO_REDE_CRIACAO).exists() else str(PASTA_LOCAL_CRIACAO)
            caminho_cr = ui.input(value=pasta_padrao_cr, label='Caminho:').props('dense outlined').classes('w-full text-xs my-2')
            ui.button('Ingerir Criações Diárias', on_click=lambda: executar_ingestao_cr(caminho_cr.value)).classes('w-full bg-[#FF7300] text-white text-xs font-bold')

        with ui.expansion('Manutenção & Integridade', icon='build').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg'):
            ui.label('Corrige codificação corrompida (mojibake) em todo o acervo.').classes('text-xs text-gray-400 mb-2')
            ui.button('Reparar Codificação Mojibake', on_click=executar_reparo_mojibake).classes('w-full bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs')

    ctx.dlg_palette = dlg_palette
    ctx.dlg_ingestao = dlg_ingestao


def criar_dialogos_sketch(ctx: "Pagina") -> None:
    """Diálogos de novo sketch e de remoção de sketch."""

    # -------------------------------------------------------------------
    # DIALOG: NOVO SKETCH
    # -------------------------------------------------------------------
    def _abrir_dialog_novo_sketch(id_dossie: str):
        with ui.dialog() as dlg_ns, ui.card().classes(f'{CLASSE_MENU_PADRAO} w-[480px] p-5'):
            ui.label('Adicionar Sketch ao Dossiê').classes('text-sm font-bold text-[#FF7300] uppercase tracking-wider mb-3')

            ui.label('Vincular célula já catalogada:').classes('text-xs text-gray-300 mb-1')
            ids_ja_no_dossie = {sk['cluster_origem_id'] for sk in listar_sketches_do_caso(id_dossie) if sk['cluster_origem_id']}
            opcoes_cluster = {
                c['id']: (DADOS.casos_cadastrados.get(c['id'], {}).get('nome') or c['hub_label'])
                for c in DADOS.cluster_info if c['id'] not in ids_ja_no_dossie
            }
            sel_cluster_ns = ui.select(opcoes_cluster, label='Célula:').props('dense outlined').classes('w-full text-xs mb-2')
            ipt_nome_ns1 = ui.input(placeholder='Nome deste sketch (ex: Núcleo Financeiro)').props('dense outlined').classes('w-full text-xs mb-3')

            def _vincular():
                if not sel_cluster_ns.value or not (ipt_nome_ns1.value or "").strip():
                    ui.notify("Selecione a célula e informe um nome.", type="warning")
                    return
                vincular_cluster_como_sketch(id_dossie, sel_cluster_ns.value, ipt_nome_ns1.value.strip())
                dlg_ns.close()
                ui.notify("Sketch vinculado.")
                ctx.workspace.refresh()

            ui.button('Vincular Célula', on_click=_vincular).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs h-8 rounded-md mb-4')

            ui.separator().classes('bg-[#2B2B2F] mb-4')

            ui.label('Ou criar um sketch manual vazio:').classes('text-xs text-gray-300 mb-1')
            ipt_nome_ns2 = ui.input(placeholder='Nome do sketch manual').props('dense outlined').classes('w-full text-xs mb-2')

            def _criar_manual():
                if not (ipt_nome_ns2.value or "").strip():
                    ui.notify("Informe um nome.", type="warning")
                    return
                criar_sketch(id_dossie, ipt_nome_ns2.value.strip(), tipo_origem="MANUAL")
                dlg_ns.close()
                ui.notify("Sketch manual criado.")
                ctx.workspace.refresh()

            ui.button('Criar Sketch Manual', on_click=_criar_manual).classes('w-full bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs h-8 rounded-md')
        dlg_ns.open()

    def _abrir_confirmacao_remocao_sketch(sk: Dict[str, Any]):
        with ui.dialog() as dlg_conf, ui.card().classes(f'{CLASSE_MENU_PADRAO} w-[420px] p-4'):
            ui.label(f'Remover o sketch "{sk["nome_sketch"]}"?').classes('text-sm text-gray-200 mb-1')
            ui.label('A célula e os dados brutos permanecem no acervo — apenas este recorte deixa de ser listado no dossiê.').classes('text-xs text-[#71717A] mb-3')
            with ui.row().classes('w-full gap-2 justify-end'):
                ui.button('Cancelar', on_click=dlg_conf.close).props('flat').classes('text-xs text-gray-400')

                def _confirmar():
                    ok, msg = remover_sketch(sk['id_sketch'])
                    dlg_conf.close()
                    ui.notify(msg, type="positive" if ok else "negative")
                    ctx.workspace.refresh()

                ui.button('Remover', on_click=_confirmar).classes('bg-red-900/60 hover:bg-red-800 text-white text-xs h-8 rounded-md px-3')
        dlg_conf.open()

    ctx.abrir_dialog_novo_sketch = _abrir_dialog_novo_sketch
    ctx.abrir_confirmacao_remocao_sketch = _abrir_confirmacao_remocao_sketch
