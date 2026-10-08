"""
Painel de análises / pareceres técnicos (criação e edição).
"""

from __future__ import annotations

import re
from typing import Any, Callable, Dict

from nicegui import ui

from lcfo_ui.config import CLASSE_MENU_PADRAO
from lcfo_ui.servicos import carregar_historico_pareceres, registrar_parecer, atualizar_parecer_historico, formatar_mencoes_forenses
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_painel_analises(ctx: "Pagina") -> None:
    """Publica ctx.renderizar_painel_analises."""
    st = ctx.st

    # -------------------------------------------------------------------
    # DIALOG: EDITAR PARECER JÁ REGISTRADO
    # -------------------------------------------------------------------
    def _abrir_edicao_parecer(h: Dict[str, Any], refresh_callback: Callable[[], None]):
        with ui.dialog() as dlg_ed, ui.card().classes(f'{CLASSE_MENU_PADRAO} w-[560px] p-4'):
            ui.label('Editar Análise').classes('text-xs font-bold text-[#FF7300] uppercase tracking-wider mb-2')
            ipt_t = ui.input(label='Título (opcional):', value=h.get('titulo', '')).props('dense outlined').classes('w-full text-xs mb-2')
            ipt_c = ui.editor(value=h['parecer']).classes('w-full mb-3').props('dense dark')

            def _salvar():
                atualizar_parecer_historico(h['id'], ipt_t.value or "", ipt_c.value or "")
                dlg_ed.close()
                ui.notify("Análise atualizada.")
                refresh_callback()

            ui.button('Salvar Alterações', on_click=_salvar).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md')
        dlg_ed.open()

    # -------------------------------------------------------------------
    # PAINEL DE ANÁLISES / PARECERES TÉCNICOS
    # -------------------------------------------------------------------
    def renderizar_painel_analises(identificador_dossie: str, dados_caso_info: Dict[str, Any], refresh_callback: Callable[[], None]):
        with ui.row().classes('w-full items-center justify-between mb-2'):
            ui.label('Análises (Pareceres Técnicos)').classes('text-sm font-bold text-gray-200')
            ui.button('+ Nova', on_click=lambda: (
                setattr(st, 'mostrar_form_nova_analise', not st.mostrar_form_nova_analise), refresh_callback()
            )).props('dense').classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-2.5 py-1 rounded-md border border-[#2B2B2F]')

        if st.mostrar_form_nova_analise:
            ipt_titulo_parecer = ui.input(placeholder='Título da análise (opcional)').props('dense outlined').classes('w-full text-xs mb-2')
            ui.label('Use @CPF_..., @TEL_... ou @PLACA_... para destacar menções.').classes('text-[11px] text-[#71717A] mb-1')

            ipt_novo_parecer = ui.editor(placeholder='Escreva a análise técnica pericial...').classes('w-full mb-2').props('dense dark')

            def _salvar_parecer():
                texto = (ipt_novo_parecer.value or "").strip()
                texto_puro = re.sub(r'<[^>]*>', '', texto).strip()
                if not texto_puro:
                    ui.notify("Escreva um parecer antes de salvar.", type="warning")
                    return

                registrar_parecer(
                    identificador_dossie,
                    ipt_titulo_parecer.value or "",
                    dados_caso_info.get("analista_responsavel", ""),
                    dados_caso_info.get("status", "Em Investigação"),
                    texto,
                )
                DADOS.recarregar()
                st.mostrar_form_nova_analise = False
                ui.notify("Análise registrada.")
                refresh_callback()
                ctx.sidebar_container.refresh()

            ui.button('Salvar Análise', on_click=_salvar_parecer).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md mb-3')

        historico_notas = carregar_historico_pareceres(identificador_dossie)
        if historico_notas:
            for h in historico_notas:
                with ui.element('div').classes('parecer-card'):
                    with ui.row().classes('w-full items-start justify-between gap-2'):
                        with ui.column().classes('gap-0.5 flex-1 min-w-0'):
                            if h.get('titulo'):
                                ui.label(h['titulo']).classes('text-sm font-bold text-white truncate')
                            ui.html(
                                f"<span style='color:#6C93B0; font-size:11px; font-weight:600;'>{h['data_registro']}</span> • "
                                f"<span style='color:#C9A66B; font-size:11px;'>Analista: <b>{h['analista'] or 'Sistema'}</b></span> • "
                                f"<span style='color:#6FA98A; font-size:11px;'>Status: <b>{h['status']}</b></span>"
                            )
                        ui.button(icon='edit', color=None).props('flat dense round size=xs').classes(
                            'text-[#71717A] hover:text-[#FF7300] shrink-0'
                        ).tooltip('Editar análise').on(
                            'click', lambda h=h: _abrir_edicao_parecer(h, refresh_callback)
                        )
                    ui.html(
                        f"<div style='color:#CCCCCC; font-size:12px; margin-top:6px; line-height:1.6;'>{formatar_mencoes_forenses(h['parecer'])}</div>"
                    )
        else:
            with ui.element('div').classes('w-full bg-[#18181A] border border-[#2B2B2F] p-4 rounded-lg flex items-center justify-center'):
                ui.label('Nenhuma análise registrada ainda.').classes('text-xs text-[#71717A]')

    ctx.renderizar_painel_analises = renderizar_painel_analises
