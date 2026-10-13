"""
Painel da mesa: Ferramentas.
"""

from __future__ import annotations

from typing import Any, Dict, Tuple

from nicegui import ui

from lcfo_ui.config import TEM_EXIF_ENGINE, extrair_metadados_foto, validar_coerencia_geografica_foto, TOLERANCIA_KM_EXIF
from lcfo_ui.servicos import ocultar_no_do_caso, restaurar_no_do_caso, listar_nos_ocultos_do_caso

from lcfo_ui.telas.contexto_mesa import ContextoMesa


def render(m: ContextoMesa, lado: str) -> None:
    ctx = m.ctx
    identificador_caso = m.identificador_caso
    dossie_persistido = m.dossie_persistido
    cluster_obj = m.cluster_obj
    df_dados = m.df_dados
    with ui.column().classes('w-full h-full pt-2 px-8 pb-8 overflow-y-auto custom-scroll max-w-4xl gap-2'):
        if not dossie_persistido:
            ui.label('Disponível apenas para sketches oficiais.').classes('text-[#71717A]')
        else:
            ui.label('Isolamento de Terceiros da Cadeia de Custódia (escopo: este sketch)').classes('text-sm font-bold text-gray-300')
            with ui.row().classes('w-full gap-2 items-center mb-2'):
                sel_no = ui.select(cluster_obj['nodes'] if cluster_obj else [], label='Selecionar Nó:').props('dense outlined').classes('w-64 text-xs')
                ipt_mot = ui.input(placeholder='Motivo do isolamento...').props('dense outlined').classes('w-80 text-xs')
                ui.button('Isolar Nó', on_click=lambda: (
                    ocultar_no_do_caso(identificador_caso, sel_no.value, ipt_mot.value),
                    ui.notify('Nó isolado!'),
                    ctx.workspace.refresh(),
                )).classes('bg-red-900/60 hover:bg-red-800 text-white text-xs h-8 rounded-md px-3 border border-red-700/50')

            ocultos = listar_nos_ocultos_do_caso(identificador_caso)
            if ocultos:
                ui.label('Nós ocultados neste sketch:').classes('text-xs font-semibold text-[#71717A] mt-2')
                for oc in ocultos:
                    with ui.row().classes('items-center justify-between w-full text-xs text-gray-400 py-1'):
                        ui.label(f"{oc['node_id']} — {oc['motivo'] or 'sem justificativa'}")
                        ui.button('Reativar', on_click=lambda nid=oc['node_id']: (
                            restaurar_no_do_caso(identificador_caso, nid), ctx.workspace.refresh()
                        )).props('dense flat').classes('text-[#FF7300]')

            if TEM_EXIF_ENGINE:
                ui.separator().classes('my-4 bg-[#2B2B2F]')
                ui.label('Perícia Forense de Imagens (EXIF)').classes('text-sm font-bold text-gray-300 mb-1')
                ui.label('Confronte o GPS da foto contra a praça declarada no sinistro.').classes('text-xs text-[#71717A] mb-3')

                if df_dados.empty or "id_assistencia" not in df_dados.columns:
                    ui.label('Sem identificadores de assistência disponíveis.').classes('text-xs text-[#71717A]')
                else:
                    opcoes_assist_exif: Dict[str, Tuple[str, str]] = {}
                    for _, r_assist in df_dados.iterrows():
                        rot = f"{r_assist.get('id_assistencia', '?')} — {r_assist.get('cidade', '?')}/{r_assist.get('uf', '?')} ({r_assist.get('data', '?')})"
                        opcoes_assist_exif[rot] = (str(r_assist.get("cidade", "")), str(r_assist.get("uf", "")))

                    sel_assist_exif = ui.select(
                        list(opcoes_assist_exif.keys()), label='Assistência:'
                    ).props('dense outlined').classes('w-full text-xs mb-2')

                    resultado_exif_box = ui.column().classes('w-full gap-2 mt-2')
                    estado_exif: Dict[str, Any] = {"bytes": None, "nome": None}

                    def _ao_subir_foto_exif(e):
                        estado_exif["bytes"] = e.content.read()
                        estado_exif["nome"] = e.name
                        ui.notify(f"Foto '{e.name}' carregada. Clique em Analisar.")

                    ui.upload(
                        on_upload=_ao_subir_foto_exif,
                        label='Foto (JPEG/PNG):',
                        auto_upload=True,
                    ).props('dense accept=".jpg,.jpeg,.png"').classes('w-full text-xs mb-2')

                    def _executar_analise_exif():
                        resultado_exif_box.clear()
                        if not estado_exif["bytes"]:
                            ui.notify("Envie uma foto antes de analisar.", type="warning")
                            return
                        if not sel_assist_exif.value:
                            ui.notify("Selecione uma assistência.", type="warning")
                            return
                        cidade_sel, uf_sel = opcoes_assist_exif[sel_assist_exif.value]
                        metadados = extrair_metadados_foto(estado_exif["bytes"])
                        with resultado_exif_box:
                            if not metadados["possui_exif"]:
                                ui.label('Metadados EXIF ausentes na imagem.').classes('text-xs text-yellow-400')
                            elif metadados.get("erro"):
                                ui.label(f"Erro ao processar: {metadados['erro']}").classes('text-xs text-red-400')
                            else:
                                with ui.row().classes('w-full gap-8'):
                                    with ui.column().classes('gap-1'):
                                        ui.label('Metadados Extraídos:').classes('text-xs font-bold text-gray-300')
                                        ui.label(f"Dispositivo: {metadados.get('fabricante') or 'N/D'} {metadados.get('modelo') or ''}").classes('text-xs text-gray-400')
                                        ui.label(f"Data: {metadados.get('data_captura') or 'N/D'}")
                                        ui.label(f"GPS: [{metadados.get('latitude')}, {metadados.get('longitude')}]").classes('text-xs text-gray-400')
                                    with ui.column().classes('gap-1'):
                                        ui.label(f"Confronto Territorial ({cidade_sel}/{uf_sel}):").classes('text-xs font-bold text-gray-300')
                                        confronto = validar_coerencia_geografica_foto(
                                            metadados_foto=metadados, cidade_declarada=cidade_sel,
                                            uf_declarada=uf_sel, tolerancia_km=TOLERANCIA_KM_EXIF
                                        )
                                        cor_confronto = 'text-red-400' if confronto["divergencia_critica"] else 'text-green-400'
                                        ui.label(confronto["parecer_exif"]).classes(f'text-xs {cor_confronto}')

                    ui.button('Analisar Foto', icon='search', on_click=_executar_analise_exif).classes('bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md')
