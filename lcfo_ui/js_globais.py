"""
Scripts JS globais: ponte do grafo (postMessage), resize de painéis, bloqueio de atalhos.
"""

from __future__ import annotations


from nicegui import ui


# =========================================================================
# PONTE DO GRAFO COM VIS.JS E POSTMESSAGE
# =========================================================================
def instalar_ponte_grafo_global():
    ui.add_body_html("""
    <script>
    (function() {
        window.__lcfoGraphIframes = window.__lcfoGraphIframes || {};

        // initialPayloadUrl: URL de um arquivo estático com o payload do
        // grafo (nós/arestas/posições), em vez do payload embutido direto
        // neste comando JS. Payloads grandes (~1000 nós) chegavam perto ou
        // estouravam o teto de tamanho de mensagem do WebSocket usado por
        // baixo do ui.run_javascript, e a mensagem era descartada em
        // silêncio — o grafo carregava com o fundo vazio, sem erro algum.
        // Só a URL (pequena) trafega pelo WebSocket; o conteúdo real chega
        // via fetch() HTTP normal, sem esse teto.
        window.lcfoRegistrarGrafoIframe = function(chave, iframeId, initialPayloadUrl) {
            // Guarda o ID (string), não o elemento: ui.html (que cria o iframe) e este
            // ui.run_javascript chegam ao navegador como mensagens separadas, e o elemento
            // pode ainda não existir no DOM neste instante. Resolver o elemento aqui
            // descartava o registro em silêncio e o iframe nunca recebia o payload
            // (grafo/mapa com fundo vazio, sem erro algum). O elemento é resolvido sob demanda.
            window.__lcfoGraphIframes[chave] = { iframeId: iframeId, initialPayloadUrl: initialPayloadUrl, sent: false, el: null };
        };

        function resolverRegistro(registro) {
            if (registro.el && registro.el.isConnected) return registro.el;
            const el = document.getElementById(registro.iframeId);
            if (el) registro.el = el;
            return el;
        }

        if (window.__lcfoBridgeInstalado) return;
        window.__lcfoBridgeInstalado = true;

        window.addEventListener('message', function(event) {
            if (!event.data || !event.data.isStreamlitMessage) return;

            if (event.data.type === 'streamlit:componentReady') {
                for (const chave in window.__lcfoGraphIframes) {
                    const registro = window.__lcfoGraphIframes[chave];
                    const el = resolverRegistro(registro);
                    if (el && el.contentWindow === event.source && !registro.sent) {
                        registro.sent = true;
                        event.source.postMessage({ type: 'streamlit:render_url', url: registro.initialPayloadUrl }, '*');
                    }
                }
                return;
            }

            if (event.data.type !== 'streamlit:setComponentValue') return;

            let chaveEncontrada = null;
            for (const chave in window.__lcfoGraphIframes) {
                const registro = window.__lcfoGraphIframes[chave];
                const el = resolverRegistro(registro);
                if (el && el.contentWindow === event.source) { chaveEncontrada = chave; break; }
            }
            if (!chaveEncontrada) return;

            const bridge = document.getElementById('lcfo-graph-bridge');
            if (!bridge) return;
            bridge.dispatchEvent(new CustomEvent('graph-message', {
                detail: JSON.stringify({ key: chaveEncontrada, value: event.data.value })
            }));
        });
    })();
    </script>
    """)

def instalar_resize_notas_global():
    ui.add_body_html("""
    <script>
    (function() {
        window.lcfoRegistrarNotasResize = function(panelId, handleId) {
            window.__lcfoNotesResizeTarget = { panelId: panelId, handleId: handleId };
        };

        if (window.__lcfoNotesResizeInit) return;
        window.__lcfoNotesResizeInit = true;

        let resizing = false;
        document.addEventListener('mousedown', function(e) {
            const target = window.__lcfoNotesResizeTarget;
            if (!target) return;
            const handle = document.getElementById(target.handleId);
            if (!handle) return;
            if (e.target === handle || handle.contains(e.target)) {
                resizing = true;
                document.body.style.userSelect = 'none';
                e.preventDefault();
            }
        });
        document.addEventListener('mousemove', function(e) {
            if (!resizing) return;
            const target = window.__lcfoNotesResizeTarget;
            if (!target) return;
            const panel = document.getElementById(target.panelId);
            if (!panel) return;
            const rect = panel.getBoundingClientRect();
            const newWidth = Math.max(280, Math.min(720, rect.right - e.clientX));
            panel.style.width = newWidth + 'px';
            panel.style.flexBasis = newWidth + 'px';
        });
        document.addEventListener('mouseup', function() {
            if (resizing) { resizing = false; document.body.style.userSelect = ''; }
        });
    })();
    </script>
    """)

def instalar_resize_sidebar_global():
    ui.add_body_html("""
    <script>
    (function() {
        window.lcfoRegistrarSidebarResize = function(panelId, handleId) {
            window.__lcfoSidebarResizeTarget = { panelId: panelId, handleId: handleId };
        };

        if (window.__lcfoSidebarResizeInit) return;
        window.__lcfoSidebarResizeInit = true;

        let resizing = false;
        document.addEventListener('mousedown', function(e) {
            const target = window.__lcfoSidebarResizeTarget;
            if (!target) return;
            const handle = document.getElementById(target.handleId);
            if (!handle) return;
            if (e.target === handle || handle.contains(e.target)) {
                resizing = true;
                document.body.style.userSelect = 'none';
                e.preventDefault();
            }
        });
        document.addEventListener('mousemove', function(e) {
            if (!resizing) return;
            const target = window.__lcfoSidebarResizeTarget;
            if (!target) return;
            const panel = document.getElementById(target.panelId);
            if (!panel) return;
            const rect = panel.getBoundingClientRect();
            const newWidth = Math.max(220, Math.min(520, e.clientX - rect.left));
            panel.style.width = newWidth + 'px';
        });
        document.addEventListener('mouseup', function() {
            if (resizing) { resizing = false; activeDrawer = null; document.body.style.userSelect = ''; }
        });
    })();
    </script>
    """)

def instalar_bloqueio_atalhos_navegador():
    ui.add_body_html("""
    <script>
    document.addEventListener('keydown', function(e) {
        const k = e.key.toLowerCase();
        if ((e.ctrlKey || e.metaKey) && ['j', 'b', 'l'].includes(k)) {
            e.preventDefault();
        }
    }, true);
    </script>
    """)
