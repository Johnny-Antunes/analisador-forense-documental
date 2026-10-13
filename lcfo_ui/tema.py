"""
CSS global — neutralização do Quasar e densidade forense.
"""

from __future__ import annotations

from nicegui import ui

CSS_GLOBAL = """
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,0,0" />
<style>
    :root, body, .q-dark {
        --background: #121214;
        --card: #18181A;
        --card-foreground: #EDEDED;
        --popover: #18181A;
        --popover-foreground: #EDEDED;
        --muted: #222225;
        --muted-foreground: #71717A;
        --border: #2B2B2F;
        --input: #2B2B2F;
        --primary: #FF7300;
        --primary-foreground: #FFFFFF;
        --ring: rgba(255, 115, 0, 0.4);

        --bg-main: #121214;
        --bg-surface: #18181A;
        --bg-input: #222225;
        --border-subtle: #2B2B2F;
        --primary-orange: #FF7300;
        --text-main: #EDEDED;

        --q-primary: #FF7300 !important;
        --q-primary-rgb: 255, 115, 0 !important;
        --q-secondary: #222225 !important;
        --q-accent: #FF7300 !important;
        --q-info: #71717A !important;
    }

    body {
        background-color: var(--bg-main) !important;
        color: var(--text-main) !important;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
        overflow: hidden !important;
        margin: 0; padding: 0;
        -webkit-font-smoothing: antialiased;
        letter-spacing: -0.01em;
    }

    html, body, #app, .nicegui-content, .q-page, .q-page-container, .q-layout {
        height: 100% !important;
    }
    .nicegui-content { padding: 0 !important; }

    .q-header {
        height: 44px !important;
        min-height: 44px !important;
        max-height: 44px !important;
        background-color: rgba(18, 18, 20, 0.98) !important;
        border-bottom: 1px solid var(--border-subtle) !important;
        display: flex !important;
        align-items: center !important;
        padding: 0 16px !important;
    }
    .q-header .q-btn {
        height: 28px !important;
        min-height: 28px !important;
        padding: 0 8px !important;
        font-size: 12px !important;
        color: #CCCCCC !important;
    }
    .q-header .q-btn:hover {
        color: #FFFFFF !important;
        background-color: #222225 !important;
    }

    .top-search-btn {
        height: 28px !important;
        min-height: 28px !important;
        line-height: 26px !important;
        padding: 0 14px !important;
        font-size: 11px !important;
        color: #888888 !important;
        background-color: #18181A !important;
        border: 1px solid #2B2B2F !important;
        border-radius: 999px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    .top-search-btn:hover {
        border-color: #44444A !important;
        color: #EDEDED !important;
    }

    .text-primary:not(.text-orange):not(.lcfo-dock-btn-active) {
        color: #CCCCCC !important;
    }
    .bg-primary {
        background-color: #FF7300 !important;
    }
    .q-focus-helper {
        background: currentColor !important;
        opacity: 0 !important;
    }
    ::selection {
        background: rgba(255, 115, 0, 0.3) !important;
        color: #FFFFFF !important;
    }

    .q-checkbox__inner--truthy .q-checkbox__bg {
        background: #FF7300 !important;
        border-color: #FF7300 !important;
    }
    .q-checkbox__bg {
        border-radius: 4px !important;
        border: 1px solid var(--border-subtle) !important;
        background: rgba(34, 34, 37, 0.4) !important;
        width: 15px !important;
        height: 15px !important;
    }
    .q-checkbox__inner:before { display: none !important; }

    .q-toggle {
        height: 28px !important;
        min-height: 28px !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    .q-toggle__inner:before { display: none !important; }
    .q-toggle__inner {
        position: relative !important;
        width: 32px !important;
        height: 18px !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    .q-toggle__track {
        position: absolute !important;
        top: 50% !important;
        left: 0 !important;
        margin-top: -9px !important;
        height: 18px !important;
        width: 32px !important;
        border-radius: 999px !important;
        background: #2B2B2F !important;
        opacity: 1 !important;
    }
    .q-toggle__inner--truthy .q-toggle__track {
        background: #FF7300 !important;
    }
    .q-toggle__thumb {
        position: absolute !important;
        top: 50% !important;
        left: 2px !important;
        margin-top: -7px !important;
        height: 14px !important;
        width: 14px !important;
        background: #FFFFFF !important;
        border-radius: 50% !important;
    }
    .q-toggle__inner--truthy .q-toggle__thumb {
        left: 16px !important;
    }

    .q-btn--dense .q-icon, .q-btn--dense i {
        font-size: 15px !important;
    }

    .q-tooltip {
        background: #18181A !important;
        color: #EDEDED !important;
        border: 1px solid #2B2B2F !important;
        font-size: 11px !important;
        padding: 4px 8px !important;
        border-radius: 6px !important;
        box-shadow: 0 4px 16px rgba(0,0,0,0.6) !important;
    }

    .lcfo-dock-btn {
        width: 32px !important;
        height: 32px !important;
        min-width: 32px !important;
        min-height: 32px !important;
        padding: 0 !important;
        border-radius: 6px !important;
        color: #A1A1AA !important;
        background: transparent !important;
        border: 1px solid transparent !important;
    }
    .lcfo-dock-btn .q-icon, .lcfo-dock-btn i {
        font-size: 15px !important;
        color: #A1A1AA !important;
    }
    .lcfo-dock-btn:hover {
        background: #27272A !important;
        color: #FFFFFF !important;
        border-color: #3F3F46 !important;
    }
    .lcfo-dock-btn:hover .q-icon, .lcfo-dock-btn:hover i {
        color: #FFFFFF !important;
    }
    .lcfo-dock-btn-active, .lcfo-dock-btn.lcfo-dock-btn-active {
        border-radius: 6px !important;
        background: rgba(255, 115, 0, 0.15) !important;
        border-color: rgba(255, 115, 0, 0.4) !important;
        color: #FF7300 !important;
    }
    .lcfo-dock-btn-active .q-icon, .lcfo-dock-btn-active i {
        color: #FF7300 !important;
    }

    .lcfo-dossie-btn {
        width: 34px !important;
        height: 34px !important;
        min-width: 34px !important;
        color: #A1A1AA !important;
        background: transparent !important;
        border: 1px solid rgba(255, 255, 255, 0.10) !important;
        border-radius: 8px !important;
    }
    .lcfo-dossie-btn .q-icon, .lcfo-dossie-btn i {
        font-size: 16px !important;
        color: #A1A1AA !important;
    }
    .lcfo-dossie-btn:hover {
        background: #27272A !important;
        color: #FFFFFF !important;
        border-color: rgba(255, 255, 255, 0.25) !important;
    }
    .lcfo-dossie-btn:hover .q-icon, .lcfo-dossie-btn:hover i {
        color: #FFFFFF !important;
    }

    .q-field--outlined .q-field__control {
        border-radius: 6px !important;
        background: #18181A !important;
        border-color: #2B2B2F !important;
    }
    .q-field--outlined.q-field--focused .q-field__control {
        border-color: #FF7300 !important;
        box-shadow: 0 0 0 2px rgba(255, 115, 0, 0.25) !important;
    }
    .q-field__native, .q-field__input {
        color: #EDEDED !important;
        font-size: 12px !important;
    }
    .q-field--dense .q-field__control, .q-field--dense .q-field__marginal {
        height: 30px !important;
        min-height: 30px !important;
    }

    .q-table thead tr, .q-table tbody td {
        height: 32px !important;
    }
    .q-table th {
        font-size: 11px !important;
        font-weight: 600 !important;
        color: #71717A !important;
        border-bottom: 1px solid #2B2B2F !important;
        padding: 6px 10px !important;
        text-transform: uppercase !important;
    }
    .q-table td {
        font-size: 11px !important;
        color: #EDEDED !important;
        border-bottom: 1px solid #1F1F23 !important;
        padding: 6px 10px !important;
    }
    .q-table tbody tr:hover {
        background: rgba(34, 34, 37, 0.5) !important;
    }

    .q-splitter__separator {
        width: 1px !important;
        background: #2B2B2F !important;
        position: relative !important;
    }
    .q-splitter__separator:hover {
        background: #FF7300 !important;
    }
    .q-splitter__separator::after {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 3px;
        height: 32px;
        pointer-events: none;
        background-image: radial-gradient(circle, #55555A 1.4px, transparent 1.4px);
        background-size: 3px 7px;
        background-repeat: repeat-y;
        background-position: center;
    }

    /* Alça de largura da gaveta lateral: mesma linguagem do divisor da tela dividida
       (linha de 1px + pega pontilhada), com área de clique de 9px. */
    .lcfo-alca-lateral { position: relative; width: 9px; cursor: ew-resize; }
    .lcfo-alca-lateral::before {
        content: ''; position: absolute; top: 0; bottom: 0; left: 4px; width: 1px;
        background: #2B2B2F; transition: background .15s;
    }
    .lcfo-alca-lateral:hover::before { background: #FF7300; }
    .lcfo-alca-lateral::after {
        content: ''; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
        width: 3px; height: 32px; pointer-events: none;
        background-image: radial-gradient(circle, #55555A 1.4px, transparent 1.4px);
        background-size: 3px 7px; background-repeat: repeat-y; background-position: center;
    }

    .custom-scroll::-webkit-scrollbar { width: 5px; height: 5px; }
    .custom-scroll::-webkit-scrollbar-track { background: transparent; }
    .custom-scroll::-webkit-scrollbar-thumb { background: #2B2B2F; border-radius: 999px; }
    .custom-scroll::-webkit-scrollbar-thumb:hover { background: #FF7300; }

    .lcfo-panel {
        background: rgba(24, 24, 26, 0.92) !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 10px !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.5) !important;
    }

    .hud-compact-row {
        border-left: 3px solid var(--primary-orange) !important;
        padding: 7px 16px;
        display: flex; align-items: center; justify-content: center;
        gap: 10px 14px; flex-wrap: wrap;
    }
    .badge-pill { font-size: 10px; font-weight: 600; padding: 2px 8px; border-radius: 999px; text-transform: uppercase; white-space: nowrap; }

    .viewport-header {
        background: rgba(18, 18, 20, 0.95);
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
        border-bottom: 1px solid var(--border-subtle);
        height: 32px;
        display: flex; align-items: center; justify-content: space-between;
        padding: 0 10px; z-index: 30;
    }

    .parecer-card { background: #18181A; border: 1px solid var(--border-subtle); border-radius: 8px; padding: 10px 12px; margin-bottom: 8px; }
    .foco-lateral-card {
        background: #222225; border: 1px solid var(--primary-orange); border-left: 3px solid var(--primary-orange);
        border-radius: 6px; padding: 8px 10px; margin-bottom: 10px; font-size: 12px;
    }
    .selecao-lateral-card { background: #222225; border: 1px solid var(--border-subtle); border-radius: 6px; padding: 8px 10px; margin-bottom: 8px; font-size: 12px; }

    .q-editor {
        background-color: #18181A !important;
        border: 1px solid #2B2B2F !important;
        border-radius: 8px !important;
    }
    .q-editor__toolbar {
        background-color: #141416 !important;
        border-bottom: 1px solid #2B2B2F !important;
    }
    .q-editor__content {
        color: #EDEDED !important;
        font-size: 12px !important;
        min-height: 110px !important;
    }

    .lcfo-sketch-card {
        background: #18181A;
        border: 1px solid #2B2B2F;
        border-radius: 12px;
        overflow: hidden;
        transition: border-color 0.15s;
    }
    .lcfo-sketch-card:hover {
        border-color: rgba(255, 115, 0, 0.5);
    }
</style>
"""


def instalar_tema() -> None:
    """Registra o CSS global uma única vez (shared=True), antes do ui.run."""
    ui.add_head_html(CSS_GLOBAL, shared=True)
