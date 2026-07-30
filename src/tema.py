# -*- coding: utf-8 -*-
"""
Tema visual (ttkbootstrap) do Divisor de Capas.
"""

from ttkbootstrap.style import ThemeDefinition

NOME_TEMA = "divisorpdf"

COR_PRIMARIA = "#2D315F"        # azul escuro (banner, ações de destaque)
COR_SECUNDARIA = "#D3782A"      # laranja (ação principal)
COR_FUNDO = "#F4F5F9"
COR_TEXTO = "#22243F"
COR_SUCESSO = "#1F7A4D"
COR_ERRO = "#C00000"

CORES = {
    "primary": COR_PRIMARIA,
    "secondary": COR_SECUNDARIA,
    "success": COR_SUCESSO,
    "info": "#3B82C4",
    "warning": "#E0A526",
    "danger": COR_ERRO,
    "light": COR_FUNDO,
    "dark": COR_TEXTO,
    "bg": COR_FUNDO,
    "fg": COR_TEXTO,
    "selectbg": COR_PRIMARIA,
    "selectfg": "#FFFFFF",
    "border": "#C9CCE0",
    "inputfg": COR_TEXTO,
    "inputbg": "#FFFFFF",
    "active": "#3B4070",
}

TEMA = ThemeDefinition(name=NOME_TEMA, colors=CORES, themetype="light")


def aplicar(root) -> "ttkbootstrap.Style":
    estilo = root.style
    estilo.register_theme(TEMA)
    estilo.theme_use(NOME_TEMA)
    estilo.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=(18, 10))
    # abas desabilitadas ("em breve") ficam pálidas, com cara de inativas —
    # o estado "disabled" é acrescentado na frente do mapa existente do
    # tema, preservando as cores da aba selecionada/ativa
    mapa_fg = list(estilo.map("TNotebook.Tab").get("foreground", []))
    mapa_bg = list(estilo.map("TNotebook.Tab").get("background", []))
    estilo.map(
        "TNotebook.Tab",
        foreground=[("disabled", "#B4B7C9")] + mapa_fg,
        background=[("disabled", "#ECEDF3")] + mapa_bg,
    )
    estilo.configure("Horizontal.TScale", troughcolor="#6870A4", sliderlength=20)
    estilo.configure("secondary.Horizontal.TScale", troughcolor="#6870A4", sliderlength=20)
    return estilo
