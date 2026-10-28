# -*- coding: utf-8 -*-
"""
Tema visual (ttkbootstrap) do ANEXT.

Layout "barra lateral": navegação em azul-marinho à esquerda, conteúdo em
branco com cards de borda fina, e o laranja reservado pra ação principal.
"""

import datetime
import os

from ttkbootstrap.style import ThemeDefinition

NOME_TEMA = "divisorpdf"

# Comentários de cada cor (valem pras duas paletas):
#   PRIMARIA   barra lateral, títulos, botões "primary"
#   SECUNDARIA ação principal (Gerar/Dividir/Juntar) e acento da sidebar
#   FUNDO_SUAVE botões secundários "macios", caixa de resultado
#   AJUDA      legendas de ajuda, sem disputar com a cor da ação principal
#   AVISO      avisos de verdade (com ⚠), distinto da cor de "clique aqui"
#   NAV_ATIVO  item selecionado / hover na barra lateral
#   NAV_TEXTO  texto dos itens não selecionados
#   NAV_ICONE  ponto de cor nos ícones do rodapé da sidebar (sem ele ficava monocromático)
#   NAV_DIVISOR linha sutil acima do rodapé da sidebar
#   ESCALA     trilho dos controles deslizantes (Padronizar tabela)
PALETA_NORMAL = {
    "PRIMARIA": "#2D315F",        # azul-marinho
    "SECUNDARIA": "#D3782A",      # laranja
    "FUNDO": "#FFFFFF",
    "FUNDO_SUAVE": "#EEF0F5",
    "BORDA": "#E3E6EF",
    "TEXTO": "#22243F",
    "SUCESSO": "#1F7A4D",
    "ERRO": "#C00000",
    "AJUDA": "#6B7085",
    "AVISO": "#8A5A00",
    "NAV_ATIVO": "#3B4070",
    "NAV_TEXTO": "#C9CCE0",
    "NAV_ICONE": "#7FA8FF",
    "NAV_DIVISOR": "#454A7D",
    "ESCALA": "#6870A4",
}

# Edição Halloween: roxo-noite no lugar do azul-marinho e abóbora no lugar do
# laranja. O fundo do conteúdo continua branco (campos, tabelas e caixas de
# texto do Tk não acompanham um tema escuro de forma confiável).
PALETA_HALLOWEEN = {
    "PRIMARIA": "#2B1646",        # roxo-noite
    "SECUNDARIA": "#E8650A",      # abóbora
    "FUNDO": "#FFFFFF",
    "FUNDO_SUAVE": "#F3EEFA",
    "BORDA": "#E6DDF2",
    "TEXTO": "#241438",
    "SUCESSO": "#1F7A4D",
    "ERRO": "#C00000",
    "AJUDA": "#74698A",
    "AVISO": "#8A5A00",
    "NAV_ATIVO": "#43246B",
    "NAV_TEXTO": "#D9CCF0",
    "NAV_ICONE": "#FFB066",
    "NAV_DIVISOR": "#5A3A86",
    "ESCALA": "#8A6BB8",
}


# período da Edição Halloween (mês, dia), inclusive nas duas pontas: do
# lançamento da 4.2 (qua 28/10/2026) até a segunda depois do dia 31
INICIO_HALLOWEEN = (10, 28)
FIM_HALLOWEEN = (11, 2)


def halloween_ativo(hoje: datetime.date | None = None) -> bool:
    """Tema de Halloween vale de 28/10 a 02/11 (INICIO/FIM_HALLOWEEN). Pra
    pré-visualizar (ou forçar) em qualquer data, defina a variável de ambiente
    ANEXT_TEMA=halloween (ou ANEXT_TEMA=normal pra garantir o visual padrão
    mesmo dentro do período)."""
    forcado = os.environ.get("ANEXT_TEMA", "").strip().lower()
    if forcado == "halloween":
        return True
    if forcado == "normal":
        return False
    hoje = hoje or datetime.date.today()
    return INICIO_HALLOWEEN <= (hoje.month, hoje.day) <= FIM_HALLOWEEN


HALLOWEEN = halloween_ativo()
_PALETA = PALETA_HALLOWEEN if HALLOWEEN else PALETA_NORMAL

COR_PRIMARIA = _PALETA["PRIMARIA"]
COR_SECUNDARIA = _PALETA["SECUNDARIA"]
COR_FUNDO = _PALETA["FUNDO"]
COR_FUNDO_SUAVE = _PALETA["FUNDO_SUAVE"]
COR_BORDA = _PALETA["BORDA"]
COR_TEXTO = _PALETA["TEXTO"]
COR_SUCESSO = _PALETA["SUCESSO"]
COR_ERRO = _PALETA["ERRO"]
COR_AJUDA = _PALETA["AJUDA"]
COR_AVISO = _PALETA["AVISO"]
COR_NAV_ATIVO = _PALETA["NAV_ATIVO"]
COR_NAV_TEXTO = _PALETA["NAV_TEXTO"]
COR_NAV_ICONE = _PALETA["NAV_ICONE"]
COR_NAV_DIVISOR = _PALETA["NAV_DIVISOR"]
COR_ESCALA = _PALETA["ESCALA"]

CORES = {
    "primary": COR_PRIMARIA,
    "secondary": COR_SECUNDARIA,
    "success": COR_SUCESSO,
    "info": "#3B82C4",
    "warning": COR_AVISO,
    "danger": COR_ERRO,
    "light": COR_FUNDO_SUAVE,
    "dark": COR_TEXTO,
    "bg": COR_FUNDO,
    "fg": COR_TEXTO,
    "selectbg": COR_PRIMARIA,
    "selectfg": "#FFFFFF",
    "border": "#D5D9E5",
    "inputfg": COR_TEXTO,
    "inputbg": "#FFFFFF",
    "active": COR_NAV_ATIVO,
}

TEMA = ThemeDefinition(name=NOME_TEMA, colors=CORES, themetype="light")


def aplicar(root) -> "ttkbootstrap.Style":
    estilo = root.style
    estilo.register_theme(TEMA)
    estilo.theme_use(NOME_TEMA)

    estilo.configure(".", font=("Segoe UI", 10))
    estilo.configure("TButton", padding=(14, 8))
    estilo.configure("TEntry", padding=(8, 6))

    # cards: borda fina e neutra, título em azul-marinho
    estilo.configure(
        "TLabelframe", bordercolor=COR_BORDA, lightcolor=COR_BORDA, darkcolor=COR_BORDA,
        borderwidth=1, relief="solid",
    )
    estilo.configure("TLabelframe.Label", foreground=COR_PRIMARIA, font=("Segoe UI", 11, "bold"))
    estilo.configure("Ajuda.TLabel", foreground=COR_AJUDA, font=("Segoe UI", 9))
    estilo.configure("Aviso.TLabel", foreground=COR_AVISO, font=("Segoe UI", 9))
    estilo.configure("Titulo.TLabel", foreground=COR_PRIMARIA, font=("Segoe UI", 16, "bold"))
    estilo.configure("Descricao.TLabel", foreground=COR_AJUDA, font=("Segoe UI", 10))

    # a navegação fica na barra lateral: o Notebook continua gerenciando as
    # páginas, mas sem desenhar as abas
    estilo.layout("TNotebook.Tab", [])
    estilo.configure(
        "TNotebook", borderwidth=0, tabmargins=0,
        bordercolor=COR_FUNDO, lightcolor=COR_FUNDO, darkcolor=COR_FUNDO,
    )

    # barra lateral
    estilo.configure("Sidebar.TFrame", background=COR_PRIMARIA)
    estilo.configure("Acento.TFrame", background=COR_SECUNDARIA)
    estilo.configure("Sidebar.TLabel", background=COR_PRIMARIA, foreground="#FFFFFF")
    estilo.configure("SidebarSuave.TLabel", background=COR_PRIMARIA, foreground=COR_NAV_TEXTO, font=("Segoe UI", 9))
    # links do rodapé (Notas de atualização / Manual rápido): ícone e texto em
    # rótulos separados, com o ícone numa coluna de largura fixa — assim o
    # texto das duas linhas começa sempre no mesmo X, não importa a largura
    # do glifo do ícone (ⓘ e ? têm larguras bem diferentes na fonte)
    estilo.configure("SidebarLink.TLabel", background=COR_PRIMARIA, foreground=COR_NAV_TEXTO, font=("Segoe UI", 9))
    # hover: só texto laranja sublinhado (cor de ação do app), sem preencher a linha
    # inteira com uma barra de fundo - um bloco tomando a largura toda da sidebar por um
    # link secundário do rodapé chamava atenção demais pro que é.
    estilo.configure(
        "SidebarLinkAtivo.TLabel", background=COR_PRIMARIA, foreground=COR_SECUNDARIA,
        font=("Segoe UI", 9, "underline"),
    )
    # ícone do link em azul mais vivo que o texto (ver COR_NAV_ICONE) - dá hierarquia entre
    # o "marcador" (ícone) e o texto secundário, em vez dos dois na mesma cor apagada
    estilo.configure("SidebarIcone.TLabel", background=COR_PRIMARIA, foreground=COR_NAV_ICONE, font=("Segoe UI", 9))
    estilo.configure("SidebarIconeAtivo.TLabel", background=COR_PRIMARIA, foreground=COR_SECUNDARIA, font=("Segoe UI", 9))
    estilo.configure("SidebarDivisor.TFrame", background=COR_NAV_DIVISOR)
    for nome, fundo, texto, fonte, preenchimento in (
        ("Nav.TButton", COR_PRIMARIA, COR_NAV_TEXTO, ("Segoe UI", 11), (20, 12)),
        ("NavAtivo.TButton", COR_NAV_ATIVO, "#FFFFFF", ("Segoe UI", 11, "bold"), (20, 12)),
    ):
        estilo.configure(
            nome, background=fundo, foreground=texto, bordercolor=fundo, lightcolor=fundo, darkcolor=fundo,
            focuscolor=fundo, focusthickness=0, borderwidth=0, relief="flat", anchor="w",
            padding=preenchimento, font=fonte,
        )
        estilo.map(
            nome,
            background=[("active", COR_NAV_ATIVO), ("pressed", COR_NAV_ATIVO)],
            bordercolor=[("active", COR_NAV_ATIVO), ("pressed", COR_NAV_ATIVO)],
            lightcolor=[("active", COR_NAV_ATIVO), ("pressed", COR_NAV_ATIVO)],
            darkcolor=[("active", COR_NAV_ATIVO), ("pressed", COR_NAV_ATIVO)],
            foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")],
        )

    estilo.configure("Horizontal.TScale", troughcolor=COR_ESCALA, sliderlength=20)
    estilo.configure("secondary.Horizontal.TScale", troughcolor=COR_ESCALA, sliderlength=20)
    return estilo
