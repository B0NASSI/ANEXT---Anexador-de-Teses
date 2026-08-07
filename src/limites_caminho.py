# -*- coding: utf-8 -*-
"""Evita que um nome de segurado/tese longo demais quebre o salvamento do
arquivo: o Windows recusa caminhos com mais de ~260 caracteres, e o erro que
sobra (ex.: "WinError 3 - não foi possível encontrar o caminho especificado")
não diz pra quem não é técnico o que de fato aconteceu.

Corta só o nome livre (nunca a pasta, que a pessoa escolheu) o quanto for
preciso pra caber com folga, em vez de deixar o salvamento falhar.
"""
from pathlib import Path

LIMITE_CAMINHO_WINDOWS = 259  # MAX_PATH (260), com 1 caractere de folga


def truncar_para_caminho(pasta: Path, nome: str, caracteres_extras: int = 0,
                          limite: int = LIMITE_CAMINHO_WINDOWS) -> str:
    """Encurta `nome` para que `pasta / nome` + `caracteres_extras` (prefixo
    e/ou sufixo fixos ao redor do nome, ex.: um ".pdf") caibam dentro de
    `limite` caracteres no total.

    Sem espaço nenhum sobrando, ainda devolve pelo menos 1 caractere — se o
    caminho mesmo assim não couber (pasta já enorme por si só), o erro que
    aparecer depois ao menos aponta a pasta certa, em vez de travar aqui.
    """
    orcamento = limite - len(str(pasta)) - 1 - caracteres_extras  # 1 = separador de pasta
    if orcamento < 1:
        orcamento = 1
    if len(nome) <= orcamento:
        return nome
    return nome[:orcamento].rstrip()
