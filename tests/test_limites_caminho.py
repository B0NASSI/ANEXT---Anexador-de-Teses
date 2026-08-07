# -*- coding: utf-8 -*-
from pathlib import Path

from limites_caminho import LIMITE_CAMINHO_WINDOWS, truncar_para_caminho


def test_nao_trunca_quando_ja_cabe():
    pasta = Path("C:/casos/tese")
    nome = "FULANO DE TAL"
    assert truncar_para_caminho(pasta, nome, caracteres_extras=4) == nome


def test_trunca_para_caber_no_limite_com_pasta_curta():
    pasta = Path("C:/casos/tese")
    nome_gigante = "A" * 500
    resultado = truncar_para_caminho(pasta, nome_gigante, caracteres_extras=4)

    caminho_final = str(pasta / resultado) + ".pdf"
    assert len(caminho_final) <= LIMITE_CAMINHO_WINDOWS
    assert len(resultado) < len(nome_gigante)


def test_trunca_mais_quando_pasta_ja_e_longa():
    pasta_curta = Path("C:/tese")
    pasta_longa = Path("C:/" + "x" * 200 + "/tese")
    nome = "A" * 100

    resultado_curta = truncar_para_caminho(pasta_curta, nome, caracteres_extras=4)
    resultado_longa = truncar_para_caminho(pasta_longa, nome, caracteres_extras=4)

    assert len(resultado_longa) < len(resultado_curta)


def test_nunca_devolve_vazio_mesmo_com_pasta_enorme():
    pasta_enorme = Path("C:/" + "x" * 400)
    resultado = truncar_para_caminho(pasta_enorme, "NOME QUALQUER", caracteres_extras=4)
    assert len(resultado) >= 1
