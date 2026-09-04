# -*- coding: utf-8 -*-
from anexo_fixo import caminho_anexo_fixo

TITULO_EXATO = "CONTESTAÇÃO ADMINISTRATIVA, EFEITO SUSPENSIVO E PRAZO PRESCRICIONAL"


def test_reconhece_titulo_exato():
    assert caminho_anexo_fixo(TITULO_EXATO).name == "ACORDAO STJ.pdf"


def test_tolera_caixa_acentos_e_espacos_extras():
    titulo = "  contestacao   administrativa, efeito suspensivo e prazo prescricional  "
    assert caminho_anexo_fixo(titulo) is not None


def test_tolera_pequeno_erro_de_digitacao():
    # falta o "d" de "administrativa" — pequeno erro de digitação
    titulo = "CONTESTAÇÃO ADMINISTRATIVA EFEITO SUSPENSIVO E PRAZO PRESCRICINAL"
    assert caminho_anexo_fixo(titulo) is not None


def test_nao_reconhece_titulo_de_outra_tese():
    assert caminho_anexo_fixo("REVISÃO DE BENEFÍCIO - AUXÍLIO ACIDENTE") is None


def test_titulo_vazio_nao_da_erro():
    assert caminho_anexo_fixo("") is None
    assert caminho_anexo_fixo("   ") is None
