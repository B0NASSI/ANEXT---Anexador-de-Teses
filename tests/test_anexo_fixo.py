# -*- coding: utf-8 -*-
from helpers import criar_pdf_paginas

from anexo_fixo import caminho_anexo_fixo, ja_esta_presente

TITULO_EXATO = "CONTESTAÇÃO ADMINISTRATIVA, EFEITO SUSPENSIVO E PRAZO PRESCRICIONAL"

_TEXTO_ACORDAO = [
    ["RECURSO ESPECIAL Nº 1.234.567 - SP", "RELATOR: MINISTRO FULANO DE TAL"],
    ["EMENTA: PREVIDENCIÁRIO. PRAZO PRESCRICIONAL. EFEITO SUSPENSIVO. RECURSO PROVIDO."],
]


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


def test_ja_esta_presente_reconhece_mesmo_documento(tmp_path):
    anexo = tmp_path / "anexo.pdf"
    candidato = tmp_path / "candidato.pdf"
    criar_pdf_paginas(anexo, _TEXTO_ACORDAO)
    criar_pdf_paginas(candidato, _TEXTO_ACORDAO)

    assert ja_esta_presente(candidato, anexo) is True


def test_ja_esta_presente_nao_confunde_documento_diferente(tmp_path):
    anexo = tmp_path / "anexo.pdf"
    candidato = tmp_path / "candidato.pdf"
    criar_pdf_paginas(anexo, _TEXTO_ACORDAO)
    criar_pdf_paginas(candidato, [["TELA DO FAP - CONSULTA DE BENEFÍCIO", "SEGURADO: FULANO DE TAL"]])

    assert ja_esta_presente(candidato, anexo) is False


def test_ja_esta_presente_arquivo_ilegivel_nao_da_erro(tmp_path):
    anexo = tmp_path / "anexo.pdf"
    candidato = tmp_path / "nao_e_pdf.pdf"
    criar_pdf_paginas(anexo, _TEXTO_ACORDAO)
    candidato.write_text("isto não é um PDF de verdade")

    assert ja_esta_presente(candidato, anexo) is False
